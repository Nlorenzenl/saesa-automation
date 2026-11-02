"""
Cifra la agenda (site/index.html) para publicarla en GitHub Pages.

La página publicada solo muestra un campo de contraseña. Los datos van cifrados
(AES-256-GCM, clave derivada de la contraseña con PBKDF2-SHA256), así que sin la
contraseña no se pueden leer, aunque el repositorio y la página sean públicos.

Uso:  CENTRALITY2_CLAVE=... python encriptar_pagina.py site/index.html public/index.html
"""
import base64
import gzip
import hashlib
import json
import os
import sys
from pathlib import Path

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

ITERACIONES = 310_000


def sal():
    # Fija por repositorio, para que "Recordar en este equipo" siga sirviendo después de cada actualización.
    return hashlib.sha256(("centrality2|" + os.environ.get("GITHUB_REPOSITORY", "local")).encode()).digest()[:16]


def cifrar(html, clave):
    s = sal()
    llave = hashlib.pbkdf2_hmac("sha256", clave.encode(), s, ITERACIONES, 32)
    iv = os.urandom(12)
    datos = AESGCM(llave).encrypt(iv, gzip.compress(html.encode("utf-8"), 9), None)
    b64 = lambda b: base64.b64encode(b).decode()
    return {"sal": b64(s), "iv": b64(iv), "it": ITERACIONES, "ct": b64(datos)}


PORTADA = """<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>Centrality 2.0</title>
<style>
  :root { --bg:#f4f6f9; --card:#fff; --tx:#1c2430; --mut:#5d6878; --bd:#d9dee6; --ac:#1d5a9e; --err:#b42318; }
  @media (prefers-color-scheme: dark) { :root { --bg:#12161c; --card:#1b2129; --tx:#e6e9ee; --mut:#9aa4b2; --bd:#2c3440; --ac:#6ea8ec; --err:#f97066; } }
  * { box-sizing:border-box; }
  body { margin:0; min-height:100vh; display:grid; place-items:center; background:var(--bg); color:var(--tx); font:15px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif; padding:16px; }
  form { width:100%; max-width:360px; background:var(--card); border:1px solid var(--bd); border-radius:12px; padding:28px 24px; }
  h1 { margin:0 0 4px; font-size:20px; }
  p { margin:0 0 18px; color:var(--mut); font-size:13px; }
  label { display:block; font-size:13px; margin-bottom:6px; }
  input[type=password] { width:100%; padding:10px 12px; border:1px solid var(--bd); border-radius:8px; background:var(--bg); color:var(--tx); font-size:15px; }
  input[type=password]:focus { outline:2px solid var(--ac); outline-offset:1px; }
  .rec { display:flex; gap:8px; align-items:center; margin:12px 0 18px; font-size:13px; color:var(--mut); }
  button { width:100%; padding:10px; border:0; border-radius:8px; background:var(--ac); color:#fff; font-size:15px; cursor:pointer; }
  button:disabled { opacity:.6; cursor:wait; }
  .err { color:var(--err); font-size:13px; min-height:18px; margin:10px 0 0; }
</style>
</head>
<body>
<form id="f" hidden>
  <h1>Centrality 2.0</h1>
  <p>Agenda de PT · actualizada __LEIDO__</p>
  <label for="c">Contraseña</label>
  <input id="c" type="password" autocomplete="current-password" required autofocus>
  <label class="rec"><input id="r" type="checkbox" checked> Recordar en este equipo</label>
  <button id="b">Entrar</button>
  <div class="err" id="e" role="alert"></div>
</form>
<script>
const P = __PAQUETE__;
const bytes = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
const aB64 = u => btoa(String.fromCharCode(...u));
const LS = "centrality2-llave";
async function derivar(clave) {
  const base = await crypto.subtle.importKey("raw", new TextEncoder().encode(clave), "PBKDF2", false, ["deriveBits"]);
  return new Uint8Array(await crypto.subtle.deriveBits({ name: "PBKDF2", hash: "SHA-256", salt: bytes(P.sal), iterations: P.it }, base, 256));
}
async function abrir(llave) {
  const k = await crypto.subtle.importKey("raw", llave, "AES-GCM", false, ["decrypt"]);
  const plano = await crypto.subtle.decrypt({ name: "AES-GCM", iv: bytes(P.iv) }, k, bytes(P.ct));
  const html = await new Response(new Blob([plano]).stream().pipeThrough(new DecompressionStream("gzip"))).text();
  document.open(); document.write(html); document.close();
}
const f = document.getElementById("f"), e = document.getElementById("e"), b = document.getElementById("b");
f.addEventListener("submit", async ev => {
  ev.preventDefault(); e.textContent = ""; b.disabled = true; b.textContent = "Abriendo…";
  try {
    const recordar = document.getElementById("r").checked;
    const llave = await derivar(document.getElementById("c").value);
    try { await abrir(llave); } catch { throw new Error("Contraseña incorrecta."); }
    try { if (recordar) localStorage.setItem(LS, aB64(llave)); } catch {}
  } catch (err) {
    e.textContent = err.message || "No se pudo abrir la agenda."; b.disabled = false; b.textContent = "Entrar";
  }
});
(async () => {
  let guardada = null;
  try { guardada = localStorage.getItem(LS); } catch {}
  if (guardada) {
    try { await abrir(bytes(guardada)); return; } catch { try { localStorage.removeItem(LS); } catch {} }
  }
  f.hidden = false; document.getElementById("c").focus();
})();
</script>
</body>
</html>
"""


def main():
    origen, destino = Path(sys.argv[1]), Path(sys.argv[2])
    clave = os.environ.get("CENTRALITY2_CLAVE")
    if not clave:
        sys.exit("Falta el secreto CENTRALITY2_CLAVE (contraseña de la página).")
    html = origen.read_text(encoding="utf-8")
    leido = ""
    if "const META = " in html:
        try:
            leido = json.loads(html.split("const META = ", 1)[1].split(";\n", 1)[0]).get("leido", "")
        except Exception:
            pass
    pagina = PORTADA.replace("__LEIDO__", leido or "desde Centrality").replace("__PAQUETE__", json.dumps(cifrar(html, clave)))
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_text(pagina, encoding="utf-8")
    (destino.parent / ".nojekyll").write_text("", encoding="utf-8")
    print(f"Página cifrada: {destino} ({destino.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    main()
