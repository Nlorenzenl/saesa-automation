"""
Centrality 2.0 - sincronización de PT desde Centrality.

Cada ejecución:
  1. Inicia sesión en Centrality (usuario y contraseña desde variables de entorno).
  2. Lee todos los PT con inicio entre DESDE y HASTA.
  3. Lee la ficha completa (condiciones y bloqueos del jefe de faena, adjuntos,
     subestación, descripción) de los PT nuevos, de los que cambiaron y de los
     que parten pronto. El resto reutiliza la ficha guardada en la ejecución anterior.
  4. Compara contra la ejecución anterior: PT nuevos, cambios de estado, cambios
     de fecha y PT que salieron del rango.
  5. Guarda la foto en datos/pts.json, agrega los cambios a datos/cambios.jsonl
     y genera site/index.html (la agenda) a partir de plantilla.html.
  6. Envía un correo con el resumen de cambios y la agenda adjunta.

Variables de entorno:
  SAESA_USER, SAESA_PASS                 credenciales de Centrality (obligatorias)
  GMAIL_USER, GMAIL_APP_PASS, EMAIL_DEST correo (si faltan, no se envía correo)
  EMAIL_CC                               copias separadas por coma (opcional)
  DESDE, HASTA                           rango de fechas de inicio, AAAA-MM-DD
  FULL_REFRESH=true                      vuelve a leer la ficha de todos los PT
  SIN_CORREO=true                        no envía correo (pruebas)
  PAGINA_URL                             enlace a la agenda publicada (se pone en el correo)
"""
import json
import os
import smtplib
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE = "https://stx.saesa.cl:8091/backend/sts/"
TZ = ZoneInfo("America/Santiago")
DIR = Path(__file__).resolve().parent
DATOS = DIR / "datos"
SITE = DIR / "site"
DESDE = os.environ.get("DESDE", "2026-09-01")
HASTA = os.environ.get("HASTA", "2026-12-31")
FULL = os.environ.get("FULL_REFRESH", "false").lower() == "true"
SIN_CORREO = os.environ.get("SIN_CORREO", "false").lower() == "true"
HILOS = 6
# Además de los nuevos y los que cambiaron, se relee la ficha de los PT que parten
# en esta ventana, porque condiciones y adjuntos se agregan cerca de la fecha.
VENTANA_DIAS_ATRAS, VENTANA_DIAS_ADELANTE = 3, 14
CAMPOS_LISTA = ("estado", "ini", "fin", "prog", "tipo", "zona", "area", "desc", "fp")


def log(msg):
    print(f"[{datetime.now(TZ):%H:%M:%S}] {msg}", flush=True)


def fecha_local(iso):
    if not iso:
        return ""
    return datetime.fromisoformat(iso.replace("Z", "+00:00")).astimezone(TZ).strftime("%Y-%m-%d")


class Centrality:
    def __init__(self, usuario, clave):
        self.s = requests.Session()
        self.s.verify = False
        self.s.headers["User-Agent"] = "Centrality2-sync/1.0"
        self.usuario, self.clave = usuario, clave

    def login(self):
        for intento in range(3):
            try:
                self.s.get(BASE + "login.php", timeout=60)
                self.s.post(BASE + "login.php", data={"backurl": "centrality.php", "usrname": self.usuario, "usrpwd": self.clave}, timeout=60)
                prueba = self.listar(0, 1)
                if isinstance(prueba, dict) and "data" in prueba:
                    log("Sesión iniciada en Centrality")
                    return
            except Exception as e:
                log(f"Intento de inicio de sesión {intento + 1} falló: {e}")
            time.sleep(10)
        raise RuntimeError("No se pudo iniciar sesión en Centrality (revisa usuario, contraseña o conexión al puerto 8091).")

    def post(self, servicio, datos, timeout=60, reintentos=3):
        ultimo = None
        for intento in range(reintentos):
            try:
                r = self.s.post(BASE + servicio, data=datos, timeout=timeout)
                return r.json()
            except Exception as e:
                ultimo = e
                time.sleep(3 * (intento + 1))
        raise RuntimeError(f"{servicio} {datos.get('action')}: {ultimo}")

    def listar(self, inicio, limite):
        desde = (datetime.fromisoformat(DESDE) - timedelta(days=1)).strftime("%Y-%m-%dT00:00:00")
        hasta = (datetime.fromisoformat(HASTA) + timedelta(days=1)).strftime("%Y-%m-%dT06:00:00")
        datos = {"start": str(inicio), "limit": str(limite), "sort": "business_id", "dir": "ASC", "in_active_box": "",
                 "business_ids": "", "origin_context_entity": "", "origin_context_identifier": "",
                 "allowed_time_frame_from": desde, "allowed_time_frame_up": hasta, "allowed_time_frame": "",
                 "entered_from": "", "entered_up": "", "source_id": "", "maneuver_element_id": "", "area_ids": "",
                 "operation_order_kind_id": "", "state": "", "creator_id": "", "applicant_ou_id": "",
                 "action": "list_planning_applicant_operation_orders"}
        return self.post("svc_dms_operation_order.php", datos, timeout=180)

    def todos(self):
        filas, inicio = [], 0
        while True:
            j = self.listar(inicio, 2000)
            datos = j.get("data") or []
            filas.extend(datos)
            if len(datos) < 2000:
                break
            inicio += 2000
        return [f for f in filas if DESDE <= fecha_local(f.get("start_date")) <= HASTA]

    def ficha(self, pt_id):
        j = self.post("svc_dms_operation_order.php", {"action": "get_operation_order_info", "operation_order_id": str(pt_id)}, timeout=60)
        return j.get("data") or {}


def de_lista(x):
    return {"id": x["id"], "pt": x["business_id"], "ini": x.get("start_date"), "fin": x.get("end_date"),
            "zona": x.get("coverage_area_caption"), "area": x.get("origin_context_entity_name"),
            "tipo": x.get("operation_order_kind"), "estado": x.get("state_caption"), "prog": x.get("work_programmer_name"),
            "fp": x.get("alternative_approval_circuit") is True or any("fuera de plazo" in (b.get("tooltip") or "").lower() for b in (x.get("badges") or [])),
            "desc": x.get("work_description") or "", "creador": x.get("creator_name") or "", "solicitante": x.get("applicant") or ""}


def con_ficha(base, d):
    g = d.get("general") or {}
    base.update({
        "ref": g.get("reference_object_name") or "",
        "desc": g.get("work_description") or base.get("desc") or "",
        "caso": g.get("case_business_id") or "",
        "creador": g.get("creator_name") or base.get("creador") or "",
        "solicitante": g.get("applicant") or base.get("solicitante") or "",
        "cond": [[o.get("objective_order"), o.get("responsibility_name"), o.get("reference_object_names"), o.get("operational_condition_name"), o.get("display"), 1 if o.get("is_mandatory") else 0] for o in (d.get("tactic_objectives") or [])],
        "adj": [[a.get("id"), a.get("name"), a.get("size_caption"), a.get("type_caption"), a.get("mime_type"), a.get("creation_date"), a.get("creator")] for a in (d.get("attachments") or []) if not a.get("deletion_date")],
        "ficha_leida": datetime.now(TZ).isoformat(timespec="seconds"),
    })
    return base


def comparar(prev, nuevo):
    cambios = []
    for pt, x in nuevo.items():
        p = prev.get(pt)
        if p is None:
            cambios.append({"pt": pt, "tipo": "nuevo", "txt": f"PT nuevo en el rango, estado {x['estado']}", "estado": x["estado"], "ini": x["ini"]})
            continue
        if p.get("estado") != x.get("estado"):
            cambios.append({"pt": pt, "tipo": "estado", "txt": f"Estado: {p.get('estado')} → {x.get('estado')}", "antes": p.get("estado"), "estado": x["estado"], "ini": x["ini"]})
        if p.get("ini") != x.get("ini") or p.get("fin") != x.get("fin"):
            fh = lambda iso: datetime.fromisoformat(iso.replace("Z", "+00:00")).astimezone(TZ).strftime("%d/%m %H:%M") if iso else "—"
            partes = []
            if p.get("ini") != x.get("ini"):
                partes.append(f"Inicio {fh(p.get('ini'))} → {fh(x.get('ini'))}")
            if p.get("fin") != x.get("fin"):
                partes.append(f"Término {fh(p.get('fin'))} → {fh(x.get('fin'))}")
            cambios.append({"pt": pt, "tipo": "fecha", "txt": " · ".join(partes), "estado": x["estado"], "ini": x["ini"]})
        if p.get("prog") != x.get("prog") and x.get("prog"):
            cambios.append({"pt": pt, "tipo": "programador", "txt": f"Programador: {p.get('prog') or 'sin asignar'} → {x.get('prog')}", "estado": x["estado"], "ini": x["ini"]})
    for pt, p in prev.items():
        if pt not in nuevo:
            cambios.append({"pt": pt, "tipo": "salio", "txt": "Ya no tiene fecha de inicio dentro del rango", "estado": p.get("estado"), "ini": p.get("ini")})
    return cambios


def generar_html(pts, cambios, leido):
    plantilla = (DIR / "plantilla.html").read_text(encoding="utf-8")
    filas = sorted(pts.values(), key=lambda d: (d.get("ini") or "", d["pt"]))
    datos = [{k: v for k, v in d.items() if k != "ficha_leida"} for d in filas]
    for d in datos:
        d.setdefault("ref", ""); d.setdefault("cond", []); d.setdefault("adj", []); d.setdefault("caso", "")
    js = json.dumps(datos, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    resumen = {}
    for c in cambios:
        if c["tipo"] != "salio":
            resumen.setdefault(c["pt"], []).append(c["txt"])
    meta = {"hoy": datetime.now(TZ).strftime("%Y-%m-%d"), "leido": leido, "desde": DESDE, "hasta": HASTA,
            "cambios": [{"pt": k, "txt": " · ".join(v)} for k, v in resumen.items()]}
    html = plantilla.replace("/*DATA*/", js).replace("/*META*/", json.dumps(meta, ensure_ascii=False).replace("</", "<\\/"))
    SITE.mkdir(exist_ok=True)
    (SITE / "index.html").write_text(html, encoding="utf-8")
    return SITE / "index.html"


def enviar_correo(cambios, pts, leido, adjunto, primera=False):
    usuario, clave, dest = os.environ.get("GMAIL_USER"), os.environ.get("GMAIL_APP_PASS"), os.environ.get("EMAIL_DEST")
    if SIN_CORREO or not (usuario and clave and dest):
        log("Correo no enviado (SIN_CORREO activo o faltan credenciales de correo)")
        return
    cc = [c.strip() for c in os.environ.get("EMAIL_CC", "").split(",") if c.strip()]
    por = {}
    for c in cambios:
        por.setdefault(c["tipo"], []).append(c)
    estados = {}
    for x in pts.values():
        estados[x["estado"]] = estados.get(x["estado"], 0) + 1
    nombres = {"nuevo": "PT nuevos", "estado": "Cambios de estado", "fecha": "Cambios de fecha", "programador": "Cambios de programador", "salio": "PT que salieron del rango"}

    def tabla(lista, con_antes=False):
        filas = "".join(
            f"<tr><td style='font-family:monospace;padding:4px 10px'>{escape(c['pt'])}</td><td style='padding:4px 10px'>{fecha_local(c.get('ini'))}</td>"
            + (f"<td style='padding:4px 10px;color:#777'>{escape(c.get('antes') or '')}</td>" if con_antes else "")
            + f"<td style='padding:4px 10px'>{escape(c.get('estado') or '')}</td><td style='padding:4px 10px;color:#555'>{escape(c['txt'])}</td></tr>"
            for c in sorted(lista, key=lambda c: (c.get("ini") or "", c["pt"]))[:300])
        mas = f"<p style='color:#777;font-size:12px'>Se muestran 300 de {len(lista)}; el resto está en la agenda adjunta.</p>" if len(lista) > 300 else ""
        cab = "<tr style='background:#f3f5f8'><th align='left' style='padding:4px 10px'>PT</th><th align='left' style='padding:4px 10px'>Inicio</th>" + ("<th align='left' style='padding:4px 10px'>Antes</th>" if con_antes else "") + "<th align='left' style='padding:4px 10px'>Estado actual</th><th align='left' style='padding:4px 10px'>Detalle</th></tr>"
        return f"<table style='border-collapse:collapse;font-size:13px;width:100%'>{cab}{filas}</table>{mas}"

    bloques = "".join(f"<h3 style='margin:22px 0 6px;color:#1d5a9e'>{nombres[k]} ({len(por[k])})</h3>{tabla(por[k], k == 'estado')}" for k in ("estado", "nuevo", "fecha", "programador", "salio") if por.get(k))
    resumen_estados = "".join(f"<tr><td style='padding:3px 10px'>{escape(k)}</td><td align='right' style='padding:3px 10px'>{v}</td></tr>" for k, v in sorted(estados.items(), key=lambda kv: -kv[1]))
    total_cambios = sum(len(v) for k, v in por.items())
    url = os.environ.get("PAGINA_URL", "")
    enlace = f'<p style="margin:0 0 14px"><a href="{escape(url)}" style="background:#1d5a9e;color:#fff;padding:8px 14px;border-radius:6px;text-decoration:none;display:inline-block">Abrir la agenda en la web</a></p>' if url else ""
    cuerpo = f"""<html><body style="font-family:Arial,sans-serif;color:#1c2430;max-width:860px">
<h2 style="margin:0 0 4px">Centrality 2.0 · actualización {leido}</h2>
{enlace}
<p style="margin:0 0 14px;color:#555">PT con inicio del {DESDE[8:]}/{DESDE[5:7]}/{DESDE[:4]} al {HASTA[8:]}/{HASTA[5:7]}/{HASTA[:4]}: {len(pts)} en total. {'Primera carga: desde la próxima actualización este correo muestra los cambios.' if primera else ('Sin cambios desde la actualización anterior.' if not total_cambios else f'{total_cambios} cambios desde la actualización anterior.')}</p>
{bloques}
<h3 style="margin:22px 0 6px;color:#1d5a9e">PT por estado en Centrality</h3>
<table style="border-collapse:collapse;font-size:13px">{resumen_estados}</table>
<p style="color:#777;font-size:12px;margin-top:20px">{'La agenda completa está en la web (botón de arriba). Pide la contraseña del equipo.' if url else 'La agenda completa va adjunta (centrality_2_0.html): ábrela en el navegador.'}</p>
</body></html>"""
    msg = MIMEMultipart()
    corto = {"estado": "cambios de estado", "nuevo": "PT nuevos", "fecha": "cambios de fecha"}
    msg["Subject"] = f"[Centrality 2.0] {leido} · " + ("primera carga" if primera else ", ".join(f"{len(por[k])} {corto[k]}" for k in corto if por.get(k)) or "sin cambios")
    msg["From"], msg["To"] = usuario, dest
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg.attach(MIMEText(cuerpo, "html", "utf-8"))
    # Los servidores de correo corporativos suelen bloquear adjuntos .html grandes;
    # si hay página web, el correo lleva solo el enlace.
    if not url or os.environ.get("ADJUNTAR_HTML", "false").lower() == "true":
        parte = MIMEApplication(adjunto.read_bytes(), Name="centrality_2_0.html")
        parte["Content-Disposition"] = 'attachment; filename="centrality_2_0.html"'
        msg.attach(parte)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as s:
        s.login(usuario, clave)
        s.sendmail(usuario, [dest] + cc, msg.as_string())
    log(f"Correo enviado a {dest}" + (f" con copia a {len(cc)}" if cc else ""))


def main():
    t0 = time.time()
    usuario, clave = os.environ.get("SAESA_USER"), os.environ.get("SAESA_PASS")
    if not usuario or not clave:
        sys.exit("Faltan SAESA_USER y SAESA_PASS.")
    DATOS.mkdir(exist_ok=True)
    archivo = DATOS / "pts.json"
    prev = json.loads(archivo.read_text(encoding="utf-8")) if archivo.exists() else {}
    primera = not prev
    c = Centrality(usuario, clave)
    c.login()
    lista = c.todos()
    log(f"{len(lista)} PT con inicio entre {DESDE} y {HASTA}")

    hoy = datetime.now(TZ).date()
    ventana = (str(hoy - timedelta(days=VENTANA_DIAS_ATRAS)), str(hoy + timedelta(days=VENTANA_DIAS_ADELANTE)))
    nuevo, pendientes = {}, []
    for x in lista:
        b = de_lista(x)
        p = prev.get(b["pt"])
        cambio = p is None or any(p.get(k) != b.get(k) for k in CAMPOS_LISTA if k != "desc")
        cerca = ventana[0] <= fecha_local(b["ini"]) <= ventana[1]
        if FULL or cambio or cerca or not (p and p.get("ficha_leida")):
            pendientes.append(b)
        else:
            for k in ("ref", "caso", "cond", "adj", "ficha_leida", "desc", "creador", "solicitante"):
                if k in p:
                    b[k] = p[k]
        nuevo[b["pt"]] = b
    log(f"Fichas a leer: {len(pendientes)} (el resto se reutiliza)")

    errores = []

    def leer(b):
        try:
            return con_ficha(b, c.ficha(b["id"]))
        except Exception as e:
            errores.append((b["pt"], str(e)[:120]))
            p = prev.get(b["pt"]) or {}
            for k in ("ref", "caso", "cond", "adj", "ficha_leida"):
                if k in p:
                    b[k] = p[k]
            return b

    with ThreadPoolExecutor(HILOS) as ex:
        for i, b in enumerate(ex.map(leer, pendientes), 1):
            nuevo[b["pt"]] = b
            if i % 200 == 0:
                log(f"  {i}/{len(pendientes)} fichas")
    if errores:
        log(f"{len(errores)} fichas no se pudieron leer; se usó la versión anterior. Ejemplos: {errores[:3]}")

    cambios = [] if primera else comparar(prev, nuevo)
    leido = datetime.now(TZ).strftime("%d/%m/%Y %H:%M")
    archivo.write_text(json.dumps(nuevo, ensure_ascii=False), encoding="utf-8")
    if cambios:
        with open(DATOS / "cambios.jsonl", "a", encoding="utf-8") as f:
            for cb in cambios:
                f.write(json.dumps(dict(cb, cuando=leido), ensure_ascii=False) + "\n")
    html = generar_html(nuevo, cambios, leido)
    log(f"Agenda generada: {html} ({html.stat().st_size // 1024} KB). Cambios: {len(cambios)}{' (primera carga)' if primera else ''}")
    enviar_correo(cambios, nuevo, leido, html, primera)
    log(f"Listo en {int(time.time() - t0)} s")


if __name__ == "__main__":
    main()
