"""
Automatización SAESA – Autorización diaria de PT's
Filtro: Estado = PCCT + bandeja de trabajo
Aprobación: SOLO PT's cuya columna Área contenga "Metropolitana"
"""

import asyncio
import smtplib
import os
import re
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

SAESA_URL     = "https://stx.saesa.cl:8091/backend/sts/login.php?backurl=%2Fbackend%2Fsts%2Fcentrality.php"
SAESA_USER    = os.environ["SAESA_USER"]
SAESA_PASS    = os.environ["SAESA_PASS"]
GMAIL_USER    = os.environ["GMAIL_USER"]
GMAIL_PASS    = os.environ["GMAIL_APP_PASS"]
EMAIL_DEST    = os.environ["EMAIL_DEST"]
TIMEOUT       = 30_000
ESTADO_FILTRO = "Revisión y Autorización PCCT"
AREA_REQUERIDA = "Metropolitana"  # Solo aprobamos PT's cuya área contenga esta palabra


async def screenshot(page, nombre):
    os.makedirs("capturas", exist_ok=True)
    path = f"capturas/{nombre}_{datetime.now().strftime('%H%M%S')}.png"
    await page.screenshot(path=path, full_page=False)
    print(f"  📸 {path}")


async def get_content_frame(page):
    for f in page.frames:
        if f.name == 'content':
            return f
    for f in page.frames:
        try:
            await f.wait_for_selector('text=Planificación', timeout=2000)
            return f
        except Exception:
            continue
    return page


async def hacer_login(page):
    print("\n[1] LOGIN")
    await page.goto(SAESA_URL, wait_until="domcontentloaded", timeout=60_000)
    await page.wait_for_timeout(3000)
    inputs = await page.query_selector_all('input[type="text"], input:not([type]), input[type="password"]')
    for inp in inputs:
        typ = (await inp.get_attribute("type") or "text").lower()
        if typ in ("text", "") and await inp.is_visible():
            await inp.fill(SAESA_USER); break
    for inp in inputs:
        if (await inp.get_attribute("type") or "").lower() == "password" and await inp.is_visible():
            await inp.fill(SAESA_PASS); break
    await page.click('input[value="Login"], button:has-text("Login"), input[type="submit"]')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2000)
    print("  ✓ Login OK")
    await screenshot(page, "01_login_ok")


async def navegar_a_permisos(page):
    print("\n[2] NAVEGACIÓN")
    await page.wait_for_selector('a:has-text("Aplicaciones"), span:has-text("Aplicaciones")', timeout=TIMEOUT)
    await page.click('a:has-text("Aplicaciones"), span:has-text("Aplicaciones")')
    await page.wait_for_timeout(1500)
    await page.wait_for_selector('a:has-text("DMS")', timeout=TIMEOUT)
    await page.click('a:has-text("DMS")')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2000)

    frame = await get_content_frame(page)
    print(f"  → frame: '{frame.name}'")
    await frame.click('text=Planificación')
    await page.wait_for_timeout(1000)
    await frame.click('text=Permisos de trabajo')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    print("  ✓ En Permisos de trabajo")
    return frame


async def aplicar_filtro_estado(page, frame):
    """Aplica solo el filtro de Estado = PCCT (que sabemos que funciona)."""
    print("\n[3] FILTRO (solo Estado PCCT)")
    await frame.click('text=Filtro')
    await page.wait_for_timeout(2500)
    await screenshot(page, "04_filtro_abierto")

    # El Estado PCCT se selecciona via clic en el arrow trigger y luego en el item
    # Sabemos que el frame 'content' tiene el arrow en y≈521 y el dropdown aparece en ese frame
    estado_ok = await frame.evaluate(f'''() => {{
        // Buscar el arrow trigger del campo Estado (y≈521 en coordenadas del iframe)
        const triggers = Array.from(document.querySelectorAll("img.x-form-arrow-trigger"));
        // El trigger de Estado es el que está más cerca del label "Estado:"
        // Basado en logs anteriores: el Estado está en y≈521 dentro del iframe
        // Intentar disparar clic en todos y ver cuál abre el dropdown correcto
        for (const t of triggers) {{
            const r = t.getBoundingClientRect();
            if (r.y > 515 && r.y < 540 && r.x > 830) {{
                t.click();
                return {{found: true, x: Math.round(r.x), y: Math.round(r.y)}};
            }}
        }}
        // Fallback: último arrow trigger visible
        const visible = triggers.filter(t => t.offsetParent);
        if (visible.length > 0) {{
            const last = visible[visible.length - 1];
            last.click();
            const r = last.getBoundingClientRect();
            return {{found: true, via: "last", x: Math.round(r.x), y: Math.round(r.y)}};
        }}
        return {{found: false}};
    }}''')
    print(f"  → Click arrow Estado: {estado_ok}")
    await page.wait_for_timeout(1500)
    await screenshot(page, "04b_dropdown_estado")

    # Seleccionar PCCT en el dropdown (funciona en frame 'content')
    pcct_ok = await frame.evaluate(f'''() => {{
        const all = Array.from(document.querySelectorAll("div,li,td,span"));
        for (const el of all) {{
            if (!el.offsetParent) continue;
            const t = (el.innerText || "").trim();
            if (t.includes("PCCT") && t.includes("Revisión") && t.length < 60) {{
                el.click();
                return {{ok: true, text: t}};
            }}
        }}
        return {{ok: false}};
    }}''')
    print(f"  → PCCT seleccionado: {pcct_ok}")
    await page.wait_for_timeout(500)
    await screenshot(page, "04c_estado_pcct")

    # Clic en Aplicar
    await frame.evaluate('''() => {
        const btns = Array.from(document.querySelectorAll("button,input[type='button'],a,span"));
        for (const btn of btns) {
            if (!btn.offsetParent) continue;
            const t = (btn.innerText || btn.value || btn.textContent || "").trim();
            if (t === "Aplicar") { btn.click(); return; }
        }
    }''')
    print("  ✓ Aplicar")
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    await screenshot(page, "05_filtro_aplicado")
    print("  ✓ Filtro aplicado")


async def leer_todos_los_pts(frame):
    """
    Lee TODOS los PT's de la lista paginada (puede haber varias páginas).
    Retorna lista de dicts: {id, area, descripcion, fila_idx}
    Solo incluye PT's cuya área contenga "Metropolitana".
    """
    pts_metropolitana = []
    pts_omitidos = []
    pagina = 1

    while True:
        print(f"  → Leyendo página {pagina}...")
        await frame.wait_for_timeout(1000)

        filas = await frame.query_selector_all('table tbody tr')
        print(f"    Filas en página {pagina}: {len(filas)}")

        for i, fila in enumerate(filas):
            celdas = await fila.query_selector_all('td')
            if len(celdas) < 3:
                continue
            id_pt   = (await celdas[0].inner_text()).strip()
            area_pt = (await celdas[2].inner_text()).strip()  # columna Área
            desc_pt = (await celdas[-1].inner_text()).strip()[:60]

            if not re.match(r'\d{4}-\d{5}', id_pt):
                continue

            if AREA_REQUERIDA.lower() in area_pt.lower():
                pts_metropolitana.append({"id": id_pt, "area": area_pt, "desc": desc_pt})
                print(f"    ✅ {id_pt} | {area_pt[:35]}")
            else:
                pts_omitidos.append({"id": id_pt, "area": area_pt})
                print(f"    ⏭️  {id_pt} | {area_pt[:35]} — OMITIDO (no es Metropolitana)")

        # Verificar si hay página siguiente
        siguiente = await frame.query_selector('button[id*="next"]:not([disabled]), .x-tbar-page-next:not(.x-item-disabled), [class*="page-next"]:not([disabled])')
        if siguiente:
            await siguiente.click()
            await frame.wait_for_timeout(2000)
            pagina += 1
        else:
            break

    return pts_metropolitana, pts_omitidos


async def aprobar_pts(page, frame):
    """
    Aprueba los PT's uno a uno.
    Estrategia: buscar cada PT por su ID en la lista y aprobarlo.
    Solo aprueba PT's cuya área sea Metropolitana.
    """
    print("\n[4] LEYENDO PT's DE LA LISTA")
    pts_metro, pts_omitidos = await leer_todos_los_pts(frame)

    print(f"\n  → PT's Metropolitana a aprobar: {len(pts_metro)}")
    print(f"  → PT's omitidos (otra zona): {len(pts_omitidos)}")

    if not pts_metro:
        print("  ✓ No hay PT's Metropolitana para aprobar hoy.")
        await screenshot(page, "06_final")
        return [], []

    print("\n[5] APROBANDO PT's METROPOLITANA")
    pts_aprobados = []
    pts_fallidos  = []

    # Volver a página 1 si navegamos
    try:
        primera_pag = await frame.query_selector('button[id*="first"], .x-tbar-page-first, [class*="page-first"]')
        if primera_pag:
            await primera_pag.click()
            await frame.wait_for_timeout(1500)
    except Exception:
        pass

    for pt in pts_metro:
        try:
            print(f"\n  → Buscando PT {pt['id']} en la lista...")

            # Buscar la fila con este ID en la tabla actual
            encontrado = False
            for intento in range(3):  # máx 3 páginas buscando
                filas = await frame.query_selector_all('table tbody tr')
                for fila in filas:
                    celdas = await fila.query_selector_all('td')
                    if not celdas:
                        continue
                    id_celda = (await celdas[0].inner_text()).strip()
                    if id_celda == pt['id']:
                        # Hacer clic en la fila para seleccionarla
                        await fila.click()
                        await page.wait_for_timeout(800)
                        encontrado = True
                        print(f"    ✓ Fila encontrada")
                        break

                if encontrado:
                    break

                # No encontrado en esta página, ir a siguiente
                siguiente = await frame.query_selector('button[id*="next"]:not([disabled]), .x-tbar-page-next:not(.x-item-disabled)')
                if siguiente:
                    await siguiente.click()
                    await frame.wait_for_timeout(1500)
                else:
                    break

            if not encontrado:
                pts_fallidos.append(f"{pt['id']} (no encontrado en tabla)")
                print(f"    ✗ PT {pt['id']} no encontrado en la tabla")
                continue

            # Clic en botón "Aprobar" de la barra
            await frame.click('a:has-text("Aprobar"), button:has-text("Aprobar")')
            await page.wait_for_timeout(1500)
            await screenshot(page, f"aprobar_{pt['id']}")

            # Aceptar el popup (sin comentario)
            aceptar_ok = await frame.evaluate('''() => {
                for (const btn of document.querySelectorAll("button,input[type='button']")) {
                    if (!btn.offsetParent) continue;
                    if ((btn.innerText || btn.value || "").trim() === "Aceptar") {
                        btn.click(); return true;
                    }
                }
                return false;
            }''')
            if not aceptar_ok:
                # Fallback en page principal
                await page.evaluate('''() => {
                    for (const btn of document.querySelectorAll("button,input[type='button']")) {
                        if (!btn.offsetParent) continue;
                        if ((btn.innerText || btn.value || "").trim() === "Aceptar") {
                            btn.click();
                        }
                    }
                }''')

            await page.wait_for_load_state("networkidle", timeout=15_000)
            await page.wait_for_timeout(2000)
            pts_aprobados.append(pt['id'])
            print(f"    ✓ {pt['id']} APROBADO")

        except PlaywrightTimeout:
            pts_fallidos.append(f"{pt['id']} (timeout)")
            print(f"    ✗ Timeout aprobando {pt['id']}")
        except Exception as e:
            msg = str(e)[:80]
            pts_fallidos.append(f"{pt['id']}: {msg}")
            print(f"    ✗ Error: {msg}")
            await screenshot(page, f"error_{pt['id']}")

    await screenshot(page, "06_final")
    return pts_aprobados, pts_fallidos


def enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos=None, error_critico=None):
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    omitidos = pts_omitidos or []

    lista_ok = "".join(
        f"<tr><td style='padding:4px 12px'>✅</td><td style='font-family:monospace;padding:4px 12px'>{pt}</td></tr>"
        for pt in pts_aprobados
    ) or "<tr><td colspan='2' style='padding:4px 12px;color:#888'>Ninguno</td></tr>"

    lista_err = "".join(
        f"<tr><td style='padding:4px 12px'>❌</td><td style='padding:4px 12px'>{pt}</td></tr>"
        for pt in pts_fallidos
    ) or "<tr><td colspan='2' style='padding:4px 12px;color:#888'>Sin errores</td></tr>"

    lista_omit = "".join(
        f"<tr><td style='padding:4px 12px'>⏭️</td><td style='padding:4px 12px;color:#888'>{pt['id']} — {pt['area'][:50]}</td></tr>"
        for pt in omitidos
    ) or "<tr><td colspan='2' style='padding:4px 12px;color:#888'>Ninguno</td></tr>"

    error_bloque = (
        f"<div style='background:#fff0f0;border-left:4px solid #c00;padding:12px;margin:16px 0'>"
        f"<strong>⚠️ Error crítico:</strong><br><code>{error_critico}</code></div>"
        if error_critico else ""
    )
    html = f"""<html><body style="font-family:Arial,sans-serif;max-width:640px;margin:auto">
      <div style="background:#003580;color:white;padding:20px;border-radius:8px 8px 0 0">
        <h2 style="margin:0">📋 Reporte PT's – SAESA/DMS</h2>
        <p style="margin:4px 0 0;opacity:.8">Autorización PCCT – Área Zonal Metropolitana</p>
      </div>
      <div style="border:1px solid #ddd;border-top:none;padding:20px;border-radius:0 0 8px 8px">
        <p><strong>Fecha:</strong> {fecha}</p>
        <p><strong>Filtro aplicado:</strong> Estado = {ESTADO_FILTRO}</p>
        <p><strong>Criterio de aprobación:</strong> Solo PT's con área que contenga "{AREA_REQUERIDA}"</p>
        {error_bloque}
        <h3 style="color:#006600">✅ PT's Aprobados ({len(pts_aprobados)})</h3>
        <table style="border-collapse:collapse;width:100%">{lista_ok}</table>
        <h3 style="color:#cc0000;margin-top:20px">❌ PT's con Error ({len(pts_fallidos)})</h3>
        <table style="border-collapse:collapse;width:100%">{lista_err}</table>
        <h3 style="color:#888;margin-top:20px">⏭️ PT's Omitidos por área ({len(omitidos)})</h3>
        <table style="border-collapse:collapse;width:100%">{lista_omit}</table>
        <p style="color:#999;font-size:11px;margin-top:20px">Bot SAESA – GitHub Actions | Lun–Vie 08:00 Chile</p>
      </div></body></html>"""

    asunto = f"[SAESA] {datetime.now().strftime('%d/%m/%Y')} – {len(pts_aprobados)} aprobados, {len(omitidos)} omitidos"
    if error_critico:
        asunto = f"[SAESA] ⚠️ ERROR {datetime.now().strftime('%d/%m/%Y')}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"] = GMAIL_USER
    msg["To"] = EMAIL_DEST
    msg.attach(MIMEText(html, "html"))
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(GMAIL_USER, GMAIL_PASS)
            s.sendmail(GMAIL_USER, EMAIL_DEST, msg.as_string())
        print(f"  ✓ Correo enviado a {EMAIL_DEST}")
    except Exception as e:
        print(f"  ✗ Error correo: {e}")


async def main():
    print(f"\n{'='*55}\n  SAESA | {datetime.now().strftime('%d/%m/%Y %H:%M')}\n{'='*55}")
    pts_aprobados, pts_fallidos, pts_omitidos, error_critico = [], [], [], None

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=["--ignore-certificate-errors", "--no-sandbox",
                  "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        )
        page = await (await browser.new_context(
            ignore_https_errors=True, viewport={"width": 1400, "height": 900}
        )).new_page()
        try:
            await hacer_login(page)
            frame = await navegar_a_permisos(page)
            await aplicar_filtro_estado(page, frame)
            pts_aprobados, pts_fallidos = await aprobar_pts(page, frame)
        except Exception as e:
            error_critico = str(e)
            print(f"\n✗ ERROR CRÍTICO: {e}")
            await screenshot(page, "error_critico")
        finally:
            await browser.close()

    print(f"\n  {len(pts_aprobados)} aprobados | {len(pts_fallidos)} errores | {len(pts_omitidos)} omitidos")
    enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos, error_critico)
    print("✓ Fin.\n")


if __name__ == "__main__":
    asyncio.run(main())
