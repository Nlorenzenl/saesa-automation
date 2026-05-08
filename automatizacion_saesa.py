"""
Automatización SAESA – Autorización diaria de PT's
Filtro: Estado = PCCT (usando el arrow trigger correcto)
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

SAESA_URL      = "https://stx.saesa.cl:8091/backend/sts/login.php?backurl=%2Fbackend%2Fsts%2Fcentrality.php"
SAESA_USER     = os.environ["SAESA_USER"]
SAESA_PASS     = os.environ["SAESA_PASS"]
GMAIL_USER     = os.environ["GMAIL_USER"]
GMAIL_PASS     = os.environ["GMAIL_APP_PASS"]
EMAIL_DEST     = os.environ["EMAIL_DEST"]
TIMEOUT        = 30_000
ESTADO_FILTRO  = "Revisión y Autorización PCCT"
AREA_REQUERIDA = "Metropolitana"


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
    """
    Aplica filtro Estado = PCCT.
    El popup del filtro tiene estas filas (de arriba a abajo):
      En bandeja de trabajo (checkbox)
      Ids de permisos de trabajo (texto)
      Entidad de contexto de origen (combo, arrow en y≈261)
      Identificador de área (texto)
      Inicio disponibilidad desde/hasta (fechas)
      Disponibilidad horaria
      Ingresada desde/hasta
      Alimentador
      Elemento de maniobra
      Áreas (combo+list-trigger, arrow en y≈495)
      Tipo de permiso de trabajo (combo, arrow en y≈521)  ← ESTE ES EL QUE SE CLICKEÓ MAL
      Estado (combo, arrow en y≈547)                      ← ESTE ES EL CORRECTO
      Creador (combo, arrow en y≈573)
      Sector requirente (combo, arrow en y≈599)
    """
    print("\n[3] FILTRO - Estado PCCT")
    await frame.click('text=Filtro')
    await page.wait_for_timeout(2500)
    await screenshot(page, "04_filtro_abierto")

    # Diagnóstico: listar TODOS los arrow triggers con sus posiciones exactas
    triggers_info = await frame.evaluate('''() => {
        const triggers = Array.from(document.querySelectorAll("img.x-form-arrow-trigger"));
        return triggers.filter(t => t.offsetParent).map(t => {
            const r = t.getBoundingClientRect();
            // Buscar el label asociado mirando el td anterior en la tabla
            let label = "";
            try {
                const row = t.closest("tr");
                if (row) {
                    const firstTd = row.querySelector("td:first-child");
                    label = firstTd ? firstTd.innerText.trim() : "";
                }
            } catch(e) {}
            return {x: Math.round(r.x), y: Math.round(r.y), label};
        });
    }''')
    print(f"  → Arrow triggers con labels:")
    for i, t in enumerate(triggers_info):
        print(f"    [{i}] y={t['y']} label='{t['label']}'")

    # Buscar el arrow trigger cuyo label contenga "Estado"
    estado_trigger = None
    for t in triggers_info:
        if 'Estado' in t['label'] or 'estado' in t['label']:
            estado_trigger = t
            print(f"  → Trigger Estado encontrado: y={t['y']}")
            break

    # Si no encontramos por label, usar la posición conocida (y≈547 según el layout)
    if not estado_trigger:
        # Buscar el que esté entre y=540 y y=560
        for t in triggers_info:
            if 540 <= t['y'] <= 560:
                estado_trigger = t
                print(f"  → Trigger Estado por posición: y={t['y']}")
                break

    if not estado_trigger and triggers_info:
        # Fallback: el penúltimo visible (antes de Creador y Sector requirente)
        visible = [t for t in triggers_info if t['y'] > 400]
        if len(visible) >= 3:
            estado_trigger = visible[-3]  # antepenúltimo
            print(f"  → Trigger Estado por índice: y={estado_trigger['y']}")

    if estado_trigger:
        print(f"  → Click en trigger Estado ({estado_trigger['x']}, {estado_trigger['y']})")
        # Usar frame.evaluate para hacer clic directamente en el elemento
        await frame.evaluate(f'''() => {{
            const triggers = Array.from(document.querySelectorAll("img.x-form-arrow-trigger"));
            const target = triggers.find(t => {{
                const r = t.getBoundingClientRect();
                return Math.round(r.y) === {estado_trigger['y']} && t.offsetParent;
            }});
            if (target) target.click();
        }}''')
        await page.wait_for_timeout(2000)
        await screenshot(page, "04b_dropdown_estado")

        # Seleccionar PCCT en el dropdown
        pcct_ok = await frame.evaluate(f'''() => {{
            const all = Array.from(document.querySelectorAll("div,li,td,span"));
            for (const el of all) {{
                if (!el.offsetParent) continue;
                const t = (el.innerText || "").trim();
                if (t.includes("PCCT") && t.includes("Revisión") && t.length < 80) {{
                    el.click();
                    return {{ok: true, text: t}};
                }}
            }}
            return {{ok: false}};
        }}''')
        print(f"  → PCCT: {pcct_ok}")
        await page.wait_for_timeout(500)
        await screenshot(page, "04c_pcct_seleccionado")
    else:
        print("  ⚠️ No se encontró trigger de Estado")

    # Clic en Aplicar
    await frame.evaluate('''() => {
        const btns = Array.from(document.querySelectorAll("button,a,span"));
        for (const btn of btns) {
            if (!btn.offsetParent) continue;
            if ((btn.innerText || btn.textContent || "").trim() === "Aplicar") {
                btn.click(); return;
            }
        }
    }''')
    print("  ✓ Aplicar")
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    await screenshot(page, "05_filtro_aplicado")

    # Verificar cuántos registros quedaron
    total = await frame.evaluate('''() => {
        const el = document.querySelector('[class*="paging-info"], .x-paging-info, [id*="paging-info"]');
        return el ? el.innerText : "no encontrado";
    }''')
    print(f"  → Total registros tras filtro: {total}")
    print("  ✓ Filtro aplicado")


async def leer_pts_pagina_actual(frame):
    """Lee los PT's de la página actual. Retorna (metropolitana[], omitidos[])."""
    pts_metro = []
    pts_omit  = []

    # ExtJS renderiza solo las filas visibles en el viewport
    # La tabla real tiene class x-grid3-row
    filas_info = await frame.evaluate('''() => {
        const filas = Array.from(document.querySelectorAll(".x-grid3-row"));
        return filas.map(f => {
            const celdas = Array.from(f.querySelectorAll(".x-grid3-cell-inner"));
            return celdas.map(c => c.innerText.trim());
        }).filter(row => row.length > 0);
    }''')

    for row in filas_info:
        if not row or not re.match(r'\d{4}-\d{5}', row[0] if row else ''):
            continue
        id_pt   = row[0]
        # Columna 2 es Área (índice 2: Id, Fecha, Área, Estado, Tipo, Desc)
        area_pt = row[2] if len(row) > 2 else ""
        estado_pt = row[3] if len(row) > 3 else ""

        # Solo procesar los que tienen estado PCCT
        if "PCCT" not in estado_pt:
            continue

        if AREA_REQUERIDA.lower() in area_pt.lower():
            pts_metro.append({"id": id_pt, "area": area_pt})
            print(f"    ✅ {id_pt} | {area_pt[:40]} | {estado_pt[:30]}")
        else:
            pts_omit.append({"id": id_pt, "area": area_pt})
            print(f"    ⏭️  {id_pt} | {area_pt[:40]} — omitido")

    return pts_metro, pts_omit


async def aprobar_pts(page, frame):
    print("\n[4] APROBANDO PT's METROPOLITANA con estado PCCT")
    pts_aprobados = []
    pts_fallidos  = []
    pts_omitidos  = []

    # Obtener total de páginas
    paginacion = await frame.evaluate('''() => {
        const info = document.querySelector(".x-paging-info, [class*='paging-info']");
        const total_el = document.querySelector("input.x-tbar-page-number");
        const pages_el = document.querySelectorAll(".x-paging-info");
        return {
            info: info ? info.innerText : "",
            currentPage: total_el ? total_el.value : "1"
        };
    }''')
    print(f"  → Paginación: {paginacion}")

    # Obtener número total de páginas del indicador "Página X de Y"
    total_paginas = await frame.evaluate('''() => {
        // Buscar el texto "de X" en la barra de paginación
        const all = Array.from(document.querySelectorAll("*"));
        for (const el of all) {
            if (!el.offsetParent || el.children.length > 2) continue;
            const t = (el.innerText || "").trim();
            if (t.match(/^de \d+$/)) return parseInt(t.replace("de ", ""));
        }
        return 1;
    }''')
    print(f"  → Total páginas: {total_paginas}")

    # Leer SOLO la primera página para identificar PT's Metropolitana PCCT
    # (el filtro PCCT + bandeja ya reduce bastante)
    print(f"\n  → Leyendo PT's en las páginas disponibles...")

    paginas_a_leer = min(total_paginas, 20)  # máximo 20 páginas para evitar timeout

    for pagina in range(1, paginas_a_leer + 1):
        print(f"  → Página {pagina}/{paginas_a_leer}")
        metro, omit = await leer_pts_pagina_actual(frame)
        pts_omitidos.extend(omit)

        # Aprobar los PT's metropolitana encontrados en ESTA página
        for pt in metro:
            try:
                print(f"\n  → Aprobando {pt['id']}...")
                # Buscar la fila en la tabla actual
                encontrado = await frame.evaluate(f'''() => {{
                    const filas = Array.from(document.querySelectorAll(".x-grid3-row"));
                    for (const fila of filas) {{
                        const celda = fila.querySelector(".x-grid3-cell-inner");
                        if (celda && celda.innerText.trim() === "{pt['id']}") {{
                            fila.click();
                            return true;
                        }}
                    }}
                    return false;
                }}''')

                if not encontrado:
                    pts_fallidos.append(f"{pt['id']} (fila no encontrada)")
                    continue

                await page.wait_for_timeout(800)

                # Clic en Aprobar
                await frame.click('a:has-text("Aprobar"), button:has-text("Aprobar")')
                await page.wait_for_timeout(1500)
                await screenshot(page, f"aprobar_{pt['id']}")

                # Aceptar popup
                aceptar = await frame.evaluate('''() => {
                    for (const btn of document.querySelectorAll("button,input[type='button']")) {
                        if (!btn.offsetParent) continue;
                        if ((btn.innerText || btn.value || "").trim() === "Aceptar") {
                            btn.click(); return true;
                        }
                    }
                    return false;
                }''')
                if not aceptar:
                    await page.evaluate('''() => {
                        for (const btn of document.querySelectorAll("button,input[type='button']")) {
                            if (!btn.offsetParent) continue;
                            if ((btn.innerText || btn.value || "").trim() === "Aceptar") btn.click();
                        }
                    }''')

                await page.wait_for_load_state("networkidle", timeout=15_000)
                await page.wait_for_timeout(1500)
                pts_aprobados.append(pt['id'])
                print(f"    ✓ {pt['id']} APROBADO")

            except Exception as e:
                msg = str(e)[:80]
                pts_fallidos.append(f"{pt['id']}: {msg}")
                print(f"    ✗ Error: {msg}")
                await screenshot(page, f"error_{pt['id']}")

        # Ir a página siguiente si hay más
        if pagina < paginas_a_leer:
            siguiente_ok = await frame.evaluate('''() => {
                // Botón siguiente: class x-tbar-page-next, no debe tener x-item-disabled
                const btn = document.querySelector(".x-tbar-page-next:not(.x-item-disabled)");
                if (btn) { btn.click(); return true; }
                return false;
            }''')
            if not siguiente_ok:
                print("  → No hay más páginas")
                break
            await page.wait_for_timeout(2000)

    await screenshot(page, "06_final")
    return pts_aprobados, pts_fallidos, pts_omitidos


def enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos=None, error_critico=None):
    fecha    = datetime.now().strftime("%d/%m/%Y %H:%M")
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
        <p><strong>Criterio:</strong> Estado = {ESTADO_FILTRO} + Área contiene "{AREA_REQUERIDA}"</p>
        {error_bloque}
        <h3 style="color:#006600">✅ PT's Aprobados ({len(pts_aprobados)})</h3>
        <table style="border-collapse:collapse;width:100%">{lista_ok}</table>
        <h3 style="color:#cc0000;margin-top:20px">❌ PT's con Error ({len(pts_fallidos)})</h3>
        <table style="border-collapse:collapse;width:100%">{lista_err}</table>
        <h3 style="color:#888;margin-top:20px">⏭️ PT's Omitidos por área ({len(omitidos)})</h3>
        <table style="border-collapse:collapse;width:100%">{lista_omit}</table>
        <p style="color:#999;font-size:11px;margin-top:20px">Bot SAESA – GitHub Actions | Lun–Vie 08:00 Chile</p>
      </div></body></html>"""

    asunto = f"[SAESA] {datetime.now().strftime('%d/%m/%Y')} – {len(pts_aprobados)} aprobados"
    if error_critico:
        asunto = f"[SAESA] ⚠️ ERROR {datetime.now().strftime('%d/%m/%Y')}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"] = GMAIL_USER
    msg["To"]   = EMAIL_DEST
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
            pts_aprobados, pts_fallidos, pts_omitidos = await aprobar_pts(page, frame)
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
