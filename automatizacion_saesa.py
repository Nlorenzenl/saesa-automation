"""
Automatización SAESA – Autorización diaria de PT's
El trigger de Estado está en y=547 exactamente (confirmado por logs)
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

# Coordenada Y exacta del trigger de Estado (confirmada por logs: label "Estado" en y=547)
ESTADO_TRIGGER_Y = 547


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
    print("\n[3] FILTRO - Estado PCCT")
    await frame.click('text=Filtro')
    await page.wait_for_timeout(2500)
    await screenshot(page, "04_filtro_abierto")

    # El trigger de Estado está en x=838, y=547 (confirmado por logs anteriores)
    # El label "Estado:" está en y=547 y su trigger arrow está en esa misma fila
    # Necesitamos hacer clic específicamente en el arrow trigger de y≈547
    click_result = await frame.evaluate('''() => {
        // Buscar TODOS los triggers y filtrar el que está en y≈547
        const triggers = Array.from(document.querySelectorAll("img.x-form-arrow-trigger, img.x-form-trigger"));
        const log = [];

        for (const t of triggers) {
            if (!t.offsetParent) continue;
            const r = t.getBoundingClientRect();
            const y = Math.round(r.y);
            const x = Math.round(r.x);
            log.push("trigger en (" + x + ", " + y + ") cls=" + t.className);

            // El trigger de Estado está en y=547 +/- 5px
            if (y >= 542 && y <= 552 && x > 800) {
                t.click();
                return {ok: true, x: x, y: y, cls: t.className, log: log};
            }
        }

        // Si no encontramos en y=547, intentar buscar por la posición del label "Estado:"
        const allEls = Array.from(document.querySelectorAll("*"));
        for (const el of allEls) {{
            if (!el.offsetParent) continue;
            const txt = (el.innerText || "").trim();
            if (txt !== "Estado:") continue;
            const lr = el.getBoundingClientRect();
            log.push(`Label Estado: en y=${Math.round(lr.y)}`);

            // Buscar el arrow trigger más cercano en Y a este label
            let closest = null, minDist = 999;
            for (const t of triggers) {{
                if (!t.offsetParent) continue;
                const tr = t.getBoundingClientRect();
                if (!t.className.includes('arrow')) continue;
                const dist = Math.abs(tr.y - lr.y);
                if (dist < minDist && tr.x > 800) {{
                    minDist = dist;
                    closest = t;
                }}
            }}
            if (closest && minDist < 20) {{
                const cr = closest.getBoundingClientRect();
                closest.click();
                return {{ok: true, via: "label_proximity", x: Math.round(cr.x), y: Math.round(cr.y), dist: minDist, log}};
            }}
        }}

        return {{ok: false, log}};
    }}''')
    print(f"  → Click trigger Estado: {click_result}")
    await page.wait_for_timeout(2000)
    await screenshot(page, "04b_dropdown_estado")

    # Verificar qué apareció en el dropdown
    dropdown_items = await frame.evaluate('''() => {
        const all = Array.from(document.querySelectorAll("div,li,td,span"));
        const visible = all.filter(el => el.offsetParent && el.children.length === 0);
        const items = visible.map(el => (el.innerText || "").trim()).filter(t => t.length > 2 && t.length < 80);
        return [...new Set(items)].slice(0, 10);
    }''')
    print(f"  → Items en dropdown: {dropdown_items}")

    # Seleccionar PCCT
    pcct_ok = await frame.evaluate(f'''() => {{
        const all = Array.from(document.querySelectorAll("div,li,td,span"));
        for (const el of all) {{
            if (!el.offsetParent) continue;
            const t = (el.innerText || "").trim();
            if (t === "{ESTADO_FILTRO}") {{ el.click(); return {{ok: true, exact: true, text: t}}; }}
        }}
        for (const el of all) {{
            if (!el.offsetParent) continue;
            const t = (el.innerText || "").trim();
            if (t.includes("PCCT") && t.includes("Revisión") && t.length < 80) {{
                el.click(); return {{ok: true, text: t}};
            }}
        }}
        return {{ok: false}};
    }}''')
    print(f"  → PCCT: {pcct_ok}")
    await page.wait_for_timeout(500)
    await screenshot(page, "04c_pcct")

    # Aplicar filtro
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

    count_info = await frame.evaluate('''() => {
        const rows = document.querySelectorAll(".x-grid3-row");
        const textos = Array.from(document.querySelectorAll("*"))
            .filter(el => el.children.length === 0 && el.offsetParent && (el.innerText || "").includes("Mostrando"))
            .map(e => e.innerText.trim());
        return {filas_visibles: rows.length, textos};
    }''')
    print(f"  → Resultado filtro: {count_info}")
    print("  ✓ Filtro aplicado")


async def aprobar_pts(page, frame):
    print("\n[4] APROBANDO PT's METROPOLITANA + PCCT")
    pts_aprobados = []
    pts_fallidos  = []
    pts_omitidos  = []

    total_paginas = await frame.evaluate('''() => {
        const all = Array.from(document.querySelectorAll("*"));
        for (const el of all) {
            if (!el.offsetParent || el.children.length > 2) continue;
            const t = (el.innerText || "").trim();
            if (/^de [0-9]+$/.test(t)) return parseInt(t.replace("de ", ""));
        }
        return 1;
    }''')
    print(f"  → Total páginas: {total_paginas}")

    # Con filtro PCCT correcto deberían ser 1-2 páginas (9 PT's)
    # Aumentamos a 10 páginas por si acaso
    paginas_a_leer = min(total_paginas, 10)

    for pagina in range(1, paginas_a_leer + 1):
        print(f"\n  → Página {pagina}/{paginas_a_leer}")

        filas_info = await frame.evaluate('''() => {
            const filas = Array.from(document.querySelectorAll(".x-grid3-row"));
            return filas.map(f => {
                const celdas = Array.from(f.querySelectorAll(".x-grid3-cell-inner"));
                return celdas.map(c => c.innerText.trim());
            });
        }''')

        # Diagnóstico de la primera fila para ver la estructura de columnas
        if filas_info:
            print(f"    Estructura primera fila ({len(filas_info[0])} cols): {filas_info[0][:6]}")

        pts_pagina = []
        for row in filas_info:
            if not row:
                continue

            # Buscar el ID (formato YYYY-NNNNN)
            id_pt = ""
            area_pt = ""
            estado_pt = ""

            for cell in row:
                if re.match(r'^\d{4}-\d{5}$', cell):
                    id_pt = cell
                elif any(x in cell for x in ["Metropolitana", "OSORNO", "Antofagasta", "Chiloé", "Copiapó", "LLVV", "SCADA", "PROTECCIONES"]):
                    area_pt = cell
                elif any(x in cell for x in ["Revisión", "Autorización", "Nueva", "Publicada", "Aprobada"]):
                    if not estado_pt:
                        estado_pt = cell

            if not id_pt:
                continue

            # Solo procesar PT's en estado PCCT
            if "PCCT" not in estado_pt:
                continue

            if AREA_REQUERIDA.lower() in area_pt.lower():
                pts_pagina.append({"id": id_pt, "area": area_pt})
                print(f"    ✅ {id_pt} | {area_pt[:40]} | {estado_pt[:35]}")
            else:
                pts_omitidos.append({"id": id_pt, "area": area_pt})
                print(f"    ⏭️  {id_pt} | {area_pt[:40]} — omitido")

        # Aprobar los PT's de esta página
        for pt in pts_pagina:
            try:
                print(f"\n    → Aprobando {pt['id']}...")
                encontrado = await frame.evaluate(f'''() => {{
                    const filas = Array.from(document.querySelectorAll(".x-grid3-row"));
                    for (const fila of filas) {{
                        const celdas = Array.from(fila.querySelectorAll(".x-grid3-cell-inner"));
                        for (const c of celdas) {{
                            if (c.innerText.trim() === "{pt['id']}") {{
                                fila.click();
                                return true;
                            }}
                        }}
                    }}
                    return false;
                }}''')

                if not encontrado:
                    pts_fallidos.append(f"{pt['id']} (no encontrado en tabla)")
                    print(f"    ✗ No encontrado en tabla")
                    continue

                await page.wait_for_timeout(800)
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
                print(f"    ✗ {msg}")
                await screenshot(page, f"error_{pt['id']}")

        # Ir a siguiente página
        if pagina < paginas_a_leer:
            sig = await frame.evaluate('''() => {
                const btn = document.querySelector(".x-tbar-page-next:not(.x-item-disabled)");
                if (btn) { btn.click(); return true; }
                return false;
            }''')
            if not sig:
                print("  → Fin de páginas")
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
        <p><strong>Criterio:</strong> Estado = {ESTADO_FILTRO} + Área "{AREA_REQUERIDA}"</p>
        {error_bloque}
        <h3 style="color:#006600">✅ Aprobados ({len(pts_aprobados)})</h3>
        <table style="border-collapse:collapse;width:100%">{lista_ok}</table>
        <h3 style="color:#cc0000;margin-top:20px">❌ Errores ({len(pts_fallidos)})</h3>
        <table style="border-collapse:collapse;width:100%">{lista_err}</table>
        <h3 style="color:#888;margin-top:20px">⏭️ Omitidos ({len(omitidos)})</h3>
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
