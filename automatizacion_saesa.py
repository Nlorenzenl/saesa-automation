"""
Automatizacion SAESA - Autorizacion diaria de PT's
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
ESTADO_FILTRO  = "Revision y Autorizacion PCCT"
AREA_REQUERIDA = "Metropolitana"

# JS puro como strings normales (sin f-string para evitar conflictos con {})
JS_CLICK_ESTADO_TRIGGER = """
() => {
    var triggers = Array.from(document.querySelectorAll("img.x-form-arrow-trigger"));
    var log = [];
    for (var i = 0; i < triggers.length; i++) {
        var t = triggers[i];
        if (!t.offsetParent) continue;
        var r = t.getBoundingClientRect();
        var ty = Math.round(r.y);
        var tx = Math.round(r.x);
        log.push("trigger y=" + ty + " x=" + tx + " cls=" + t.className);
        if (ty >= 542 && ty <= 555 && tx > 800) {
            t.click();
            return {ok: true, x: tx, y: ty, log: log};
        }
    }
    // Fallback: buscar label "Estado:" y el trigger mas cercano
    var allEls = Array.from(document.querySelectorAll("label, td, div, span"));
    for (var j = 0; j < allEls.length; j++) {
        var el = allEls[j];
        if (!el.offsetParent) continue;
        var txt = (el.innerText || "").trim();
        if (txt !== "Estado:") continue;
        var lr = el.getBoundingClientRect();
        log.push("Label Estado: en y=" + Math.round(lr.y));
        var closest = null;
        var minDist = 999;
        for (var k = 0; k < triggers.length; k++) {
            var t2 = triggers[k];
            if (!t2.offsetParent) continue;
            var tr2 = t2.getBoundingClientRect();
            var dist = Math.abs(tr2.y - lr.y);
            if (dist < minDist && tr2.x > 800) {
                minDist = dist;
                closest = t2;
            }
        }
        if (closest && minDist < 25) {
            var cr = closest.getBoundingClientRect();
            closest.click();
            return {ok: true, via: "label", x: Math.round(cr.x), y: Math.round(cr.y), dist: minDist, log: log};
        }
    }
    return {ok: false, log: log};
}
"""

JS_SELECT_PCCT = """
() => {
    var all = Array.from(document.querySelectorAll("div,li,td,span"));
    for (var i = 0; i < all.length; i++) {
        var el = all[i];
        if (!el.offsetParent) continue;
        var t = (el.innerText || "").trim();
        if (t === "Revisi\u00f3n y Autorizaci\u00f3n PCCT") { el.click(); return {ok: true, exact: true, text: t}; }
    }
    for (var j = 0; j < all.length; j++) {
        var el2 = all[j];
        if (!el2.offsetParent) continue;
        var t2 = (el2.innerText || "").trim();
        if (t2.indexOf("PCCT") >= 0 && t2.indexOf("Revisi") >= 0 && t2.length < 80) {
            el2.click(); return {ok: true, text: t2};
        }
    }
    return {ok: false};
}
"""

JS_CLICK_APLICAR = """
() => {
    var btns = Array.from(document.querySelectorAll("button,a,span"));
    for (var i = 0; i < btns.length; i++) {
        var btn = btns[i];
        if (!btn.offsetParent) continue;
        var t = (btn.innerText || btn.textContent || "").trim();
        if (t === "Aplicar") { btn.click(); return true; }
    }
    return false;
}
"""

JS_COUNT_RESULT = """
() => {
    var rows = document.querySelectorAll(".x-grid3-row");
    var textos = Array.from(document.querySelectorAll("*"))
        .filter(function(el) { return el.children.length === 0 && el.offsetParent && (el.innerText || "").indexOf("Mostrando") >= 0; })
        .map(function(e) { return e.innerText.trim(); });
    return {filas: rows.length, textos: textos};
}
"""

JS_GET_TOTAL_PAGES = """
() => {
    var all = Array.from(document.querySelectorAll("*"));
    for (var i = 0; i < all.length; i++) {
        var el = all[i];
        if (!el.offsetParent || el.children.length > 2) continue;
        var t = (el.innerText || "").trim();
        if (/^de [0-9]+$/.test(t)) return parseInt(t.replace("de ", ""));
    }
    return 1;
}
"""

JS_READ_ROWS = """
() => {
    var filas = Array.from(document.querySelectorAll(".x-grid3-row"));
    return filas.map(function(f) {
        var celdas = Array.from(f.querySelectorAll(".x-grid3-cell-inner"));
        return celdas.map(function(c) { return c.innerText.trim(); });
    });
}
"""

JS_NEXT_PAGE = """
() => {
    var btn = document.querySelector(".x-tbar-page-next:not(.x-item-disabled)");
    if (btn) { btn.click(); return true; }
    return false;
}
"""

JS_CLICK_APROBAR_POPUP = """
() => {
    var btns = Array.from(document.querySelectorAll("button,input"));
    for (var i = 0; i < btns.length; i++) {
        var btn = btns[i];
        if (!btn.offsetParent) continue;
        var t = (btn.innerText || btn.value || "").trim();
        if (t === "Aceptar") { btn.click(); return true; }
    }
    return false;
}
"""


async def screenshot(page, nombre):
    os.makedirs("capturas", exist_ok=True)
    path = "capturas/" + nombre + "_" + datetime.now().strftime("%H%M%S") + ".png"
    await page.screenshot(path=path, full_page=False)
    print("  captura: " + path)


async def get_content_frame(page):
    for f in page.frames:
        if f.name == "content":
            return f
    for f in page.frames:
        try:
            await f.wait_for_selector("text=Planificacion", timeout=2000)
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
            await inp.fill(SAESA_USER)
            break
    for inp in inputs:
        if (await inp.get_attribute("type") or "").lower() == "password" and await inp.is_visible():
            await inp.fill(SAESA_PASS)
            break
    await page.click('input[value="Login"], button:has-text("Login"), input[type="submit"]')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2000)
    print("  Login OK")


async def navegar_a_permisos(page):
    print("\n[2] NAVEGACION")
    await page.wait_for_selector('a:has-text("Aplicaciones"), span:has-text("Aplicaciones")', timeout=TIMEOUT)
    await page.click('a:has-text("Aplicaciones"), span:has-text("Aplicaciones")')
    await page.wait_for_timeout(1500)
    await page.wait_for_selector('a:has-text("DMS")', timeout=TIMEOUT)
    await page.click('a:has-text("DMS")')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2000)
    frame = await get_content_frame(page)
    print("  frame: " + frame.name)
    await frame.click("text=Planificacion")
    await page.wait_for_timeout(1000)
    await frame.click("text=Permisos de trabajo")
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    print("  En Permisos de trabajo")
    return frame


async def aplicar_filtro_estado(page, frame):
    print("\n[3] FILTRO - Estado PCCT")
    await frame.click("text=Filtro")
    await page.wait_for_timeout(2500)
    await screenshot(page, "04_filtro_abierto")

    # Clic en trigger de Estado (y=547 confirmado)
    r = await frame.evaluate(JS_CLICK_ESTADO_TRIGGER)
    print("  trigger Estado: " + str(r))
    await page.wait_for_timeout(2000)
    await screenshot(page, "04b_dropdown")

    # Seleccionar PCCT
    r2 = await frame.evaluate(JS_SELECT_PCCT)
    print("  PCCT: " + str(r2))
    await page.wait_for_timeout(500)

    # Aplicar
    await frame.evaluate(JS_CLICK_APLICAR)
    print("  Aplicar OK")
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    await screenshot(page, "05_filtro_aplicado")

    r3 = await frame.evaluate(JS_COUNT_RESULT)
    print("  Resultado: " + str(r3))


async def aprobar_pts(page, frame):
    print("\n[4] APROBANDO PT's")
    pts_aprobados = []
    pts_fallidos  = []
    pts_omitidos  = []

    total_paginas = await frame.evaluate(JS_GET_TOTAL_PAGES)
    print("  Total paginas: " + str(total_paginas))
    paginas = min(total_paginas, 10)

    for pagina in range(1, paginas + 1):
        print("\n  Pagina " + str(pagina) + "/" + str(paginas))

        filas_info = await frame.evaluate(JS_READ_ROWS)
        if filas_info:
            print("  Filas: " + str(len(filas_info)) + ", cols primera: " + str(len(filas_info[0])))

        pts_pagina = []
        for row in filas_info:
            if not row:
                continue
            id_pt = ""
            area_pt = ""
            estado_pt = ""
            for cell in row:
                if re.match(r"^\d{4}-\d{5}$", cell):
                    id_pt = cell
                elif any(x in cell for x in ["Metropolitana", "OSORNO", "Antofagasta", "Chiloe", "Copiapo", "LLVV", "SCADA", "PROTECCIONES", "TEMUCO", "Puerto Montt"]):
                    area_pt = cell
                elif "PCCT" in cell and not estado_pt:
                    estado_pt = cell
                elif "Revision" in cell or "Autorizacion" in cell:
                    if not estado_pt:
                        estado_pt = cell

            if not id_pt:
                continue
            if "PCCT" not in estado_pt:
                continue

            if AREA_REQUERIDA.lower() in area_pt.lower():
                pts_pagina.append({"id": id_pt, "area": area_pt})
                print("    OK: " + id_pt + " | " + area_pt[:40])
            else:
                pts_omitidos.append({"id": id_pt, "area": area_pt})
                print("    SKIP: " + id_pt + " | " + area_pt[:40])

        # Aprobar PT's de esta pagina
        for pt in pts_pagina:
            try:
                print("\n    Aprobando " + pt["id"] + "...")

                # JS para encontrar y clicar la fila
                js_find = "() => { var filas = Array.from(document.querySelectorAll('.x-grid3-row')); for (var i=0; i<filas.length; i++) { var celdas = Array.from(filas[i].querySelectorAll('.x-grid3-cell-inner')); for (var j=0; j<celdas.length; j++) { if (celdas[j].innerText.trim() === '" + pt["id"] + "') { filas[i].click(); return true; } } } return false; }"

                encontrado = await frame.evaluate(js_find)
                if not encontrado:
                    pts_fallidos.append(pt["id"] + " (no encontrado)")
                    continue

                await page.wait_for_timeout(800)
                await frame.click('a:has-text("Aprobar"), button:has-text("Aprobar")')
                await page.wait_for_timeout(1500)

                aceptar = await frame.evaluate(JS_CLICK_APROBAR_POPUP)
                if not aceptar:
                    await page.evaluate(JS_CLICK_APROBAR_POPUP)

                await page.wait_for_load_state("networkidle", timeout=15_000)
                await page.wait_for_timeout(1500)
                pts_aprobados.append(pt["id"])
                print("    APROBADO: " + pt["id"])

            except Exception as e:
                msg = str(e)[:80]
                pts_fallidos.append(pt["id"] + ": " + msg)
                print("    ERROR: " + msg)
                await screenshot(page, "error_" + pt["id"])

        if pagina < paginas:
            sig = await frame.evaluate(JS_NEXT_PAGE)
            if not sig:
                print("  Fin paginas")
                break
            await page.wait_for_timeout(2000)

    await screenshot(page, "06_final")
    return pts_aprobados, pts_fallidos, pts_omitidos


def enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos=None, error_critico=None):
    fecha    = datetime.now().strftime("%d/%m/%Y %H:%M")
    omitidos = pts_omitidos or []
    lista_ok = "".join(
        "<tr><td style='padding:4px 12px'>OK</td><td style='font-family:monospace;padding:4px 12px'>" + pt + "</td></tr>"
        for pt in pts_aprobados
    ) or "<tr><td colspan='2' style='padding:4px 12px;color:#888'>Ninguno</td></tr>"
    lista_err = "".join(
        "<tr><td style='padding:4px 12px'>ERR</td><td style='padding:4px 12px'>" + pt + "</td></tr>"
        for pt in pts_fallidos
    ) or "<tr><td colspan='2' style='padding:4px 12px;color:#888'>Sin errores</td></tr>"
    lista_omit = "".join(
        "<tr><td style='padding:4px 12px'>SKIP</td><td style='padding:4px 12px;color:#888'>" + pt["id"] + " - " + pt["area"][:50] + "</td></tr>"
        for pt in omitidos
    ) or "<tr><td colspan='2' style='padding:4px 12px;color:#888'>Ninguno</td></tr>"
    error_bloque = (
        "<div style='background:#fff0f0;border-left:4px solid #c00;padding:12px;margin:16px 0'>"
        "<strong>Error critico:</strong><br><code>" + str(error_critico) + "</code></div>"
        if error_critico else ""
    )
    html = (
        "<html><body style='font-family:Arial,sans-serif;max-width:640px;margin:auto'>"
        "<div style='background:#003580;color:white;padding:20px;border-radius:8px 8px 0 0'>"
        "<h2 style='margin:0'>Reporte PT's - SAESA/DMS</h2>"
        "<p style='margin:4px 0 0;opacity:.8'>Autorizacion PCCT - Area Zonal Metropolitana</p>"
        "</div>"
        "<div style='border:1px solid #ddd;border-top:none;padding:20px;border-radius:0 0 8px 8px'>"
        "<p><strong>Fecha:</strong> " + fecha + "</p>"
        "<p><strong>Criterio:</strong> Estado = PCCT + Area Metropolitana</p>"
        + error_bloque +
        "<h3 style='color:#006600'>PT's Aprobados (" + str(len(pts_aprobados)) + ")</h3>"
        "<table style='border-collapse:collapse;width:100%'>" + lista_ok + "</table>"
        "<h3 style='color:#cc0000;margin-top:20px'>PT's con Error (" + str(len(pts_fallidos)) + ")</h3>"
        "<table style='border-collapse:collapse;width:100%'>" + lista_err + "</table>"
        "<h3 style='color:#888;margin-top:20px'>PT's Omitidos (" + str(len(omitidos)) + ")</h3>"
        "<table style='border-collapse:collapse;width:100%'>" + lista_omit + "</table>"
        "<p style='color:#999;font-size:11px;margin-top:20px'>Bot SAESA - GitHub Actions</p>"
        "</div></body></html>"
    )
    asunto = "[SAESA] " + datetime.now().strftime("%d/%m/%Y") + " - " + str(len(pts_aprobados)) + " aprobados"
    if error_critico:
        asunto = "[SAESA] ERROR " + datetime.now().strftime("%d/%m/%Y")
    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"] = GMAIL_USER
    msg["To"]   = EMAIL_DEST
    msg.attach(MIMEText(html, "html"))
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(GMAIL_USER, GMAIL_PASS)
            s.sendmail(GMAIL_USER, EMAIL_DEST, msg.as_string())
        print("  Correo enviado a " + EMAIL_DEST)
    except Exception as e:
        print("  Error correo: " + str(e))


async def main():
    print("\n" + "="*50 + "\n  SAESA | " + datetime.now().strftime("%d/%m/%Y %H:%M") + "\n" + "="*50)
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
            print("\nERROR CRITICO: " + str(e))
            await screenshot(page, "error_critico")
        finally:
            await browser.close()
    print("\n  " + str(len(pts_aprobados)) + " aprobados | " + str(len(pts_fallidos)) + " errores | " + str(len(pts_omitidos)) + " omitidos")
    enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos, error_critico)
    print("Fin.\n")


if __name__ == "__main__":
    asyncio.run(main())
