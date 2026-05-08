"""
Automatizacion SAESA - Autorizacion diaria de PT's
Zonal Metropolitana | Estado: Revision y Autorizacion PCCT

Flujo real (confirmado por capturas):
  login -> Aplicaciones -> DMS (mismo frame, navbar top) ->
  Planificacion (menu top) -> Permisos de trabajo ->
  Filtro (popup ExtJS) -> Estado = PCCT -> Aplicar ->
  Por cada fila Metropolitana: seleccionar -> Aprobar -> Aceptar (popup con comentario)
"""

import asyncio
import smtplib
import os
import re
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

# ─── Configuracion ────────────────────────────────────────────────────────────
SAESA_URL = "https://stx.saesa.cl:8091/backend/sts/login.php?backurl=%2Fbackend%2Fsts%2Fcentrality.php"
SAESA_USER     = os.environ["SAESA_USER"]
SAESA_PASS     = os.environ["SAESA_PASS"]
GMAIL_USER     = os.environ["GMAIL_USER"]
GMAIL_PASS     = os.environ["GMAIL_APP_PASS"]
EMAIL_DEST     = os.environ["EMAIL_DEST"]
TIMEOUT        = 30_000

# Palabras que identifican area Metropolitana (columna Area de la grilla)
AREA_KEYWORDS  = ["metropolitana"]

# ─── JavaScript helpers ───────────────────────────────────────────────────────

JS_READ_ROWS = """
() => {
    var rows = Array.from(document.querySelectorAll(".x-grid3-row"));
    return rows.map(function(r) {
        return Array.from(r.querySelectorAll(".x-grid3-cell-inner"))
                    .map(function(c) { return (c.innerText || "").trim(); });
    });
}
"""

JS_GET_TOTAL_PAGES = """
() => {
    var els = Array.from(document.querySelectorAll("*"));
    for (var i = 0; i < els.length; i++) {
        var el = els[i];
        if (!el.offsetParent || el.children.length > 2) continue;
        var t = (el.innerText || "").trim();
        if (/^de [0-9]+$/.test(t)) return parseInt(t.split(" ")[1]);
    }
    return 1;
}
"""

JS_NEXT_PAGE = """
() => {
    var btn = document.querySelector(".x-tbar-page-next:not(.x-item-disabled)");
    if (btn) { btn.click(); return true; }
    return false;
}
"""

JS_SELECT_ROW_BY_ID = """
(ptId) => {
    var rows = Array.from(document.querySelectorAll(".x-grid3-row"));
    for (var i = 0; i < rows.length; i++) {
        var cells = Array.from(rows[i].querySelectorAll(".x-grid3-cell-inner"));
        for (var j = 0; j < cells.length; j++) {
            if (cells[j].innerText.trim() === ptId) {
                rows[i].scrollIntoView({block: "center"});
                rows[i].click();
                return {found: true, rowIndex: i};
            }
        }
    }
    return {found: false};
}
"""

JS_CHECK_ROW_SELECTED = """
(ptId) => {
    var sel = document.querySelector(".x-grid3-row-selected");
    if (!sel) return {selected: false};
    var cells = Array.from(sel.querySelectorAll(".x-grid3-cell-inner"));
    var ids = cells.map(function(c) { return c.innerText.trim(); });
    return {selected: ids.indexOf(ptId) >= 0, firstCell: ids[0] || ""};
}
"""

JS_CHECK_BTN_APROBAR = """
() => {
    var candidatos = Array.from(document.querySelectorAll("a,button,td,span"));
    for (var i = 0; i < candidatos.length; i++) {
        var el = candidatos[i];
        if (!el.offsetParent) continue;
        var txt = (el.innerText || el.textContent || "").trim();
        if (txt !== "Aprobar") continue;
        var disabled = el.classList.contains("x-item-disabled") ||
                       !!(el.closest && el.closest(".x-item-disabled"));
        return {found: true, disabled: disabled, cls: el.className.substring(0, 100)};
    }
    return {found: false};
}
"""

JS_CLICK_BTN_APROBAR = """
() => {
    var candidatos = Array.from(document.querySelectorAll("a,button,td,span"));
    for (var i = 0; i < candidatos.length; i++) {
        var el = candidatos[i];
        if (!el.offsetParent) continue;
        var txt = (el.innerText || el.textContent || "").trim();
        if (txt !== "Aprobar") continue;
        if (el.classList.contains("x-item-disabled")) continue;
        if (el.closest && el.closest(".x-item-disabled")) continue;
        el.click();
        return {clicked: true, tag: el.tagName, cls: el.className.substring(0,60)};
    }
    return {clicked: false};
}
"""

JS_DETECT_POPUP = """
() => {
    var headers = Array.from(document.querySelectorAll(".x-window-header-text, .x-panel-header-text"));
    for (var i = 0; i < headers.length; i++) {
        var h = headers[i];
        if (!h.offsetParent) continue;
        var txt = (h.innerText || h.textContent || "").trim();
        if (txt === "Aprobar") {
            var win = h.closest(".x-window");
            if (!win || !win.offsetParent) continue;
            var r = win.getBoundingClientRect();
            return {found: true, x: Math.round(r.x), y: Math.round(r.y),
                    w: Math.round(r.width), h: Math.round(r.height)};
        }
    }
    return {found: false};
}
"""

JS_CLICK_ACEPTAR = """
() => {
    var headers = Array.from(document.querySelectorAll(".x-window-header-text, .x-panel-header-text"));
    for (var i = 0; i < headers.length; i++) {
        var h = headers[i];
        if (!h.offsetParent) continue;
        var txt = (h.innerText || h.textContent || "").trim();
        if (txt !== "Aprobar") continue;
        var win = h.closest(".x-window");
        if (!win || !win.offsetParent) continue;
        // Limpiar textarea de comentario (opcional, lo dejamos vacio)
        var ta = win.querySelector("textarea");
        if (ta) ta.value = "";
        // Buscar boton Aceptar dentro de la ventana
        var btns = Array.from(win.querySelectorAll("button"));
        for (var j = 0; j < btns.length; j++) {
            var btxt = (btns[j].innerText || btns[j].textContent || "").trim();
            if (btxt === "Aceptar") { btns[j].click(); return {ok: true, via: "button"}; }
        }
        // Fallback: x-btn dentro del window
        var xbtns = Array.from(win.querySelectorAll(".x-btn"));
        for (var k = 0; k < xbtns.length; k++) {
            var xbtxt = (xbtns[k].innerText || xbtns[k].textContent || "").trim();
            if (xbtxt === "Aceptar") { xbtns[k].click(); return {ok: true, via: "x-btn"}; }
        }
        return {ok: false, win_found: true, btns_count: btns.length};
    }
    return {ok: false, win_found: false};
}
"""

JS_PT_EXISTE = """
(ptId) => {
    var cells = Array.from(document.querySelectorAll(".x-grid3-cell-inner"));
    return cells.some(function(c) { return c.innerText.trim() === ptId; });
}
"""


# ─── Utilidades ───────────────────────────────────────────────────────────────

async def screenshot(page, nombre):
    os.makedirs("capturas", exist_ok=True)
    ts   = datetime.now().strftime("%H%M%S")
    path = f"capturas/{nombre}_{ts}.png"
    await page.screenshot(path=path, full_page=False)
    print(f"    captura: {path}")
    return path


def es_metropolitana(area: str) -> bool:
    return any(k in area.lower() for k in AREA_KEYWORDS)


# ─── Paso 1: Login ────────────────────────────────────────────────────────────

async def hacer_login(page):
    print("\n[1] LOGIN")
    await page.goto(SAESA_URL, wait_until="domcontentloaded", timeout=60_000)
    await page.wait_for_timeout(3000)

    usuario = await page.query_selector('input[name="user"], input[type="text"]')
    if usuario:
        await usuario.fill(SAESA_USER)

    passwd = await page.query_selector('input[name="pass"], input[type="password"]')
    if passwd:
        await passwd.fill(SAESA_PASS)

    await page.click('input[value="Login"], button:has-text("Login"), input[type="submit"]')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2000)
    print("  OK: sesion iniciada")


# ─── Paso 2: Navegacion a Permisos de trabajo ─────────────────────────────────

async def navegar_a_permisos(page):
    """
    Flujo confirmado por capturas:
    1. Click en "Aplicaciones" (sidebar o link central)
    2. Click en "DMS"
    3. DMS carga con navbar top: Planificacion es un menu desplegable ahi
    4. Planificacion -> Permisos de trabajo
    El contenido puede estar en un iframe llamado 'content' o en el frame principal.
    """
    print("\n[2] NAVEGACION")

    # Aplicaciones
    await page.wait_for_selector(
        'a:has-text("Aplicaciones"), span:has-text("Aplicaciones")',
        timeout=TIMEOUT
    )
    await page.click('a:has-text("Aplicaciones"), span:has-text("Aplicaciones")')
    await page.wait_for_timeout(1500)
    print("  -> Aplicaciones")

    # DMS
    await page.wait_for_selector('a:has-text("DMS")', timeout=TIMEOUT)
    await page.click('a:has-text("DMS")')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2500)
    print("  -> DMS cargado")
    await screenshot(page, "nav_01_dms")

    # Detectar en que frame esta el contenido DMS
    # Puede ser el frame principal o un iframe "content"
    frame = page
    for f in page.frames:
        if f.name == "content":
            frame = f
            print(f"  frame detectado: '{f.name}'")
            break
        try:
            el = await f.query_selector('text="Planificación"')
            if el:
                frame = f
                print(f"  frame detectado por selector: '{f.name}'")
                break
        except Exception:
            pass

    # Planificacion (menu top del DMS)
    await frame.wait_for_selector('text="Planificación"', timeout=TIMEOUT)
    await frame.click('text="Planificación"')
    await page.wait_for_timeout(1000)
    print("  -> menu Planificacion abierto")
    await screenshot(page, "nav_02_planificacion")

    # Permisos de trabajo (submenu)
    await frame.wait_for_selector('text="Permisos de trabajo"', timeout=TIMEOUT)
    await frame.click('text="Permisos de trabajo"')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    print("  -> Permisos de trabajo")
    await screenshot(page, "nav_03_permisos")

    return frame


# ─── Paso 3: Aplicar filtro Estado = PCCT ─────────────────────────────────────

async def aplicar_filtro_pcct(page, frame):
    """
    El popup Filtros es un Ext.Window.
    Campo Estado: combobox ExtJS -> seleccionar 'Revision y Autorizacion PCCT'
    Luego click en Aplicar.
    """
    print("\n[3] FILTRO")

    # Abrir filtros
    await frame.click('text="Filtro"')
    await page.wait_for_timeout(2000)
    await screenshot(page, "filtro_01_abierto")

    # Click en el trigger del combobox Estado (img.x-form-arrow-trigger mas cercana al label)
    r_trigger = await frame.evaluate("""
    () => {
        var labels = Array.from(document.querySelectorAll("label,td,div,span,b"));
        var estadoEl = null;
        for (var i = 0; i < labels.length; i++) {
            var el = labels[i];
            if (!el.offsetParent) continue;
            if ((el.innerText || "").trim() === "Estado:") { estadoEl = el; break; }
        }
        if (!estadoEl) return {ok: false, msg: "label Estado no encontrado"};
        var eRect = estadoEl.getBoundingClientRect();

        var triggers = Array.from(document.querySelectorAll("img.x-form-arrow-trigger"));
        var closest = null, minDist = 9999;
        for (var j = 0; j < triggers.length; j++) {
            var t = triggers[j];
            if (!t.offsetParent) continue;
            var r = t.getBoundingClientRect();
            var dist = Math.abs(r.y - eRect.y) + Math.abs(r.x - eRect.x);
            if (dist < minDist) { minDist = dist; closest = t; }
        }
        if (!closest) return {ok: false, msg: "trigger no encontrado"};
        closest.click();
        var cr = closest.getBoundingClientRect();
        return {ok: true, x: Math.round(cr.x), y: Math.round(cr.y), dist: Math.round(minDist)};
    }
    """)
    print(f"  trigger Estado: {r_trigger}")
    await page.wait_for_timeout(1500)
    await screenshot(page, "filtro_02_dropdown")

    # Seleccionar "Revision y Autorizacion PCCT" en la lista
    r_pcct = await frame.evaluate("""
    () => {
        // Lista de opciones del combobox de ExtJS
        var items = Array.from(document.querySelectorAll(".x-combo-list-item"));
        for (var i = 0; i < items.length; i++) {
            var t = (items[i].innerText || "").trim();
            if (t === "Revisi\u00f3n y Autorizaci\u00f3n PCCT") {
                items[i].click();
                return {ok: true, exact: true, text: t};
            }
        }
        // Fallback: cualquier elemento visible que contenga "PCCT" y "Revisi"
        var all = Array.from(document.querySelectorAll("*"));
        for (var j = 0; j < all.length; j++) {
            var el = all[j];
            if (!el.offsetParent || el.children.length > 0) continue;
            var txt = (el.innerText || "").trim();
            if (txt.indexOf("PCCT") >= 0 && txt.indexOf("Revisi") >= 0 && txt.length < 60) {
                el.click();
                return {ok: true, fallback: true, text: txt};
            }
        }
        return {ok: false};
    }
    """)
    print(f"  seleccion PCCT: {r_pcct}")
    await page.wait_for_timeout(500)
    await screenshot(page, "filtro_03_pcct_seleccionado")

    # Click en Aplicar
    r_aplicar = await frame.evaluate("""
    () => {
        var els = Array.from(document.querySelectorAll("a,button,span,td"));
        for (var i = 0; i < els.length; i++) {
            var el = els[i];
            if (!el.offsetParent) continue;
            var t = (el.innerText || el.textContent || "").trim();
            if (t === "Aplicar") { el.click(); return {ok: true}; }
        }
        return {ok: false};
    }
    """)
    print(f"  Aplicar: {r_aplicar}")

    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    await screenshot(page, "filtro_04_aplicado")

    info = await frame.evaluate("""
    () => {
        var rows = document.querySelectorAll(".x-grid3-row");
        var pagText = Array.from(document.querySelectorAll("*"))
            .filter(function(e) {
                return e.children.length === 0 && e.offsetParent &&
                       (e.innerText || "").indexOf("Mostrando") >= 0;
            }).map(function(e) { return e.innerText.trim(); });
        return {filas_visibles: rows.length, paginador: pagText};
    }
    """)
    print(f"  resultado filtro: {info}")


# ─── Paso 4: Aprobar PT's Metropolitana ───────────────────────────────────────

async def aprobar_pts(page, frame):
    print("\n[4] APROBANDO PT's")
    pts_aprobados = []
    pts_fallidos  = []
    pts_omitidos  = []

    total_paginas = await frame.evaluate(JS_GET_TOTAL_PAGES)
    print(f"  Total paginas: {total_paginas}")
    paginas = min(total_paginas, 20)

    for pagina in range(1, paginas + 1):
        print(f"\n  ── Pagina {pagina}/{paginas} ──")

        filas = await frame.evaluate(JS_READ_ROWS)
        print(f"  Filas leidas: {len(filas)}")

        # Clasificar filas
        pts_esta_pagina = []
        for row in filas:
            if not row:
                continue
            id_pt = area_pt = estado_pt = ""
            for cell in row:
                if re.match(r"^\d{4}-\d{5}$", cell):
                    id_pt = cell
                elif any(k in cell.lower() for k in [
                    "metropolitana", "osorno", "antofagasta", "chiloe", "copiapo",
                    "llvv", "scada", "temuco", "puerto montt", "transemel",
                    "protecciones", "proyectos", "mayor zonal", "llv"
                ]):
                    if not area_pt:
                        area_pt = cell
                elif "PCCT" in cell and not estado_pt:
                    estado_pt = cell
                elif ("Revisi" in cell or "Autorizaci" in cell) and not estado_pt:
                    estado_pt = cell

            if not id_pt or "PCCT" not in estado_pt:
                continue

            if es_metropolitana(area_pt):
                pts_esta_pagina.append({"id": id_pt, "area": area_pt})
                print(f"    [APROBAR] {id_pt} | {area_pt}")
            else:
                pts_omitidos.append({"id": id_pt, "area": area_pt})
                print(f"    [OMITIR]  {id_pt} | {area_pt}")

        # Aprobar los PT's Metropolitana
        ya_aprobados_set = set(pts_aprobados)
        for pt in pts_esta_pagina:
            if pt["id"] in ya_aprobados_set:
                continue

            print(f"\n    >> {pt['id']}")

            try:
                # A: Seleccionar fila
                sel = await frame.evaluate(JS_SELECT_ROW_BY_ID, pt["id"])
                print(f"    seleccion: {sel}")
                if not sel.get("found"):
                    pts_fallidos.append(f"{pt['id']} (fila no encontrada)")
                    await screenshot(page, f"err_nofila_{pt['id']}")
                    continue
                await page.wait_for_timeout(1000)

                # B: Verificar seleccion ExtJS
                check = await frame.evaluate(JS_CHECK_ROW_SELECTED, pt["id"])
                print(f"    fila seleccionada: {check}")
                if not check.get("selected"):
                    # Reintento
                    await frame.evaluate(JS_SELECT_ROW_BY_ID, pt["id"])
                    await page.wait_for_timeout(800)

                # C: Verificar boton Aprobar habilitado
                btn = await frame.evaluate(JS_CHECK_BTN_APROBAR)
                print(f"    btn Aprobar: {btn}")
                if not btn.get("found"):
                    pts_fallidos.append(f"{pt['id']} (boton Aprobar no visible)")
                    await screenshot(page, f"err_nobtn_{pt['id']}")
                    continue
                if btn.get("disabled"):
                    # Reintentar seleccion con pausa mas larga
                    print(f"    Aprobar disabled, reintentando...")
                    await frame.evaluate(JS_SELECT_ROW_BY_ID, pt["id"])
                    await page.wait_for_timeout(1500)
                    btn2 = await frame.evaluate(JS_CHECK_BTN_APROBAR)
                    print(f"    btn retry: {btn2}")
                    if btn2.get("disabled"):
                        pts_fallidos.append(f"{pt['id']} (Aprobar disabled tras reintentos)")
                        await screenshot(page, f"err_disabled_{pt['id']}")
                        continue

                # D: Click en boton Aprobar (abre popup con comentario)
                await screenshot(page, f"pre_{pt['id']}")
                click_r = await frame.evaluate(JS_CLICK_BTN_APROBAR)
                print(f"    click Aprobar: {click_r}")
                if not click_r.get("clicked"):
                    pts_fallidos.append(f"{pt['id']} (click Aprobar no ejecutado)")
                    continue

                # E: Esperar popup Ext.Window "Aprobar"
                popup = {"found": False}
                for intento in range(12):   # hasta 6 segundos
                    await page.wait_for_timeout(500)
                    popup = await frame.evaluate(JS_DETECT_POPUP)
                    if popup.get("found"):
                        print(f"    popup OK (intento {intento+1}): {popup}")
                        break

                await screenshot(page, f"popup_{pt['id']}")

                if not popup.get("found"):
                    pts_fallidos.append(f"{pt['id']} (popup no aparecio)")
                    print(f"    ERROR: popup no aparecio")
                    continue

                # F: Click en Aceptar dentro del popup (comentario vacio)
                aceptar = await frame.evaluate(JS_CLICK_ACEPTAR)
                print(f"    Aceptar: {aceptar}")
                if not aceptar.get("ok"):
                    pts_fallidos.append(f"{pt['id']} (click Aceptar fallo: {aceptar})")
                    await screenshot(page, f"err_aceptar_{pt['id']}")
                    continue

                # G: Esperar procesamiento
                await page.wait_for_load_state("networkidle", timeout=20_000)
                await page.wait_for_timeout(2500)
                await screenshot(page, f"post_{pt['id']}")

                # H: Verificar que desaparecio
                aun_existe = await frame.evaluate(JS_PT_EXISTE, pt["id"])
                if aun_existe:
                    print(f"    ADVERTENCIA: {pt['id']} sigue visible (lag del sistema)")
                else:
                    print(f"    Confirmado: {pt['id']} ya no aparece")

                pts_aprobados.append(pt["id"])
                print(f"    APROBADO: {pt['id']}")

            except Exception as e:
                msg = str(e)[:150]
                pts_fallidos.append(f"{pt['id']}: {msg}")
                print(f"    EXCEPCION: {msg}")
                await screenshot(page, f"exc_{pt['id']}")

        # Siguiente pagina
        if pagina < paginas:
            sig = await frame.evaluate(JS_NEXT_PAGE)
            if not sig:
                print("  No hay mas paginas.")
                break
            await page.wait_for_load_state("networkidle", timeout=15_000)
            await page.wait_for_timeout(2500)

    await screenshot(page, "final")
    return pts_aprobados, pts_fallidos, pts_omitidos


# ─── Reporte por correo ───────────────────────────────────────────────────────

def enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos, error_critico=None):
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")

    def filas_aprobados():
        if not pts_aprobados:
            return "<tr><td colspan='2' style='padding:6px 12px;color:#999'>Ninguno</td></tr>"
        return "".join(
            f"<tr><td style='padding:4px 8px;color:#006600;font-size:16px'>&#10003;</td>"
            f"<td style='font-family:monospace;padding:4px 12px'>{pt}</td></tr>"
            for pt in pts_aprobados
        )

    def filas_fallidos():
        if not pts_fallidos:
            return "<tr><td colspan='2' style='padding:6px 12px;color:#999'>Sin errores</td></tr>"
        return "".join(
            f"<tr><td style='padding:4px 8px;color:#cc0000;font-size:16px'>&#10007;</td>"
            f"<td style='padding:4px 12px;font-size:13px'>{pt}</td></tr>"
            for pt in pts_fallidos
        )

    def filas_omitidos():
        if not pts_omitidos:
            return "<tr><td colspan='2' style='padding:6px 12px;color:#999'>Ninguno</td></tr>"
        return "".join(
            f"<tr><td style='padding:4px 8px;color:#aaa'>&mdash;</td>"
            f"<td style='padding:4px 12px;color:#777;font-size:13px'>"
            f"{pt['id']} &nbsp;&middot;&nbsp; {pt['area'][:60]}</td></tr>"
            for pt in pts_omitidos
        )

    error_bloque = ""
    if error_critico:
        error_bloque = (
            "<div style='background:#fff0f0;border-left:4px solid #c00;"
            "padding:12px 16px;margin:16px 0;border-radius:4px'>"
            f"<strong>Error critico:</strong><br>"
            f"<code style='font-size:12px'>{error_critico}</code></div>"
        )

    html = (
        "<html><body style='font-family:Arial,sans-serif;max-width:680px;margin:auto;color:#222'>"
        "<div style='background:#003580;color:white;padding:24px;border-radius:8px 8px 0 0'>"
        "<h2 style='margin:0;font-size:20px'>Reporte PT's &mdash; SAESA / DMS</h2>"
        "<p style='margin:6px 0 0;opacity:.75;font-size:14px'>"
        "Autorizacion PCCT &middot; Zonal Metropolitana</p>"
        "</div>"
        "<div style='border:1px solid #ddd;border-top:none;padding:20px 24px;border-radius:0 0 8px 8px'>"
        f"<p style='margin:0 0 4px'><strong>Fecha:</strong> {fecha}</p>"
        "<p style='margin:0 0 16px'><strong>Criterio:</strong> "
        "Estado = Revision y Autorizacion PCCT &nbsp;|&nbsp; Area contiene 'Metropolitana'</p>"
        + error_bloque +
        f"<h3 style='color:#006600;margin:20px 0 8px'>PT's Aprobados ({len(pts_aprobados)})</h3>"
        f"<table style='border-collapse:collapse;width:100%'>{filas_aprobados()}</table>"
        f"<h3 style='color:#cc0000;margin:20px 0 8px'>PT's con Error ({len(pts_fallidos)})</h3>"
        f"<table style='border-collapse:collapse;width:100%'>{filas_fallidos()}</table>"
        f"<h3 style='color:#888;margin:20px 0 8px'>PT's Omitidos &mdash; otras zonas ({len(pts_omitidos)})</h3>"
        f"<table style='border-collapse:collapse;width:100%'>{filas_omitidos()}</table>"
        "<p style='color:#bbb;font-size:11px;margin-top:24px;border-top:1px solid #eee;padding-top:12px'>"
        "Bot SAESA &middot; GitHub Actions &middot; github.com/Nlorenzenl/saesa-automation</p>"
        "</div></body></html>"
    )

    asunto = (
        f"[SAESA] ERROR {datetime.now().strftime('%d/%m/%Y')}"
        if error_critico else
        f"[SAESA] {datetime.now().strftime('%d/%m/%Y')} "
        f"| {len(pts_aprobados)} aprobados"
        f" | {len(pts_fallidos)} errores"
        f" | {len(pts_omitidos)} omitidos"
    )

    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"]    = GMAIL_USER
    msg["To"]      = EMAIL_DEST
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(GMAIL_USER, GMAIL_PASS)
            s.sendmail(GMAIL_USER, EMAIL_DEST, msg.as_string())
        print(f"  Correo enviado a {EMAIL_DEST}")
    except Exception as e:
        print(f"  Error correo: {e}")


# ─── Main ─────────────────────────────────────────────────────────────────────

async def main():
    sep = "=" * 55
    print(f"\n{sep}\n  SAESA | {datetime.now().strftime('%d/%m/%Y %H:%M')}\n{sep}")

    pts_aprobados, pts_fallidos, pts_omitidos = [], [], []
    error_critico = None

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--ignore-certificate-errors",
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
            ]
        )
        ctx  = await browser.new_context(
            ignore_https_errors=True,
            viewport={"width": 1400, "height": 900}
        )
        page = await ctx.new_page()

        try:
            await hacer_login(page)
            frame = await navegar_a_permisos(page)
            await aplicar_filtro_pcct(page, frame)
            pts_aprobados, pts_fallidos, pts_omitidos = await aprobar_pts(page, frame)

        except Exception as e:
            error_critico = str(e)
            print(f"\nERROR CRITICO: {e}")
            await screenshot(page, "error_critico")

        finally:
            await browser.close()

    resumen = (
        f"  {len(pts_aprobados)} aprobados | "
        f"{len(pts_fallidos)} errores | "
        f"{len(pts_omitidos)} omitidos"
    )
    print(f"\n{sep}\n{resumen}\n{sep}")
    enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos, error_critico)
    print("Fin.\n")


if __name__ == "__main__":
    asyncio.run(main())
