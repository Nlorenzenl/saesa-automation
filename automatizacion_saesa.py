"""
Automatizacion SAESA - Aprobacion diaria de PT's
Zonal Metropolitana | Estado exacto: Revision y Autorizacion PCCT

Flujo:
login -> Aplicaciones -> DMS -> Planificacion -> Permisos de trabajo ->
Filtro -> Estado = Revision y Autorizacion PCCT -> Aplicar ->
Aprobar solo PTs con Area Metropolitana -> Enviar correo resumen.
"""

import asyncio
import smtplib
import os
import re
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from playwright.async_api import async_playwright

# ─── CONFIGURACION ────────────────────────────────────────────────────────────

SAESA_URL = "https://stx.saesa.cl:8091/backend/sts/login.php?backurl=%2Fbackend%2Fsts%2Fcentrality.php"

SAESA_USER = os.environ["SAESA_USER"]
SAESA_PASS = os.environ["SAESA_PASS"]

GMAIL_USER = os.environ["GMAIL_USER"]
GMAIL_PASS = os.environ["GMAIL_APP_PASS"]
EMAIL_DEST = os.environ["EMAIL_DEST"]

TIMEOUT = 30_000

# Primero probar en True. Cuando esté validado, cambiar a False en GitHub Secret.
DRY_RUN = os.environ.get("DRY_RUN", "true").lower() == "true"

ESTADO_EXACTO = "Revisión y Autorización PCCT"
AREA_KEYWORDS = ["metropolitana"]
MAX_APROBACIONES = int(os.environ.get("MAX_APROBACIONES", "50"))

# ─── JAVASCRIPT HELPERS ───────────────────────────────────────────────────────

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

JS_REFRESH_GRID = """
() => {
    var btn = document.querySelector(".x-tbar-loading");
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
        return {found: true, disabled: disabled, cls: String(el.className).substring(0, 100)};
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
        return {clicked: true, tag: el.tagName, cls: String(el.className).substring(0,60)};
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

        var ta = win.querySelector("textarea");
        if (ta) {
            ta.value = "";
            ta.dispatchEvent(new Event("input", { bubbles: true }));
            ta.dispatchEvent(new Event("change", { bubbles: true }));
        }

        var btns = Array.from(win.querySelectorAll("button"));
        for (var j = 0; j < btns.length; j++) {
            var btxt = (btns[j].innerText || btns[j].textContent || "").trim();
            if (btxt === "Aceptar") {
                btns[j].click();
                return {ok: true, via: "button"};
            }
        }

        var xbtns = Array.from(win.querySelectorAll(".x-btn"));
        for (var k = 0; k < xbtns.length; k++) {
            var xbtxt = (xbtns[k].innerText || xbtns[k].textContent || "").trim();
            if (xbtxt === "Aceptar") {
                xbtns[k].click();
                return {ok: true, via: "x-btn"};
            }
        }

        return {ok: false, win_found: true};
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

# ─── UTILIDADES ───────────────────────────────────────────────────────────────

async def screenshot(page, nombre):
    os.makedirs("capturas", exist_ok=True)
    ts = datetime.now().strftime("%H%M%S")
    safe = re.sub(r"[^a-zA-Z0-9_-]", "_", nombre)
    path = f"capturas/{safe}_{ts}.png"
    await page.screenshot(path=path, full_page=False)
    print(f"    captura: {path}")
    return path


def es_metropolitana(area: str) -> bool:
    return any(k in area.lower() for k in AREA_KEYWORDS)


def normalizar(txt: str) -> str:
    return " ".join((txt or "").strip().split())


def extraer_info_fila(row):
    id_pt = ""
    area_pt = ""
    estado_pt = ""

    for cell in row:
        c = normalizar(cell)

        if re.match(r"^\d{4}-\d{5}$", c):
            id_pt = c
            continue

        if "Revisión y Autorización" in c or "Revision y Autorizacion" in c:
            estado_pt = c
            continue

        posibles_areas = [
            "metropolitana", "osorno", "antofagasta", "chiloe", "chiloé",
            "copiapo", "copiapó", "llvv", "scada", "temuco", "puerto montt",
            "transemel", "protecciones", "proyectos", "mayor zonal",
            "zonal", "mantenimiento"
        ]

        if any(k in c.lower() for k in posibles_areas) and not area_pt:
            area_pt = c

    return id_pt, area_pt, estado_pt


# ─── LOGIN ────────────────────────────────────────────────────────────────────

async def hacer_login(page):
    print("\\n[1] LOGIN")

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
    await page.wait_for_timeout(2500)

    print("  OK: sesión iniciada")


# ─── NAVEGACION ───────────────────────────────────────────────────────────────

async def navegar_a_permisos(page):
    print("\\n[2] NAVEGACION")

    await page.wait_for_selector(
        'a:has-text("Aplicaciones"), span:has-text("Aplicaciones")',
        timeout=TIMEOUT
    )
    await page.click('a:has-text("Aplicaciones"), span:has-text("Aplicaciones")')
    await page.wait_for_timeout(1500)
    print("  -> Aplicaciones")

    await page.wait_for_selector('a:has-text("DMS")', timeout=TIMEOUT)
    await page.click('a:has-text("DMS")')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)

    print("  -> DMS cargado")
    await screenshot(page, "nav_01_dms")

    frame = page

    for f in page.frames:
        try:
            if f.name == "content":
                frame = f
                print("  frame detectado: content")
                break

            el = await f.query_selector('text="Planificación"')
            if el:
                frame = f
                print(f"  frame detectado por selector: {f.name}")
                break
        except Exception:
            pass

    await frame.wait_for_selector('text="Planificación"', timeout=TIMEOUT)
    await frame.click('text="Planificación"')
    await page.wait_for_timeout(1000)

    print("  -> menú Planificación abierto")
    await screenshot(page, "nav_02_planificacion")

    await frame.wait_for_selector('text="Permisos de trabajo"', timeout=TIMEOUT)
    await frame.click('text="Permisos de trabajo"')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3500)

    print("  -> Permisos de trabajo")
    await screenshot(page, "nav_03_permisos")

    return frame


# ─── FILTRO PCCT ──────────────────────────────────────────────────────────────

async def aplicar_filtro_pcct(page, frame):
    print("\\n[3] FILTRO")

    await frame.click('text="Filtro"')
    await page.wait_for_timeout(2000)
    await screenshot(page, "filtro_01_abierto")

    r_trigger = await frame.evaluate("""
    () => {
        var labels = Array.from(document.querySelectorAll("label,td,div,span,b"));
        var estadoEl = null;

        for (var i = 0; i < labels.length; i++) {
            var el = labels[i];
            if (!el.offsetParent) continue;
            if ((el.innerText || "").trim() === "Estado:") {
                estadoEl = el;
                break;
            }
        }

        if (!estadoEl) return {ok: false, msg: "label Estado no encontrado"};

        var eRect = estadoEl.getBoundingClientRect();
        var triggers = Array.from(document.querySelectorAll("img.x-form-arrow-trigger"));

        var closest = null;
        var minDist = 999999;

        for (var j = 0; j < triggers.length; j++) {
            var t = triggers[j];
            if (!t.offsetParent) continue;

            var r = t.getBoundingClientRect();
            var dist = Math.abs(r.y - eRect.y) + Math.abs(r.x - eRect.x);

            if (dist < minDist) {
                minDist = dist;
                closest = t;
            }
        }

        if (!closest) return {ok: false, msg: "trigger Estado no encontrado"};

        closest.click();

        var cr = closest.getBoundingClientRect();
        return {ok: true, x: Math.round(cr.x), y: Math.round(cr.y), dist: Math.round(minDist)};
    }
    """)

    print(f"  trigger Estado: {r_trigger}")
    await page.wait_for_timeout(1500)
    await screenshot(page, "filtro_02_dropdown")

    r_pcct = await frame.evaluate("""
    () => {
        var objetivo = "Revisión y Autorización PCCT";

        var items = Array.from(document.querySelectorAll(".x-combo-list-item"));
        for (var i = 0; i < items.length; i++) {
            var t = (items[i].innerText || "").trim();
            if (t === objetivo) {
                items[i].click();
                return {ok: true, exact: true, text: t};
            }
        }

        var all = Array.from(document.querySelectorAll("*"));
        for (var j = 0; j < all.length; j++) {
            var el = all[j];
            if (!el.offsetParent || el.children.length > 0) continue;

            var txt = (el.innerText || "").trim();

            if (txt === objetivo) {
                el.click();
                return {ok: true, fallback_exact: true, text: txt};
            }
        }

        return {ok: false};
    }
    """)

    print(f"  selección PCCT: {r_pcct}")
    await page.wait_for_timeout(700)
    await screenshot(page, "filtro_03_pcct_seleccionado")

    if not r_pcct.get("ok"):
        raise RuntimeError("No se pudo seleccionar Estado = Revisión y Autorización PCCT")

    r_aplicar = await frame.evaluate("""
    () => {
        var els = Array.from(document.querySelectorAll("a,button,span,td"));
        for (var i = 0; i < els.length; i++) {
            var el = els[i];
            if (!el.offsetParent) continue;

            var t = (el.innerText || el.textContent || "").trim();

            if (t === "Aplicar") {
                el.click();
                return {ok: true};
            }
        }

        return {ok: false};
    }
    """)

    print(f"  Aplicar: {r_aplicar}")

    if not r_aplicar.get("ok"):
        raise RuntimeError("No se pudo presionar Aplicar en el filtro")

    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3500)
    await screenshot(page, "filtro_04_aplicado")

    info = await frame.evaluate("""
    () => {
        var rows = document.querySelectorAll(".x-grid3-row");

        var pagText = Array.from(document.querySelectorAll("*"))
            .filter(function(e) {
                return e.children.length === 0 &&
                       e.offsetParent &&
                       (e.innerText || "").indexOf("Mostrando") >= 0;
            })
            .map(function(e) { return e.innerText.trim(); });

        return {filas_visibles: rows.length, paginador: pagText};
    }
    """)

    print(f"  resultado filtro: {info}")


# ─── APROBAR PTs ──────────────────────────────────────────────────────────────

async def aprobar_pts(page, frame):
    print("\\n[4] APROBANDO PTs")
    print(f"  DRY_RUN: {DRY_RUN}")
    print(f"  MAX_APROBACIONES: {MAX_APROBACIONES}")

    pts_aprobados = []
    pts_fallidos = []
    pts_omitidos = []

    total_paginas = await frame.evaluate(JS_GET_TOTAL_PAGES)
    paginas = min(total_paginas, 20)

    print(f"  Total páginas: {total_paginas}")

    for pagina in range(1, paginas + 1):
        print(f"\\n  ── Página {pagina}/{paginas} ──")

        filas = await frame.evaluate(JS_READ_ROWS)
        print(f"  Filas leídas: {len(filas)}")

        pts_esta_pagina = []

        for row in filas:
            if not row:
                continue

            id_pt, area_pt, estado_pt = extraer_info_fila(row)

            if not id_pt:
                continue

            if estado_pt.strip() != ESTADO_EXACTO:
                pts_omitidos.append({
                    "id": id_pt,
                    "area": area_pt or "Sin área detectada",
                    "motivo": f"Estado no corresponde: {estado_pt or 'Sin estado detectado'}"
                })
                print(f"    [OMITIR ESTADO] {id_pt} | {estado_pt}")
                continue

            if es_metropolitana(area_pt):
                pts_esta_pagina.append({"id": id_pt, "area": area_pt, "estado": estado_pt})
                print(f"    [APROBAR] {id_pt} | {area_pt} | {estado_pt}")
            else:
                pts_omitidos.append({
                    "id": id_pt,
                    "area": area_pt or "Sin área detectada",
                    "motivo": "Área no Metropolitana"
                })
                print(f"    [OMITIR AREA] {id_pt} | {area_pt}")

        for pt in pts_esta_pagina:
            if len(pts_aprobados) >= MAX_APROBACIONES:
                print("    LÍMITE DE SEGURIDAD ALCANZADO")
                pts_fallidos.append(f"Límite de seguridad alcanzado: {MAX_APROBACIONES}")
                return pts_aprobados, pts_fallidos, pts_omitidos

            if pt["id"] in [x.get("id") if isinstance(x, dict) else x for x in pts_aprobados]:
                continue

            print(f"\\n    >> Procesando {pt['id']}")

            try:
                if DRY_RUN:
                    print(f"    [DRY RUN] {pt['id']} NO fue aprobado realmente")
                    pts_aprobados.append({
                        "id": pt["id"],
                        "area": pt["area"],
                        "estado": pt["estado"],
                        "modo": "SIMULADO"
                    })
                    continue

                sel = await frame.evaluate(JS_SELECT_ROW_BY_ID, pt["id"])
                print(f"    selección: {sel}")

                if not sel.get("found"):
                    pts_fallidos.append(f"{pt['id']} - fila no encontrada")
                    await screenshot(page, f"err_nofila_{pt['id']}")
                    continue

                await page.wait_for_timeout(1000)

                check = await frame.evaluate(JS_CHECK_ROW_SELECTED, pt["id"])
                print(f"    fila seleccionada: {check}")

                if not check.get("selected"):
                    await frame.evaluate(JS_SELECT_ROW_BY_ID, pt["id"])
                    await page.wait_for_timeout(1000)

                    check2 = await frame.evaluate(JS_CHECK_ROW_SELECTED, pt["id"])
                    print(f"    fila seleccionada retry: {check2}")

                    if not check2.get("selected"):
                        pts_fallidos.append(f"{pt['id']} - no se pudo seleccionar")
                        await screenshot(page, f"err_select_{pt['id']}")
                        continue

                btn = await frame.evaluate(JS_CHECK_BTN_APROBAR)
                print(f"    botón Aprobar: {btn}")

                if not btn.get("found"):
                    pts_fallidos.append(f"{pt['id']} - botón Aprobar no visible")
                    await screenshot(page, f"err_nobtn_{pt['id']}")
                    continue

                if btn.get("disabled"):
                    print("    Aprobar deshabilitado, reintentando selección")
                    await frame.evaluate(JS_SELECT_ROW_BY_ID, pt["id"])
                    await page.wait_for_timeout(1500)

                    btn2 = await frame.evaluate(JS_CHECK_BTN_APROBAR)
                    print(f"    botón retry: {btn2}")

                    if btn2.get("disabled"):
                        pts_fallidos.append(f"{pt['id']} - botón Aprobar deshabilitado")
                        await screenshot(page, f"err_disabled_{pt['id']}")
                        continue

                await screenshot(page, f"pre_{pt['id']}")

                click_r = await frame.evaluate(JS_CLICK_BTN_APROBAR)
                print(f"    click Aprobar: {click_r}")

                if not click_r.get("clicked"):
                    pts_fallidos.append(f"{pt['id']} - click Aprobar no ejecutado")
                    await screenshot(page, f"err_click_aprobar_{pt['id']}")
                    continue

                popup = {"found": False}

                for intento in range(14):
                    await page.wait_for_timeout(500)
                    popup = await frame.evaluate(JS_DETECT_POPUP)

                    if popup.get("found"):
                        print(f"    popup OK intento {intento + 1}: {popup}")
                        break

                await screenshot(page, f"popup_{pt['id']}")

                if not popup.get("found"):
                    pts_fallidos.append(f"{pt['id']} - popup Aprobar no apareció")
                    continue

                aceptar = await frame.evaluate(JS_CLICK_ACEPTAR)
                print(f"    Aceptar: {aceptar}")

                if not aceptar.get("ok"):
                    pts_fallidos.append(f"{pt['id']} - click Aceptar falló: {aceptar}")
                    await screenshot(page, f"err_aceptar_{pt['id']}")
                    continue

                await page.wait_for_load_state("networkidle", timeout=20_000)
                await page.wait_for_timeout(2500)

                refresh = await frame.evaluate(JS_REFRESH_GRID)
                print(f"    refresh grilla: {refresh}")

                await page.wait_for_timeout(2500)
                await screenshot(page, f"post_{pt['id']}")

                aun_existe = await frame.evaluate(JS_PT_EXISTE, pt["id"])

                if aun_existe:
                    print(f"    ADVERTENCIA: {pt['id']} sigue visible")
                else:
                    print(f"    Confirmado: {pt['id']} ya no aparece")

                pts_aprobados.append({
                    "id": pt["id"],
                    "area": pt["area"],
                    "estado": pt["estado"],
                    "modo": "REAL"
                })

                print(f"    APROBADO: {pt['id']}")

            except Exception as e:
                msg = str(e)[:250]
                pts_fallidos.append(f"{pt['id']} - {msg}")
                print(f"    EXCEPCIÓN: {msg}")
                await screenshot(page, f"exc_{pt['id']}")

        if pagina < paginas:
            sig = await frame.evaluate(JS_NEXT_PAGE)

            if not sig:
                print("  No hay más páginas")
                break

            await page.wait_for_load_state("networkidle", timeout=15_000)
            await page.wait_for_timeout(3000)

    await screenshot(page, "final")
    return pts_aprobados, pts_fallidos, pts_omitidos


# ─── CORREO ───────────────────────────────────────────────────────────────────

def enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos, error_critico=None):
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")

    def filas_aprobados():
        if not pts_aprobados:
            return "<tr><td style='padding:6px 12px;color:#999'>Ninguno</td></tr>"

        html = ""

        for pt in pts_aprobados:
            if isinstance(pt, dict):
                modo = pt.get("modo", "")
                etiqueta = "SIMULADO" if modo == "SIMULADO" else "REAL"
                html += (
                    "<tr>"
                    "<td style='padding:4px 8px;color:#006600;font-size:16px'>&#10003;</td>"
                    f"<td style='font-family:monospace;padding:4px 12px'>{pt.get('id')}</td>"
                    f"<td style='padding:4px 12px'>{pt.get('area','')}</td>"
                    f"<td style='padding:4px 12px'><strong>{etiqueta}</strong></td>"
                    "</tr>"
                )
            else:
                html += (
                    "<tr>"
                    "<td style='padding:4px 8px;color:#006600;font-size:16px'>&#10003;</td>"
                    f"<td style='font-family:monospace;padding:4px 12px'>{pt}</td>"
                    "<td></td><td></td>"
                    "</tr>"
                )

        return html

    def filas_fallidos():
        if not pts_fallidos:
            return "<tr><td style='padding:6px 12px;color:#999'>Sin errores</td></tr>"

        return "".join(
            "<tr>"
            "<td style='padding:4px 8px;color:#cc0000;font-size:16px'>&#10007;</td>"
            f"<td style='padding:4px 12px;font-size:13px'>{pt}</td>"
            "</tr>"
            for pt in pts_fallidos
        )

    def filas_omitidos():
        if not pts_omitidos:
            return "<tr><td style='padding:6px 12px;color:#999'>Ninguno</td></tr>"

        html = ""

        for pt in pts_omitidos:
            html += (
                "<tr>"
                "<td style='padding:4px 8px;color:#aaa'>&mdash;</td>"
                f"<td style='font-family:monospace;padding:4px 12px;color:#777'>{pt.get('id','')}</td>"
                f"<td style='padding:4px 12px;color:#777'>{pt.get('area','')[:80]}</td>"
                f"<td style='padding:4px 12px;color:#777'>{pt.get('motivo','')}</td>"
                "</tr>"
            )

        return html

    error_bloque = ""

    if error_critico:
        error_bloque = (
            "<div style='background:#fff0f0;border-left:4px solid #c00;"
            "padding:12px 16px;margin:16px 0;border-radius:4px'>"
            "<strong>Error crítico:</strong><br>"
            f"<code style='font-size:12px'>{error_critico}</code>"
            "</div>"
        )

    modo_txt = "DRY RUN / SIMULACIÓN" if DRY_RUN else "REAL"

    html = (
        "<html><body style='font-family:Arial,sans-serif;max-width:760px;margin:auto;color:#222'>"
        "<div style='background:#003580;color:white;padding:24px;border-radius:8px 8px 0 0'>"
        "<h2 style='margin:0;font-size:20px'>Reporte PTs — SAESA / DMS</h2>"
        "<p style='margin:6px 0 0;opacity:.8;font-size:14px'>"
        "Aprobación PCCT · Zonal Metropolitana</p>"
        "</div>"

        "<div style='border:1px solid #ddd;border-top:none;padding:20px 24px;border-radius:0 0 8px 8px'>"
        f"<p><strong>Fecha:</strong> {fecha}</p>"
        f"<p><strong>Modo:</strong> {modo_txt}</p>"
        "<p><strong>Criterio:</strong> Estado exacto = Revisión y Autorización PCCT "
        " | Área contiene Metropolitana</p>"
        + error_bloque +

        f"<h3 style='color:#006600;margin:20px 0 8px'>PTs Aprobados ({len(pts_aprobados)})</h3>"
        "<table style='border-collapse:collapse;width:100%'>"
        "<tr style='background:#f6f6f6'><th></th><th>PT</th><th>Área</th><th>Modo</th></tr>"
        f"{filas_aprobados()}</table>"

        f"<h3 style='color:#cc0000;margin:20px 0 8px'>PTs con Error ({len(pts_fallidos)})</h3>"
        f"<table style='border-collapse:collapse;width:100%'>{filas_fallidos()}</table>"

        f"<h3 style='color:#888;margin:20px 0 8px'>PTs Omitidos ({len(pts_omitidos)})</h3>"
        "<table style='border-collapse:collapse;width:100%'>"
        "<tr style='background:#f6f6f6'><th></th><th>PT</th><th>Área</th><th>Motivo</th></tr>"
        f"{filas_omitidos()}</table>"

        "<p style='color:#bbb;font-size:11px;margin-top:24px;border-top:1px solid #eee;padding-top:12px'>"
        "Bot SAESA · GitHub Actions · github.com/Nlorenzenl/saesa-automation</p>"
        "</div></body></html>"
    )

    asunto = (
        f"[SAESA] ERROR {datetime.now().strftime('%d/%m/%Y')}"
        if error_critico else
        f"[SAESA] {datetime.now().strftime('%d/%m/%Y')} | "
        f"{len(pts_aprobados)} aprobados | "
        f"{len(pts_fallidos)} errores | "
        f"{len(pts_omitidos)} omitidos | {modo_txt}"
    )

    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"] = GMAIL_USER
    msg["To"] = EMAIL_DEST
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(GMAIL_USER, GMAIL_PASS)
            s.sendmail(GMAIL_USER, EMAIL_DEST, msg.as_string())

        print(f"  Correo enviado a {EMAIL_DEST}")

    except Exception as e:
        print(f"  Error enviando correo: {e}")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

async def main():
    sep = "=" * 65

    print(f"\\n{sep}")
    print(f"  SAESA AUTOMATION | {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"  DRY_RUN: {DRY_RUN}")
    print(f"{sep}")

    pts_aprobados = []
    pts_fallidos = []
    pts_omitidos = []
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

        ctx = await browser.new_context(
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
            print(f"\\nERROR CRÍTICO: {e}")

            try:
                await screenshot(page, "error_critico")
            except Exception:
                pass

        finally:
            await browser.close()

    print(f"\\n{sep}")
    print(f"  {len(pts_aprobados)} aprobados | {len(pts_fallidos)} errores | {len(pts_omitidos)} omitidos")
    print(f"{sep}")

    enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos, error_critico)

    print("Fin.\\n")


if __name__ == "__main__":
    asyncio.run(main())
