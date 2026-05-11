"""
Automatizacion SAESA - Autorizacion diaria de PT's
Zonal Metropolitana | Estado: Revision y Autorizacion PCCT

FIXES v3:
- Filtro: trigger buscaba el de 'Areas' en vez de 'Estado' (distY < 15px fix)
- Filtro: verifica que el panel se cierre antes de operar el grid
- Filtro: estrategia alternativa escribiendo en el input si el dropdown falla
- Seleccion fila: usa page.mouse.click() con coordenadas reales (nativo Playwright)
- Popup Aprobar: busca .x-window con titulo "Aprobar" y clickea Aceptar dentro
"""

import asyncio
import smtplib
import os
import re
from datetime import datetime
from zoneinfo import ZoneInfo
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

# ─── Configuracion ────────────────────────────────────────────────────────────
SAESA_URL  = "https://stx.saesa.cl:8091/backend/sts/login.php?backurl=%2Fbackend%2Fsts%2Fcentrality.php"
SAESA_USER = os.environ["SAESA_USER"]
SAESA_PASS = os.environ["SAESA_PASS"]
GMAIL_USER = os.environ["GMAIL_USER"]
GMAIL_PASS = os.environ["GMAIL_APP_PASS"]
EMAIL_DEST  = os.environ["EMAIL_DEST"]
# Destinatarios adicionales (CC)
EMAIL_CC    = ["alexis.aedo@saesa.cl", "jorge.canete@saesa.cl"]
TIMEOUT     = 30_000
AREA_KEYWORDS = ["metropolitana"]
TZ_CHILE    = ZoneInfo("America/Santiago")

# ─── JS helpers ───────────────────────────────────────────────────────────────

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

# Devuelve las coordenadas del centro de la fila para hacer click nativo
JS_GET_ROW_COORDS = """
(ptId) => {
    var rows = Array.from(document.querySelectorAll(".x-grid3-row"));
    for (var i = 0; i < rows.length; i++) {
        var cells = Array.from(rows[i].querySelectorAll(".x-grid3-cell-inner"));
        for (var j = 0; j < cells.length; j++) {
            if (cells[j].innerText.trim() === ptId) {
                var r = rows[i].getBoundingClientRect();
                rows[i].scrollIntoView({block: "center"});
                var r2 = rows[i].getBoundingClientRect();
                return {
                    found: true,
                    x: Math.round(r2.left + r2.width / 2),
                    y: Math.round(r2.top + r2.height / 2),
                    rowIndex: i
                };
            }
        }
    }
    return {found: false};
}
"""

JS_CHECK_ROW_SELECTED = """
(ptId) => {
    var sel = document.querySelector(".x-grid3-row-selected");
    if (!sel) return {selected: false, reason: "no selected row"};
    var cells = Array.from(sel.querySelectorAll(".x-grid3-cell-inner"));
    var texts = cells.map(function(c) { return c.innerText.trim(); });
    return {selected: texts.indexOf(ptId) >= 0, firstCell: texts[0] || ""};
}
"""

JS_CHECK_BTN_APROBAR = """
() => {
    var els = Array.from(document.querySelectorAll("a,button,td,span"));
    for (var i = 0; i < els.length; i++) {
        var el = els[i];
        if (!el.offsetParent) continue;
        var txt = (el.innerText || el.textContent || "").trim();
        if (txt !== "Aprobar") continue;
        var disabled = el.classList.contains("x-item-disabled") ||
                       !!(el.closest && el.closest(".x-item-disabled"));
        return {found: true, disabled: disabled, cls: el.className.substring(0,80)};
    }
    return {found: false};
}
"""

JS_CLICK_BTN_APROBAR = """
() => {
    var els = Array.from(document.querySelectorAll("a,button,td,span"));
    for (var i = 0; i < els.length; i++) {
        var el = els[i];
        if (!el.offsetParent) continue;
        var txt = (el.innerText || el.textContent || "").trim();
        if (txt !== "Aprobar") continue;
        if (el.classList.contains("x-item-disabled")) continue;
        if (el.closest && el.closest(".x-item-disabled")) continue;
        el.click();
        return {clicked: true, tag: el.tagName};
    }
    return {clicked: false};
}
"""

JS_DETECT_POPUP_APROBAR = """
() => {
    var headers = Array.from(document.querySelectorAll(
        ".x-window-header-text,.x-panel-header-text"
    ));
    for (var i = 0; i < headers.length; i++) {
        var h = headers[i];
        if (!h.offsetParent) continue;
        if ((h.innerText || h.textContent || "").trim() !== "Aprobar") continue;
        var win = h.closest(".x-window");
        if (!win || !win.offsetParent) continue;
        var r = win.getBoundingClientRect();
        return {found: true, x: Math.round(r.x), y: Math.round(r.y),
                w: Math.round(r.width), h: Math.round(r.height)};
    }
    return {found: false};
}
"""

JS_CLICK_ACEPTAR_EN_POPUP = """
() => {
    var headers = Array.from(document.querySelectorAll(
        ".x-window-header-text,.x-panel-header-text"
    ));
    for (var i = 0; i < headers.length; i++) {
        var h = headers[i];
        if (!h.offsetParent) continue;
        if ((h.innerText || h.textContent || "").trim() !== "Aprobar") continue;
        var win = h.closest(".x-window");
        if (!win || !win.offsetParent) continue;
        // Limpiar comentario (opcional)
        var ta = win.querySelector("textarea");
        if (ta) ta.value = "";
        // Click en Aceptar
        var btns = Array.from(win.querySelectorAll("button,.x-btn"));
        for (var j = 0; j < btns.length; j++) {
            var t = (btns[j].innerText || btns[j].textContent || "").trim();
            if (t === "Aceptar") { btns[j].click(); return {ok: true}; }
        }
        return {ok: false, win_found: true, btns: btns.length};
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

JS_FILTRO_VISIBLE = """
() => {
    var wins = Array.from(document.querySelectorAll(".x-window,.x-panel"));
    for (var w = 0; w < wins.length; w++) {
        if (!wins[w].offsetParent) continue;
        var wt = wins[w].innerText || "";
        if (wt.indexOf("En bandeja de trabajo") >= 0 && wt.indexOf("Aplicar") >= 0)
            return true;
    }
    return false;
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


# ─── Paso 2: Navegacion ───────────────────────────────────────────────────────

async def navegar_a_permisos(page):
    print("\n[2] NAVEGACION")
    await page.wait_for_selector(
        'a:has-text("Aplicaciones"), span:has-text("Aplicaciones")', timeout=TIMEOUT)
    await page.click('a:has-text("Aplicaciones"), span:has-text("Aplicaciones")')
    await page.wait_for_timeout(1500)
    print("  -> Aplicaciones")

    await page.wait_for_selector('a:has-text("DMS")', timeout=TIMEOUT)
    await page.click('a:has-text("DMS")')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2500)
    print("  -> DMS cargado")

    # Detectar frame donde vive el contenido DMS
    frame = page
    for f in page.frames:
        if f.name == "content":
            frame = f
            print(f"  frame: '{f.name}'")
            break
        try:
            el = await f.query_selector('text="Planificación"')
            if el:
                frame = f
                print(f"  frame detectado: '{f.name}'")
                break
        except Exception:
            pass

    await frame.wait_for_selector('text="Planificación"', timeout=TIMEOUT)
    await frame.click('text="Planificación"')
    await page.wait_for_timeout(1000)
    print("  -> menu Planificacion")

    await frame.wait_for_selector('text="Permisos de trabajo"', timeout=TIMEOUT)
    await frame.click('text="Permisos de trabajo"')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    print("  -> Permisos de trabajo OK")
    await screenshot(page, "nav_permisos")
    return frame


# ─── Paso 3: Filtro Estado = PCCT ────────────────────────────────────────────

async def aplicar_filtro_pcct(page, frame):
    """
    FIX PRINCIPAL: el trigger de 'Areas' estaba siendo clickeado en vez del de 'Estado'.
    Solucion: buscar el trigger cuyo centro Y difiera menos de 15px del centro Y del label 'Estado:'
    y que ademas este a la derecha del label (mismo renglon).
    """
    print("\n[3] FILTRO")

    await frame.click('text="Filtro"')
    await page.wait_for_timeout(2000)
    await screenshot(page, "filtro_01_abierto")
    # Cerrar cualquier combo abierto antes de operar
    await page.keyboard.press("Escape")
    await page.wait_for_timeout(300)

    # ── Click en el trigger del combobox Estado (no Areas) ────────────────────
    r_trigger = await frame.evaluate("""
    () => {
        // Encontrar el panel Filtros
        var filtroWin = null;
        var wins = Array.from(document.querySelectorAll(".x-window,.x-panel"));
        for (var w = 0; w < wins.length; w++) {
            var wt = wins[w].innerText || "";
            if (wt.indexOf("En bandeja de trabajo") >= 0 && wt.indexOf("Estado:") >= 0) {
                filtroWin = wins[w]; break;
            }
        }
        if (!filtroWin) return {ok: false, msg: "panel Filtros no encontrado"};

        // Encontrar label "Estado:" dentro del panel
        var labels = Array.from(filtroWin.querySelectorAll("*"));
        var estadoEl = null;
        for (var i = 0; i < labels.length; i++) {
            var el = labels[i];
            if (!el.offsetParent || el.children.length > 0) continue;
            if ((el.innerText || "").trim() === "Estado:") { estadoEl = el; break; }
        }
        if (!estadoEl) return {ok: false, msg: "label Estado: no encontrado"};

        var eRect = estadoEl.getBoundingClientRect();
        var eCenterY = eRect.top + eRect.height / 2;

        // Buscar trigger en la misma linea horizontal (distY < 15px) y a la derecha
        var triggers = Array.from(filtroWin.querySelectorAll("img.x-form-arrow-trigger"));
        var debug = [];
        var best = null, bestDistY = 9999;
        for (var j = 0; j < triggers.length; j++) {
            var t = triggers[j];
            if (!t.offsetParent) continue;
            var tr = t.getBoundingClientRect();
            var tCenterY = tr.top + tr.height / 2;
            var distY = Math.abs(tCenterY - eCenterY);
            debug.push({x: Math.round(tr.x), y: Math.round(tr.y), distY: Math.round(distY)});
            // Mismo renglon (distY < 15) y a la derecha del label
            if (distY < 15 && tr.left > eRect.left && distY < bestDistY) {
                bestDistY = distY;
                best = t;
            }
        }

        if (!best) {
            // Fallback: menor distY sin restriccion de X
            for (var k = 0; k < triggers.length; k++) {
                var t2 = triggers[k];
                if (!t2.offsetParent) continue;
                var tr2 = t2.getBoundingClientRect();
                var d2 = Math.abs((tr2.top + tr2.height/2) - eCenterY);
                if (d2 < bestDistY) { bestDistY = d2; best = t2; }
            }
            if (!best) return {ok: false, msg: "ningún trigger encontrado", debug: debug};
        }

        var br = best.getBoundingClientRect();
        best.click();
        return {
            ok: true,
            x: Math.round(br.x), y: Math.round(br.y),
            labelCenterY: Math.round(eCenterY),
            distY: Math.round(bestDistY),
            debug: debug
        };
    }
    """)
    print(f"  trigger Estado: {r_trigger}")
    await page.wait_for_timeout(1500)
    await screenshot(page, "filtro_02_dropdown")

    # ── Seleccionar PCCT en el combo ──────────────────────────────────────────
    r_pcct = await frame.evaluate("""
    () => {
        // Lista desplegable ExtJS
        var items = Array.from(document.querySelectorAll(".x-combo-list-item"));
        for (var i = 0; i < items.length; i++) {
            var t = (items[i].innerText || "").trim();
            if (t === "Revisi\u00f3n y Autorizaci\u00f3n PCCT") {
                items[i].click();
                return {ok: true, method: "exact", text: t};
            }
        }
        for (var j = 0; j < items.length; j++) {
            var t2 = (items[j].innerText || "").trim();
            if (t2.indexOf("PCCT") >= 0 && t2.indexOf("Revisi") >= 0) {
                items[j].click();
                return {ok: true, method: "partial", text: t2};
            }
        }
        var disponibles = items.map(function(e) { return (e.innerText||"").trim(); });
        return {ok: false, disponibles: disponibles.slice(0, 30)};
    }
    """)
    print(f"  seleccion PCCT: {r_pcct}")
    await page.wait_for_timeout(800)
    await screenshot(page, "filtro_03_pcct")

    # Si no se pudo seleccionar con el combo, intentar escribir en el input
    if not r_pcct.get("ok"):
        print("  combo fallo, intentando escribir en input Estado...")
        await frame.evaluate("""
        () => {
            var filtroWin = null;
            var wins = Array.from(document.querySelectorAll(".x-window,.x-panel"));
            for (var w=0;w<wins.length;w++) {
                if ((wins[w].innerText||"").indexOf("Estado:")>=0 &&
                    (wins[w].innerText||"").indexOf("Aplicar")>=0) {
                    filtroWin=wins[w]; break;
                }
            }
            if (!filtroWin) return;
            // Encontrar label Estado y el input mas cercano
            var labels=Array.from(filtroWin.querySelectorAll("*"));
            var estadoEl=null;
            for (var i=0;i<labels.length;i++) {
                var el=labels[i];
                if (!el.offsetParent||el.children.length>0) continue;
                if ((el.innerText||"").trim()==="Estado:") {estadoEl=el;break;}
            }
            if (!estadoEl) return;
            var eY=estadoEl.getBoundingClientRect().top;
            var inputs=Array.from(filtroWin.querySelectorAll("input[type=text],input:not([type])"));
            var best=null,bestD=999;
            for (var j=0;j<inputs.length;j++) {
                var inp=inputs[j];
                if (!inp.offsetParent) continue;
                var d=Math.abs(inp.getBoundingClientRect().top-eY);
                if (d<bestD){bestD=d;best=inp;}
            }
            if (best) { best.focus(); best.value=""; }
        }
        """)
        await page.wait_for_timeout(300)
        await page.keyboard.type("PCCT", delay=80)
        await page.wait_for_timeout(1000)
        await screenshot(page, "filtro_alt_input")
        r_pcct2 = await frame.evaluate("""
        () => {
            var items=Array.from(document.querySelectorAll(".x-combo-list-item"));
            for (var i=0;i<items.length;i++) {
                var t=(items[i].innerText||"").trim();
                if (t.indexOf("PCCT")>=0&&t.indexOf("Revisi")>=0) {
                    items[i].click(); return {ok:true,text:t};
                }
            }
            return {ok:false};
        }
        """)
        print(f"  seleccion alt: {r_pcct2}")

    # ── Click en Aplicar ──────────────────────────────────────────────────────
    r_aplicar = await frame.evaluate("""
    () => {
        var filtroWin = null;
        var wins = Array.from(document.querySelectorAll(".x-window,.x-panel"));
        for (var w=0;w<wins.length;w++) {
            var wt=wins[w].innerText||"";
            if (wt.indexOf("Aplicar")>=0&&wt.indexOf("Limpiar")>=0&&wins[w].offsetParent) {
                filtroWin=wins[w]; break;
            }
        }
        if (!filtroWin) return {ok:false,msg:"panel no encontrado"};
        var btns=Array.from(filtroWin.querySelectorAll("a,button,span,td"));
        for (var i=0;i<btns.length;i++) {
            var b=btns[i];
            if (!b.offsetParent) continue;
            var t=(b.innerText||b.textContent||"").trim();
            if (t==="Aplicar") { b.click(); return {ok:true}; }
        }
        return {ok:false};
    }
    """)
    print(f"  Aplicar: {r_aplicar}")

    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    await screenshot(page, "filtro_04_aplicado")

    # ── Verificar que el panel Filtros se cerro ───────────────────────────────
    filtro_abierto = await frame.evaluate(JS_FILTRO_VISIBLE)
    if filtro_abierto:
        print("  ADVERTENCIA: panel Filtros sigue visible, cerrando con Escape...")
        await page.keyboard.press("Escape")
        await page.wait_for_timeout(1000)
        await screenshot(page, "filtro_05_cerrado")

    # ── Contar resultados ─────────────────────────────────────────────────────
    info = await frame.evaluate("""
    () => {
        var rows = document.querySelectorAll(".x-grid3-row");
        var pag = Array.from(document.querySelectorAll("*"))
            .filter(function(e) {
                return e.children.length===0 && e.offsetParent &&
                       (e.innerText||"").indexOf("Mostrando")>=0;
            }).map(function(e) { return e.innerText.trim(); });
        return {filas: rows.length, paginador: pag};
    }
    """)
    print(f"  resultado filtro: {info}")
    pag_str = str(info.get("paginador", []))
    if "343" in pag_str:
        print("  ADVERTENCIA: el filtro no se aplico (343 resultados = sin filtro)")


# ─── Paso 4: Aprobar PT's ────────────────────────────────────────────────────

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
        print(f"  Filas: {len(filas)}")

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
                    "llvv", "llv", "scada", "temuco", "puerto montt", "transemel",
                    "protecciones", "proyectos", "mayor zonal", "area llv"
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

        ya_aprobados_set = set(pts_aprobados)
        for pt in pts_esta_pagina:
            if pt["id"] in ya_aprobados_set:
                continue

            print(f"\n    >> {pt['id']}")
            try:
                # A: Obtener coordenadas reales de la fila y hacer click nativo
                coords = await frame.evaluate(JS_GET_ROW_COORDS, pt["id"])
                print(f"    coords fila: {coords}")
                if not coords.get("found"):
                    pts_fallidos.append(f"{pt['id']} (fila no encontrada)")
                    await screenshot(page, f"err_nofila_{pt['id']}")
                    continue

                # Click nativo con page.mouse usando coordenadas absolutas de la pagina
                await page.mouse.click(coords["x"], coords["y"])
                await page.wait_for_timeout(1000)

                # B: Verificar seleccion en ExtJS
                check = await frame.evaluate(JS_CHECK_ROW_SELECTED, pt["id"])
                print(f"    seleccionada: {check}")

                if not check.get("selected"):
                    # Segundo intento
                    await page.mouse.click(coords["x"], coords["y"])
                    await page.wait_for_timeout(800)
                    check2 = await frame.evaluate(JS_CHECK_ROW_SELECTED, pt["id"])
                    print(f"    seleccion retry: {check2}")

                # C: Verificar boton Aprobar
                btn = await frame.evaluate(JS_CHECK_BTN_APROBAR)
                print(f"    btn Aprobar: {btn}")
                if not btn.get("found"):
                    pts_fallidos.append(f"{pt['id']} (boton Aprobar no visible)")
                    await screenshot(page, f"err_nobtn_{pt['id']}")
                    continue
                if btn.get("disabled"):
                    print(f"    disabled, reintentando click en fila...")
                    await page.mouse.click(coords["x"], coords["y"])
                    await page.wait_for_timeout(1200)
                    btn2 = await frame.evaluate(JS_CHECK_BTN_APROBAR)
                    print(f"    btn retry: {btn2}")
                    if btn2.get("disabled"):
                        pts_fallidos.append(f"{pt['id']} (Aprobar disabled)")
                        await screenshot(page, f"err_disabled_{pt['id']}")
                        continue

                # D: Click en boton Aprobar
                await screenshot(page, f"pre_{pt['id']}")
                click_r = await frame.evaluate(JS_CLICK_BTN_APROBAR)
                print(f"    click Aprobar: {click_r}")
                if not click_r.get("clicked"):
                    pts_fallidos.append(f"{pt['id']} (click Aprobar fallo)")
                    continue

                # E: Esperar popup Ext.Window "Aprobar"
                popup = {"found": False}
                for intento in range(14):   # hasta 7 segundos
                    await page.wait_for_timeout(500)
                    popup = await frame.evaluate(JS_DETECT_POPUP_APROBAR)
                    if popup.get("found"):
                        print(f"    popup OK (intento {intento+1}): {popup}")
                        break

                await screenshot(page, f"popup_{pt['id']}")

                if not popup.get("found"):
                    pts_fallidos.append(f"{pt['id']} (popup no aparecio)")
                    print(f"    ERROR: popup no aparecio")
                    continue

                # F: Click en Aceptar
                aceptar = await frame.evaluate(JS_CLICK_ACEPTAR_EN_POPUP)
                print(f"    Aceptar: {aceptar}")
                if not aceptar.get("ok"):
                    pts_fallidos.append(f"{pt['id']} (Aceptar fallo: {aceptar})")
                    await screenshot(page, f"err_aceptar_{pt['id']}")
                    continue

                # G: Esperar procesamiento
                await page.wait_for_load_state("networkidle", timeout=20_000)
                await page.wait_for_timeout(2500)
                await screenshot(page, f"post_{pt['id']}")

                # H: Confirmar desaparicion
                aun_existe = await frame.evaluate(JS_PT_EXISTE, pt["id"])
                if aun_existe:
                    print(f"    ADVERTENCIA: {pt['id']} sigue visible (lag)")
                else:
                    print(f"    Confirmado: {pt['id']} desaparecio")

                pts_aprobados.append(pt["id"])
                print(f"    APROBADO: {pt['id']}")

            except Exception as e:
                msg = str(e)[:150]
                pts_fallidos.append(f"{pt['id']}: {msg}")
                print(f"    EXCEPCION: {msg}")
                await screenshot(page, f"exc_{pt['id']}")

        if pagina < paginas:
            sig = await frame.evaluate(JS_NEXT_PAGE)
            if not sig:
                print("  No hay mas paginas.")
                break
            await page.wait_for_load_state("networkidle", timeout=15_000)
            await page.wait_for_timeout(2500)

    await screenshot(page, "final")
    return pts_aprobados, pts_fallidos, pts_omitidos


# ─── Reporte correo ───────────────────────────────────────────────────────────

def enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos, error_critico=None):
    # Hora real en Chile
    ahora_chile = datetime.now(TZ_CHILE)
    fecha       = ahora_chile.strftime("%d/%m/%Y")
    hora        = ahora_chile.strftime("%H:%M")
    hora_ampm   = ahora_chile.strftime("%I:%M %p").lower()  # ej: 08:01 am

    def filas(items, tipo):
        if not items:
            return "<tr><td colspan='2' style='padding:6px 12px;color:#999'>Ninguno</td></tr>"
        if tipo == "ok":
            return "".join(
                f"<tr><td style='padding:4px 8px;color:#006600'>&#10003;</td>"
                f"<td style='font-family:monospace;padding:4px 12px'>{pt}</td></tr>"
                for pt in items)
        if tipo == "err":
            return "".join(
                f"<tr><td style='padding:4px 8px;color:#cc0000'>&#10007;</td>"
                f"<td style='padding:4px 12px;font-size:13px'>{pt}</td></tr>"
                for pt in items)
        if tipo == "omit":
            return "".join(
                f"<tr><td style='padding:4px 8px;color:#aaa'>&mdash;</td>"
                f"<td style='padding:4px 12px;color:#777;font-size:13px'>"
                f"{pt['id']} &middot; {pt['area'][:60]}</td></tr>"
                for pt in items)
        return ""

    error_bloque = ""
    if error_critico:
        error_bloque = (
            "<div style='background:#fff0f0;border-left:4px solid #c00;"
            "padding:12px 16px;margin:16px 0;border-radius:4px'>"
            f"<strong>Error critico:</strong><br>"
            f"<code style='font-size:12px'>{error_critico}</code></div>")

    html = (
        "<html><body style='font-family:Arial,sans-serif;max-width:680px;margin:auto;color:#222'>"
        "<div style='background:#003580;color:white;padding:24px;border-radius:8px 8px 0 0'>"
        "<h2 style='margin:0'>Reporte PT&rsquo;s &mdash; Centrality / DMS</h2>"
        "<p style='margin:6px 0 0;opacity:.75;font-size:14px'>Autorizacion PCCT &middot; Zonal Metropolitana</p>"
        "</div>"
        "<div style='border:1px solid #ddd;border-top:none;padding:20px 24px;border-radius:0 0 8px 8px'>"
        f"<p><strong>Fecha:</strong> {fecha} {hora}</p>"
        "<p><strong>Criterio:</strong> Estado = PCCT &nbsp;|&nbsp; Area contiene &lsquo;Metropolitana&rsquo;</p>"
        + error_bloque
        + f"<h3 style='color:#006600;margin:20px 0 8px'>Aprobados ({len(pts_aprobados)})</h3>"
        f"<table style='border-collapse:collapse;width:100%'>{filas(pts_aprobados,'ok')}</table>"
        f"<h3 style='color:#cc0000;margin:20px 0 8px'>Con Error ({len(pts_fallidos)})</h3>"
        f"<table style='border-collapse:collapse;width:100%'>{filas(pts_fallidos,'err')}</table>"
        f"<h3 style='color:#888;margin:20px 0 8px'>Omitidos &mdash; otras zonas ({len(pts_omitidos)})</h3>"
        f"<table style='border-collapse:collapse;width:100%'>{filas(pts_omitidos,'omit')}</table>"
        "</div></body></html>"
    )

    # Asunto con hora Chile
    if error_critico:
        asunto = f"[Reporte PTS Centrality] ERROR {fecha} {hora_ampm}"
    else:
        asunto = (
            f"[Reporte PTS Centrality] {fecha} {hora_ampm} "
            f"| {len(pts_aprobados)} aprobados"
            f" | {len(pts_fallidos)} errores"
            f" | {len(pts_omitidos)} omitidos"
        )

    # Destinatarios: principal + CC
    todos = [EMAIL_DEST] + EMAIL_CC
    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"]    = GMAIL_USER
    msg["To"]      = EMAIL_DEST
    msg["Cc"]      = ", ".join(EMAIL_CC)
    msg.attach(MIMEText(html, "html"))
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(GMAIL_USER, GMAIL_PASS)
            s.sendmail(GMAIL_USER, todos, msg.as_string())
        print(f"  Correo enviado a {todos}")
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
            args=["--ignore-certificate-errors","--no-sandbox",
                  "--disable-setuid-sandbox","--disable-dev-shm-usage"])
        ctx  = await browser.new_context(
            ignore_https_errors=True, viewport={"width": 1400, "height": 900})
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

    print(f"\n{sep}\n  {len(pts_aprobados)} aprobados | {len(pts_fallidos)} errores | {len(pts_omitidos)} omitidos\n{sep}")
    enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos, error_critico)
    print("Fin.\n")


if __name__ == "__main__":
    asyncio.run(main())
