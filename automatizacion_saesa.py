import asyncio
import os
import re
import smtplib
from datetime import datetime
from zoneinfo import ZoneInfo
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from playwright.async_api import async_playwright


# =============================================================================
# CONFIG
# =============================================================================

SAESA_URL = "https://stx.saesa.cl:8091/backend/sts/login.php?backurl=%2Fbackend%2Fsts%2Fcentrality.php"
OPAT_URL  = "https://opat.cl/agendaopat/"

SAESA_USER = os.environ["SAESA_USER"]
SAESA_PASS = os.environ["SAESA_PASS"]

OPAT_USER  = os.environ["OPAT_USER"]
OPAT_PASS  = os.environ["OPAT_PASS"]

GMAIL_USER = os.environ["GMAIL_USER"]
GMAIL_PASS = os.environ["GMAIL_APP_PASS"]
EMAIL_DEST = os.environ["EMAIL_DEST"]
EMAIL_CC   = ["nicolas.lorenzen@saesa.cl"]

DRY_RUN          = os.environ.get("DRY_RUN", "true").lower() == "true"
MAX_APROBACIONES = int(os.environ.get("MAX_APROBACIONES", "50"))

TIMEOUT   = 30_000
TZ_CHILE  = ZoneInfo("America/Santiago")

ESTADO_EXACTO = "Revisión y Autorización PCCT"
AREA_KEYWORDS = ["metropolitana"]


# =============================================================================
# JS HELPERS — CENTRALITY
# =============================================================================

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
        return {found: true, disabled: disabled};
    }
    return {found: false};
}
"""

JS_CLICK_BTN_APROBAR = """
() => {
    const candidatos = Array.from(
        document.querySelectorAll("button.x-btn-text, .x-btn-text, button")
    ).filter(el => el.offsetParent);
    for (const el of candidatos) {
        const txt = (el.innerText || el.textContent || "").trim();
        if (txt === "Aprobar") {
            el.click();
            return {clicked: true, tag: el.tagName};
        }
    }
    return {clicked: false};
}
"""

JS_DETECT_POPUP = """
() => {
    var headers = Array.from(document.querySelectorAll(
        ".x-window-header-text, .x-panel-header-text"
    ));
    for (var i = 0; i < headers.length; i++) {
        var h = headers[i];
        if (!h.offsetParent) continue;
        var txt = (h.innerText || h.textContent || "").trim();
        if (txt === "Aprobar") {
            var win = h.closest(".x-window");
            if (!win || !win.offsetParent) continue;
            var r = win.getBoundingClientRect();
            return {found: true, x: Math.round(r.x), y: Math.round(r.y)};
        }
    }
    return {found: false};
}
"""

JS_CLICK_ACEPTAR = """
() => {
    var headers = Array.from(document.querySelectorAll(
        ".x-window-header-text, .x-panel-header-text"
    ));
    for (var i = 0; i < headers.length; i++) {
        var h = headers[i];
        if (!h.offsetParent) continue;
        if ((h.innerText || h.textContent || "").trim() !== "Aprobar") continue;
        var win = h.closest(".x-window");
        if (!win || !win.offsetParent) continue;
        var ta = win.querySelector("textarea");
        if (ta) {
            ta.value = "";
            ta.dispatchEvent(new Event("input", {bubbles: true}));
            ta.dispatchEvent(new Event("change", {bubbles: true}));
        }
        var btns = Array.from(win.querySelectorAll("button,.x-btn"));
        for (var j = 0; j < btns.length; j++) {
            var t = (btns[j].innerText || btns[j].textContent || "").trim();
            if (t === "Aceptar") { btns[j].click(); return {ok: true}; }
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

# Lee el detalle completo del PT seleccionado en el panel inferior de Centrality
JS_LEER_DETALLE_PT = """
() => {
    // El detalle aparece en el panel inferior con etiquetas y valores
    var detalle = {};

    // Buscar todas las celdas de la tabla de detalle
    var celdas = Array.from(document.querySelectorAll(
        ".x-panel-body table td, .x-grid3-body td"
    ));

    var campos = [
        "Identificador", "Fecha de recepción", "Tipo de permiso de trabajo",
        "Instalación", "Instalación a intervenir", "Detalle de instalación",
        "Solicitante de PT", "Estado", "Comentarios del cierre",
        "Creador", "Desde", "Hasta",
        "Modificación al esquema eléctrico", "Observaciones", "Elemento referencia",
        "Jerarquía de red", "Área de cobertura", "Área de mantenimiento",
        "Dirección", "Sector afectado", "Descripción del trabajo general",
        "Sectores sin suministro", "Modifica instalaciones", "Área",
        "Caso", "Editable por", "Responsable del caso",
        "Fuera de plazo", "Requiere planificación de faena", "Requiere evaluación de riesgos"
    ];

    // Buscar en el DOM del panel de detalle
    var panelDetalle = document.querySelector(".x-panel-body");
    if (!panelDetalle) return detalle;

    var tds = Array.from(panelDetalle.querySelectorAll("td"));

    for (var i = 0; i < tds.length - 1; i++) {
        var label = (tds[i].innerText || "").trim().replace(/:$/, "");
        if (campos.indexOf(label) >= 0) {
            var valor = (tds[i + 1].innerText || "").trim();
            detalle[label] = valor;
        }
    }

    return detalle;
}
"""


# =============================================================================
# UTILS
# =============================================================================

async def screenshot(page, nombre):
    os.makedirs("capturas", exist_ok=True)
    ts   = datetime.now().strftime("%H%M%S")
    safe = re.sub(r"[^a-zA-Z0-9_-]", "_", nombre)
    path = f"capturas/{safe}_{ts}.png"
    await page.screenshot(path=path, full_page=False)
    print(f"    captura: {path}")
    return path


def normalizar(txt):
    return " ".join((txt or "").strip().split())


def es_metropolitana(area):
    return any(k in (area or "").lower() for k in AREA_KEYWORDS)


def es_estado_pcct_exacto(estado):
    return normalizar(estado) == ESTADO_EXACTO


def extraer_info_fila(row):
    id_pt = area_pt = estado_pt = ""
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
            "zonal", "mantenimiento",
        ]
        if any(k in c.lower() for k in posibles_areas) and not area_pt:
            area_pt = c
    return id_pt, area_pt, estado_pt


def determinar_tipo_trabajo(tipo_pt_texto):
    """Mapea el tipo de PT de Centrality al tipo en OPAT."""
    t = (tipo_pt_texto or "").upper()
    if "DESCONEX" in t:
        return "DESCONEXIÓN"
    if "INTERVEN" in t:
        return "INTERVENCIÓN"
    return "DESCONEXIÓN"  # default


def parsear_fecha_hora(texto):
    """
    Parsea 'DD/MM/YYYY HH:MM:SS' o 'DD/MM/YYYY HH:MM' o '01/06/2026 09:00:00'
    Devuelve (fecha_mm_dd_yyyy, hora_hh_mm_ampm) para OPAT.
    Ej: ('06/01/2026', '09:00 AM')
    """
    try:
        texto = normalizar(texto)
        # Intentar varios formatos
        for fmt in ["%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"]:
            try:
                dt = datetime.strptime(texto, fmt)
                fecha = dt.strftime("%m/%d/%Y")   # OPAT usa mm/dd/yyyy
                hora  = dt.strftime("%I:%M %p")    # ej: 09:00 AM
                return fecha, hora
            except ValueError:
                continue
    except Exception:
        pass
    return None, None


# =============================================================================
# LOGIN CENTRALITY
# =============================================================================

async def hacer_login_centrality(page):
    print("\n[1] LOGIN CENTRALITY")
    await page.goto(SAESA_URL, wait_until="domcontentloaded", timeout=60_000)
    await page.wait_for_timeout(3000)
    usuario = await page.query_selector('input[name="user"], input[type="text"]')
    if usuario:
        await usuario.fill(SAESA_USER)
    password = await page.query_selector('input[name="pass"], input[type="password"]')
    if password:
        await password.fill(SAESA_PASS)
    await page.click('input[value="Login"], button:has-text("Login"), input[type="submit"]')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2500)
    print("  OK: sesión Centrality iniciada")


# =============================================================================
# NAVEGACIÓN CENTRALITY
# =============================================================================

async def navegar_a_permisos(page):
    print("\n[2] NAVEGACION")
    await page.wait_for_selector(
        'a:has-text("Aplicaciones"), span:has-text("Aplicaciones")', timeout=TIMEOUT)
    await page.click('a:has-text("Aplicaciones"), span:has-text("Aplicaciones")')
    await page.wait_for_timeout(1500)
    await page.wait_for_selector('a:has-text("DMS")', timeout=TIMEOUT)
    await page.click('a:has-text("DMS")')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    print("  -> DMS cargado")

    frame = page
    for f in page.frames:
        try:
            if f.name == "content":
                frame = f
                break
            el = await f.query_selector('text="Planificación"')
            if el:
                frame = f
                break
        except Exception:
            pass

    await frame.wait_for_selector('text="Planificación"', timeout=TIMEOUT)
    await frame.click('text="Planificación"')
    await page.wait_for_timeout(1000)
    await frame.wait_for_selector('text="Permisos de trabajo"', timeout=TIMEOUT)
    await frame.click('text="Permisos de trabajo"')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3500)
    print("  -> Permisos de trabajo")
    return frame


# =============================================================================
# FILTRO PCCT
# =============================================================================

async def aplicar_filtro_pcct(page, frame):
    print("\n[3] FILTRO")
    await frame.click('text=Filtro')
    await page.wait_for_timeout(2000)

    r_estado = await frame.evaluate("""
    () => {
        const win = Array.from(document.querySelectorAll(".x-window"))
            .filter(w => w.offsetParent && (w.innerText || "").includes("Filtros"))[0];
        if (!win) return {ok:false, msg:"No encontré ventana Filtros"};
        const labels = Array.from(win.querySelectorAll("label,td,div,span,b"))
            .filter(el => el.offsetParent);
        let estadoLabel = null;
        for (const el of labels) {
            if ((el.innerText || "").trim() === "Estado:") { estadoLabel = el; break; }
        }
        if (!estadoLabel) return {ok:false, msg:"No encontré label Estado:"};
        const lr = estadoLabel.getBoundingClientRect();
        const wr = win.getBoundingClientRect();
        const y = lr.y + lr.height / 2;
        const x = wr.right - 22;
        const el = document.elementFromPoint(x, y);
        if (!el) return {ok:false, msg:"elementFromPoint no encontró elemento"};
        el.click();
        return {ok:true, x:Math.round(x), y:Math.round(y)};
    }
    """)
    print(f"  trigger Estado: {r_estado}")
    if not r_estado.get("ok"):
        raise RuntimeError(f"No se pudo abrir combo Estado: {r_estado}")

    await page.wait_for_timeout(1500)

    r_pcct = await frame.evaluate("""
    () => {
        const items = Array.from(document.querySelectorAll(".x-combo-list-item"))
            .filter(el => el.offsetParent);
        for (const item of items) {
            const raw = (item.innerText || "").trim();
            if (raw.indexOf("PCCT") >= 0 && raw.indexOf("FP") < 0 && raw.indexOf("JACCT") < 0) {
                item.scrollIntoView({block: "center"});
                var e1 = new MouseEvent("mousedown", {bubbles:true, cancelable:true});
                var e2 = new MouseEvent("mouseup",   {bubbles:true, cancelable:true});
                var e3 = new MouseEvent("click",     {bubbles:true, cancelable:true});
                item.dispatchEvent(e1);
                item.dispatchEvent(e2);
                item.dispatchEvent(e3);
                return {ok:true, texto:raw};
            }
        }
        return {ok:false};
    }
    """)
    print(f"  selección PCCT: {r_pcct}")
    if not r_pcct.get("ok"):
        raise RuntimeError(f"No se pudo seleccionar Estado PCCT: {r_pcct}")

    await page.wait_for_timeout(1000)

    aplicar_btn = frame.locator("button.x-btn-text.apply", has_text="Aplicar").first
    await aplicar_btn.click(timeout=5000, force=True)
    await page.wait_for_timeout(8000)

    filtro_abierto = await frame.evaluate("""
    () => {
        const win = Array.from(document.querySelectorAll(".x-window"))
            .filter(w => w.offsetParent && (w.innerText || "").includes("Filtros"))[0];
        return !!win;
    }
    """)
    if filtro_abierto:
        await aplicar_btn.dblclick(timeout=5000, force=True)
        await page.wait_for_timeout(8000)

    info = await frame.evaluate("""
    () => {
        const rows = document.querySelectorAll(".x-grid3-row");
        const pag = Array.from(document.querySelectorAll("*"))
            .filter(e => e.children.length===0 && e.offsetParent &&
                         (e.innerText||"").indexOf("Mostrando")>=0)
            .map(e => e.innerText.trim());
        return {filas: rows.length, paginador: pag};
    }
    """)
    print(f"  resultado filtro: {info}")


# =============================================================================
# SELECCIÓN DE FILA
# =============================================================================

async def seleccionar_fila_pt(page, frame, pt_id):
    try:
        row   = frame.locator(".x-grid3-row", has_text=pt_id).first
        await row.scroll_into_view_if_needed(timeout=5000)
        await page.wait_for_timeout(500)
        await row.click(timeout=5000, force=True)
        await page.wait_for_timeout(1200)
        celda = frame.locator(".x-grid3-cell-inner", has_text=pt_id).first
        await celda.click(timeout=5000, force=True)
        await page.wait_for_timeout(1200)
        selected = await row.evaluate("""
        (el) => el.classList.contains("x-grid3-row-selected") ||
                el.className.includes("selected")
        """)
        return {"found": True, "selected": selected}
    except Exception as e:
        return {"found": False, "selected": False, "error": str(e)}


# =============================================================================
# LEER DETALLE COMPLETO DEL PT EN CENTRALITY
# =============================================================================

async def leer_detalle_pt(page, frame, pt_id):
    """
    Selecciona la fila y lee todos los campos del panel inferior de detalle.
    Devuelve un dict con los datos necesarios para OPAT.
    """
    try:
        # Asegurarse de que la fila esté seleccionada
        await seleccionar_fila_pt(page, frame, pt_id)
        await page.wait_for_timeout(1500)

        detalle = await frame.evaluate(JS_LEER_DETALLE_PT)
        print(f"    detalle leído: {list(detalle.keys())}")

        # Extraer campos para OPAT
        desde_raw = detalle.get("Desde", "")
        hasta_raw = detalle.get("Hasta", "")

        fecha_inicio, hora_inicio = parsear_fecha_hora(desde_raw)
        fecha_fin,    hora_fin    = parsear_fecha_hora(hasta_raw)

        tipo_pt_texto = detalle.get("Tipo de permiso de trabajo", "")
        tipo_trabajo  = determinar_tipo_trabajo(tipo_pt_texto)

        se_linea      = detalle.get("Elemento referencia", "")
        componentes   = detalle.get("Detalle de instalación", "")
        descripcion   = detalle.get("Descripción del trabajo general", "")

        return {
            "id":            pt_id,
            "fecha_inicio":  fecha_inicio,
            "hora_inicio":   hora_inicio,
            "fecha_fin":     fecha_fin,
            "hora_fin":      hora_fin,
            "tipo_trabajo":  tipo_trabajo,
            "se_linea":      se_linea,
            "componentes":   componentes,
            "descripcion":   descripcion,
            "raw":           detalle,
        }
    except Exception as e:
        print(f"    ERROR leyendo detalle: {e}")
        return None


# =============================================================================
# LOGIN OPAT
# =============================================================================

async def hacer_login_opat(opat_page):
    print("\n[OPAT] LOGIN")
    await opat_page.goto(OPAT_URL, wait_until="domcontentloaded", timeout=60_000)
    await opat_page.wait_for_timeout(3000)
    await screenshot(opat_page, "opat_login")

    # Buscar campos de login
    # Campo "Nombre de Usuario" (placeholder "Ej. admin")
    usuario = await opat_page.query_selector(
        'input[placeholder*="admin"], input[placeholder*="Usuario"], '
        'input[placeholder*="usuario"], input[type="text"]'
    )
    if usuario:
        await usuario.fill(OPAT_USER)

    # Campo Contraseña
    password = await opat_page.query_selector('input[type="password"]')
    if password:
        await password.fill(OPAT_PASS)

    await opat_page.click('button:has-text("Ingresar al Sistema")')
    await opat_page.wait_for_load_state("networkidle", timeout=30_000)
    await opat_page.wait_for_timeout(2000)
    await screenshot(opat_page, "opat_post_login")
    print("  OK: sesión OPAT iniciada")


# =============================================================================
# SUBIR PT A OPAT
# =============================================================================

async def cerrar_modal_opat(opat_page):
    """Cierra el modal de OPAT si está abierto, via JS directo."""
    try:
        await opat_page.evaluate("""
        () => {
            // Cerrar modal Bootstrap via JS
            var modal = document.getElementById('modalEditarPT');
            if (modal && modal.classList.contains('show')) {
                // Usar Bootstrap JS si está disponible
                if (window.bootstrap && window.bootstrap.Modal) {
                    var m = window.bootstrap.Modal.getInstance(modal);
                    if (m) { m.hide(); return 'bootstrap_hide'; }
                }
                // Fallback: manipular DOM directamente
                modal.classList.remove('show');
                modal.style.display = 'none';
                modal.setAttribute('aria-hidden', 'true');
                modal.removeAttribute('aria-modal');
                // Quitar backdrop
                var backdrops = document.querySelectorAll('.modal-backdrop');
                backdrops.forEach(function(b) { b.remove(); });
                document.body.classList.remove('modal-open');
                document.body.style.overflow = '';
                document.body.style.paddingRight = '';
                return 'dom_manipulated';
            }
            return 'not_open';
        }
        """)
        await opat_page.wait_for_timeout(800)
    except Exception as e:
        print(f"  cerrar_modal_opat error: {e}")


async def abrir_nuevo_pt_opat(opat_page):
    """Abre el formulario Nuevo PT via JS para evitar problemas de intercepción."""
    try:
        # Primero cerrar cualquier modal abierto
        await cerrar_modal_opat(opat_page)
        await opat_page.wait_for_timeout(500)

        # Intentar click normal primero
        try:
            await opat_page.click(
                'button:has-text("Nuevo PT"), a:has-text("Nuevo PT")',
                timeout=5000,
                force=True
            )
        except Exception:
            # Fallback: llamar la función JS directamente
            await opat_page.evaluate("() => { if (typeof abrirNuevoPT === 'function') abrirNuevoPT(); }")

        await opat_page.wait_for_timeout(2000)

        # Verificar que el modal se abrió
        modal_visible = await opat_page.evaluate("""
        () => {
            var modal = document.getElementById('modalEditarPT');
            return modal && modal.classList.contains('show');
        }
        """)
        return modal_visible
    except Exception as e:
        print(f"  abrir_nuevo_pt_opat error: {e}")
        return False


async def subir_pt_a_opat(opat_page, datos):
    """
    Abre el formulario Nuevo PT en OPAT y llena todos los campos.
    Devuelve True si se guardó exitosamente, False en caso contrario.
    """
    pt_id = datos["id"]
    print(f"\n[OPAT] Subiendo PT {pt_id}...")

    try:
        # Asegurarse de estar en la página principal de OPAT
        if "agendaopat" not in opat_page.url:
            await opat_page.goto(OPAT_URL, wait_until="domcontentloaded", timeout=30_000)
            await opat_page.wait_for_timeout(2000)

        # Abrir formulario Nuevo PT (con cierre de modal previo)
        modal_ok = await abrir_nuevo_pt_opat(opat_page)
        print(f"  modal abierto: {modal_ok}")
        await screenshot(opat_page, f"opat_form_{pt_id}")

        # Trabajar dentro del modal #modalEditarPT
        modal = opat_page.locator('#modalEditarPT')

        # ── N° PT ──────────────────────────────────────────────────────────────
        campo_npt = modal.locator(
            'input[placeholder*="Automático"], input[placeholder*="ingrese"], input[placeholder*="PT"]'
        ).first
        await campo_npt.fill("")
        await campo_npt.fill(pt_id)
        await opat_page.wait_for_timeout(500)
        print(f"  N° PT: {pt_id}")

        # ── ZONA → Metropolitana ───────────────────────────────────────────────
        zona_select = modal.locator('select').filter(has_text="Seleccione Zona").first
        await zona_select.select_option(label="Metropolitana")
        await opat_page.wait_for_timeout(500)
        print("  Zona: Metropolitana")

        # ── FECHA INICIO ──────────────────────────────────────────────────────
        if datos.get("fecha_inicio"):
            fecha_str = datetime.strptime(datos["fecha_inicio"], "%m/%d/%Y").strftime("%Y-%m-%d")
            await modal.locator('input[type="date"]').first.fill(fecha_str)
            await opat_page.wait_for_timeout(300)
            print(f"  Fecha inicio: {datos['fecha_inicio']}")

        # ── HORA INICIO ───────────────────────────────────────────────────────
        if datos.get("hora_inicio"):
            dt_hora = datetime.strptime(datos["hora_inicio"], "%I:%M %p")
            await modal.locator('input[type="time"]').first.fill(dt_hora.strftime("%H:%M"))
            await opat_page.wait_for_timeout(300)
            print(f"  Hora inicio: {datos['hora_inicio']}")

        # ── FECHA FIN ─────────────────────────────────────────────────────────
        if datos.get("fecha_fin"):
            fecha_fin_str = datetime.strptime(datos["fecha_fin"], "%m/%d/%Y").strftime("%Y-%m-%d")
            await modal.locator('input[type="date"]').nth(1).fill(fecha_fin_str)
            await opat_page.wait_for_timeout(300)
            print(f"  Fecha fin: {datos['fecha_fin']}")

        # ── HORA FIN ──────────────────────────────────────────────────────────
        if datos.get("hora_fin"):
            dt_hora_fin = datetime.strptime(datos["hora_fin"], "%I:%M %p")
            await modal.locator('input[type="time"]').nth(1).fill(dt_hora_fin.strftime("%H:%M"))
            await opat_page.wait_for_timeout(300)
            print(f"  Hora fin: {datos['hora_fin']}")

        # ── TIPO TRABAJO ──────────────────────────────────────────────────────
        tipo_trabajo = datos.get("tipo_trabajo", "DESCONEXIÓN")
        tipo_select  = modal.locator('select').filter(has_text="DESCONEXIÓN").first
        await tipo_select.select_option(label=tipo_trabajo)
        await opat_page.wait_for_timeout(500)
        print(f"  Tipo trabajo: {tipo_trabajo}")

        # ── SE O LÍNEA ────────────────────────────────────────────────────────
        if datos.get("se_linea"):
            se_input = modal.locator('input[placeholder*="SE"], input[placeholder*="Línea"]').first
            await se_input.fill(datos["se_linea"])
            await opat_page.wait_for_timeout(300)
            print(f"  SE o Línea: {datos['se_linea']}")

        # ── ÁREA ZONAL ────────────────────────────────────────────────────────
        area_input = modal.locator('input[placeholder*="Metropolitana"]').first
        await area_input.fill("Área Mtto Zonal Metropolitana")
        await opat_page.wait_for_timeout(300)
        print("  Área Zonal: Área Mtto Zonal Metropolitana")

        # ── COMPONENTES ───────────────────────────────────────────────────────
        if datos.get("componentes"):
            comp_input = modal.locator('input[placeholder*="52J"]').first
            await comp_input.fill(datos["componentes"])
            await opat_page.wait_for_timeout(300)
            print(f"  Componentes: {datos['componentes'][:50]}")

        # ── DESCRIPCIÓN / OBJETIVO GENERAL ───────────────────────────────────
        if datos.get("descripcion"):
            desc_textarea = modal.locator('textarea').first
            await desc_textarea.fill(datos["descripcion"])
            await opat_page.wait_for_timeout(300)
            print(f"  Descripción: {datos['descripcion'][:60]}...")

        await screenshot(opat_page, f"opat_pre_guardar_{pt_id}")

        # ── GUARDAR Y CERRAR ──────────────────────────────────────────────────
        guardar_btn = modal.locator('button:has-text("Guardar y Cerrar"), button:has-text("Guardar")').first
        await guardar_btn.click(timeout=10_000)
        await opat_page.wait_for_timeout(3000)
        await screenshot(opat_page, f"opat_post_guardar_{pt_id}")

        # Verificar que el modal se cerró
        modal_aun_abierto = await opat_page.evaluate("""
        () => {
            var m = document.getElementById('modalEditarPT');
            return m && m.classList.contains('show');
        }
        """)

        if modal_aun_abierto:
            print(f"  ADVERTENCIA: modal sigue abierto para {pt_id}")
            await cerrar_modal_opat(opat_page)
            return False

        print(f"  ✓ PT {pt_id} subido a OPAT exitosamente")
        return True

    except Exception as e:
        print(f"  ERROR subiendo {pt_id} a OPAT: {e}")
        await screenshot(opat_page, f"opat_error_{pt_id}")
        # Intentar cerrar modal para no bloquear el siguiente PT
        await cerrar_modal_opat(opat_page)
        return False


# =============================================================================
# APROBAR PTS EN CENTRALITY Y SUBIR A OPAT
# =============================================================================

async def aprobar_pts(page, frame, opat_page):
    print("\n[4] APROBANDO PTs y subiendo a OPAT")
    print(f"  DRY_RUN: {DRY_RUN}")

    pts_aprobados = []
    pts_fallidos  = []
    pts_omitidos  = []

    total_paginas = await frame.evaluate(JS_GET_TOTAL_PAGES)
    paginas = min(total_paginas, 20)
    print(f"  Total páginas: {total_paginas}")

    for pagina in range(1, paginas + 1):
        print(f"\n  ── Página {pagina}/{paginas} ──")
        await page.wait_for_timeout(1500)

        filas = await frame.evaluate(JS_READ_ROWS)
        print(f"  Filas leídas: {len(filas)}")

        pts_esta_pagina = []
        for row in filas:
            if not row:
                continue
            id_pt, area_pt, estado_pt = extraer_info_fila(row)
            if not id_pt:
                continue
            if not es_estado_pcct_exacto(estado_pt):
                pts_omitidos.append({
                    "id": id_pt,
                    "area": area_pt or "Sin área",
                    "motivo": f"Estado: {estado_pt or 'sin estado'}"
                })
                print(f"    [OMITIR ESTADO] {id_pt}")
                continue
            if es_metropolitana(area_pt):
                pts_esta_pagina.append({"id": id_pt, "area": area_pt, "estado": estado_pt})
                print(f"    [APROBAR] {id_pt} | {area_pt}")
            else:
                pts_omitidos.append({
                    "id": id_pt,
                    "area": area_pt or "Sin área",
                    "motivo": "Área no Metropolitana"
                })
                print(f"    [OMITIR AREA] {id_pt} | {area_pt}")

        for pt in pts_esta_pagina:
            if len(pts_aprobados) >= MAX_APROBACIONES:
                print("    LÍMITE DE SEGURIDAD ALCANZADO")
                return pts_aprobados, pts_fallidos, pts_omitidos

            print(f"\n    >> Procesando {pt['id']}")

            try:
                if DRY_RUN:
                    print(f"    [DRY RUN] {pt['id']}")
                    pts_aprobados.append({
                        "id": pt["id"], "area": pt["area"],
                        "estado": pt["estado"], "opat": False, "opat_dry": True
                    })
                    continue

                # ── 1. LEER DETALLE ANTES DE APROBAR ─────────────────────────
                print("    Leyendo detalle del PT...")
                datos_opat = await leer_detalle_pt(page, frame, pt["id"])

                # ── 2. SELECCIONAR FILA ───────────────────────────────────────
                sel = await seleccionar_fila_pt(page, frame, pt["id"])
                print(f"    selección: {sel}")
                if not sel.get("found"):
                    pts_fallidos.append(f"{pt['id']} - fila no encontrada")
                    continue
                if not sel.get("selected"):
                    pts_fallidos.append(f"{pt['id']} - fila no seleccionada")
                    continue

                # ── 3. VALIDAR BOTÓN APROBAR ──────────────────────────────────
                btn = await frame.evaluate(JS_CHECK_BTN_APROBAR)
                print(f"    botón Aprobar: {btn}")
                if not btn.get("found"):
                    pts_fallidos.append(f"{pt['id']} - botón Aprobar no visible")
                    continue
                if btn.get("disabled"):
                    pts_fallidos.append(f"{pt['id']} - botón Aprobar deshabilitado")
                    continue

                await screenshot(page, f"pre_{pt['id']}")

                # ── 4. CLICK APROBAR ──────────────────────────────────────────
                click_r = await frame.evaluate(JS_CLICK_BTN_APROBAR)
                print(f"    click Aprobar: {click_r}")
                if not click_r.get("clicked"):
                    btn_loc = frame.locator("button.x-btn-text", has_text="Aprobar").first
                    await btn_loc.click(timeout=5000, force=True)

                await page.wait_for_timeout(2000)

                # ── 5. DETECTAR POPUP ─────────────────────────────────────────
                popup = {"found": False}
                for intento in range(14):
                    await page.wait_for_timeout(700)
                    popup = await frame.evaluate(JS_DETECT_POPUP)
                    if popup.get("found"):
                        print(f"    popup OK intento {intento+1}")
                        break

                await screenshot(page, f"popup_{pt['id']}")
                if not popup.get("found"):
                    pts_fallidos.append(f"{pt['id']} - popup no apareció")
                    continue

                # ── 6. CLICK ACEPTAR ──────────────────────────────────────────
                aceptar = await frame.evaluate(JS_CLICK_ACEPTAR)
                print(f"    Aceptar: {aceptar}")
                if not aceptar.get("ok"):
                    pts_fallidos.append(f"{pt['id']} - Aceptar falló")
                    continue

                # ── 7. ESPERAR PROCESAMIENTO CENTRALITY ───────────────────────
                print("    esperando procesamiento Centrality...")
                await page.wait_for_timeout(5000)
                for _ in range(20):
                    popup_abierto = await frame.evaluate("""
                    () => Array.from(document.querySelectorAll(".x-window"))
                        .some(w => w.offsetParent &&
                            ((w.innerText||"").includes("Aprobar") ||
                             (w.innerText||"").includes("Confirm")))
                    """)
                    if not popup_abierto:
                        break
                    await page.wait_for_timeout(1000)

                await frame.evaluate(JS_REFRESH_GRID)
                await page.wait_for_timeout(5000)

                # ── 8. VERIFICAR DESAPARICIÓN ─────────────────────────────────
                desaparecio = False
                for intento in range(20):
                    if not await frame.evaluate(JS_PT_EXISTE, pt["id"]):
                        desaparecio = True
                        break
                    await page.wait_for_timeout(1500)

                await screenshot(page, f"post_{pt['id']}")

                if not desaparecio:
                    pts_fallidos.append(f"{pt['id']} - sigue visible")
                    continue

                print(f"    ✓ APROBADO en Centrality: {pt['id']}")

                # ── 9. SUBIR A OPAT ───────────────────────────────────────────
                opat_ok = False
                if datos_opat and opat_page:
                    opat_ok = await subir_pt_a_opat(opat_page, datos_opat)
                else:
                    print(f"    ADVERTENCIA: sin datos para OPAT ({pt['id']})")

                pts_aprobados.append({
                    "id":    pt["id"],
                    "area":  pt["area"],
                    "estado": pt["estado"],
                    "opat":  opat_ok
                })

            except Exception as e:
                msg = str(e)[:250]
                pts_fallidos.append(f"{pt['id']} - {msg}")
                print(f"    EXCEPCIÓN: {msg}")
                await screenshot(page, f"exc_{pt['id']}")

        if len(pts_aprobados) >= MAX_APROBACIONES:
            break

        if pagina < paginas:
            sig = await frame.evaluate(JS_NEXT_PAGE)
            if not sig:
                break
            await page.wait_for_timeout(4000)

    await screenshot(page, "final")
    return pts_aprobados, pts_fallidos, pts_omitidos


# =============================================================================
# CORREO
# =============================================================================

def enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos, error_critico=None):
    ahora_chile = datetime.now(TZ_CHILE)
    fecha       = ahora_chile.strftime("%d/%m/%Y")
    hora        = ahora_chile.strftime("%H:%M")
    hora_ampm   = ahora_chile.strftime("%I:%M %p").lower()

    def filas_aprobados():
        if not pts_aprobados:
            return "<tr><td colspan='4' style='padding:6px 12px;color:#999'>Ninguno</td></tr>"
        html = ""
        for pt in pts_aprobados:
            opat_dry = pt.get("opat_dry", False)
            if opat_dry:
                opat_icon = "<span style='color:#888'>— (simulado)</span>"
            elif pt.get("opat"):
                opat_icon = "<span style='color:#006600;font-size:16px'>&#10003;</span>"
            else:
                opat_icon = "<span style='color:#cc0000;font-size:16px'>&#10007;</span>"
            html += (
                "<tr>"
                "<td style='padding:4px 8px;color:#006600;font-size:16px'>&#10003;</td>"
                f"<td style='font-family:monospace;padding:4px 12px'>{pt.get('id','')}</td>"
                f"<td style='padding:4px 12px'>{pt.get('area','')}</td>"
                f"<td style='padding:4px 12px;text-align:center'>{opat_icon}</td>"
                "</tr>"
            )
        return html

    def filas_fallidos():
        if not pts_fallidos:
            return "<tr><td colspan='2' style='padding:6px 12px;color:#999'>Sin errores</td></tr>"
        return "".join(
            "<tr>"
            "<td style='padding:4px 8px;color:#cc0000;font-size:16px'>&#10007;</td>"
            f"<td style='padding:4px 12px;font-size:13px'>{pt}</td>"
            "</tr>"
            for pt in pts_fallidos
        )

    def filas_omitidos():
        if not pts_omitidos:
            return "<tr><td colspan='4' style='padding:6px 12px;color:#999'>Ninguno</td></tr>"
        html = ""
        for pt in pts_omitidos:
            html += (
                "<tr>"
                "<td style='padding:4px 8px;color:#aaa'>&mdash;</td>"
                f"<td style='font-family:monospace;padding:4px 12px;color:#777'>{pt.get('id','')}</td>"
                f"<td style='padding:4px 12px;color:#777'>{pt.get('area','')}</td>"
                f"<td style='padding:4px 12px;color:#777'>{pt.get('motivo','')}</td>"
                "</tr>"
            )
        return html

    error_bloque = ""
    if error_critico:
        error_bloque = (
            "<div style='background:#fff0f0;border-left:4px solid #c00;"
            "padding:12px 16px;margin:16px 0;border-radius:4px'>"
            f"<strong>Error crítico:</strong><br><code style='font-size:12px'>{error_critico}</code>"
            "</div>"
        )

    # Contar subidos a OPAT
    opat_ok_count = sum(1 for pt in pts_aprobados if pt.get("opat"))

    html = (
        "<html><body style='font-family:Arial,sans-serif;max-width:760px;margin:auto;color:#222'>"
        "<div style='background:#003580;color:white;padding:24px;border-radius:8px 8px 0 0'>"
        "<h2 style='margin:0;font-size:20px'>Reporte PT's &mdash; Centrality / DMS</h2>"
        "<p style='margin:6px 0 0;opacity:.8;font-size:14px'>"
        "Aprobación PCCT &middot; Zonal Metropolitana</p>"
        "</div>"
        "<div style='border:1px solid #ddd;border-top:none;padding:20px 24px;border-radius:0 0 8px 8px'>"
        f"<p><strong>Fecha:</strong> {fecha} {hora}</p>"
        "<p><strong>Criterio:</strong> Estado = Revisión y Autorización PCCT"
        " | Área contiene Metropolitana</p>"
        + error_bloque
        + f"<h3 style='color:#006600;margin:20px 0 8px'>"
        f"PTs Aprobados ({len(pts_aprobados)}) &nbsp;"
        f"<span style='font-size:13px;color:#555'>— Subidos a OPAT: {opat_ok_count}</span></h3>"
        "<table style='border-collapse:collapse;width:100%'>"
        "<tr style='background:#f6f6f6'>"
        "<th style='padding:6px 8px'>✓</th>"
        "<th style='padding:6px 12px;text-align:left'>PT</th>"
        "<th style='padding:6px 12px;text-align:left'>Área</th>"
        "<th style='padding:6px 12px;text-align:center'>OPAT</th>"
        "</tr>"
        f"{filas_aprobados()}</table>"
        f"<p style='font-size:12px;color:#888;margin-top:6px'>"
        f"&#10003; = subido a OPAT &nbsp;&nbsp; &#10007; = error al subir</p>"
        f"<h3 style='color:#cc0000;margin:20px 0 8px'>PTs con Error ({len(pts_fallidos)})</h3>"
        f"<table style='border-collapse:collapse;width:100%'>{filas_fallidos()}</table>"
        f"<h3 style='color:#888;margin:20px 0 8px'>PTs Omitidos ({len(pts_omitidos)})</h3>"
        "<table style='border-collapse:collapse;width:100%'>"
        "<tr style='background:#f6f6f6'><th></th><th>PT</th><th>Área</th><th>Motivo</th></tr>"
        f"{filas_omitidos()}</table>"
        "</div></body></html>"
    )

    if error_critico:
        asunto = f"[Reporte PTS Centrality] ERROR {fecha} {hora_ampm}"
    else:
        asunto = (
            f"[Reporte PTS Centrality] {fecha} {hora_ampm}"
            f" | {len(pts_aprobados)} aprobados"
            f" | {opat_ok_count} en OPAT"
            f" | {len(pts_fallidos)} errores"
            f" | {len(pts_omitidos)} omitidos"
        )

    todos = [EMAIL_DEST] + EMAIL_CC
    msg   = MIMEMultipart("alternative")
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
        print(f"  Error enviando correo: {e}")


# =============================================================================
# MAIN
# =============================================================================

async def main():
    sep = "=" * 65
    print(f"\n{sep}")
    print(f"  SAESA AUTOMATION | {datetime.now(TZ_CHILE).strftime('%d/%m/%Y %H:%M')}")
    print(f"  DRY_RUN: {DRY_RUN}")
    print(f"{sep}")

    pts_aprobados = []
    pts_fallidos  = []
    pts_omitidos  = []
    error_critico = None

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--ignore-certificate-errors",
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
            ],
        )

        # Dos contextos independientes: uno para Centrality, otro para OPAT
        ctx_centrality = await browser.new_context(
            ignore_https_errors=True,
            viewport={"width": 1400, "height": 900},
        )
        ctx_opat = await browser.new_context(
            ignore_https_errors=True,
            viewport={"width": 1400, "height": 900},
        )

        page_centrality = await ctx_centrality.new_page()
        page_opat       = await ctx_opat.new_page()

        try:
            # Login en ambos sistemas en paralelo
            await asyncio.gather(
                hacer_login_centrality(page_centrality),
                hacer_login_opat(page_opat),
            )

            frame = await navegar_a_permisos(page_centrality)
            await aplicar_filtro_pcct(page_centrality, frame)
            pts_aprobados, pts_fallidos, pts_omitidos = await aprobar_pts(
                page_centrality, frame, page_opat
            )

        except Exception as e:
            error_critico = str(e)
            print(f"\nERROR CRÍTICO: {e}")
            try:
                await screenshot(page_centrality, "error_critico")
            except Exception:
                pass

        finally:
            await browser.close()

    print(f"\n{sep}")
    print(f"  {len(pts_aprobados)} aprobados | {len(pts_fallidos)} errores | {len(pts_omitidos)} omitidos")
    print(f"{sep}")

    enviar_reporte(pts_aprobados, pts_fallidos, pts_omitidos, error_critico)
    print("Fin.\n")


if __name__ == "__main__":
    asyncio.run(main())
