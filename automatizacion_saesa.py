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

NEOMANTE_URL  = "https://neomante.coordinador.cl/login?next=%2F"
NEOMANTE_USER = os.environ["NEOMANTE_USER"]
NEOMANTE_PASS = os.environ["NEOMANTE_PASS"]

GMAIL_USER = os.environ["GMAIL_USER"]
GMAIL_PASS = os.environ["GMAIL_APP_PASS"]
EMAIL_DEST = os.environ["EMAIL_DEST"]
EMAIL_CC   = ["nicolas.lorenzen@saesa.cl"]

DRY_RUN          = os.environ.get("DRY_RUN", "true").lower() == "true"
MAX_APROBACIONES = int(os.environ.get("MAX_APROBACIONES", "1"))

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
    # Reintentar hasta 3 veces si hay timeout
    for intento in range(3):
        try:
            await opat_page.goto(OPAT_URL, wait_until="domcontentloaded", timeout=120_000)
            break
        except Exception as e:
            print(f"  OPAT goto intento {intento+1} falló: {e}")
            if intento == 2:
                raise
            await opat_page.wait_for_timeout(5000)
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
            for intento in range(3):
                try:
                    await opat_page.goto(OPAT_URL, wait_until="domcontentloaded", timeout=120_000)
                    break
                except Exception as e:
                    print(f"  OPAT goto intento {intento+1} falló: {e}")
                    if intento == 2:
                        raise
                    await opat_page.wait_for_timeout(5000)
            await opat_page.wait_for_timeout(2000)

        # Abrir formulario Nuevo PT (con cierre de modal previo)
        modal_ok = await abrir_nuevo_pt_opat(opat_page)
        print(f"  modal abierto: {modal_ok}")
        await screenshot(opat_page, f"opat_form_{pt_id}")

        # Trabajar dentro del modal #modalEditarPT via JavaScript
        # Usar JS evaluate para llenar campos directamente, evitando problemas de intercepción

        async def fill_by_id(field_id, value):
            """Llena un campo por su ID usando JS."""
            result = await opat_page.evaluate(f"""
            (val) => {{
                var el = document.getElementById('{field_id}');
                if (!el) return 'not_found';
                el.focus();
                el.value = val;
                el.dispatchEvent(new Event('input', {{bubbles:true}}));
                el.dispatchEvent(new Event('change', {{bubbles:true}}));
                return 'ok';
            }}
            """, value)
            return result

        async def select_by_id(field_id, value):
            """Selecciona una opción en un select por su ID usando JS."""
            result = await opat_page.evaluate(f"""
            (val) => {{
                var el = document.getElementById('{field_id}');
                if (!el) return 'not_found';
                for (var i=0; i<el.options.length; i++) {{
                    if (el.options[i].text.trim() === val ||
                        el.options[i].value === val) {{
                        el.selectedIndex = i;
                        el.dispatchEvent(new Event('change', {{bubbles:true}}));
                        return 'ok:' + el.options[i].text;
                    }}
                }}
                // Fallback: buscar por texto parcial
                for (var j=0; j<el.options.length; j++) {{
                    if (el.options[j].text.includes(val)) {{
                        el.selectedIndex = j;
                        el.dispatchEvent(new Event('change', {{bubbles:true}}));
                        return 'partial:' + el.options[j].text;
                    }}
                }}
                return 'not_found_value';
            }}
            """, value)
            return result

        # Primero obtener los IDs reales de los campos del formulario
        campos_info = await opat_page.evaluate("""
        () => {
            var modal = document.getElementById('modalEditarPT');
            if (!modal) return {};
            var inputs = Array.from(modal.querySelectorAll('input,select,textarea'));
            var info = {};
            inputs.forEach(function(el) {
                var label = '';
                if (el.id) {
                    var labelEl = document.querySelector('label[for="' + el.id + '"]');
                    if (labelEl) label = labelEl.innerText.trim();
                }
                info[el.id || el.name || 'no-id-' + el.type] = {
                    type: el.type || el.tagName,
                    placeholder: el.placeholder || '',
                    label: label,
                    value: el.value || ''
                };
            });
            return info;
        }
        """)
        print(f"  campos_info: {list(campos_info.keys())}")

        # ── N° PT ──────────────────────────────────────────────────────────────
        # Buscar input de N° PT (placeholder "Automático o ingrese")
        r = await opat_page.evaluate(f"""
        (ptId) => {{
            var modal = document.getElementById('modalEditarPT');
            if (!modal) return 'no modal';
            var inputs = Array.from(modal.querySelectorAll('input[type="text"],input:not([type])'));
            for (var i=0; i<inputs.length; i++) {{
                var ph = (inputs[i].placeholder || '').toLowerCase();
                if (ph.includes('autom') || ph.includes('ingrese') || ph.includes('n°') || ph.includes('npt')) {{
                    inputs[i].value = ptId;
                    inputs[i].dispatchEvent(new Event('input', {{bubbles:true}}));
                    inputs[i].dispatchEvent(new Event('change', {{bubbles:true}}));
                    return 'ok:' + inputs[i].id;
                }}
            }}
            // Fallback: primer input de texto visible
            for (var j=0; j<inputs.length; j++) {{
                if (inputs[j].offsetParent) {{
                    inputs[j].value = ptId;
                    inputs[j].dispatchEvent(new Event('input', {{bubbles:true}}));
                    inputs[j].dispatchEvent(new Event('change', {{bubbles:true}}));
                    return 'fallback:' + inputs[j].id;
                }}
            }}
            return 'not_found';
        }}
        """, pt_id)
        print(f"  N° PT: {pt_id} → {r}")
        await opat_page.wait_for_timeout(300)

        # ── ZONA → Metropolitana ───────────────────────────────────────────────
        r = await opat_page.evaluate("""
        () => {
            var modal = document.getElementById('modalEditarPT');
            var selects = Array.from(modal.querySelectorAll('select'));
            for (var i=0; i<selects.length; i++) {
                var opts = Array.from(selects[i].options).map(o => o.text.trim());
                if (opts.some(o => o.includes('Zona') || o.includes('Metropolitana') || o.includes('Norte') || o.includes('Sur'))) {
                    for (var j=0; j<selects[i].options.length; j++) {
                        if (selects[i].options[j].text.trim() === 'Metropolitana') {
                            selects[i].selectedIndex = j;
                            selects[i].dispatchEvent(new Event('change', {bubbles:true}));
                            return 'ok:' + selects[i].id;
                        }
                    }
                }
            }
            return 'not_found';
        }
        """)
        print(f"  Zona: Metropolitana → {r}")
        await opat_page.wait_for_timeout(300)

        # ── FECHA INICIO ──────────────────────────────────────────────────────
        if datos.get("fecha_inicio"):
            fecha_str = datetime.strptime(datos["fecha_inicio"], "%m/%d/%Y").strftime("%Y-%m-%d")
            r = await opat_page.evaluate(f"""
            (val) => {{
                var modal = document.getElementById('modalEditarPT');
                var dates = Array.from(modal.querySelectorAll('input[type="date"]'));
                if (dates[0]) {{
                    dates[0].value = val;
                    dates[0].dispatchEvent(new Event('change', {{bubbles:true}}));
                    return 'ok:' + dates[0].id;
                }}
                return 'not_found';
            }}
            """, fecha_str)
            print(f"  Fecha inicio: {datos['fecha_inicio']} → {r}")
            await opat_page.wait_for_timeout(300)

        # ── HORA INICIO ───────────────────────────────────────────────────────
        if datos.get("hora_inicio"):
            dt_hora = datetime.strptime(datos["hora_inicio"], "%I:%M %p")
            hora_str = dt_hora.strftime("%H:%M")
            r = await opat_page.evaluate(f"""
            (val) => {{
                var modal = document.getElementById('modalEditarPT');
                var times = Array.from(modal.querySelectorAll('input[type="time"]'));
                if (times[0]) {{
                    times[0].value = val;
                    times[0].dispatchEvent(new Event('change', {{bubbles:true}}));
                    return 'ok:' + times[0].id;
                }}
                return 'not_found';
            }}
            """, hora_str)
            print(f"  Hora inicio: {datos['hora_inicio']} → {r}")
            await opat_page.wait_for_timeout(300)

        # ── FECHA FIN ─────────────────────────────────────────────────────────
        if datos.get("fecha_fin"):
            fecha_fin_str = datetime.strptime(datos["fecha_fin"], "%m/%d/%Y").strftime("%Y-%m-%d")
            r = await opat_page.evaluate(f"""
            (val) => {{
                var modal = document.getElementById('modalEditarPT');
                var dates = Array.from(modal.querySelectorAll('input[type="date"]'));
                if (dates[1]) {{
                    dates[1].value = val;
                    dates[1].dispatchEvent(new Event('change', {{bubbles:true}}));
                    return 'ok:' + dates[1].id;
                }}
                return 'not_found';
            }}
            """, fecha_fin_str)
            print(f"  Fecha fin: {datos['fecha_fin']} → {r}")
            await opat_page.wait_for_timeout(300)

        # ── HORA FIN ──────────────────────────────────────────────────────────
        if datos.get("hora_fin"):
            dt_hora_fin = datetime.strptime(datos["hora_fin"], "%I:%M %p")
            hora_fin_str = dt_hora_fin.strftime("%H:%M")
            r = await opat_page.evaluate(f"""
            (val) => {{
                var modal = document.getElementById('modalEditarPT');
                var times = Array.from(modal.querySelectorAll('input[type="time"]'));
                if (times[1]) {{
                    times[1].value = val;
                    times[1].dispatchEvent(new Event('change', {{bubbles:true}}));
                    return 'ok:' + times[1].id;
                }}
                return 'not_found';
            }}
            """, hora_fin_str)
            print(f"  Hora fin: {datos['hora_fin']} → {r}")
            await opat_page.wait_for_timeout(300)

        # ── TIPO TRABAJO ──────────────────────────────────────────────────────
        tipo_trabajo = datos.get("tipo_trabajo", "DESCONEXIÓN")
        r = await opat_page.evaluate(f"""
        (tipo) => {{
            var modal = document.getElementById('modalEditarPT');
            var selects = Array.from(modal.querySelectorAll('select'));
            for (var i=0; i<selects.length; i++) {{
                var opts = Array.from(selects[i].options).map(o => o.text.trim());
                if (opts.includes('DESCONEXIÓN') || opts.includes('INTERVENCIÓN')) {{
                    for (var j=0; j<selects[i].options.length; j++) {{
                        if (selects[i].options[j].text.trim() === tipo) {{
                            selects[i].selectedIndex = j;
                            selects[i].dispatchEvent(new Event('change', {{bubbles:true}}));
                            return 'ok:' + selects[i].id;
                        }}
                    }}
                }}
            }}
            return 'not_found';
        }}
        """, tipo_trabajo)
        print(f"  Tipo trabajo: {tipo_trabajo} → {r}")
        await opat_page.wait_for_timeout(300)

        # ── SE O LÍNEA ────────────────────────────────────────────────────────
        if datos.get("se_linea"):
            r = await opat_page.evaluate(f"""
            (val) => {{
                var modal = document.getElementById('modalEditarPT');
                // SE o Línea: buscar input de texto cerca del label "SE o Línea"
                var labels = Array.from(modal.querySelectorAll('label'));
                for (var i=0; i<labels.length; i++) {{
                    var lt = (labels[i].innerText || '').trim();
                    if (lt.includes('SE') || lt.includes('L') && lt.includes('nea')) {{
                        var forId = labels[i].getAttribute('for');
                        if (forId) {{
                            var el = document.getElementById(forId);
                            if (el) {{
                                el.value = val;
                                el.dispatchEvent(new Event('input', {{bubbles:true}}));
                                el.dispatchEvent(new Event('change', {{bubbles:true}}));
                                return 'ok:' + forId;
                            }}
                        }}
                    }}
                }}
                // Fallback: buscar por placeholder vacío entre los inputs de texto
                var inputs = Array.from(modal.querySelectorAll('input[type="text"]'));
                for (var j=0; j<inputs.length; j++) {{
                    var ph = (inputs[j].placeholder || '').trim();
                    if (ph === '' && inputs[j].offsetParent) {{
                        inputs[j].value = val;
                        inputs[j].dispatchEvent(new Event('input', {{bubbles:true}}));
                        inputs[j].dispatchEvent(new Event('change', {{bubbles:true}}));
                        return 'fallback_empty:' + inputs[j].id;
                    }}
                }}
                return 'not_found';
            }}
            """, datos["se_linea"])
            print(f"  SE o Línea: {datos['se_linea']} → {r}")
            await opat_page.wait_for_timeout(300)

        # ── ÁREA ZONAL ────────────────────────────────────────────────────────
        r = await opat_page.evaluate("""
        (val) => {
            var modal = document.getElementById('modalEditarPT');
            var inputs = Array.from(modal.querySelectorAll('input[type="text"]'));
            for (var i=0; i<inputs.length; i++) {
                var ph = (inputs[i].placeholder || '').toLowerCase();
                if (ph.includes('metropolitana') || ph.includes('area') || ph.includes('área')) {
                    inputs[i].value = val;
                    inputs[i].dispatchEvent(new Event('input', {bubbles:true}));
                    inputs[i].dispatchEvent(new Event('change', {bubbles:true}));
                    return 'ok:' + inputs[i].id;
                }
            }
            return 'not_found';
        }
        """, "Área Mtto Zonal Metropolitana")
        print(f"  Área Zonal → {r}")
        await opat_page.wait_for_timeout(300)

        # ── COMPONENTES ───────────────────────────────────────────────────────
        if datos.get("componentes"):
            r = await opat_page.evaluate(f"""
            (val) => {{
                var modal = document.getElementById('modalEditarPT');
                var inputs = Array.from(modal.querySelectorAll('input'));
                for (var i=0; i<inputs.length; i++) {{
                    var ph = (inputs[i].placeholder || '').toLowerCase();
                    if (ph.includes('52j') || ph.includes('componente')) {{
                        inputs[i].value = val;
                        inputs[i].dispatchEvent(new Event('input', {{bubbles:true}}));
                        inputs[i].dispatchEvent(new Event('change', {{bubbles:true}}));
                        return 'ok:' + inputs[i].id;
                    }}
                }}
                return 'not_found';
            }}
            """, datos["componentes"])
            print(f"  Componentes → {r}")
            await opat_page.wait_for_timeout(300)

        # ── DESCRIPCIÓN / OBJETIVO GENERAL ───────────────────────────────────
        if datos.get("descripcion"):
            r = await opat_page.evaluate(f"""
            (val) => {{
                var modal = document.getElementById('modalEditarPT');
                var textareas = Array.from(modal.querySelectorAll('textarea'));
                if (textareas[0]) {{
                    textareas[0].value = val;
                    textareas[0].dispatchEvent(new Event('input', {{bubbles:true}}));
                    textareas[0].dispatchEvent(new Event('change', {{bubbles:true}}));
                    return 'ok:' + textareas[0].id;
                }}
                return 'not_found';
            }}
            """, datos["descripcion"])
            print(f"  Descripción → {r}")
            await opat_page.wait_for_timeout(300)

        await screenshot(opat_page, f"opat_pre_guardar_{pt_id}")

        # ── GUARDAR Y CERRAR ──────────────────────────────────────────────────
        # Usar JS para hacer click en el botón Guardar y Cerrar dentro del modal
        r_guardar = await opat_page.evaluate("""
        () => {
            var modal = document.getElementById('modalEditarPT');
            if (!modal) return 'no_modal';
            var btns = Array.from(modal.querySelectorAll('button'));
            for (var i=0; i<btns.length; i++) {
                var t = (btns[i].innerText || btns[i].textContent || '').trim();
                if (t.includes('Guardar')) {
                    btns[i].click();
                    return 'ok:' + t;
                }
            }
            return 'not_found';
        }
        """)
        print(f"  Guardar: {r_guardar}")
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

async def aprobar_pts(page, frame, opat_page, neo_page):
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
                        "estado": pt["estado"], "opat": False, "opat_dry": True,
                        "cen": None
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

                # Crear aviso CEN en Neomante
                numero_cen = None
                if opat_ok and neo_page:
                    numero_cen = await crear_aviso_cen(neo_page, datos_opat)
                    # Actualizar campo modAvisoCen en OPAT con el número obtenido
                    if numero_cen and opat_page:
                        try:
                            await opat_page.evaluate(f"""
                            (num) => {{
                                var el = document.getElementById('modAvisoCen');
                                if (!el) {{
                                    // Buscar por el aviso CEN en el modal abierto más reciente
                                    var inputs = Array.from(document.querySelectorAll('input'));
                                    for (var i=0; i<inputs.length; i++) {{
                                        if ((inputs[i].placeholder||'').toLowerCase().includes('cen') ||
                                            (inputs[i].id||'').toLowerCase().includes('cen') ||
                                            (inputs[i].name||'').toLowerCase().includes('cen')) {{
                                            el = inputs[i]; break;
                                        }}
                                    }}
                                }}
                                if (el) {{
                                    el.value = num;
                                    el.dispatchEvent(new Event('input', {{bubbles:true}}));
                                    el.dispatchEvent(new Event('change', {{bubbles:true}}));
                                    return 'ok';
                                }}
                                return 'not_found';
                            }}
                            """, numero_cen)
                            print(f"    Aviso CEN {numero_cen} guardado en OPAT")
                        except Exception as e_cen:
                            print(f"    ADVERTENCIA: no se pudo guardar CEN en OPAT: {e_cen}")

                pts_aprobados.append({
                    "id":        pt["id"],
                    "area":      pt["area"],
                    "estado":    pt["estado"],
                    "opat":      opat_ok,
                    "cen":       numero_cen,
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
# NEOMANTE — HELPERS
# =============================================================================

NEOMANTE_TRABAJO_SOBRE_MAP = {
    # Paños
    "int ": "Paños", "alim": "Paños", "alimentador": "Paños",
    "paño": "Paños", "pano": "Paños", "bco": "Paños",
    # Transformador
    "trafo": "Transformador", "transformador": "Transformador",
    "t2d": "Transformador", " at": "Transformador", " ht": "Transformador",
    "autotransf": "Transformador",
    # Secciones de barra
    "barra": "Secciones de barra", "secc": "Secciones de barra",
    "sección": "Secciones de barra", "seccion": "Secciones de barra",
    # Scada
    "scada": "Scada", "utr": "Scada", "control y": "Scada",
    # Compensadores
    "compensador": "Compensadores", "reactor": "Compensadores",
    # Medidores
    "medidor": "Medidores de facturación",
}


def determinar_trabajo_sobre(instalacion_a_intervenir: str, detalle_instalacion: str) -> str:
    """Determina la opción 'Trabajo Sobre' de Neomante según datos del PT."""
    textos = [
        (instalacion_a_intervenir or "").lower(),
        (detalle_instalacion or "").lower(),
    ]
    for texto in textos:
        for keyword, trabajo in NEOMANTE_TRABAJO_SOBRE_MAP.items():
            if keyword in texto:
                return trabajo
    return "Otros equipos"


def extraer_codigo_componente(detalle_instalacion: str) -> str:
    """
    Extrae el código del componente de la última parte del texto.
    Ej: 'INT 12kV BCO 1A' → '1A'
         'T2D AT1' → 'AT1'
         'BARRA 220kV' → '220kV'
    """
    if not detalle_instalacion:
        return ""
    partes = detalle_instalacion.strip().split()
    if partes:
        return partes[-1]  # última palabra
    return detalle_instalacion.strip()


# =============================================================================
# NEOMANTE — LOGIN Y SWITCH EMPRESA
# =============================================================================

async def hacer_login_neomante(neo_page):
    print("\n[NEOMANTE] LOGIN")

    for intento in range(3):
        try:
            await neo_page.goto(NEOMANTE_URL, wait_until="domcontentloaded", timeout=120_000)
            break
        except Exception as e:
            print(f"  goto intento {intento+1} falló: {e}")
            if intento == 2:
                raise
            await neo_page.wait_for_timeout(5000)

    await neo_page.wait_for_timeout(2000)
    await screenshot(neo_page, "neomante_login")

    # ── Seleccionar pestaña "Coordinado" via JS ───────────────────────────────
    # El login tiene 4 pestañas: Centro de Control | Empresa | Coordinado | Coordinador
    # Debemos hacer click en "Coordinado" (exactamente, no "Coordinador")
    r_tab = await neo_page.evaluate("""
    () => {
        // Buscar todos los elementos clickeables que digan exactamente "Coordinado"
        var all = Array.from(document.querySelectorAll('*'));
        for (var i=0; i<all.length; i++) {
            var el = all[i];
            var t = (el.innerText || el.textContent || '').trim();
            if (t === 'Coordinado' && el.children.length === 0) {
                el.click();
                return 'ok:' + el.tagName + ':' + el.className;
            }
        }
        return 'not_found';
    }
    """)
    print(f"  tab Coordinado: {r_tab}")
    await neo_page.wait_for_timeout(2000)
    await screenshot(neo_page, "neomante_pre_fill")

    # ── Email — ID exacto confirmado: email_coordinado ────────────────────────
    r_email = await neo_page.evaluate("""
    (val) => {
        var el = document.getElementById('email_coordinado');
        if (el) {
            el.value = val;
            el.dispatchEvent(new Event('input', {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
            return 'ok_by_id:email_coordinado';
        }
        return 'not_found';
    }
    """, NEOMANTE_USER)
    print(f"  email fill: {r_email}")
    await neo_page.wait_for_timeout(300)

    # ── Contraseña — ID exacto confirmado: password_coordinado ───────────────
    # IMPORTANTE: NO usar password_centro_control (está oculto)
    # El campo visible para Coordinado es password_coordinado
    r_pass = await neo_page.evaluate("""
    (val) => {
        var el = document.getElementById('password_coordinado');
        if (el) {
            el.value = val;
            el.dispatchEvent(new Event('input', {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
            return 'ok_by_id:password_coordinado';
        }
        // Fallback: primer input password visible
        var inputs = Array.from(document.querySelectorAll('input[type="password"]'));
        for (var i=0; i<inputs.length; i++) {
            if (inputs[i].offsetParent) {
                inputs[i].value = val;
                inputs[i].dispatchEvent(new Event('input', {bubbles:true}));
                inputs[i].dispatchEvent(new Event('change', {bubbles:true}));
                return 'ok_visible:' + inputs[i].id;
            }
        }
        return 'not_found';
    }
    """, NEOMANTE_PASS)
    print(f"  pass fill: {r_pass}")
    await neo_page.wait_for_timeout(300)

    # ── Click en Ingresar — botón visible del form Coordinado ─────────────────
    r_btn = await neo_page.evaluate("""
    () => {
        // Buscar el form que contiene email_coordinado
        var emailEl = document.getElementById('email_coordinado');
        if (emailEl) {
            var form = emailEl.closest('form');
            if (form) {
                var btn = form.querySelector('button[type="submit"]');
                if (btn) { btn.click(); return 'ok_form:' + btn.id; }
            }
        }
        // Fallback: botón visible
        var btns = Array.from(document.querySelectorAll('button'));
        for (var i=0; i<btns.length; i++) {
            if ((btns[i].innerText||'').trim()==='Ingresar' && btns[i].offsetParent) {
                btns[i].click();
                return 'ok_visible:' + btns[i].id;
            }
        }
        return 'not_found';
    }
    """)
    print(f"  btn Ingresar: {r_btn}")
    await neo_page.wait_for_load_state("networkidle", timeout=30_000)
    await neo_page.wait_for_timeout(2000)
    await screenshot(neo_page, "neomante_post_login")
    print("  OK: sesión Neomante iniciada")

    # ── Switch a SOCIEDAD TRANSMISORA METROPOLITANA S.A. ─────────────────────
    print("  Haciendo switch a SOCIEDAD TRANSMISORA METROPOLITANA S.A...")
    try:
        # El nombre de empresa aparece en el header — puede tener texto ligeramente distinto
        # Usar JS para encontrar y clickear cualquier elemento que contenga ese texto
        r_switch_menu = await neo_page.evaluate("""
        () => {
            var all = Array.from(document.querySelectorAll('*'));
            for (var i=0; i<all.length; i++) {
                var t = (all[i].innerText || '').trim();
                if (t.includes('TRANSMISIÓN DEL SUR') || t.includes('TRANSMISION DEL SUR')) {
                    all[i].click();
                    return 'ok:' + all[i].tagName + ':' + t.substring(0,50);
                }
            }
            return 'not_found';
        }
        """)
        print(f"  switch menu: {r_switch_menu}")
        await neo_page.wait_for_timeout(2000)
        await screenshot(neo_page, "neomante_switch_menu")

        # Hacer click en la opción SOCIEDAD TRANSMISORA METROPOLITANA S.A.
        r_switch = await neo_page.evaluate("""
        () => {
            var all = Array.from(document.querySelectorAll('a, li, div, span, button'));
            for (var i=0; i<all.length; i++) {
                var t = (all[i].innerText || all[i].textContent || '').trim();
                if (t.includes('TRANSMISORA METROPOLITANA')) {
                    all[i].click();
                    return 'ok:' + all[i].tagName + ':' + t.substring(0,60);
                }
            }
            return 'not_found';
        }
        """)
        print(f"  switch empresa: {r_switch}")
        await neo_page.wait_for_load_state("networkidle", timeout=20_000)
        await neo_page.wait_for_timeout(2000)
        await screenshot(neo_page, "neomante_post_switch")
        print("  OK: switch a SOCIEDAD TRANSMISORA METROPOLITANA S.A.")
    except Exception as e:
        print(f"  ADVERTENCIA switch: {e}")


# =============================================================================
# NEOMANTE — CREAR AVISO CEN
# =============================================================================

async def crear_aviso_cen(neo_page, datos):
    """
    Crea el aviso al CEN en Neomante con los datos del PT.
    Devuelve el número de aviso (ej: '2026061675') o None si falló.
    """
    pt_id = datos["id"]
    print(f"\n[NEOMANTE] Creando aviso CEN para PT {pt_id}...")

    try:
        # Ir al home de Neomante
        await neo_page.goto("https://neomante.coordinador.cl/", wait_until="domcontentloaded", timeout=60_000)
        await neo_page.wait_for_timeout(2000)

        # Click en Desconexión/Intervención → Subestación
        await neo_page.click('text="Subestación"', timeout=TIMEOUT)
        await neo_page.wait_for_load_state("networkidle", timeout=20_000)
        await neo_page.wait_for_timeout(1500)
        await screenshot(neo_page, f"neo_01_tipo_{pt_id}")

        # ── PASO 1: Tipo de Solicitud ──────────────────────────────────────────
        tipo_solicitud = datos.get("tipo_trabajo", "DESCONEXIÓN")
        if "INTERV" in tipo_solicitud.upper():
            await neo_page.click('text="Intervención"', timeout=10_000)
        else:
            await neo_page.click('text="Desconexión"', timeout=10_000)
        await neo_page.wait_for_timeout(500)
        print(f"  Tipo solicitud: {tipo_solicitud}")

        # Origen Interno
        await neo_page.click('text="Origen Interno"', timeout=10_000)
        await neo_page.wait_for_timeout(500)

        # Programada
        await neo_page.click('text="Programada"', timeout=10_000)
        await neo_page.wait_for_timeout(500)

        # Siguiente
        await neo_page.click('button:has-text("Siguiente")', timeout=10_000)
        await neo_page.wait_for_load_state("networkidle", timeout=20_000)
        await neo_page.wait_for_timeout(1500)
        await screenshot(neo_page, f"neo_02_subestacion_{pt_id}")

        # ── PASO 2: Subestación ───────────────────────────────────────────────
        # Buscar la subestación por el nombre del Elemento referencia
        ssee_raw = datos.get("se_linea", "")
        # Limpiar el nombre: quitar "(Subestación)" etc.
        ssee_nombre = re.sub(r'\s*\(.*?\)', '', ssee_raw).strip().upper()
        print(f"  Buscando subestación: {ssee_nombre}")

        # Seleccionar "Mis subestaciones" primero, luego "Todas las subestaciones" si no aparece
        seleccionada = False
        for opcion in ["Mis subestaciones", "Todas las subestaciones"]:
            try:
                await neo_page.click(f'text="{opcion}"', timeout=5000)
                await neo_page.wait_for_timeout(1000)

                # Buscar en el select
                resultado = await neo_page.evaluate(f"""
                (nombre) => {{
                    var selects = Array.from(document.querySelectorAll('select'));
                    for (var i=0; i<selects.length; i++) {{
                        var opts = Array.from(selects[i].options);
                        // Buscar coincidencia exacta primero
                        for (var j=0; j<opts.length; j++) {{
                            if (opts[j].text.toUpperCase().includes(nombre)) {{
                                selects[i].selectedIndex = j;
                                selects[i].dispatchEvent(new Event('change', {{bubbles:true}}));
                                return 'ok:' + opts[j].text;
                            }}
                        }}
                    }}
                    return 'not_found';
                }}
                """, ssee_nombre)
                print(f"  Subestación ({opcion}): {resultado}")

                if "ok:" in resultado:
                    seleccionada = True
                    break
            except Exception:
                continue

        if not seleccionada:
            print(f"  ADVERTENCIA: subestación no encontrada, seleccionando primera disponible")
            await neo_page.evaluate("""
            () => {
                var s = document.querySelector('select');
                if (s && s.options.length > 1) {
                    s.selectedIndex = 1;
                    s.dispatchEvent(new Event('change', {bubbles:true}));
                }
            }
            """)

        await neo_page.click('button:has-text("Siguiente")', timeout=10_000)
        await neo_page.wait_for_load_state("networkidle", timeout=20_000)
        await neo_page.wait_for_timeout(1500)
        await screenshot(neo_page, f"neo_03_trabajo_sobre_{pt_id}")

        # ── PASO 3: Trabajo Sobre ─────────────────────────────────────────────
        instalacion = datos.get("raw", {}).get("Instalación a intervenir", "")
        detalle     = datos.get("componentes", "")
        trabajo_sobre = determinar_trabajo_sobre(instalacion, detalle)
        print(f"  Trabajo sobre: {trabajo_sobre} (instalacion='{instalacion}', detalle='{detalle}')")

        await neo_page.click(f'text="{trabajo_sobre}"', timeout=10_000)
        await neo_page.wait_for_timeout(500)

        await neo_page.click('button:has-text("Siguiente")', timeout=10_000)
        await neo_page.wait_for_load_state("networkidle", timeout=20_000)
        await neo_page.wait_for_timeout(1500)
        await screenshot(neo_page, f"neo_04_elementos_{pt_id}")

        # ── PASO 4: Seleccionar Elementos ─────────────────────────────────────
        codigo = extraer_codigo_componente(detalle)
        print(f"  Buscando elemento con código: '{codigo}'")

        elemento_marcado = await neo_page.evaluate(f"""
        (codigo) => {{
            var checkboxes = Array.from(document.querySelectorAll('input[type="checkbox"]'));
            var labels = Array.from(document.querySelectorAll('label'));
            var items = [];

            // Construir lista de (checkbox, texto)
            for (var i=0; i<labels.length; i++) {{
                var lbl = labels[i];
                var txt = (lbl.innerText || '').trim().toUpperCase();
                if (txt.length < 3) continue;
                var cb = null;
                if (lbl.htmlFor) {{
                    cb = document.getElementById(lbl.htmlFor);
                }} else {{
                    cb = lbl.querySelector('input[type="checkbox"]');
                }}
                if (!cb) continue;
                items.push({{cb: cb, txt: txt}});
            }}

            // También checkboxes directos
            for (var j=0; j<checkboxes.length; j++) {{
                var sibling = checkboxes[j].nextSibling;
                var txt2 = sibling ? (sibling.textContent || '').trim().toUpperCase() : '';
                if (txt2.length > 2) items.push({{cb: checkboxes[j], txt: txt2}});
            }}

            if (items.length === 0) return 'no_items';

            var codigoUpper = (codigo || '').toUpperCase();

            // Buscar por coincidencia del código
            if (codigoUpper.length > 0) {{
                for (var k=0; k<items.length; k++) {{
                    if (items[k].txt.includes(codigoUpper)) {{
                        items[k].cb.checked = true;
                        items[k].cb.dispatchEvent(new Event('change', {{bubbles:true}}));
                        return 'ok_match:' + items[k].txt;
                    }}
                }}
            }}

            // Fallback: primer elemento disponible
            items[0].cb.checked = true;
            items[0].cb.dispatchEvent(new Event('change', {{bubbles:true}}));
            return 'ok_first:' + items[0].txt;
        }}
        """, codigo)
        print(f"  Elemento seleccionado: {elemento_marcado}")

        await neo_page.click('button:has-text("Siguiente")', timeout=10_000)
        await neo_page.wait_for_load_state("networkidle", timeout=20_000)
        await neo_page.wait_for_timeout(1500)
        await screenshot(neo_page, f"neo_05_riesgo_{pt_id}")

        # ── PASO 5: Riesgo del Trabajo ────────────────────────────────────────
        riesgo_ta = await neo_page.query_selector('textarea')
        if riesgo_ta:
            await riesgo_ta.fill("Trabajo sin riesgo para el sistema.")
        await neo_page.click('button:has-text("Siguiente")', timeout=10_000)
        await neo_page.wait_for_load_state("networkidle", timeout=20_000)
        await neo_page.wait_for_timeout(1000)

        # ── PASO 6: Consumo → "No tiene consumo afectado" ─────────────────────
        await neo_page.click('text="No tiene consumo afectado"', timeout=10_000)
        await neo_page.wait_for_timeout(500)
        await neo_page.click('button:has-text("Siguiente")', timeout=10_000)
        await neo_page.wait_for_load_state("networkidle", timeout=20_000)
        await neo_page.wait_for_timeout(1000)
        await screenshot(neo_page, f"neo_06_tipo_trabajo_{pt_id}")

        # ── PASO 7: Tipo Trabajo ──────────────────────────────────────────────
        # Seleccionar "Otro tipo de trabajo" en el dropdown
        resultado_tipo = await neo_page.evaluate("""
        () => {
            var selects = Array.from(document.querySelectorAll('select'));
            for (var i=0; i<selects.length; i++) {
                var opts = Array.from(selects[i].options);
                for (var j=0; j<opts.length; j++) {
                    if (opts[j].text.toLowerCase().includes('otro tipo')) {
                        selects[i].selectedIndex = j;
                        selects[i].dispatchEvent(new Event('change', {bubbles:true}));
                        return 'ok:' + opts[j].text;
                    }
                }
            }
            return 'not_found';
        }
        """)
        print(f"  Tipo trabajo: {resultado_tipo}")

        # Describir el trabajo
        desc = datos.get("descripcion", "")
        textareas = await neo_page.query_selector_all('textarea')
        if len(textareas) >= 1:
            await textareas[0].fill(desc[:500])  # limitar a 500 chars

        await neo_page.click('button:has-text("Siguiente")', timeout=10_000)
        await neo_page.wait_for_load_state("networkidle", timeout=20_000)
        await neo_page.wait_for_timeout(1000)

        # ── PASO 8: Trabajo Afecta → solo Siguiente ───────────────────────────
        await neo_page.click('button:has-text("Siguiente")', timeout=10_000)
        await neo_page.wait_for_load_state("networkidle", timeout=20_000)
        await neo_page.wait_for_timeout(1000)

        # ── PASO 9: Comentario Adicional ──────────────────────────────────────
        textarea_comentario = await neo_page.query_selector('textarea')
        if textarea_comentario:
            await textarea_comentario.fill(
                "Trabajos enmarcados en el plan de mantenimiento anual de STM"
            )
        await neo_page.click('button:has-text("Siguiente")', timeout=10_000)
        await neo_page.wait_for_load_state("networkidle", timeout=20_000)
        await neo_page.wait_for_timeout(1500)
        await screenshot(neo_page, f"neo_07_fecha_{pt_id}")

        # ── PASO 10: Fecha / Hora del Trabajo ─────────────────────────────────
        # Seleccionar "Ninguno de los antecedentes anteriores"
        await neo_page.click('text="Ninguno de los antecedentes anteriores"', timeout=10_000)
        await neo_page.wait_for_timeout(500)

        # Llenar fecha y hora de inicio
        if datos.get("fecha_inicio") and datos.get("hora_inicio"):
            dt_inicio = datetime.strptime(
                datos["fecha_inicio"] + " " + datos["hora_inicio"],
                "%m/%d/%Y %I:%M %p"
            )
            fecha_hora_inicio = dt_inicio.strftime("%d-%m-%Y %H:%M")
            # El campo es un input de texto con formato dd-mm-yyyy HH:MM
            await neo_page.evaluate(f"""
            (val) => {{
                var inputs = Array.from(document.querySelectorAll('input'));
                // Buscar inputs de fecha/hora (los que tienen el ícono de reloj)
                for (var i=0; i<inputs.length; i++) {{
                    var inp = inputs[i];
                    if (!inp.offsetParent) continue;
                    var parent = inp.closest('div,td,tr');
                    if (!parent) continue;
                    var parentTxt = (parent.innerText || '').toLowerCase();
                    if (parentTxt.includes('inicio') || parentTxt.includes('start')) {{
                        inp.value = val;
                        inp.dispatchEvent(new Event('input', {{bubbles:true}}));
                        inp.dispatchEvent(new Event('change', {{bubbles:true}}));
                        return 'ok:' + inp.id;
                    }}
                }}
                // Fallback: primer input visible
                for (var j=0; j<inputs.length; j++) {{
                    if (inputs[j].offsetParent && inputs[j].type !== 'hidden') {{
                        inputs[j].value = val;
                        inputs[j].dispatchEvent(new Event('change', {{bubbles:true}}));
                        return 'fallback:' + inputs[j].id;
                    }}
                }}
                return 'not_found';
            }}
            """, fecha_hora_inicio)
            print(f"  Fecha/hora inicio: {fecha_hora_inicio}")

            # Aplicar con el calendario
            try:
                await neo_page.click('button:has-text("Aplicar")', timeout=5000)
                await neo_page.wait_for_timeout(500)
            except Exception:
                pass

        # Llenar fecha y hora de fin
        if datos.get("fecha_fin") and datos.get("hora_fin"):
            dt_fin = datetime.strptime(
                datos["fecha_fin"] + " " + datos["hora_fin"],
                "%m/%d/%Y %I:%M %p"
            )
            fecha_hora_fin = dt_fin.strftime("%d-%m-%Y %H:%M")
            await neo_page.evaluate(f"""
            (val) => {{
                var inputs = Array.from(document.querySelectorAll('input'));
                for (var i=0; i<inputs.length; i++) {{
                    var inp = inputs[i];
                    if (!inp.offsetParent) continue;
                    var parent = inp.closest('div,td,tr');
                    if (!parent) continue;
                    var parentTxt = (parent.innerText || '').toLowerCase();
                    if (parentTxt.includes('fin') || parentTxt.includes('término') || parentTxt.includes('termino')) {{
                        inp.value = val;
                        inp.dispatchEvent(new Event('input', {{bubbles:true}}));
                        inp.dispatchEvent(new Event('change', {{bubbles:true}}));
                        return 'ok:' + inp.id;
                    }}
                }}
                // Fallback: segundo input visible
                var visibles = Array.from(document.querySelectorAll('input')).filter(function(i) {{
                    return i.offsetParent && i.type !== 'hidden';
                }});
                if (visibles.length >= 2) {{
                    visibles[1].value = val;
                    visibles[1].dispatchEvent(new Event('change', {{bubbles:true}}));
                    return 'fallback2:' + visibles[1].id;
                }}
                return 'not_found';
            }}
            """, fecha_hora_fin)
            print(f"  Fecha/hora fin: {fecha_hora_fin}")

            try:
                await neo_page.click('button:has-text("Aplicar")', timeout=5000)
                await neo_page.wait_for_timeout(500)
            except Exception:
                pass

        await screenshot(neo_page, f"neo_08_pre_enviar_{pt_id}")

        # ── PASO 11: Crear y Enviar al Coordinador ────────────────────────────
        await neo_page.click('button:has-text("Crear y Enviar al Coordinador")', timeout=10_000)
        await neo_page.wait_for_timeout(1500)

        # Confirmar popup "¿Está seguro de enviar la Solicitud?"
        await neo_page.click('button:has-text("Aceptar")', timeout=10_000)
        await neo_page.wait_for_load_state("networkidle", timeout=30_000)
        await neo_page.wait_for_timeout(2000)
        await screenshot(neo_page, f"neo_09_resultado_{pt_id}")

        # ── PASO 12: Capturar número de aviso ────────────────────────────────
        numero_aviso = await neo_page.evaluate("""
        () => {
            // Buscar el mensaje de éxito: "Solicitud creada con exito, Número: XXXXXXXXXX"
            var els = Array.from(document.querySelectorAll('*'));
            for (var i=0; i<els.length; i++) {
                var txt = (els[i].innerText || '').trim();
                if (txt.includes('Solicitud creada') && txt.includes('Número:')) {
                    var match = txt.match(/Número:\s*(\d+)/);
                    if (match) return match[1];
                }
                // También buscar solo el número en el banner de éxito
                if (txt.includes('creada con exito')) {
                    var match2 = txt.match(/(\d{10,})/);
                    if (match2) return match2[1];
                }
            }
            return null;
        }
        """)

        if numero_aviso:
            print(f"  ✓ Aviso CEN creado: {numero_aviso}")
            return numero_aviso
        else:
            print(f"  ADVERTENCIA: no se pudo capturar el número de aviso")
            return None

    except Exception as e:
        print(f"  ERROR creando aviso CEN: {e}")
        await screenshot(neo_page, f"neo_error_{pt_id}")
        return None


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
            return "<tr><td colspan='5' style='padding:6px 12px;color:#999'>Ninguno</td></tr>"
        html = ""
        for pt in pts_aprobados:
            opat_dry = pt.get("opat_dry", False)
            if opat_dry:
                opat_icon = "<span style='color:#888'>— (simulado)</span>"
            elif pt.get("opat"):
                opat_icon = "<span style='color:#006600;font-size:16px'>&#10003;</span>"
            else:
                opat_icon = "<span style='color:#cc0000;font-size:16px'>&#10007;</span>"

            cen = pt.get("cen")
            if opat_dry:
                cen_cell = "<span style='color:#888'>— (simulado)</span>"
            elif cen:
                cen_cell = f"<span style='font-family:monospace;color:#006600;font-weight:bold'>{cen}</span>"
            else:
                cen_cell = "<span style='color:#cc0000;font-size:12px'>No se pudo realizar aviso al CEN,<br>realizarlo manualmente</span>"

            html += (
                "<tr>"
                "<td style='padding:4px 8px;color:#006600;font-size:16px'>&#10003;</td>"
                f"<td style='font-family:monospace;padding:4px 12px'>{pt.get('id','')}</td>"
                f"<td style='padding:4px 12px'>{pt.get('area','')}</td>"
                f"<td style='padding:4px 12px;text-align:center'>{opat_icon}</td>"
                f"<td style='padding:4px 12px;text-align:center'>{cen_cell}</td>"
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
    cen_ok_count  = sum(1 for pt in pts_aprobados if pt.get("cen"))

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
        f"<span style='font-size:13px;color:#555'>— OPAT: {opat_ok_count} &nbsp;|&nbsp; Avisos CEN: {cen_ok_count}</span></h3>"
        "<table style='border-collapse:collapse;width:100%'>"
        "<tr style='background:#f6f6f6'>"
        "<th style='padding:6px 8px'>✓</th>"
        "<th style='padding:6px 12px;text-align:left'>PT</th>"
        "<th style='padding:6px 12px;text-align:left'>Área</th>"
        "<th style='padding:6px 12px;text-align:center'>OPAT</th>"
        "<th style='padding:6px 12px;text-align:center'>Aviso CEN</th>"
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
            f" | {cen_ok_count} avisos CEN"
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
        ctx_neomante = await browser.new_context(
            ignore_https_errors=True,
            viewport={"width": 1400, "height": 900},
        )

        page_centrality = await ctx_centrality.new_page()
        page_opat       = await ctx_opat.new_page()
        page_neomante   = await ctx_neomante.new_page()

        try:
            # Login en los tres sistemas en paralelo
            await asyncio.gather(
                hacer_login_centrality(page_centrality),
                hacer_login_opat(page_opat),
                hacer_login_neomante(page_neomante),
            )

            frame = await navegar_a_permisos(page_centrality)
            await aplicar_filtro_pcct(page_centrality, frame)
            pts_aprobados, pts_fallidos, pts_omitidos = await aprobar_pts(
                page_centrality, frame, page_opat, page_neomante
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
