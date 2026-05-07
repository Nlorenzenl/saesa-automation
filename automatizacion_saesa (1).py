"""
Automatización SAESA – Autorización diaria de PT's (Permisos de Trabajo)
Flujo: Login → Aplicaciones → DMS → Planificación → Permisos de trabajo
       → Filtro (Zonales / Área Zonal Metropolitana / Revisión y Autorización PCCT)
       → Aprobar uno a uno → Reporte por correo
"""

import asyncio
import smtplib
import os
import re
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN  (variables de entorno – definidas como Secrets en GitHub)
# ─────────────────────────────────────────────────────────────────────────────
SAESA_URL   = "https://stx.saesa.cl:8091/backend/sts/login.php?backurl=%2Fbackend%2Fsts%2Fcentrality.php"
SAESA_USER  = os.environ["SAESA_USER"]
SAESA_PASS  = os.environ["SAESA_PASS"]
GMAIL_USER  = os.environ["GMAIL_USER"]
GMAIL_PASS  = os.environ["GMAIL_APP_PASS"]
EMAIL_DEST  = os.environ["EMAIL_DEST"]

TIMEOUT     = 30_000   # ms – tiempo máximo de espera por elemento
AREA_ZONAL  = "Area Zonal Metropolitana"
ESTADO_FILTRO = "Revisión y Autorización PCCT"


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
async def screenshot(page, nombre):
    """Guarda captura de pantalla para diagnóstico (se sube como artefacto en CI)."""
    os.makedirs("capturas", exist_ok=True)
    path = f"capturas/{nombre}_{datetime.now().strftime('%H%M%S')}.png"
    await page.screenshot(path=path, full_page=False)
    print(f"  📸 Captura: {path}")


async def esperar_y_click(page, selector, descripcion, timeout=TIMEOUT):
    print(f"  → Esperando: {descripcion}...")
    await page.wait_for_selector(selector, timeout=timeout)
    await page.click(selector)
    print(f"  ✓ Clic en: {descripcion}")


# ─────────────────────────────────────────────────────────────────────────────
# PASO 1 – LOGIN
# ─────────────────────────────────────────────────────────────────────────────
async def hacer_login(page):
    print("\n[1] LOGIN")
    await page.goto(SAESA_URL, wait_until="networkidle", timeout=60_000)
    await page.wait_for_selector('input[name="usuario"]', timeout=TIMEOUT)
    await page.fill('input[name="usuario"]', SAESA_USER)
    await page.fill('input[name="password"]', SAESA_PASS)
    await page.click('input[type="submit"], button[type="submit"], input[value="Login"], button:has-text("Login")')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    print("  ✓ Login exitoso")
    await screenshot(page, "01_login_ok")


# ─────────────────────────────────────────────────────────────────────────────
# PASO 2 – NAVEGAR A DMS → PLANIFICACIÓN → PERMISOS DE TRABAJO
# ─────────────────────────────────────────────────────────────────────────────
async def navegar_a_permisos(page):
    print("\n[2] NAVEGACIÓN → DMS → Permisos de trabajo")

    # Clic en "Aplicaciones" en el menú lateral izquierdo
    await esperar_y_click(page, 'a:has-text("Aplicaciones"), span:has-text("Aplicaciones")', "Aplicaciones")
    await page.wait_for_timeout(1500)

    # Clic en "DMS"
    await esperar_y_click(page, 'a:has-text("DMS")', "DMS")
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2000)
    await screenshot(page, "02_dms")

    # El DMS carga en un iframe – necesitamos trabajar dentro de él
    # Esperar a que aparezca el menú "Planificación" dentro del frame
    frame = None
    for f in page.frames:
        try:
            await f.wait_for_selector('text=Planificación', timeout=5000)
            frame = f
            break
        except PlaywrightTimeout:
            continue

    if frame is None:
        # Si no hay iframe, operar en la página principal
        frame = page

    # Clic en menú "Planificación"
    await frame.click('text=Planificación')
    await page.wait_for_timeout(1000)

    # Clic en "Permisos de trabajo" del dropdown
    await frame.click('text=Permisos de trabajo')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2000)
    await screenshot(page, "03_permisos_trabajo")
    print("  ✓ En pantalla Permisos de trabajo")

    return frame


# ─────────────────────────────────────────────────────────────────────────────
# PASO 3 – APLICAR FILTRO
# ─────────────────────────────────────────────────────────────────────────────
async def aplicar_filtro(page, frame):
    print("\n[3] FILTRO")

    # Clic en botón "Filtro"
    await frame.click('text=Filtro')
    await page.wait_for_timeout(1500)
    await screenshot(page, "04_filtro_abierto")

    # ── Campo "Áreas": seleccionar "Zonales" ──────────────────────────────────
    # El select de Áreas está junto a la etiqueta "Áreas:"
    areas_select = await frame.query_selector('select')   # puede haber varios selects
    # Buscar el select correcto por label cercano
    # Estrategia: buscar todos los <select> y elegir el que tiene opción "Zonales"
    selects = await frame.query_selector_all('select')
    for sel in selects:
        options = await sel.inner_text()
        if 'Zonales' in options:
            await sel.select_option(label='Zonales')
            print("  ✓ Áreas → Zonales")
            break
    await page.wait_for_timeout(1000)

    # ── Checkbox del área zonal: clic en el ícono de lista a la derecha ───────
    # El botón es el pequeño checkbox/ícono a la derecha del select de Zonales
    # Buscar el botón que abre el popup "Editar lista"
    try:
        # Intentar clic en el botón inmediato a la derecha (normalmente un <input type="button"> o un pequeño ícono)
        await frame.click('input[type="checkbox"][title*="área"], input.area-check, span.area-selector, td:has-text("Zonales") ~ td input', timeout=5000)
    except PlaywrightTimeout:
        # Fallback: buscar botón que abre el diálogo de lista
        btns = await frame.query_selector_all('input[type="button"], button')
        for btn in btns:
            txt = await btn.get_attribute('value') or await btn.inner_text()
            if '...' in txt or 'lista' in txt.lower() or txt.strip() == '':
                await btn.click()
                break
    await page.wait_for_timeout(1500)

    # ── En el popup "Editar lista": desmarcar "Todos", marcar "Area Zonal Metropolitana" ──
    # Primero desmarcar "Todos"
    try:
        todos_cb = await frame.query_selector('input[type="checkbox"] + text=Todos, label:has-text("Todos") input')
        if todos_cb:
            if await todos_cb.is_checked():
                await todos_cb.click()
        await page.wait_for_timeout(500)
    except Exception:
        pass

    # Marcar "Area Zonal Metropolitana"
    checkboxes = await frame.query_selector_all('input[type="checkbox"]')
    for cb in checkboxes:
        # Buscar el label o texto cercano
        parent = await cb.evaluate_handle('el => el.parentElement')
        texto = await parent.inner_text() if parent else ''
        if AREA_ZONAL.lower() in texto.lower() or 'Metropolitana' in texto:
            if not await cb.is_checked():
                await cb.click()
            print(f"  ✓ Seleccionada: {texto.strip()}")
            break

    # Clic en "Aceptar" del popup
    await frame.click('button:has-text("Aceptar"), input[value="Aceptar"]')
    await page.wait_for_timeout(1000)

    # ── Campo "Estado": seleccionar "Revisión y Autorización PCCT" ───────────
    estado_selects = await frame.query_selector_all('select')
    for sel in estado_selects:
        options = await sel.inner_text()
        if 'PCCT' in options or 'Revisión' in options:
            await sel.select_option(label=ESTADO_FILTRO)
            print(f"  ✓ Estado → {ESTADO_FILTRO}")
            break
    await page.wait_for_timeout(500)

    # ── Clic en "Aplicar" ────────────────────────────────────────────────────
    await frame.click('button:has-text("Aplicar"), input[value="Aplicar"], a:has-text("Aplicar")')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2000)
    await screenshot(page, "05_filtro_aplicado")
    print("  ✓ Filtro aplicado")


# ─────────────────────────────────────────────────────────────────────────────
# PASO 4 – LEER LISTA DE PT's
# ─────────────────────────────────────────────────────────────────────────────
async def obtener_pts(frame):
    """Retorna lista de dicts con {id, descripcion} de los PT's visibles."""
    pts = []
    try:
        filas = await frame.query_selector_all('table tbody tr, tr[class*="row"], tr[class*="fila"]')
        for fila in filas:
            celdas = await fila.query_selector_all('td')
            if len(celdas) >= 2:
                id_pt = (await celdas[0].inner_text()).strip()
                desc  = (await celdas[-1].inner_text()).strip()
                if id_pt and re.match(r'\d{4}-\d{5}', id_pt):
                    pts.append({"id": id_pt, "descripcion": desc[:80]})
    except Exception as e:
        print(f"  ⚠️ No se pudo leer la lista: {e}")
    return pts


# ─────────────────────────────────────────────────────────────────────────────
# PASO 5 – APROBAR PT's UNO A UNO
# ─────────────────────────────────────────────────────────────────────────────
async def aprobar_pts(page, frame):
    print("\n[4] APROBACIÓN DE PT's")
    pts_aprobados = []
    pts_fallidos  = []
    iteracion = 0
    max_iter = 100  # seguridad para evitar loop infinito

    while iteracion < max_iter:
        iteracion += 1

        # Leer la primera fila de la tabla (siempre seleccionamos la primera disponible)
        try:
            primera_fila = await frame.query_selector('table tbody tr:first-child td, tr[class*="row"]:first-child td')
            if primera_fila is None:
                print("  ✓ No hay más PT's en la lista. Proceso completado.")
                break

            # Obtener ID del primer PT
            primera_celda = await frame.query_selector('table tbody tr:first-child td:first-child')
            id_pt = (await primera_celda.inner_text()).strip() if primera_celda else f"PT-{iteracion}"

            # Si la celda no parece un ID válido, terminamos
            if not re.match(r'\d{4}-\d{5}', id_pt):
                print(f"  ✓ Lista vacía o sin PT's válidos (celda: '{id_pt}'). Proceso completado.")
                break

            print(f"  → Aprobando PT {id_pt} ({iteracion})...")

            # Seleccionar la primera fila haciendo clic en ella
            await frame.click('table tbody tr:first-child')
            await page.wait_for_timeout(800)

            # Clic en botón "Aprobar" de la barra de herramientas
            await frame.click('a:has-text("Aprobar"), button:has-text("Aprobar"), input[value="Aprobar"]')
            await page.wait_for_timeout(1500)

            # Aparece popup "Aprobar" con campo de comentarios opcional
            # Clic directo en "Aceptar" (sin agregar comentario)
            await frame.click('button:has-text("Aceptar"), input[value="Aceptar"]')
            await page.wait_for_timeout(2000)

            # Esperar a que la fila desaparezca (el PT aprobado se va de la lista)
            await page.wait_for_load_state("networkidle", timeout=15_000)
            await page.wait_for_timeout(1500)

            pts_aprobados.append(id_pt)
            print(f"  ✓ PT {id_pt} aprobado")

        except PlaywrightTimeout:
            print(f"  ⚠️ Timeout en iteración {iteracion}, la lista podría estar vacía.")
            break
        except Exception as e:
            pts_fallidos.append(f"PT-{iteracion} (error: {str(e)[:60]})")
            print(f"  ✗ Error en iteración {iteracion}: {e}")
            await screenshot(page, f"error_{iteracion}")
            # Intentar continuar con el siguiente
            await page.wait_for_timeout(2000)

    await screenshot(page, "06_final")
    return pts_aprobados, pts_fallidos


# ─────────────────────────────────────────────────────────────────────────────
# PASO 6 – ENVIAR REPORTE POR CORREO
# ─────────────────────────────────────────────────────────────────────────────
def enviar_reporte(pts_aprobados, pts_fallidos, error_critico=None):
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    total_ok  = len(pts_aprobados)
    total_err = len(pts_fallidos)

    lista_ok  = "".join(f"<tr><td style='padding:4px 12px;'>✅</td><td style='padding:4px 12px;font-family:monospace'>{pt}</td></tr>"
                        for pt in pts_aprobados) or \
                "<tr><td colspan='2' style='padding:4px 12px;color:#888'>Ninguno</td></tr>"
    lista_err = "".join(f"<tr><td style='padding:4px 12px;'>❌</td><td style='padding:4px 12px;'>{pt}</td></tr>"
                        for pt in pts_fallidos) or \
                "<tr><td colspan='2' style='padding:4px 12px;color:#888'>Sin errores</td></tr>"

    error_bloque = ""
    if error_critico:
        error_bloque = f"""
        <div style='background:#fff0f0;border-left:4px solid #c00;padding:12px;margin:16px 0;border-radius:4px'>
          <strong>⚠️ Error crítico:</strong><br><code>{error_critico}</code>
        </div>"""

    html = f"""
    <html><body style="font-family:Arial,sans-serif;max-width:640px;margin:auto;color:#222">
      <div style="background:#003580;color:white;padding:20px;border-radius:8px 8px 0 0">
        <h2 style="margin:0">📋 Reporte Diario de PT's – SAESA/DMS</h2>
        <p style="margin:6px 0 0;opacity:0.8">Autorización PCCT – Área Zonal Metropolitana</p>
      </div>
      <div style="border:1px solid #ddd;border-top:none;padding:20px;border-radius:0 0 8px 8px">
        <p><strong>Fecha de ejecución:</strong> {fecha}</p>
        <p><strong>Filtro aplicado:</strong> Zonales → Área Zonal Metropolitana | Estado: Revisión y Autorización PCCT</p>
        {error_bloque}
        <h3 style="color:#006600">✅ PT's Aprobados ({total_ok})</h3>
        <table style="border-collapse:collapse;width:100%">{lista_ok}</table>

        <h3 style="color:#cc0000;margin-top:24px">❌ PT's con Error ({total_err})</h3>
        <table style="border-collapse:collapse;width:100%">{lista_err}</table>

        <hr style="margin:24px 0;border:none;border-top:1px solid #eee">
        <p style="color:#999;font-size:12px">
          Enviado automáticamente por el bot SAESA – GitHub Actions<br>
          Ejecución programada: Lunes a Viernes 08:00 hrs (Chile)
        </p>
      </div>
    </body></html>
    """

    asunto = f"[SAESA] PT's del {datetime.now().strftime('%d/%m/%Y')} – {total_ok} aprobados"
    if error_critico:
        asunto = f"[SAESA] ⚠️ ERROR – {datetime.now().strftime('%d/%m/%Y')}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"]    = GMAIL_USER
    msg["To"]      = EMAIL_DEST
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_USER, GMAIL_PASS)
            server.sendmail(GMAIL_USER, EMAIL_DEST, msg.as_string())
        print("\n  ✓ Reporte enviado a:", EMAIL_DEST)
    except Exception as e:
        print(f"\n  ✗ Error al enviar correo: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
async def main():
    print(f"\n{'='*60}")
    print(f"  SAESA – Autorización PT's  |  {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"{'='*60}")

    pts_aprobados = []
    pts_fallidos  = []
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
        context = await browser.new_context(
            ignore_https_errors=True,
            viewport={"width": 1400, "height": 900}
        )
        page = await context.new_page()

        try:
            await hacer_login(page)
            frame = await navegar_a_permisos(page)
            await aplicar_filtro(page, frame)
            pts_aprobados, pts_fallidos = await aprobar_pts(page, frame)

        except Exception as e:
            error_critico = str(e)
            print(f"\n✗ ERROR CRÍTICO: {e}")
            await screenshot(page, "error_critico")

        finally:
            await browser.close()

    print(f"\n{'─'*60}")
    print(f"  Resultado: {len(pts_aprobados)} aprobados | {len(pts_fallidos)} errores")
    print(f"{'─'*60}")

    enviar_reporte(pts_aprobados, pts_fallidos, error_critico)
    print("\n✓ Proceso finalizado.\n")


if __name__ == "__main__":
    asyncio.run(main())
