"""
Automatización SAESA – Autorización diaria de PT's (Permisos de Trabajo)
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


async def screenshot(page, nombre):
    os.makedirs("capturas", exist_ok=True)
    path = f"capturas/{nombre}_{datetime.now().strftime('%H%M%S')}.png"
    await page.screenshot(path=path, full_page=False)
    print(f"  📸 {path}")


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

    frame = page
    for f in page.frames:
        try:
            await f.wait_for_selector('text=Planificación', timeout=5000)
            frame = f
            print("  → iframe detectado")
            break
        except PlaywrightTimeout:
            continue

    await frame.click('text=Planificación')
    await page.wait_for_timeout(1000)
    await frame.click('text=Permisos de trabajo')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    await screenshot(page, "03_permisos_trabajo")
    print("  ✓ En Permisos de trabajo")
    return frame


async def aplicar_filtro(page, frame):
    print("\n[3] FILTRO")
    await frame.click('text=Filtro')
    await page.wait_for_timeout(2500)
    await screenshot(page, "04_filtro_abierto")

    # ── Diagnóstico completo del DOM del popup ────────────────────────────────
    # El popup usa widgets custom (divs/tables), no <select> estándar
    # Vamos a inspeccionar todo el contenido del popup
    popup_html = await frame.evaluate('''() => {
        // Buscar el popup del filtro por su título
        const titles = Array.from(document.querySelectorAll('*')).filter(
            el => el.innerText && el.innerText.trim() === 'Filtros'
        );
        if (titles.length > 0) {
            const popup = titles[0].closest('div, table, form') || titles[0].parentElement;
            return popup ? popup.outerHTML.substring(0, 3000) : 'no popup encontrado';
        }
        return 'titulo Filtros no encontrado';
    }''')
    print(f"  → HTML popup (primeros 3000 chars):\n{popup_html[:2000]}")

    # ── Buscar todos los elementos interactivos dentro del popup ─────────────
    elementos = await frame.evaluate('''() => {
        const result = [];
        // Buscar elementos con texto "Zonales", "Áreas", "Estado", etc.
        const all = Array.from(document.querySelectorAll('select, option, input, div[class], span[class], td'));
        for (const el of all) {
            const text = (el.innerText || el.textContent || el.value || '').trim();
            if (text && (
                text.includes('Zonales') || text.includes('Áreas') || text.includes('Estado') ||
                text.includes('PCCT') || text.includes('Revisión') || text.includes('Metropolitana')
            )) {
                const rect = el.getBoundingClientRect();
                result.push({
                    tag: el.tagName,
                    cls: el.className || '',
                    text: text.substring(0, 60),
                    x: Math.round(rect.x),
                    y: Math.round(rect.y),
                    w: Math.round(rect.width),
                    h: Math.round(rect.height)
                });
            }
        }
        return result.slice(0, 30);
    }''')
    print(f"  → Elementos relevantes encontrados: {len(elementos)}")
    for el in elementos:
        print(f"    {el['tag']} cls='{el['cls'][:30]}' text='{el['text']}' pos=({el['x']},{el['y']}) size={el['w']}x{el['h']}")

    # ── Buscar el elemento "Áreas" y el select/widget asociado ───────────────
    # Intentar hacer clic en el widget de Áreas usando evaluación JS directa
    seleccionado_zonales = await frame.evaluate('''() => {
        // Buscar todos los <select> incluyendo los ocultos
        const selects = Array.from(document.querySelectorAll('select'));
        console.log("Selects totales (incluyendo ocultos):", selects.length);
        for (const sel of selects) {
            const opts = Array.from(sel.options).map(o => o.text);
            if (opts.some(o => o.includes('Zonales'))) {
                // Seleccionar Zonales
                for (let i = 0; i < sel.options.length; i++) {
                    if (sel.options[i].text.includes('Zonales')) {
                        sel.selectedIndex = i;
                        sel.dispatchEvent(new Event('change', {bubbles: true}));
                        return {ok: true, msg: "Zonales seleccionado en select index " + i};
                    }
                }
            }
        }
        return {ok: false, msg: "No se encontró select con Zonales. Total selects: " + selects.length};
    }''')
    print(f"  → Selección Zonales via JS: {seleccionado_zonales}")
    await page.wait_for_timeout(1000)
    await screenshot(page, "04b_post_zonales")

    # ── Buscar y hacer clic en el checkbox/botón de lista a la derecha ────────
    # Inspeccionar la fila de "Áreas" para encontrar el botón de lista
    btn_lista_info = await frame.evaluate('''() => {
        // Buscar todas las celdas/divs que contengan "Áreas"
        const all = Array.from(document.querySelectorAll('td, div, span, label'));
        for (const el of all) {
            const text = (el.innerText || el.textContent || '').trim();
            if (text === 'Áreas:' || text === 'Áreas') {
                // Buscar elementos hermanos o dentro del mismo row/parent
                const row = el.closest('tr') || el.parentElement;
                if (row) {
                    const inputs = row.querySelectorAll('input, button, img');
                    const result = [];
                    for (const inp of inputs) {
                        const rect = inp.getBoundingClientRect();
                        result.push({
                            tag: inp.tagName,
                            type: inp.type || '',
                            cls: inp.className || '',
                            x: Math.round(rect.x),
                            y: Math.round(rect.y),
                            visible: rect.width > 0
                        });
                    }
                    return {found: true, row_tag: row.tagName, inputs: result};
                }
            }
        }
        return {found: false};
    }''')
    print(f"  → Fila Áreas: {btn_lista_info}")

    # Si encontramos el botón de lista, hacer clic en él
    if btn_lista_info.get('found') and btn_lista_info.get('inputs'):
        for inp in btn_lista_info['inputs']:
            if inp['visible'] and inp['type'] in ('checkbox', 'button', 'image', ''):
                x, y = inp['x'] + 5, inp['y'] + 5
                print(f"  → Click en botón lista en ({x}, {y})")
                await page.mouse.click(x, y)
                await page.wait_for_timeout(1500)
                await screenshot(page, "04c_popup_lista")

                # Verificar si abrió el popup
                popup_lista = await frame.query_selector('text=Editar lista')
                if popup_lista:
                    print("  ✓ Popup 'Editar lista' abierto")
                    # Seleccionar Metropolitana via JS
                    resultado = await frame.evaluate('''() => {
                        const cbs = Array.from(document.querySelectorAll("input[type='checkbox']"));
                        let desmarcados = 0, marcados = 0;
                        for (const cb of cbs) {
                            const parent = cb.parentElement;
                            const text = parent ? (parent.innerText || parent.textContent || '').trim() : '';
                            if (text === 'Todos' || text.includes('Todos')) {
                                if (cb.checked) { cb.click(); desmarcados++; }
                            }
                            if (text.includes('Metropolitana')) {
                                if (!cb.checked) { cb.click(); marcados++; }
                            }
                        }
                        return {desmarcados, marcados};
                    }''')
                    print(f"  → Resultado selección: {resultado}")
                    await page.wait_for_timeout(500)
                    await screenshot(page, "04d_metropolitana")

                    # Clic en Aceptar del popup
                    aceptar = await frame.evaluate('''() => {
                        const btns = Array.from(document.querySelectorAll("button, input[type='button']"));
                        for (const btn of btns) {
                            const text = (btn.innerText || btn.value || '').trim();
                            if (text === 'Aceptar' && btn.offsetParent !== null) {
                                const rect = btn.getBoundingClientRect();
                                return {x: Math.round(rect.x), y: Math.round(rect.y), text};
                            }
                        }
                        return null;
                    }''')
                    if aceptar:
                        await page.mouse.click(aceptar['x'] + 5, aceptar['y'] + 5)
                        print("  ✓ Aceptar popup")
                    await page.wait_for_timeout(1000)
                break

    await screenshot(page, "04e_post_popup")

    # ── Seleccionar Estado = "Revisión y Autorización PCCT" ──────────────────
    estado_result = await frame.evaluate(f'''() => {{
        const selects = Array.from(document.querySelectorAll("select"));
        for (const sel of selects) {{
            const opts = Array.from(sel.options).map(o => o.text);
            if (opts.some(o => o.includes('PCCT'))) {{
                for (let i = 0; i < sel.options.length; i++) {{
                    if (sel.options[i].text.includes('PCCT')) {{
                        sel.selectedIndex = i;
                        sel.dispatchEvent(new Event('change', {{bubbles: true}}));
                        return {{ok: true, opcion: sel.options[i].text}};
                    }}
                }}
            }}
        }}
        return {{ok: false, msg: "Select PCCT no encontrado"}};
    }}''')
    print(f"  → Estado PCCT: {estado_result}")
    await page.wait_for_timeout(500)
    await screenshot(page, "04f_estado")

    # ── Clic en "Aplicar" ────────────────────────────────────────────────────
    aplicar_result = await frame.evaluate('''() => {
        const btns = Array.from(document.querySelectorAll("button, input[type='button'], a"));
        for (const btn of btns) {
            const text = (btn.innerText || btn.value || btn.textContent || '').trim();
            if (text === 'Aplicar' && btn.offsetParent !== null) {
                const rect = btn.getBoundingClientRect();
                return {x: Math.round(rect.x), y: Math.round(rect.y)};
            }
        }
        return null;
    }''')
    if aplicar_result:
        await page.mouse.click(aplicar_result['x'] + 5, aplicar_result['y'] + 5)
        print("  ✓ Clic en Aplicar")
    else:
        # fallback
        await frame.click('text=Aplicar')

    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    await screenshot(page, "05_filtro_aplicado")
    print("  ✓ Filtro aplicado")


async def aprobar_pts(page, frame):
    print("\n[4] APROBANDO PT's")
    pts_aprobados, pts_fallidos = [], []
    iteracion = 0

    while iteracion < 100:
        iteracion += 1
        try:
            primera_celda = await frame.query_selector('table tbody tr:first-child td:first-child')
            if primera_celda is None:
                print("  ✓ Lista vacía.")
                break
            id_pt = (await primera_celda.inner_text()).strip()
            if not re.match(r'\d{4}-\d{5}', id_pt):
                print(f"  ✓ Sin PT's válidos (celda='{id_pt}')")
                break

            print(f"  → PT {id_pt} (iter {iteracion})")
            await frame.click('table tbody tr:first-child')
            await page.wait_for_timeout(800)

            # Clic en Aprobar (botón en barra de herramientas)
            await frame.click('a:has-text("Aprobar"), button:has-text("Aprobar")')
            await page.wait_for_timeout(1500)

            # Clic en Aceptar del popup via JS (más confiable)
            await frame.evaluate('''() => {
                const btns = Array.from(document.querySelectorAll("button, input[type='button']"));
                for (const btn of btns) {
                    const text = (btn.innerText || btn.value || '').trim();
                    if (text === 'Aceptar' && btn.offsetParent !== null) {
                        btn.click();
                        return true;
                    }
                }
                return false;
            }''')

            await page.wait_for_load_state("networkidle", timeout=15_000)
            await page.wait_for_timeout(2000)
            pts_aprobados.append(id_pt)
            print(f"  ✓ {id_pt} aprobado")

        except PlaywrightTimeout:
            print(f"  ⚠️ Timeout iter {iteracion}")
            break
        except Exception as e:
            msg = str(e)[:80]
            pts_fallidos.append(f"iter-{iteracion}: {msg}")
            print(f"  ✗ {msg}")
            await screenshot(page, f"error_{iteracion}")
            await page.wait_for_timeout(2000)

    await screenshot(page, "06_final")
    return pts_aprobados, pts_fallidos


def enviar_reporte(pts_aprobados, pts_fallidos, error_critico=None):
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    lista_ok = "".join(
        f"<tr><td style='padding:4px 12px'>✅</td><td style='font-family:monospace;padding:4px 12px'>{pt}</td></tr>"
        for pt in pts_aprobados
    ) or "<tr><td colspan='2' style='padding:4px 12px;color:#888'>Ninguno</td></tr>"
    lista_err = "".join(
        f"<tr><td style='padding:4px 12px'>❌</td><td style='padding:4px 12px'>{pt}</td></tr>"
        for pt in pts_fallidos
    ) or "<tr><td colspan='2' style='padding:4px 12px;color:#888'>Sin errores</td></tr>"
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
        {error_bloque}
        <h3 style="color:#006600">✅ Aprobados ({len(pts_aprobados)})</h3>
        <table style="border-collapse:collapse;width:100%">{lista_ok}</table>
        <h3 style="color:#cc0000;margin-top:20px">❌ Errores ({len(pts_fallidos)})</h3>
        <table style="border-collapse:collapse;width:100%">{lista_err}</table>
        <p style="color:#999;font-size:11px;margin-top:20px">Bot SAESA – GitHub Actions | Lun–Vie 08:00 Chile</p>
      </div></body></html>"""

    asunto = f"[SAESA] {datetime.now().strftime('%d/%m/%Y')} – {len(pts_aprobados)} PT's aprobados"
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
    pts_aprobados, pts_fallidos, error_critico = [], [], None

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
            await aplicar_filtro(page, frame)
            pts_aprobados, pts_fallidos = await aprobar_pts(page, frame)
        except Exception as e:
            error_critico = str(e)
            print(f"\n✗ ERROR CRÍTICO: {e}")
            await screenshot(page, "error_critico")
        finally:
            await browser.close()

    print(f"\n  {len(pts_aprobados)} aprobados | {len(pts_fallidos)} errores")
    enviar_reporte(pts_aprobados, pts_fallidos, error_critico)
    print("✓ Fin.\n")


if __name__ == "__main__":
    asyncio.run(main())
