"""
Automatización SAESA – Autorización diaria de PT's
Framework: ExtJS (Anachronics DMS)
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
    print("  ✓ En Permisos de trabajo")
    return frame


async def aplicar_filtro(page, frame):
    print("\n[3] FILTRO (ExtJS)")
    await frame.click('text=Filtro')
    await page.wait_for_timeout(2500)
    await screenshot(page, "04_filtro_abierto")

    # ── El sistema usa ExtJS. Los widgets son inputs con imagen-trigger al lado.
    # Buscar los triggers por su clase CSS específica de ExtJS ─────────────────
    info = await frame.evaluate('''() => {
        const result = {};

        // Áreas: campo con x-form-arrow-trigger (flecha del combo "Zonales")
        // y x-form-list-trigger (el cuadradito que abre "Editar lista")
        const arrowTriggers = Array.from(document.querySelectorAll('img.x-form-arrow-trigger'));
        const listTriggers  = Array.from(document.querySelectorAll('img.x-form-list-trigger'));
        const allInputs     = Array.from(document.querySelectorAll('input.x-form-text'));

        result.arrowTriggers = arrowTriggers.map(el => {
            const r = el.getBoundingClientRect();
            return {x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), visible: r.width > 0};
        });
        result.listTriggers = listTriggers.map(el => {
            const r = el.getBoundingClientRect();
            return {x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), visible: r.width > 0};
        });
        result.textInputs = allInputs.filter(el => el.offsetParent).map(el => {
            const r = el.getBoundingClientRect();
            return {x: Math.round(r.x), y: Math.round(r.y), cls: el.className, val: el.value};
        });
        return result;
    }''')

    print(f"  → Arrow triggers: {len(info['arrowTriggers'])}")
    for i, t in enumerate(info['arrowTriggers']):
        print(f"    arrow[{i}] pos=({t['x']},{t['y']}) visible={t['visible']}")

    print(f"  → List triggers: {len(info['listTriggers'])}")
    for i, t in enumerate(info['listTriggers']):
        print(f"    list[{i}] pos=({t['x']},{t['y']}) visible={t['visible']}")

    print(f"  → Text inputs visibles: {len(info['textInputs'])}")
    for i, t in enumerate(info['textInputs']):
        print(f"    input[{i}] pos=({t['x']},{t['y']}) val='{t['val'][:30]}'")

    # ── PASO A: Click en la flecha del combo "Áreas" para abrir dropdown ─────
    # El combo de Áreas (Zonales/CIREN/etc.) tiene una x-form-arrow-trigger
    # Según el log anterior, el campo Áreas está en y≈495 y su arrow en x:690,y:495
    # Buscamos el arrow trigger visible en esa zona (y entre 480-510)
    areas_arrow = None
    for t in info['arrowTriggers']:
        if t['visible'] and 480 <= t['y'] <= 515:
            areas_arrow = t
            break

    if not areas_arrow and info['arrowTriggers']:
        # Fallback: usar el que esté más cerca de y=495
        visible = [t for t in info['arrowTriggers'] if t['visible']]
        if visible:
            areas_arrow = min(visible, key=lambda t: abs(t['y'] - 495))

    if areas_arrow:
        print(f"  → Click flecha Áreas en ({areas_arrow['x']+5}, {areas_arrow['y']+5})")
        await page.mouse.click(areas_arrow['x'] + 5, areas_arrow['y'] + 5)
        await page.wait_for_timeout(1500)
        await screenshot(page, "04b_dropdown_areas")

        # Buscar y hacer clic en "Zonales" en el dropdown
        zonales_clicked = await frame.evaluate('''() => {
            // El dropdown de ExtJS crea una lista flotante con class x-combo-list
            const items = Array.from(document.querySelectorAll(
                '.x-combo-list-item, .x-list-item, div[class*="combo"] div'
            ));
            for (const item of items) {
                if ((item.innerText || item.textContent || '').trim() === 'Zonales') {
                    item.click();
                    return {ok: true, text: item.innerText};
                }
            }
            // Alternativa: buscar cualquier elemento visible con texto "Zonales"
            const all = Array.from(document.querySelectorAll('div, li, span'));
            for (const el of all) {
                const text = (el.innerText || el.textContent || '').trim();
                if (text === 'Zonales' && el.offsetParent) {
                    el.click();
                    return {ok: true, via: 'fallback', text};
                }
            }
            return {ok: false};
        }''')
        print(f"  → Zonales clicked: {zonales_clicked}")
        await page.wait_for_timeout(1000)
        await screenshot(page, "04c_zonales_seleccionado")
    else:
        print("  ⚠️ No se encontró flecha de Áreas")

    # ── PASO B: Click en x-form-list-trigger para abrir "Editar lista" ────────
    # Según el log: cls='x-form-trigger x-form-list-trigger' en x:838, y:495
    list_trigger = None
    for t in info['listTriggers']:
        if t['visible']:
            list_trigger = t
            break

    if list_trigger:
        print(f"  → Click list-trigger en ({list_trigger['x']+5}, {list_trigger['y']+5})")
        await page.mouse.click(list_trigger['x'] + 5, list_trigger['y'] + 5)
        await page.wait_for_timeout(2000)
        await screenshot(page, "04d_popup_lista")

        # Verificar si abrió el popup "Editar lista"
        popup_abierto = await frame.evaluate('''() => {
            const els = Array.from(document.querySelectorAll('*'));
            return els.some(el => (el.innerText || '').includes('Editar lista') && el.offsetParent);
        }''')
        print(f"  → Popup 'Editar lista' abierto: {popup_abierto}")

        if popup_abierto:
            # Desmarcar Todos y marcar Metropolitana
            resultado = await frame.evaluate('''() => {
                const cbs = Array.from(document.querySelectorAll("input[type='checkbox']"));
                let desmarcados = 0, marcados = 0;
                for (const cb of cbs) {
                    if (!cb.offsetParent) continue;
                    const parent = cb.closest('div, td, li, label') || cb.parentElement;
                    const text = (parent ? parent.innerText : '').trim();
                    if (text.includes('Todos') && cb.checked) {
                        cb.click(); desmarcados++;
                    }
                }
                // Re-leer para marcar Metropolitana
                const cbs2 = Array.from(document.querySelectorAll("input[type='checkbox']"));
                for (const cb of cbs2) {
                    if (!cb.offsetParent) continue;
                    const parent = cb.closest('div, td, li, label') || cb.parentElement;
                    const text = (parent ? parent.innerText : '').trim();
                    if (text.includes('Metropolitana') && !cb.checked) {
                        cb.click(); marcados++;
                    }
                }
                return {desmarcados, marcados, totalCbs: cbs.length};
            }''')
            print(f"  → Selección Metropolitana: {resultado}")
            await page.wait_for_timeout(500)
            await screenshot(page, "04e_metropolitana")

            # Clic en Aceptar del popup
            aceptar_pos = await frame.evaluate('''() => {
                const btns = Array.from(document.querySelectorAll("button, input[type='button'], a"));
                for (const btn of btns) {
                    if (!btn.offsetParent) continue;
                    const text = (btn.innerText || btn.value || btn.textContent || '').trim();
                    if (text === 'Aceptar') {
                        const r = btn.getBoundingClientRect();
                        return {x: Math.round(r.x), y: Math.round(r.y)};
                    }
                }
                return null;
            }''')
            if aceptar_pos:
                await page.mouse.click(aceptar_pos['x'] + 20, aceptar_pos['y'] + 5)
                print("  ✓ Aceptar popup lista")
            await page.wait_for_timeout(1000)
    else:
        print("  ⚠️ No se encontró list-trigger — usando checkbox cuadrado a la derecha de Áreas")
        # Fallback: el checkbox cuadrado a la derecha del campo áreas (x≈1035, y≈495 según captura)
        await page.mouse.click(1035, 495)
        await page.wait_for_timeout(1500)
        await screenshot(page, "04d_fallback_checkbox")

    await screenshot(page, "04f_post_lista")

    # ── PASO C: Seleccionar Estado = "Revisión y Autorización PCCT" ───────────
    # El campo Estado tiene una x-form-arrow-trigger. Según el log: y≈521
    # Hay múltiples arrow triggers — buscamos el de Estado (y entre 515-535)
    estado_arrow = None
    info2 = await frame.evaluate('''() => {
        const arrowTriggers = Array.from(document.querySelectorAll('img.x-form-arrow-trigger'));
        return arrowTriggers.filter(el => el.offsetParent).map(el => {
            const r = el.getBoundingClientRect();
            return {x: Math.round(r.x), y: Math.round(r.y), visible: r.width > 0};
        });
    }''')

    for t in info2:
        if t['visible'] and 510 <= t['y'] <= 540:
            estado_arrow = t
            break

    if not estado_arrow and info2:
        # Usar el segundo arrow trigger visible (el primero es Áreas, el segundo es Estado)
        visible = [t for t in info2 if t['visible']]
        if len(visible) >= 2:
            estado_arrow = visible[1]
        elif visible:
            estado_arrow = visible[0]

    if estado_arrow:
        print(f"  → Click flecha Estado en ({estado_arrow['x']+5}, {estado_arrow['y']+5})")
        await page.mouse.click(estado_arrow['x'] + 5, estado_arrow['y'] + 5)
        await page.wait_for_timeout(1500)
        await screenshot(page, "04g_dropdown_estado")

        # Hacer clic en "Revisión y Autorización PCCT" del dropdown
        pcct_clicked = await frame.evaluate('''() => {
            const items = Array.from(document.querySelectorAll(
                ".x-combo-list-item, .x-list-item, div[class*='combo'] div, div[class*='list'] div"
            ));
            for (const item of items) {
                if (!item.offsetParent) continue;
                const text = (item.innerText || item.textContent || '').trim();
                if (text.includes('PCCT') && text.includes('Revisión')) {
                    item.click();
                    return {ok: true, text};
                }
            }
            // Fallback más amplio
            const all = Array.from(document.querySelectorAll('div, li, td'));
            for (const el of all) {
                if (!el.offsetParent) continue;
                const text = (el.innerText || '').trim();
                if (text.includes('PCCT') && text.includes('Revisión') && text.length < 60) {
                    el.click();
                    return {ok: true, via: 'fallback', text};
                }
            }
            return {ok: false};
        }''')
        print(f"  → PCCT clicked: {pcct_clicked}")
        await page.wait_for_timeout(500)
        await screenshot(page, "04h_estado_seleccionado")
    else:
        print("  ⚠️ No se encontró flecha de Estado")

    # ── PASO D: Clic en "Aplicar" ─────────────────────────────────────────────
    aplicar_pos = await frame.evaluate('''() => {
        const btns = Array.from(document.querySelectorAll("button, input[type='button'], a, span"));
        for (const btn of btns) {
            if (!btn.offsetParent) continue;
            const text = (btn.innerText || btn.value || btn.textContent || '').trim();
            if (text === 'Aplicar') {
                const r = btn.getBoundingClientRect();
                return {x: Math.round(r.x), y: Math.round(r.y)};
            }
        }
        return null;
    }''')

    if aplicar_pos:
        await page.mouse.click(aplicar_pos['x'] + 20, aplicar_pos['y'] + 5)
        print("  ✓ Aplicar")
    else:
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
                print(f"  ✓ Sin PT's (celda='{id_pt}')")
                break

            print(f"  → PT {id_pt} (iter {iteracion})")
            await frame.click('table tbody tr:first-child')
            await page.wait_for_timeout(800)

            # Clic en botón Aprobar de la toolbar
            await frame.click('a:has-text("Aprobar"), button:has-text("Aprobar")')
            await page.wait_for_timeout(1500)

            # Clic en Aceptar del popup via coordenadas JS
            aceptar = await frame.evaluate('''() => {
                const btns = Array.from(document.querySelectorAll("button, input[type='button']"));
                for (const btn of btns) {
                    if (!btn.offsetParent) continue;
                    const text = (btn.innerText || btn.value || '').trim();
                    if (text === 'Aceptar') {
                        const r = btn.getBoundingClientRect();
                        return {x: Math.round(r.x), y: Math.round(r.y)};
                    }
                }
                return null;
            }''')
            if aceptar:
                await page.mouse.click(aceptar['x'] + 20, aceptar['y'] + 5)
            else:
                await frame.click('button:has-text("Aceptar")')

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
