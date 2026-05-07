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
    await screenshot(page, "00_pagina_cargada")

    inputs = await page.query_selector_all('input[type="text"], input:not([type]), input[type="password"]')
    print(f"  → Inputs: {len(inputs)}")
    for i, inp in enumerate(inputs):
        name = await inp.get_attribute("name") or ""
        id_  = await inp.get_attribute("id") or ""
        typ  = await inp.get_attribute("type") or "text"
        print(f"    [{i}] name='{name}' id='{id_}' type='{typ}'")

    for inp in inputs:
        typ = (await inp.get_attribute("type") or "text").lower()
        if typ in ("text", "") and await inp.is_visible():
            await inp.fill(SAESA_USER)
            print("  ✓ Usuario")
            break

    for inp in inputs:
        if (await inp.get_attribute("type") or "").lower() == "password" and await inp.is_visible():
            await inp.fill(SAESA_PASS)
            print("  ✓ Contraseña")
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
    await screenshot(page, "02_dms")

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

    # Diagnóstico de selects
    selects = await frame.query_selector_all('select')
    print(f"  → Selects: {len(selects)}")
    for i, sel in enumerate(selects):
        opts = await sel.evaluate('el => Array.from(el.options).map(o => o.text)')
        print(f"    select[{i}]: {opts[:6]}")

    # Seleccionar Zonales
    for sel in selects:
        opts = await sel.evaluate('el => Array.from(el.options).map(o => o.text)')
        if any('Zonales' in o or 'CIREN' in o for o in opts):
            await sel.select_option(label='Zonales')
            print("  ✓ Áreas → Zonales")
            await page.wait_for_timeout(1000)
            break

    # Diagnóstico de checkboxes
    cbs = await frame.query_selector_all('input[type="checkbox"]')
    print(f"  → Checkboxes: {len(cbs)}")
    for i, cb in enumerate(cbs):
        visible = await cb.is_visible()
        parent_text = await cb.evaluate('el => el.parentElement ? el.parentElement.innerText.trim() : ""')
        print(f"    cb[{i}] visible={visible} parent='{parent_text[:50]}'")

    # Abrir popup de lista (checkbox que no es "bandeja de trabajo")
    abierto_popup = False
    for i, cb in enumerate(cbs):
        if not await cb.is_visible():
            continue
        parent_text = await cb.evaluate('el => el.parentElement ? el.parentElement.innerText.trim() : ""')
        if 'bandeja' in parent_text.lower():
            continue
        print(f"  → Click cb[{i}] ('{parent_text[:30]}')")
        await cb.click()
        await page.wait_for_timeout(1500)
        popup = await frame.query_selector('text=Editar lista')
        if popup:
            abierto_popup = True
            print("  ✓ Popup 'Editar lista' OK")
            break
        else:
            await cb.click()  # revertir
            await page.wait_for_timeout(300)

    await screenshot(page, "04b_popup")

    if abierto_popup:
        popup_cbs = await frame.query_selector_all('input[type="checkbox"]')
        # Desmarcar Todos
        for cb in popup_cbs:
            if not await cb.is_visible():
                continue
            parent_text = await cb.evaluate('el => el.parentElement ? el.parentElement.innerText.trim() : ""')
            if 'Todos' in parent_text and await cb.is_checked():
                await cb.click()
                await page.wait_for_timeout(400)
                print("  ✓ Todos desmarcado")
                break

        # Marcar Metropolitana
        popup_cbs = await frame.query_selector_all('input[type="checkbox"]')
        for cb in popup_cbs:
            if not await cb.is_visible():
                continue
            parent_text = await cb.evaluate('el => el.parentElement ? el.parentElement.innerText.trim() : ""')
            if 'Metropolitana' in parent_text:
                if not await cb.is_checked():
                    await cb.click()
                print(f"  ✓ Marcado: '{parent_text}'")
                break

        await screenshot(page, "04c_metropolitana")

        # Aceptar popup
        btns = await frame.query_selector_all('button, input[type="button"]')
        for btn in btns:
            if not await btn.is_visible():
                continue
            txt = await btn.inner_text()
            val = await btn.get_attribute('value') or ''
            if 'Aceptar' in txt or 'Aceptar' in val:
                await btn.click()
                print("  ✓ Aceptar popup")
                break
        await page.wait_for_timeout(1000)
    else:
        print("  ⚠️ No se abrió popup — continuando sin filtro de área")

    await screenshot(page, "04d_post_popup")

    # Estado = Revisión y Autorización PCCT
    selects = await frame.query_selector_all('select')
    for sel in selects:
        opts = await sel.evaluate('el => Array.from(el.options).map(o => o.text)')
        if any('PCCT' in o for o in opts):
            for opt in opts:
                if ESTADO_FILTRO in opt:
                    await sel.select_option(label=opt)
                    print(f"  ✓ Estado → '{opt}'")
                    break
            break
    await page.wait_for_timeout(500)
    await screenshot(page, "04e_estado")

    # Aplicar
    btns = await frame.query_selector_all('button, input[type="button"], a')
    for btn in btns:
        if not await btn.is_visible():
            continue
        txt = await btn.inner_text()
        val = await btn.get_attribute('value') or ''
        if 'Aplicar' in txt or 'Aplicar' in val:
            await btn.click()
            print("  ✓ Aplicar")
            break

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
            await frame.click('a:has-text("Aprobar"), button:has-text("Aprobar")')
            await page.wait_for_timeout(1500)

            btns = await frame.query_selector_all('button, input[type="button"]')
            for btn in btns:
                if not await btn.is_visible():
                    continue
                txt = await btn.inner_text()
                val = await btn.get_attribute('value') or ''
                if 'Aceptar' in txt or 'Aceptar' in val:
                    await btn.click()
                    break

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
