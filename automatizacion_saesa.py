"""
Automatización SAESA – Autorización diaria de PT's
Framework: ExtJS (Anachronics DMS)
Los dropdowns de ExtJS se renderizan en el documento PRINCIPAL, no en el iframe.
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


async def click_en_texto_dropdown(page, texto_buscado):
    """
    Busca y hace clic en un item de dropdown ExtJS.
    Los dropdowns de ExtJS se renderizan en el documento PRINCIPAL (page),
    incluso cuando el trigger está dentro de un iframe.
    Busca en TODOS los frames disponibles.
    """
    for intentos in range(3):
        # Buscar en todos los frames (incluyendo la página principal)
        for ctx in [page] + list(page.frames):
            try:
                result = await ctx.evaluate(f'''() => {{
                    const textos = [
                        '.x-combo-list-item',
                        '.x-list-item', 
                        'div.x-combo-list div',
                        '[class*="combo-list"] div',
                        '[class*="list-item"]',
                        '.x-boundlist-item'
                    ];
                    for (const sel of textos) {{
                        const items = Array.from(document.querySelectorAll(sel));
                        for (const item of items) {{
                            if (!item.offsetParent) continue;
                            const t = (item.innerText || item.textContent || '').trim();
                            if (t === "{texto_buscado}") {{
                                const r = item.getBoundingClientRect();
                                item.click();
                                return {{ok: true, x: r.x, y: r.y, sel}};
                            }}
                        }}
                    }}
                    // Fallback: cualquier div/li visible con ese texto exacto
                    const all = Array.from(document.querySelectorAll('div, li, td, span'));
                    for (const el of all) {{
                        if (!el.offsetParent) continue;
                        const t = (el.innerText || '').trim();
                        if (t === "{texto_buscado}" && el.children.length === 0) {{
                            const r = el.getBoundingClientRect();
                            el.click();
                            return {{ok: true, via: 'fallback_leaf', x: r.x, y: r.y}};
                        }}
                    }}
                    return {{ok: false}};
                }}''')
                if result.get('ok'):
                    print(f"    ✓ '{texto_buscado}' encontrado y clickeado en frame '{getattr(ctx, 'name', 'main')}'")
                    return True
            except Exception:
                continue
        await asyncio.sleep(0.5)
    print(f"    ⚠️ '{texto_buscado}' no encontrado en ningún frame tras 3 intentos")
    return False


async def hacer_login(page):
    print("\n[1] LOGIN")
    await page.goto(SAESA_URL, wait_until="domcontentloaded", timeout=60_000)
    await page.wait_for_timeout(3000)
    inputs = await page.query_selector_all('input[type="text"], input:not([type]), input[type="password"]')
    for inp in inputs:
        typ = (await inp.get_attribute("type") or "text").lower()
        if typ in ("text", "") and await inp.is_visible():
            await inp.fill(SAESA_USER); break
    for inp in inputs:
        if (await inp.get_attribute("type") or "").lower() == "password" and await inp.is_visible():
            await inp.fill(SAESA_PASS); break
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
            print(f"  → iframe: '{f.name}' url='{f.url[:60]}'")
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

    # ── PASO A: Seleccionar "Zonales" en el combo Áreas ───────────────────────
    # El arrow trigger de Áreas está en arrow[6] pos=(690,495)
    # Hacemos clic directamente con mouse en esas coordenadas
    print("  → Click en flecha combo Áreas (690, 495)")
    await page.mouse.click(690, 495)
    await page.wait_for_timeout(2000)
    await screenshot(page, "04b_dropdown_areas")

    # El dropdown ExtJS se renderiza en el documento principal de la página
    # (no en el iframe). Buscamos "Zonales" en TODOS los contextos.
    ok_zonales = await click_en_texto_dropdown(page, "Zonales")
    await page.wait_for_timeout(1000)
    await screenshot(page, "04c_zonales")

    if not ok_zonales:
        # Diagnóstico: qué hay visible en el dropdown ahora
        debug = await page.evaluate('''() => {
            const all = Array.from(document.querySelectorAll("div, li"));
            const visible = all.filter(el => el.offsetParent && el.children.length === 0);
            return visible.slice(0, 20).map(el => ({
                tag: el.tagName,
                cls: el.className.substring(0, 40),
                text: (el.innerText || '').trim().substring(0, 50)
            }));
        }''')
        print(f"  → Elementos leaf visibles en page: {debug[:10]}")

    # ── PASO B: Click en list-trigger para abrir "Editar lista" ──────────────
    # list-trigger está en pos=(838,495)
    print("  → Click en list-trigger (838, 495)")
    await page.mouse.click(838, 495)
    await page.wait_for_timeout(2000)
    await screenshot(page, "04d_popup_lista")

    # Verificar si abrió en algún frame
    popup_frame = None
    for ctx in [page] + list(page.frames):
        try:
            tiene = await ctx.evaluate('''() =>
                Array.from(document.querySelectorAll("*")).some(
                    el => el.offsetParent && (el.innerText || "").includes("Editar lista")
                )
            ''')
            if tiene:
                popup_frame = ctx
                print(f"  ✓ Popup 'Editar lista' encontrado en frame '{getattr(ctx, 'name', 'main')}'")
                break
        except Exception:
            continue

    if popup_frame:
        # Desmarcar Todos
        await popup_frame.evaluate('''() => {
            const cbs = Array.from(document.querySelectorAll("input[type='checkbox']"));
            for (const cb of cbs) {
                if (!cb.offsetParent) continue;
                const p = cb.closest("div,td,li,label") || cb.parentElement;
                if (p && p.innerText.includes("Todos") && cb.checked) cb.click();
            }
        }''')
        await page.wait_for_timeout(300)

        # Marcar Metropolitana
        marcado = await popup_frame.evaluate('''() => {
            const cbs = Array.from(document.querySelectorAll("input[type='checkbox']"));
            for (const cb of cbs) {
                if (!cb.offsetParent) continue;
                const p = cb.closest("div,td,li,label") || cb.parentElement;
                if (p && p.innerText.includes("Metropolitana") && !cb.checked) {
                    cb.click();
                    return true;
                }
            }
            return false;
        }''')
        print(f"  → Metropolitana marcada: {marcado}")
        await page.wait_for_timeout(300)
        await screenshot(page, "04e_metropolitana")

        # Clic en Aceptar del popup
        aceptar = await popup_frame.evaluate('''() => {
            const btns = Array.from(document.querySelectorAll("button,input[type='button']"));
            for (const btn of btns) {
                if (!btn.offsetParent) continue;
                const t = (btn.innerText || btn.value || "").trim();
                if (t === "Aceptar") {
                    const r = btn.getBoundingClientRect();
                    return {x: Math.round(r.x), y: Math.round(r.y)};
                }
            }
            return null;
        }''')
        if aceptar:
            await page.mouse.click(aceptar['x'] + 20, aceptar['y'] + 5)
            print("  ✓ Aceptar popup lista")
        await page.wait_for_timeout(1000)
    else:
        print("  ⚠️ Popup 'Editar lista' no detectado")

    await screenshot(page, "04f_post_lista")

    # ── PASO C: Seleccionar Estado = PCCT ─────────────────────────────────────
    # arrow[7] pos=(838,521) es la flecha del Estado
    print("  → Click en flecha combo Estado (838, 521)")
    await page.mouse.click(838, 521)
    await page.wait_for_timeout(2000)
    await screenshot(page, "04g_dropdown_estado")

    ok_pcct = await click_en_texto_dropdown(page, ESTADO_FILTRO)
    if not ok_pcct:
        # Intentar con texto parcial
        for ctx in [page] + list(page.frames):
            try:
                r = await ctx.evaluate('''() => {
                    const all = Array.from(document.querySelectorAll("div,li,td"));
                    for (const el of all) {
                        if (!el.offsetParent) continue;
                        const t = (el.innerText || "").trim();
                        if (t.includes("PCCT") && t.includes("Revisión")) {
                            el.click();
                            return {ok: true, text: t};
                        }
                    }
                    return {ok: false};
                }''')
                if r.get('ok'):
                    print(f"  ✓ PCCT encontrado: '{r['text']}'")
                    ok_pcct = True
                    break
            except Exception:
                continue
    await page.wait_for_timeout(500)
    await screenshot(page, "04h_estado")

    # ── PASO D: Clic en Aplicar ────────────────────────────────────────────────
    for ctx in [frame, page]:
        try:
            aplicar = await ctx.evaluate('''() => {
                const btns = Array.from(document.querySelectorAll("button,input[type='button'],a,span"));
                for (const btn of btns) {
                    if (!btn.offsetParent) continue;
                    const t = (btn.innerText || btn.value || btn.textContent || "").trim();
                    if (t === "Aplicar") {
                        const r = btn.getBoundingClientRect();
                        return {x: Math.round(r.x), y: Math.round(r.y)};
                    }
                }
                return null;
            }''')
            if aplicar:
                await page.mouse.click(aplicar['x'] + 20, aplicar['y'] + 5)
                print("  ✓ Aplicar")
                break
        except Exception:
            continue

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

            # Aceptar popup (buscar en todos los frames)
            for ctx in [page] + list(page.frames):
                try:
                    aceptar = await ctx.evaluate('''() => {
                        const btns = Array.from(document.querySelectorAll("button,input[type='button']"));
                        for (const btn of btns) {
                            if (!btn.offsetParent) continue;
                            const t = (btn.innerText || btn.value || "").trim();
                            if (t === "Aceptar") {
                                const r = btn.getBoundingClientRect();
                                return {x: Math.round(r.x), y: Math.round(r.y)};
                            }
                        }
                        return null;
                    }''')
                    if aceptar:
                        await page.mouse.click(aceptar['x'] + 20, aceptar['y'] + 5)
                        break
                except Exception:
                    continue

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
