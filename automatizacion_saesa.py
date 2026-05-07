"""
Automatización SAESA – Autorización diaria de PT's
Solución: usar la API de ExtJS directamente via JS para manipular los combos
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


async def get_content_frame(page):
    """Obtiene el iframe 'content' donde vive el DMS."""
    for f in page.frames:
        if f.name == 'content':
            return f
    # Fallback: cualquier frame con Planificación
    for f in page.frames:
        try:
            await f.wait_for_selector('text=Planificación', timeout=2000)
            return f
        except Exception:
            continue
    return page


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


async def navegar_a_permisos(page):
    print("\n[2] NAVEGACIÓN")
    await page.wait_for_selector('a:has-text("Aplicaciones"), span:has-text("Aplicaciones")', timeout=TIMEOUT)
    await page.click('a:has-text("Aplicaciones"), span:has-text("Aplicaciones")')
    await page.wait_for_timeout(1500)
    await page.wait_for_selector('a:has-text("DMS")', timeout=TIMEOUT)
    await page.click('a:has-text("DMS")')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2000)

    frame = await get_content_frame(page)
    print(f"  → frame: '{frame.name}'")

    await frame.click('text=Planificación')
    await page.wait_for_timeout(1000)
    await frame.click('text=Permisos de trabajo')
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(3000)
    print("  ✓ En Permisos de trabajo")
    return frame


async def aplicar_filtro(page, frame):
    print("\n[3] FILTRO")
    await frame.click('text=Filtro')
    await page.wait_for_timeout(2500)
    await screenshot(page, "04_filtro_abierto")

    # ── Usar API ExtJS directamente para manipular los componentes ─────────────
    # ExtJS guarda todos los componentes en Ext.ComponentMgr (o Ext.getCmp)
    # Podemos iterar todos los componentes y encontrar los combos por su store/value

    resultado = await frame.evaluate('''() => {
        const log = [];
        try {
            // Obtener todos los componentes ExtJS registrados
            const cm = Ext.ComponentMgr || Ext.ComponentManager;
            if (!cm) return {error: "No Ext.ComponentMgr"};

            const all = [];
            cm.each ? cm.each(c => all.push(c)) : Object.values(cm.map || {}).forEach(c => all.push(c));

            log.push("Total componentes: " + all.length);

            // Buscar combos (ComboBox)
            const combos = all.filter(c => c.isXType && c.isXType("combo"));
            log.push("Combos: " + combos.length);

            const results = [];
            combos.forEach((c, i) => {
                try {
                    const store = c.store;
                    let storeData = [];
                    if (store && store.data) {
                        store.data.each ? store.data.each(r => storeData.push(r.data)) : null;
                    }
                    results.push({
                        i, id: c.id,
                        value: c.getValue ? c.getValue() : null,
                        rawValue: c.getRawValue ? c.getRawValue() : null,
                        storeData: storeData.slice(0,5),
                        hidden: c.hidden
                    });
                } catch(e) { results.push({i, error: e.message}); }
            });
            return {log, combos: results};
        } catch(e) {
            return {error: e.message, log};
        }
    }''')
    print(f"  → ExtJS API: {resultado}")
    await page.wait_for_timeout(500)

    # ── Estrategia directa: usar setValue() de ExtJS en los combos ────────────
    set_result = await frame.evaluate('''() => {
        const log = [];
        try {
            const cm = Ext.ComponentMgr || Ext.ComponentManager;
            const all = [];
            cm.each ? cm.each(c => all.push(c)) : Object.values(cm.map || {}).forEach(c => all.push(c));

            // Encontrar el combo de Áreas (tiene opciones como CIREN, Zonales, etc.)
            let areasCombo = null;
            let estadoCombo = null;

            all.forEach(c => {
                if (!c.isXType || !c.isXType("combo") || c.hidden) return;
                try {
                    const store = c.store;
                    if (!store) return;
                    let hasZonales = false, hasPCCT = false;
                    if (store.data && store.data.each) {
                        store.data.each(r => {
                            const v = JSON.stringify(r.data || {});
                            if (v.includes("Zonales")) hasZonales = true;
                            if (v.includes("PCCT")) hasPCCT = true;
                        });
                    }
                    if (hasZonales) areasCombo = c;
                    if (hasPCCT) estadoCombo = c;
                } catch(e) {}
            });

            if (areasCombo) {
                log.push("areasCombo encontrado: " + areasCombo.id);
                // Buscar el record de "Zonales"
                let zonalesRecord = null;
                areasCombo.store.data.each(r => {
                    if (JSON.stringify(r.data).includes("Zonales")) zonalesRecord = r;
                });
                if (zonalesRecord) {
                    areasCombo.setValue(zonalesRecord.get(areasCombo.valueField || "id"));
                    areasCombo.fireEvent("select", areasCombo, zonalesRecord, 0);
                    log.push("Zonales seleccionado: " + JSON.stringify(zonalesRecord.data));
                } else {
                    log.push("Record Zonales no encontrado");
                }
            } else {
                log.push("areasCombo NO encontrado");
            }

            if (estadoCombo) {
                log.push("estadoCombo encontrado: " + estadoCombo.id);
                let pcctRecord = null;
                estadoCombo.store.data.each(r => {
                    if (JSON.stringify(r.data).includes("PCCT")) pcctRecord = r;
                });
                if (pcctRecord) {
                    estadoCombo.setValue(pcctRecord.get(estadoCombo.valueField || "id"));
                    estadoCombo.fireEvent("select", estadoCombo, pcctRecord, 0);
                    log.push("PCCT seleccionado: " + JSON.stringify(pcctRecord.data).substring(0, 100));
                }
            } else {
                log.push("estadoCombo NO encontrado");
            }

            return {log};
        } catch(e) {
            return {error: e.message, log};
        }
    }''')
    print(f"  → setValue ExtJS: {set_result}")
    await page.wait_for_timeout(1500)
    await screenshot(page, "04b_post_extjs_set")

    # ── Después de seleccionar Zonales, hacer clic en list-trigger ────────────
    # El list-trigger abre el popup "Editar lista" para elegir áreas específicas
    # Coordenadas dentro del iframe content: (838, 495)
    # Usar frame.click con selector CSS específico
    list_trigger_clicked = await frame.evaluate('''() => {
        const lt = document.querySelector("img.x-form-list-trigger");
        if (lt && lt.offsetParent) {
            lt.click();
            return {ok: true};
        }
        return {ok: false};
    }''')
    print(f"  → list-trigger click: {list_trigger_clicked}")
    await page.wait_for_timeout(2000)
    await screenshot(page, "04c_popup_lista")

    # Verificar popup en frame content
    popup_ok = await frame.evaluate('''() =>
        Array.from(document.querySelectorAll("*")).some(
            el => el.offsetParent && (el.innerText || "").trim() === "Editar lista"
        )
    ''')
    print(f"  → Popup 'Editar lista' en frame content: {popup_ok}")

    if popup_ok:
        # Desmarcar Todos, marcar Metropolitana
        r = await frame.evaluate('''() => {
            const cbs = Array.from(document.querySelectorAll("input[type='checkbox']"));
            let desmarcados = 0, marcados = 0;
            // Desmarcar Todos
            for (const cb of cbs) {
                if (!cb.offsetParent) continue;
                const p = cb.closest("div,td,li,tr") || cb.parentElement;
                if (p && p.innerText.trim() === "Todos" && cb.checked) { cb.click(); desmarcados++; }
            }
            // Marcar Metropolitana
            for (const cb of document.querySelectorAll("input[type='checkbox']")) {
                if (!cb.offsetParent) continue;
                const p = cb.closest("div,td,li,tr") || cb.parentElement;
                if (p && p.innerText.includes("Metropolitana") && !cb.checked) { cb.click(); marcados++; }
            }
            return {desmarcados, marcados};
        }''')
        print(f"  → Selección Metropolitana: {r}")
        await page.wait_for_timeout(300)
        await screenshot(page, "04d_metropolitana")

        # Aceptar
        aceptar = await frame.evaluate('''() => {
            for (const btn of document.querySelectorAll("button,input[type='button']")) {
                if (!btn.offsetParent) continue;
                if ((btn.innerText || btn.value || "").trim() === "Aceptar") {
                    btn.click(); return true;
                }
            }
            return false;
        }''')
        print(f"  → Aceptar popup: {aceptar}")
        await page.wait_for_timeout(1000)

    await screenshot(page, "04e_post_lista")

    # ── Seleccionar Estado PCCT si el ExtJS set no funcionó ───────────────────
    # (Ya lo intentamos via ExtJS arriba, pero como fallback hacemos clic manual)
    estado_check = await frame.evaluate('''() => {
        // Verificar si ya hay valor en el campo Estado
        const triggers = Array.from(document.querySelectorAll("img.x-form-arrow-trigger"));
        const estadoTrigger = triggers.find(t => {
            const r = t.getBoundingClientRect();
            return r.y > 515 && r.y < 535 && r.x > 830;
        });
        if (estadoTrigger) {
            const input = estadoTrigger.previousElementSibling;
            return {val: input ? input.value : "no input", found: true};
        }
        return {found: false};
    }''')
    print(f"  → Estado actual: {estado_check}")

    # Clic en Aplicar
    aplicar_ok = await frame.evaluate('''() => {
        const btns = Array.from(document.querySelectorAll("button,input[type='button'],a,span"));
        for (const btn of btns) {
            if (!btn.offsetParent) continue;
            const t = (btn.innerText || btn.value || btn.textContent || "").trim();
            if (t === "Aplicar") { btn.click(); return true; }
        }
        return false;
    }''')
    print(f"  → Aplicar: {aplicar_ok}")

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

            # Aceptar popup en frame content
            aceptar = await frame.evaluate('''() => {
                for (const btn of document.querySelectorAll("button,input[type='button']")) {
                    if (!btn.offsetParent) continue;
                    if ((btn.innerText || btn.value || "").trim() === "Aceptar") { btn.click(); return true; }
                }
                return false;
            }''')
            if not aceptar:
                # Fallback en page
                await page.evaluate('''() => {
                    for (const btn of document.querySelectorAll("button,input[type='button']")) {
                        if (!btn.offsetParent) continue;
                        if ((btn.innerText || btn.value || "").trim() === "Aceptar") { btn.click(); return; }
                    }
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
