"""
Verificador SAT Guatemala
Flujo: Captcha → Formulario (Tipo=Documento, TipoDoc=FEL, NIT, Serie, Número) → Buscar → Descargar PDF
"""

import re
import asyncio
import logging
import time
from pathlib import Path
from dataclasses import dataclass
from typing import Optional

from playwright.async_api import async_playwright, Page, Browser, BrowserContext, Frame

logger = logging.getLogger(__name__)

SAT_URL = "https://portal.sat.gob.gt/portal/verificador-integrado/"
SCREENSHOTS_DIR = Path("downloads")
SCREENSHOTS_DIR.mkdir(exist_ok=True)


@dataclass
class ResultadoVerificacion:
    valida: Optional[bool]
    mensaje: str
    detalles: str = ""
    captcha_screenshot: Optional[Path] = None
    esperando_captcha: bool = False
    pdf_path: Optional[Path] = None
    razon_social_sat: str = ""  # Razón social oficial del SAT
    monto: str = ""
    fecha_emision: str = ""
    hora_emision: str = ""  # HH:MM
    tipo_comida: str = ""   # Desayuno/Almuerzo/Cena (solo si es alimentos)
    establecimiento: str = ""  # Nombre del establecimiento
    departamento: str = ""    # Departamento extraído de la dirección del emisor


class VerificadorSAT:

    def __init__(self):
        self._playwright = None
        self._browser: Optional[Browser] = None
        self._context: Optional[BrowserContext] = None
        self._page: Optional[Page] = None
        self._datos_pendientes: dict = {}
        self._sesion_activa: bool = False
        self._sesion_timestamp: float = 0.0

    # ──────────────────────────────────────────────
    # Ciclo de vida
    # ──────────────────────────────────────────────

    async def iniciar(self):
        self._playwright = await async_playwright().start()
        candidatos = [
            Path.home() / "AppData/Local/ms-playwright/chromium-1208/chrome-win64/chrome.exe",
            Path("C:/Program Files/Google/Chrome/Application/chrome.exe"),
            Path("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"),
            Path("C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"),
            Path("C:/Program Files/Microsoft/Edge/Application/msedge.exe"),
        ]
        exe_path = next((str(p) for p in candidatos if p.exists()), None)
        if not exe_path:
            raise RuntimeError("No se encontró Chrome/Chromium instalado.")
        logger.info(f"Usando Chrome: {exe_path}")
        self._browser = await self._playwright.chromium.launch(headless=True, executable_path=exe_path)
        self._context = await self._browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800},
            accept_downloads=True,
        )
        self._page = await self._context.new_page()
        logger.info("Navegador iniciado")

    def esta_viva(self) -> bool:
        """True si el navegador/pagina siguen utilizables. Detecta navegadores
        que crashearon o fueron cerrados por fuera (el SO, el usuario, etc.).
        Es sincrono a proposito: solo consulta flags locales de Playwright."""
        try:
            if self._page is None or self._browser is None:
                return False
            if self._page.is_closed():
                return False
            if not self._browser.is_connected():
                return False
            return True
        except Exception:
            return False

    async def cerrar(self):
        for obj in [self._page, self._context, self._browser]:
            try:
                if obj: await obj.close()
            except Exception:
                pass
        try:
            if self._playwright: await self._playwright.stop()
        except Exception:
            pass

    # ──────────────────────────────────────────────
    # Helpers de frames
    # ──────────────────────────────────────────────

    def _frames(self) -> list:
        return self._page.frames

    async def _frame_formulario(self) -> Optional[Frame]:
        """Retorna el frame que contiene el formulario (tiene select con Documento/Contribuyente)"""
        for frame in self._frames():
            try:
                sel = frame.locator("select").first
                if await sel.count() == 0:
                    continue
                opts = await sel.evaluate("el => Array.from(el.options).map(o => o.text.toLowerCase())")
                if any("documento" in o or "contribuyente" in o for o in opts):
                    return frame
            except Exception:
                continue
        return None

    async def _frame_captcha(self) -> Optional[Frame]:
        """Retorna el frame que contiene el input del captcha"""
        IGNORAR = ["su mensaje", "buscar", "search", "correo"]
        for frame in self._frames():
            try:
                inputs = frame.locator('input[type="text"], input:not([type])')
                count = await inputs.count()
                for i in range(count):
                    ph = (await inputs.nth(i).get_attribute("placeholder") or "").lower()
                    if not any(ig in ph for ig in IGNORAR):
                        return frame
            except Exception:
                continue
        return None

    # ──────────────────────────────────────────────
    # Detección
    # ──────────────────────────────────────────────

    async def _hay_captcha(self) -> bool:
        """True si hay imagen de captcha visible en algún frame"""
        for frame in self._frames():
            try:
                txt = await frame.evaluate("() => document.body.innerText")
                if "ingrese los caracteres" in txt.lower():
                    return True
            except Exception:
                continue
        return False

    async def _screenshot_captcha(self, destino: Path) -> bool:
        """Captura el captcha SAT lo mas legible posible.

        Estrategia:
        1. Localiza el iframe del captcha y obtiene su bounding box en la pagina.
        2. Recorta esa region completa (captcha + input + boton) y la guarda
           ampliada al doble con PIL si esta disponible.
        3. Fallback: screenshot completo de la pagina (comportamiento previo).

        El captcha SAT es muy pequeño en un screenshot full-page — recortar solo
        el iframe lo hace mucho mas legible, especialmente en celulares.
        """
        page = self._page
        try:
            # Buscar el iframe del captcha — el que tiene "consultaFacturas" en src
            iframes = await page.query_selector_all("iframe")
            captcha_handle = None
            for iframe in iframes:
                try:
                    src = await iframe.get_attribute("src") or ""
                    if "consultaFacturas" in src or "captcha" in src.lower():
                        captcha_handle = iframe
                        break
                except Exception:
                    continue
            if captcha_handle:
                # Screenshot del element + padding
                try:
                    box = await captcha_handle.bounding_box()
                    if box and box["width"] > 50 and box["height"] > 50:
                        clip = {
                            "x": max(0, box["x"] - 10),
                            "y": max(0, box["y"] - 10),
                            "width": box["width"] + 20,
                            "height": box["height"] + 20,
                        }
                        await page.screenshot(path=str(destino), clip=clip)
                        # Ampliar 2x para legibilidad si PIL esta disponible
                        try:
                            from PIL import Image as _Image
                            img = _Image.open(destino)
                            img = img.resize(
                                (img.width * 2, img.height * 2),
                                _Image.LANCZOS,
                            )
                            img.save(destino)
                        except Exception:
                            pass  # mantener tamaño original si PIL falla
                        return True
                except Exception as e:
                    logger.debug(f"_screenshot_captcha clip fallo: {e}")
        except Exception as e:
            logger.debug(f"_screenshot_captcha localizacion fallo: {e}")
        # Fallback: screenshot completo
        await page.screenshot(path=str(destino))
        return False

    async def _click_nueva_consulta(self):
        """Click en 'Nueva Consulta' via PrimeFaces.ab() directo (más rápido y confiable)"""
        page = self._page

        for frame in self._frames():
            try:
                # Ir directo a PrimeFaces.ab() — sabemos que es lo que funciona
                ajax_ok = await frame.evaluate("""
                    () => {
                        if (typeof PrimeFaces !== 'undefined') {
                            PrimeFaces.ab({
                                s: 'formContent:btnNuevaConsulta',
                                f: 'formContent',
                                p: 'formContent:btnNuevaConsulta',
                                u: 'formContent'
                            });
                            return true;
                        }
                        return false;
                    }
                """)
                if ajax_ok:
                    logger.info("Nueva Consulta via PrimeFaces.ab()")
                    await page.wait_for_timeout(2000)
                    return
            except Exception:
                pass

            # Fallback: click directo en el botón
            try:
                btn = frame.locator('#formContent\\:btnNuevaConsulta')
                if await btn.count() > 0:
                    await btn.click(force=True)
                    logger.info("Nueva Consulta via click directo")
                    await page.wait_for_timeout(2000)
                    return
            except Exception:
                continue

        logger.warning("No se encontró 'Nueva Consulta'")

    # ──────────────────────────────────────────────
    # Punto de entrada principal
    # ──────────────────────────────────────────────

    async def verificar(self, nit: str, serie: str, numero: str, user_id: str) -> ResultadoVerificacion:
        try:
            if not self._page:
                await self.iniciar()

            self._datos_pendientes = {"nit": nit, "serie": serie, "numero": numero}

            # Invalidar sesiones viejas (>20 min) para evitar estados corruptos
            if self._sesion_activa and (time.time() - self._sesion_timestamp) > 1200:
                logger.info("Sesión expirada (>20 min), forzando navegación fresca...")
                self._sesion_activa = False
                self._sesion_timestamp = 0.0

            # Si ya tenemos sesión activa, verificar que seguimos en SAT
            if self._sesion_activa:
                url_actual = self._page.url or ""
                if "sat.gob.gt" in url_actual:
                    logger.info("Sesión activa, buscando formulario...")
                    # Intentar encontrar el formulario con espera generosa
                    for intento in range(8):
                        frame = await self._frame_formulario()
                        if frame:
                            logger.info("Formulario listo, procediendo")
                            return await self._llenar_y_buscar(nit, serie, numero, user_id)
                        await self._page.wait_for_timeout(500)

                    # No encontrado, intentar Nueva Consulta (sin navegar)
                    logger.info("Formulario no visible, intentando Nueva Consulta...")
                    await self._click_nueva_consulta()
                    for intento in range(8):
                        frame = await self._frame_formulario()
                        if frame:
                            logger.info("Formulario recuperado tras Nueva Consulta")
                            return await self._llenar_y_buscar(nit, serie, numero, user_id)
                        await self._page.wait_for_timeout(500)

                    # Último intento: reload (NO goto — preserva cookies de sesión)
                    logger.info("Intentando reload en vez de goto...")
                    await self._page.reload(wait_until="domcontentloaded", timeout=30000)
                    await self._page.wait_for_timeout(2000)
                    if not await self._hay_captcha():
                        frame = await self._frame_formulario()
                        if frame:
                            logger.info("Formulario recuperado tras reload")
                            return await self._llenar_y_buscar(nit, serie, numero, user_id)

                logger.info("Sesión perdida, navegando de nuevo")
                self._sesion_activa = False

            # Navegar al portal
            logger.info(f"Navegando a {SAT_URL}")
            await self._page.goto(SAT_URL, wait_until="networkidle", timeout=30000)
            await self._page.wait_for_timeout(2000)

            if await self._hay_captcha():
                logger.info("Captcha detectado")
                screenshot = SCREENSHOTS_DIR / f"captcha_{user_id}.png"
                await self._screenshot_captcha(screenshot)
                return ResultadoVerificacion(
                    valida=None,
                    mensaje="El portal SAT pide un captcha.\nEscribeme el texto que ves en la imagen:",
                    captcha_screenshot=screenshot,
                    esperando_captcha=True,
                )

            return await self._llenar_y_buscar(nit, serie, numero, user_id)

        except Exception as e:
            logger.error(f"Error verificar: {e}", exc_info=True)
            return ResultadoVerificacion(valida=None, mensaje=f"Error al verificar en SAT: {e}")

    async def verificar_con_captcha(self, texto: str, user_id: str) -> ResultadoVerificacion:
        try:
            page = self._page
            await self._llenar_captcha(texto)
            await self._click_boton_captcha()

            await page.wait_for_load_state("networkidle", timeout=15000)
            await page.wait_for_timeout(1500)

            # ¿Sigue el captcha? (incorrecto)
            if await self._hay_captcha():
                screenshot = SCREENSHOTS_DIR / f"captcha_{user_id}.png"
                await self._screenshot_captcha(screenshot)
                return ResultadoVerificacion(
                    valida=None,
                    mensaje="Captcha incorrecto, intenta de nuevo:",
                    captcha_screenshot=screenshot,
                    esperando_captcha=True,
                )

            # Formulario cargado
            datos = self._datos_pendientes
            return await self._llenar_y_buscar(datos["nit"], datos["serie"], datos["numero"], user_id)

        except Exception as e:
            logger.error(f"Error captcha: {e}", exc_info=True)
            return ResultadoVerificacion(valida=None, mensaje=f"Error enviando captcha: {e}")

    # ──────────────────────────────────────────────
    # Captcha
    # ──────────────────────────────────────────────

    @staticmethod
    def _sanitizar_captcha(texto: str) -> str:
        """Quita ruido invisible que algunos teclados de celular agregan al
        copiar/escribir captchas (espacios fantasma, zero-width chars, etc.).

        SAT a veces rechaza captchas correctos porque el cliente envia
        algo como ' ggj ' o 'ggj\\u200b' que el usuario nunca ve.
        """
        if not texto:
            return ""
        # Quitar zero-width y otros invisibles unicode
        invisibles = (
            "​",  # ZERO WIDTH SPACE
            "‌",  # ZERO WIDTH NON-JOINER
            "‍",  # ZERO WIDTH JOINER
            "⁠",  # WORD JOINER
            "﻿",  # BOM
            "\xa0",    # NBSP
        )
        limpio = texto
        for ch in invisibles:
            limpio = limpio.replace(ch, "")
        # strip de espacios normales + tabulaciones + saltos de linea
        return limpio.strip()

    async def _llenar_captcha(self, texto: str):
        """Llena el input del captcha usando frame directo (cross-origin safe)"""
        texto_original = texto
        texto = self._sanitizar_captcha(texto)
        if texto != texto_original:
            logger.warning(
                f"Captcha sanitizado: {texto_original!r} -> {texto!r} "
                f"(longitudes {len(texto_original)} -> {len(texto)})"
            )
        frame = await self._frame_captcha()
        if not frame:
            # fallback: click en posición fija y keyboard
            vp = self._page.viewport_size or {"width": 1280, "height": 800}
            await self._page.mouse.click(vp["width"] // 2, int(vp["height"] * 0.48))
            await self._page.wait_for_timeout(200)
            await self._page.keyboard.press("Control+a")
            await self._page.keyboard.type(texto)
            logger.info(f"Captcha '{texto}' escrito en posición fija")
            return

        IGNORAR = ["su mensaje", "buscar", "search", "correo"]
        inputs = frame.locator('input[type="text"], input:not([type])')
        count = await inputs.count()
        for i in range(count):
            inp = inputs.nth(i)
            ph = (await inp.get_attribute("placeholder") or "").lower()
            if any(ig in ph for ig in IGNORAR):
                continue
            await inp.click(click_count=3)
            await inp.fill(texto)
            # Disparar eventos PrimeFaces
            await frame.evaluate("""
                () => {
                    const inp = document.querySelector('input[type="text"]:not([placeholder*="mensaje"])');
                    if (inp) ['input','change','keyup','blur'].forEach(e =>
                        inp.dispatchEvent(new Event(e, {bubbles:true}))
                    );
                }
            """)
            logger.info(f"Captcha '{texto}' llenado en frame {frame.url}")
            return

    async def _click_boton_captcha(self):
        """Click en el botón Verificar del captcha por coordenadas"""
        for frame in self._frames():
            try:
                txt = await frame.evaluate("() => document.body.innerText")
                if "ingrese los caracteres" not in txt.lower():
                    continue
                coords = await frame.evaluate("""
                    () => {
                        const btns = Array.from(document.querySelectorAll('button, input[type="submit"], a'));
                        for (const btn of btns) {
                            const t = (btn.textContent || btn.value || btn.title || '').toLowerCase();
                            if (t.includes('verificar') || t.includes('submit')) {
                                const r = btn.getBoundingClientRect();
                                if (r.width > 0) return {x: r.left + r.width/2, y: r.top + r.height/2};
                            }
                        }
                        return null;
                    }
                """)
                if coords:
                    # Convertir a coordenadas de página
                    iframe_rect = await self._page.evaluate("""
                        () => {
                            for (const f of document.querySelectorAll('iframe')) {
                                try {
                                    const r = f.getBoundingClientRect();
                                    if (r.width > 0 && r.height > 0) return {left: r.left, top: r.top};
                                } catch(e) {}
                            }
                            return {left: 0, top: 0};
                        }
                    """)
                    px = coords['x'] + iframe_rect.get('left', 0)
                    py = coords['y'] + iframe_rect.get('top', 0)
                    await self._page.mouse.click(px, py)
                    logger.info(f"Click botón captcha en ({px:.0f}, {py:.0f})")
                    return
            except Exception:
                continue

        # Fallback: Enter
        await self._page.keyboard.press("Enter")
        logger.info("Botón captcha: Enter fallback")

    # ──────────────────────────────────────────────
    # Formulario principal
    # ──────────────────────────────────────────────

    async def _llenar_y_buscar(self, nit: str, serie: str, numero: str, user_id: str) -> ResultadoVerificacion:
        page = self._page
        await page.wait_for_timeout(500)

        # Esperar a que el frame del formulario esté disponible. El SAT a veces
        # tarda en cargar el iframe; somos pacientes (hasta ~9s) y, si no
        # aparece, recargamos la pagina UNA vez y volvemos a esperar antes de
        # rendirnos. Asi una carga lenta no tumba la verificacion.
        async def _esperar_frame(intentos: int, espera_ms: int):
            for intento in range(intentos):
                f = await self._frame_formulario()
                if f:
                    return f
                await page.wait_for_timeout(espera_ms)
                logger.info(f"Esperando formulario... intento {intento + 1}/{intentos}")
            return None

        frame = await _esperar_frame(12, 750)  # ~9s

        if not frame:
            logger.info("Formulario no aparecio — recargando pagina y reintentando...")
            try:
                await page.reload(wait_until="domcontentloaded", timeout=30000)
                await page.wait_for_timeout(1500)
            except Exception as e:
                logger.warning(f"Reload fallo: {e}")
            # Si tras el reload aparece captcha, hay que pedirlo de nuevo.
            if await self._hay_captcha():
                screenshot = SCREENSHOTS_DIR / f"captcha_{user_id}.png"
                await self._screenshot_captcha(screenshot)
                return ResultadoVerificacion(
                    valida=None,
                    mensaje="El portal SAT pide un captcha.\nEscribeme el texto que ves en la imagen:",
                    captcha_screenshot=screenshot,
                    esperando_captcha=True,
                )
            frame = await _esperar_frame(10, 800)  # ~8s mas tras el reload

        if not frame:
            screenshot = SCREENSHOTS_DIR / f"debug_{user_id}.png"
            await page.screenshot(path=str(screenshot), full_page=True)
            return ResultadoVerificacion(
                valida=None,
                mensaje="El portal SAT no cargó el formulario (puede estar lento o caído). Intenta de nuevo.",
                captcha_screenshot=screenshot,
            )

        # 1. Seleccionar "Documento" en primer dropdown y esperar que aparezca el segundo
        # Si falla, recargar la página y reintentar una vez (sesión puede estar en estado raro)
        for intento_doc in range(2):
            ok_doc = await self._click_opcion_primefaces(frame, page, 0, "Documento")
            if not ok_doc and intento_doc == 0:
                logger.warning("Dropdown 'Documento' falló, recargando página y reintentando...")
                self._sesion_activa = False
                self._sesion_timestamp = 0.0
                await page.reload(wait_until="domcontentloaded", timeout=30000)
                await page.wait_for_timeout(2000)
                if await self._hay_captcha():
                    logger.info("Captcha apareció tras reload en retry dropdown")
                    break
                frame = await self._frame_formulario()
                if not frame:
                    break
                continue
            break

        # Esperar hasta 4s a que el segundo dropdown (selTipoDocumento) esté visible
        for _ in range(8):
            await page.wait_for_timeout(500)
            frame = await self._frame_formulario()
            if frame:
                cnt = await frame.locator(".ui-selectonemenu-trigger").count()
                if cnt >= 2:
                    break
        if not frame:
            return ResultadoVerificacion(valida=None, mensaje="SAT no respondió tras seleccionar Documento.")

        # 2. Seleccionar "FEL" en segundo dropdown y esperar que aparezcan los campos de serie/número
        for intento_fel in range(3):
            await self._click_opcion_primefaces(frame, page, 1, "FEL")
            # Esperar hasta 3s a que aparezcan los campos serie y número (5+ inputs en total)
            for _ in range(6):
                await page.wait_for_timeout(500)
                frame = await self._frame_formulario()
                if not frame:
                    continue
                n_inputs = await frame.evaluate("""
                    () => Array.from(document.querySelectorAll('input[type="text"]'))
                              .filter(i => i.getBoundingClientRect().width > 0).length
                """)
                if n_inputs >= 5:
                    logger.info(f"FEL activo: {n_inputs} inputs visibles")
                    break
            else:
                logger.warning(f"FEL intento {intento_fel+1}: solo {n_inputs} inputs, reintentando...")
                continue
            break
        else:
            return ResultadoVerificacion(
                valida=None,
                mensaje="No se pudo activar el formulario FEL en el SAT. Intenta de nuevo.",
            )
        if not frame:
            return ResultadoVerificacion(valida=None, mensaje="SAT no respondió tras seleccionar FEL.")

        # 3. Limpiar todos los campos de texto antes de llenar (evita datos residuales)
        try:
            await frame.evaluate("""
                () => {
                    const inputs = document.querySelectorAll('input[type="text"]');
                    inputs.forEach(inp => {
                        if (inp.getBoundingClientRect().width > 0) {
                            inp.value = '';
                            ['input', 'change'].forEach(e =>
                                inp.dispatchEvent(new Event(e, {bubbles: true}))
                            );
                        }
                    });
                }
            """)
            logger.info("Campos de texto limpiados")
        except Exception as e:
            logger.warning(f"Error limpiando campos: {e}")
        await page.wait_for_timeout(300)

        # 4. Detectar posiciones reales de los inputs visibles
        nit_form = nit.replace("-", "")
        mapa = await self._detectar_inputs(frame)
        logger.info(f"Mapa de inputs detectados: {mapa}")

        pos_nit = mapa.get("nit_emisor")
        pos_serie = mapa.get("serie")
        pos_numero = mapa.get("numero")

        if pos_nit is not None:
            await self._llenar_input_posicion(frame, pos_nit, nit_form)
            await page.wait_for_timeout(300)
        else:
            logger.warning(f"NIT emisor NO detectado en el mapa de inputs")
        if pos_serie is not None:
            await self._llenar_input_posicion(frame, pos_serie, serie)
            await page.wait_for_timeout(300)
        if pos_numero is not None:
            await self._llenar_input_posicion(frame, pos_numero, numero)
            await page.wait_for_timeout(300)

        # Verificar que los campos se llenaron correctamente
        try:
            valores = await frame.evaluate("""
                () => {
                    const inputs = document.querySelectorAll('input[type="text"]');
                    const visibles = Array.from(inputs).filter(i => i.getBoundingClientRect().width > 0);
                    return visibles.map(i => ({val: i.value, id: i.id}));
                }
            """)
            logger.info(f"Valores en formulario antes de buscar: {valores}")
            # Verificar que los 3 campos tienen datos
            campos_llenos = [v for v in valores if v.get('val', '').strip()]
            if len(campos_llenos) < 3:
                logger.warning(f"Solo {len(campos_llenos)}/3 campos llenos, reintentando llenado...")
                # Reintentar el llenado
                frame = await self._frame_formulario()
                if frame:
                    mapa2 = await self._detectar_inputs(frame)
                    if mapa2.get("nit_emisor") is not None:
                        await self._llenar_input_posicion(frame, mapa2["nit_emisor"], nit_form)
                    if mapa2.get("serie") is not None:
                        await self._llenar_input_posicion(frame, mapa2["serie"], serie)
                    if mapa2.get("numero") is not None:
                        await self._llenar_input_posicion(frame, mapa2["numero"], numero)
                    await page.wait_for_timeout(500)
                    # Verificar de nuevo
                    valores2 = await frame.evaluate("""
                        () => {
                            const inputs = document.querySelectorAll('input[type="text"]');
                            const visibles = Array.from(inputs).filter(i => i.getBoundingClientRect().width > 0);
                            return visibles.map(i => i.value);
                        }
                    """)
                    logger.info(f"Valores despues de reintento: {valores2}")
        except Exception:
            pass

        # 5. Verificacion final: si el NIT sigue vacio, no buscar
        try:
            valores_final = await frame.evaluate("""
                () => {
                    const inputs = document.querySelectorAll('input[type="text"]');
                    const visibles = Array.from(inputs).filter(i => i.getBoundingClientRect().width > 0);
                    return visibles.map(i => i.value);
                }
            """)
            campos_con_dato = [v for v in valores_final if v.strip()]
            if len(campos_con_dato) < 3:
                logger.error(f"Solo {len(campos_con_dato)}/3 campos llenos: {valores_final}. Abortando busqueda.")
                await self._click_nueva_consulta()
                self._sesion_activa = True
                self._sesion_timestamp = time.time()
                return ResultadoVerificacion(
                    valida=None,
                    mensaje=f"Error: el formulario no se lleno correctamente. Solo {len(campos_con_dato)} de 3 campos tienen datos. Intenta de nuevo.",
                )
        except Exception:
            pass

        # 6. Click en Buscar
        await self._click_boton_frame(frame, page, ["buscar", "search"])
        logger.info("Click en Buscar ejecutado")

        # 6. Esperar resultado — polling activo en vez de networkidle (más rápido)
        resultado_listo = False
        for _ in range(30):  # max 30 segundos
            await page.wait_for_timeout(1000)
            for f in self._frames():
                try:
                    txt = await f.evaluate("() => document.body.innerText.toLowerCase()")
                    if "detalle del documento" in txt or "monto total" in txt:
                        resultado_listo = True
                        break
                    if "no existe" in txt or "no encontrad" in txt or "invalid" in txt or "no fue posible" in txt or "ingrese nit" in txt:
                        resultado_listo = True
                        break
                except Exception:
                    continue
            if resultado_listo:
                break

        # 6. Leer texto de TODOS los frames para detectar errores y extraer datos
        todos_textos = []
        texto = ""
        for f in self._frames():
            try:
                t = await f.evaluate("() => document.body.innerText")
                todos_textos.append(t)
                # Para extracción de datos, usar el frame con contenido SAT relevante
                tl = t.lower()
                if "detalle del documento" in tl or "monto total" in tl or "no existe" in tl or "fecha emisi" in tl or "no fue posible" in tl or "ingrese nit" in tl:
                    texto = t
            except Exception:
                continue
        # Fallback: si no encontró frame SAT, usar el más largo
        if not texto:
            texto = max(todos_textos, key=len) if todos_textos else ""

        # Concatenar todos los textos para detección de errores (el error puede estar en cualquier frame)
        texto_completo = "\n".join(todos_textos).lower()

        screenshot = SCREENSHOTS_DIR / f"resultado_{user_id}.png"
        await page.screenshot(path=str(screenshot), full_page=True)

        # Si sigue procesando o no hay datos útiles, reportar timeout
        if "procesando" in texto_completo or "espere" in texto_completo:
            return ResultadoVerificacion(
                valida=None,
                mensaje="El SAT esta tardando mucho en responder. Quieres intentar de nuevo?",
                captcha_screenshot=screenshot,
            )
        if any(p in texto_completo for p in ["no existe", "inválid", "invalida", "no encontrad", "no fue posible", "ingrese nit", "ingrese el nit"]):
            # Click "Nueva Consulta" para dejar listo para la siguiente
            await self._click_nueva_consulta()
            self._sesion_activa = True
            self._sesion_timestamp = time.time()
            lineas = [l.strip() for l in texto.split("\n") if len(l.strip()) > 2]
            return ResultadoVerificacion(
                valida=False, mensaje="Factura NO válida o no encontrada en SAT",
                detalles="\n".join(lineas[:20]),
                captcha_screenshot=screenshot,
            )

        # 7. Extraer monto, fecha, hora, establecimiento y razón social del texto
        monto = self._extraer_monto(texto)
        fecha = self._extraer_fecha_emision(texto)
        razon_social_sat = self._extraer_razon_social(texto)
        hora_emision = self._extraer_hora_emision(texto)
        establecimiento = self._extraer_establecimiento(texto)
        departamento = self._extraer_departamento(texto)
        tipo_comida = self._clasificar_comida(
            hora_emision, razon_social_sat, establecimiento,
            texto_fel=texto, monto=monto,
        )
        logger.info(f"Monto: {monto}, Fecha: {fecha}, Hora: {hora_emision}, Establecimiento: {establecimiento}, Depto: {departamento}, Comida: {tipo_comida}, Emisor: {razon_social_sat}")

        # 8. Descargar PDF via HTTP externo PRIMERO (ViewState aún no consumido)
        pdf_path = None
        fr = await self._encontrar_frame_resultado()
        if fr:
            export_data = await fr.evaluate("""
                () => {
                    const imgs = document.querySelectorAll('img[src*="pdf"]');
                    let exportLink = null;
                    for (const img of imgs) {
                        const a = img.closest('a');
                        if (a) { exportLink = a; break; }
                    }
                    if (!exportLink) return null;
                    const form = exportLink.closest('form') || document.querySelector('form');
                    if (!form) return null;
                    const vs = form.querySelector('input[name="javax.faces.ViewState"]');
                    if (!vs) return null;
                    const hiddens = {};
                    form.querySelectorAll('input[type="hidden"]').forEach(inp => {
                        if (inp.name && inp.name !== 'javax.faces.ViewState') {
                            hiddens[inp.name] = inp.value;
                        }
                    });
                    return {
                        formId: form.id || 'formContent',
                        formAction: form.action || window.location.href,
                        exportId: exportLink.id,
                        viewState: vs.value,
                        hiddens: hiddens
                    };
                }
            """)
            if export_data:
                logger.info(f"Export data: {export_data['exportId']}")
                pdf_path = await self._descargar_pdf_externo(export_data, monto, fecha, hora_emision, establecimiento, tipo_comida)
                # Preferir razón social del PDF oficial (formato fijo) sobre el
                # regex sobre el HTML de la página (que puede cambiar).
                if pdf_path:
                    rs_pdf = self._extraer_razon_social_pdf(pdf_path)
                    if rs_pdf:
                        razon_social_sat = rs_pdf

        # 9. DESPUÉS click "Nueva Consulta" via PrimeFaces.ab (funciona con ViewState del browser)
        await self._click_nueva_consulta()
        form_found = False
        for _ in range(10):
            frame = await self._frame_formulario()
            if frame:
                form_found = True
                break
            await page.wait_for_timeout(500)
        self._sesion_activa = form_found
        if form_found:
            self._sesion_timestamp = time.time()
        logger.info(f"Nueva Consulta: formulario {'encontrado' if form_found else 'NO encontrado'}")

        valida = True
        msg = "Factura VÁLIDA en el portal SAT"

        lineas = [l.strip() for l in texto.split("\n") if len(l.strip()) > 2]
        return ResultadoVerificacion(
            valida=valida, mensaje=msg,
            detalles="\n".join(lineas[:20]),
            captcha_screenshot=screenshot,
            pdf_path=pdf_path,
            razon_social_sat=razon_social_sat,
            monto=monto,
            fecha_emision=fecha,
            hora_emision=hora_emision,
            tipo_comida=tipo_comida,
            establecimiento=establecimiento,
            departamento=departamento,
        )

    async def _click_opcion_primefaces(self, frame: Frame, page: Page, idx: int, texto: str) -> bool:
        """
        Selecciona una opción en un PrimeFaces SelectOneMenu.
        Intenta trigger visual → click opción → si falla, JS directo.
        Retorna True si la selección tuvo éxito.
        """
        # 0. Si la opcion ya esta seleccionada (label del dropdown coincide),
        #    igual disparamos el change del <select> oculto para que PrimeFaces
        #    pinte los campos dependientes (el caso FEL preseleccionado).
        try:
            label_actual = await frame.locator(".ui-selectonemenu-label").nth(idx).text_content()
            if label_actual and texto.strip().lower() in label_actual.strip().lower():
                logger.info(f"Dropdown {idx} ya muestra '{label_actual.strip()}', forzando change")
                forzado = await self._forzar_change_primefaces(frame, idx, texto)
                if forzado:
                    return True
        except Exception as e:
            logger.debug(f"No se pudo leer label del dropdown {idx}: {e}")

        try:
            # 1. Click en el trigger del dropdown N
            trigger = frame.locator(".ui-selectonemenu-trigger").nth(idx)
            if await trigger.count() > 0:
                await trigger.click(force=True)
                logger.info(f"Trigger dropdown {idx} clickeado")
                await page.wait_for_timeout(800)
            else:
                label = frame.locator(".ui-selectonemenu-label").nth(idx)
                if await label.count() > 0:
                    await label.click(force=True)
                    await page.wait_for_timeout(800)

            # 2. Intentar clicar la opción en el panel overlay
            try:
                opcion = frame.locator(f".ui-selectonemenu-item:has-text('{texto}')").first
                if await opcion.count() > 0:
                    await opcion.click(force=True, timeout=3000)
                    logger.info(f"Opción '{texto}' seleccionada vía overlay")
                    return True
            except Exception as e:
                logger.info(f"Click overlay '{texto}' falló: {e}, usando JS...")

        except Exception as e:
            logger.info(f"Trigger {idx} falló: {e}, usando JS...")

        # 3. SIEMPRE intentar JS como fallback — selecciona por value + dispara PrimeFaces AJAX
        resultado = await frame.evaluate("""
            (texto) => {
                const selects = document.querySelectorAll('select');
                for (const sel of selects) {
                    for (const opt of sel.options) {
                        if (opt.text.trim() === texto) {
                            sel.value = opt.value;
                            ['change','input'].forEach(e =>
                                sel.dispatchEvent(new Event(e, {bubbles:true}))
                            );
                            // Disparar PrimeFaces change handler
                            const widgetVar = sel.id.replace(/:/g, '\\\\:');
                            if (typeof PrimeFaces !== 'undefined') {
                                try { PrimeFaces.ab({s: sel.id, e: 'valueChange', p: sel.id, u: sel.id.split(':')[0]}); } catch(ex) {}
                            }
                            return true;
                        }
                    }
                }
                return false;
            }
        """, texto)
        logger.info(f"Opción '{texto}' via JS fallback: {resultado}")
        return bool(resultado)

    async def _forzar_change_primefaces(self, frame: Frame, idx: int, texto: str) -> bool:
        """Fuerza el `valueChange` AJAX de PrimeFaces sobre el N-esimo select del
        formulario. Util cuando la opcion ya esta preseleccionada (FEL default)
        y sin embargo el panel dependiente no aparecio."""
        try:
            ok = await frame.evaluate(
                """
                ({idx, texto}) => {
                    const selects = Array.from(document.querySelectorAll('select'));
                    if (idx >= selects.length) return false;
                    const sel = selects[idx];
                    let target = null;
                    for (const opt of sel.options) {
                        if (opt.text.trim().toLowerCase() === texto.trim().toLowerCase()) {
                            target = opt;
                            break;
                        }
                    }
                    if (target) {
                        sel.value = target.value;
                    }
                    ['change','input'].forEach(e =>
                        sel.dispatchEvent(new Event(e, {bubbles: true}))
                    );
                    if (typeof PrimeFaces !== 'undefined') {
                        try {
                            PrimeFaces.ab({
                                s: sel.id, e: 'valueChange', p: sel.id,
                                u: sel.id.split(':')[0],
                            });
                        } catch (ex) { return false; }
                    }
                    return true;
                }
                """,
                {"idx": idx, "texto": texto},
            )
            logger.info(f"forzar_change dropdown {idx} '{texto}': {ok}")
            return bool(ok)
        except Exception as e:
            logger.warning(f"forzar_change dropdown {idx} fallo: {e}")
            return False

    async def _detectar_inputs(self, frame: Frame) -> dict:
        """Detecta la posición de cada campo visible buscando el texto label asociado"""
        try:
            info = await frame.evaluate("""
                () => {
                    const inputs = Array.from(document.querySelectorAll('input[type="text"]'));
                    const result = [];
                    for (let i = 0; i < inputs.length; i++) {
                        const inp = inputs[i];
                        const rect = inp.getBoundingClientRect();
                        if (rect.width === 0 || rect.height === 0) continue;
                        // Buscar label: por for=, o texto cercano en el DOM
                        let labelText = '';
                        if (inp.id) {
                            const lbl = document.querySelector('label[for="' + inp.id + '"]');
                            if (lbl) labelText = lbl.textContent.trim();
                        }
                        if (!labelText) {
                            // Buscar en la misma fila (tr) o contenedor padre
                            let parent = inp.closest('tr') || inp.parentElement;
                            if (parent) {
                                const labels = parent.querySelectorAll('label, td:first-child, span');
                                for (const l of labels) {
                                    const t = l.textContent.trim();
                                    if (t.length > 2 && t.length < 80) { labelText = t; break; }
                                }
                            }
                        }
                        result.push({idx: result.length, id: inp.id, name: inp.name, label: labelText.toLowerCase()});
                    }
                    return result;
                }
            """)
            logger.info(f"Inputs visibles en frame: {info}")

            mapa = {}
            for item in info:
                lbl = item.get("label", "")
                name = item.get("name", "").lower()
                ident = item.get("id", "").lower()
                combined = f"{lbl} {name} {ident}"

                if ("emisor" in combined or "nit_emisor" in combined or "nitemis" in combined) and "consulta" not in combined and "opcional" not in combined:
                    mapa["nit_emisor"] = item["idx"]
                elif "serie" in combined:
                    mapa["serie"] = item["idx"]
                elif ("mero doc" in combined or "numero doc" in combined
                      or "mero de doc" in combined or "ndocumento" in combined
                      or "numero_doc" in combined or "numdoc" in combined):
                    mapa["numero"] = item["idx"]

            # Si falta algún campo, intentar por posiciones fijas
            if "nit_emisor" not in mapa or "serie" not in mapa or "numero" not in mapa:
                logger.warning(f"Deteccion incompleta: {mapa}. Intentando posiciones fijas...")
                # Buscar primer input después de "consulta"/"opcional"
                skip = 0
                for item in info:
                    lbl = item.get("label", "")
                    if "consulta" in lbl or "opcional" in lbl:
                        skip = item["idx"] + 1
                if skip + 2 < len(info):
                    if "nit_emisor" not in mapa:
                        mapa["nit_emisor"] = skip
                    if "serie" not in mapa:
                        mapa["serie"] = skip + 1
                    if "numero" not in mapa:
                        mapa["numero"] = skip + 2
                elif len(info) >= 3:
                    if "nit_emisor" not in mapa:
                        mapa["nit_emisor"] = len(info) - 3
                    if "serie" not in mapa:
                        mapa["serie"] = len(info) - 2
                    if "numero" not in mapa:
                        mapa["numero"] = len(info) - 1

            return mapa
        except Exception as e:
            logger.warning(f"Error detectando inputs: {e}")
            return {"nit_emisor": 1, "serie": 2, "numero": 3}

    async def _llenar_input_posicion(self, frame: Frame, posicion: int, valor: str):
        """Llena el input visible en la posición dada — limpia primero y dispara eventos PrimeFaces"""
        if not valor:
            return
        try:
            inputs = frame.locator("input[type='text']")
            count = await inputs.count()
            visibles = []
            for i in range(count):
                inp = inputs.nth(i)
                if await inp.is_visible():
                    visibles.append(inp)
            if posicion < len(visibles):
                inp = visibles[posicion]
                # Limpiar campo primero (triple click + borrar)
                await inp.click(click_count=3, force=True)
                await inp.press("Backspace")
                await inp.fill("")
                # Llenar con el nuevo valor
                await inp.fill(valor)
                # Disparar eventos para PrimeFaces
                inp_id = await inp.get_attribute("id") or ""
                await frame.evaluate(f"""
                    () => {{
                        const inp = document.querySelectorAll('input[type="text"]');
                        const visibles = Array.from(inp).filter(i => i.getBoundingClientRect().width > 0);
                        const el = visibles[{posicion}];
                        if (el) {{
                            ['input', 'change', 'keyup', 'blur'].forEach(e =>
                                el.dispatchEvent(new Event(e, {{bubbles: true}}))
                            );
                        }}
                    }}
                """)
                logger.info(f"Input posición {posicion} = '{valor}' (limpiado + eventos)")
            else:
                logger.warning(f"No hay input visible en posición {posicion} (total: {len(visibles)})")
        except Exception as e:
            logger.warning(f"Error llenando posición {posicion}: {e}")

    async def _click_boton_frame(self, frame: Frame, page: Page, keywords: list):
        """Click en un botón del frame por coordenadas (evita problemas de disabled en PrimeFaces)"""
        try:
            coords = await frame.evaluate(f"""
                (keywords) => {{
                    const btns = Array.from(document.querySelectorAll('button, input[type="submit"], a.ui-button'));
                    for (const btn of btns) {{
                        const t = (btn.textContent || btn.value || btn.title || '').toLowerCase();
                        if (keywords.some(k => t.includes(k))) {{
                            const r = btn.getBoundingClientRect();
                            if (r.width > 0) return {{x: r.left + r.width/2, y: r.top + r.height/2, t: t.trim()}};
                        }}
                    }}
                    return null;
                }}
            """, keywords)

            if coords:
                # Obtener offset del iframe en la página principal
                iframe_offset = await page.evaluate("""
                    () => {
                        for (const f of document.querySelectorAll('iframe')) {
                            try {
                                const r = f.getBoundingClientRect();
                                if (r.width > 100) return {left: r.left, top: r.top};
                            } catch(e) {}
                        }
                        return {left: 0, top: 0};
                    }
                """)
                px = coords['x'] + iframe_offset.get('left', 0)
                py = coords['y'] + iframe_offset.get('top', 0)
                await page.mouse.click(px, py)
                logger.info(f"Click '{coords['t']}' en ({px:.0f}, {py:.0f})")
            else:
                await page.keyboard.press("Enter")
                logger.info(f"Botón {keywords} no encontrado, Enter enviado")
        except Exception as e:
            logger.warning(f"Error click boton {keywords}: {e}")
            await page.keyboard.press("Enter")

    def _extraer_monto(self, texto: str) -> str:
        """Extrae el monto total del texto de resultado SAT"""
        # Buscar "Monto Total: GTQ 200.00" o similar
        m = re.search(r'Monto\s+Total[:\s]*(?:GTQ?\s*)?([0-9,]+\.?\d*)', texto, re.IGNORECASE)
        if m:
            raw = m.group(1).replace(",", "")
            # Quitar decimales si son solo ceros (.00, .0000, etc.)
            raw = re.sub(r'\.0+$', '', raw)
            return f"Q{raw}"
        return "Q0"

    def _extraer_fecha_emision(self, texto: str) -> str:
        """Extrae la fecha de emisión y la formatea como DD.MM.YYYY"""
        # Buscar "Fecha Emisión: 30/03/2026 11:47:21" o "Fecha Emisi" con variaciones
        m = re.search(r'Fecha\s+Emisi[oó]n[:\s]*(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})', texto, re.IGNORECASE)
        if m:
            return f"{m.group(1).zfill(2)}.{m.group(2).zfill(2)}.{m.group(3)}"
        return "00.00.0000"

    def _extraer_hora_emision(self, texto: str) -> str:
        """Extrae la hora de emisión (HH:MM) del texto SAT"""
        m = re.search(r'Fecha\s+Emisi[oó]n[:\s]*\d{1,2}[/\-]\d{1,2}[/\-]\d{4}\s+(\d{1,2}):(\d{2})', texto, re.IGNORECASE)
        if m:
            return f"{m.group(1).zfill(2)}:{m.group(2)}"
        return ""

    def _extraer_departamento(self, texto: str) -> str:
        """Extrae el departamento de la dirección del emisor en la FEL.
        La dirección viene después de la línea 'Serie: X Número de DTE: Y'
        y termina antes de 'Numero Acceso:'. El departamento es el último
        elemento tras la última coma."""
        lineas = [l.strip() for l in texto.split('\n') if l.strip()]
        # Encontrar índice de la línea de Serie/DTE
        idx_serie = -1
        for i, l in enumerate(lineas):
            if re.search(r'Serie:.*N[úu]mero de DTE:', l, re.IGNORECASE):
                idx_serie = i
                break
        if idx_serie == -1:
            return ""
        # Recoger líneas de dirección hasta "Numero Acceso"
        dir_lineas = []
        for l in lineas[idx_serie + 1:]:
            if re.search(r'Numero Acceso|NIT Receptor|Fecha y hora', l, re.IGNORECASE):
                break
            dir_lineas.append(l)
        if not dir_lineas:
            return ""
        dir_completa = ' '.join(dir_lineas)
        partes = [p.strip() for p in dir_completa.split(',') if p.strip()]
        if partes:
            return partes[-1].upper()
        return ""

    def _extraer_establecimiento(self, texto: str) -> str:
        """Extrae el nombre del establecimiento del texto SAT"""
        m = re.search(r'Establecimiento[:\s]*(.+)', texto, re.IGNORECASE)
        if m:
            nombre = m.group(1).strip()
            # Limpiar y acortar
            nombre = re.sub(r'\s+', ' ', nombre)
            return nombre[:50]
        return ""

    def _extraer_razon_social(self, texto: str) -> str:
        """Extrae la razón social del emisor del texto SAT (ej: 'Emisor: 99499533-CORPORACION TEXGASO...')"""
        m = re.search(r'Emisor[:\s]*\d+-(.+)', texto)
        if m:
            return m.group(1).strip()
        return ""

    def _extraer_razon_social_pdf(self, pdf_path) -> str:
        """Extrae la razón social del emisor del PDF oficial descargado del SAT.
        Más estable que el regex sobre el texto HTML porque el PDF tiene
        formato fijo. Usa PyMuPDF (ya es dependencia del proyecto).

        Retorna "" si no se logra extraer; el caller debe hacer fallback al
        texto HTML."""
        try:
            import fitz  # PyMuPDF
        except Exception as e:
            logger.warning(f"PyMuPDF no disponible para leer PDF: {e}")
            return ""
        try:
            if not pdf_path or not Path(pdf_path).exists():
                return ""
            doc = fitz.open(str(pdf_path))
            texto = ""
            for page in doc:
                texto += page.get_text() + "\n"
                if len(texto) > 3000:  # con la primera página suele bastar
                    break
            doc.close()
            if not texto:
                return ""
            # Patrones típicos del PDF SAT (en orden de preferencia):
            # 1. "Nombre del Emisor: INDUSTRIA DE HAMBURGUESAS SOCIEDAD ANONIMA"
            # 2. "Emisor Nombre: ..."
            # 3. "Emisor: NIT-NOMBRE"
            patrones = [
                r'Nombre\s+(?:Comercial\s+o\s+Raz[oó]n\s+Social\s+)?(?:del\s+)?Emisor\s*:?\s*(.+)',
                r'Emisor\s+Nombre\s*:?\s*(.+)',
                r'Raz[oó]n\s+Social\s+(?:del\s+)?Emisor\s*:?\s*(.+)',
                r'Emisor\s*:\s*\d+\s*-\s*(.+)',
            ]
            for pat in patrones:
                m = re.search(pat, texto, re.IGNORECASE)
                if m:
                    candidato = m.group(1).strip()
                    # Cortar en salto de línea o siguiente etiqueta
                    candidato = re.split(r'\n|NIT\s*:|Direcci[oó]n', candidato)[0].strip()
                    # Filtrar basura: muy corto, numérico, etiqueta repetida
                    if len(candidato) >= 5 and not candidato.isdigit():
                        logger.info(f"Razón social desde PDF SAT: {candidato}")
                        return candidato
            return ""
        except Exception as e:
            logger.warning(f"Error leyendo razón social del PDF: {e}")
            return ""

    def _clasificar_comida(self, hora: str, razon_social: str, establecimiento: str,
                           texto_fel: str = "", monto: str = "") -> str:
        """Clasifica como Gasolina/Hospedaje/Desayuno/Almuerzo/Cena.

        Orden de prioridad:
        1. Gasolina / Hospedaje por nombre del emisor (sin hora).
        2. Escaneo del texto FEL: si los ítems/descripción dicen explícitamente
           'desayuno', 'almuerzo', 'cena', etc. → usar eso directamente.
        3. Fallback: nombre del negocio en keywords_alimentos + hora de emisión.

        Reglas adicionales para Hospedaje:
        - Debe tener señal clara del nombre (hotel/hostal/etc.) Y NO tener
          señales de comida (para evitar "Restaurante del Hotel X").
        - Monto mínimo Q100: un hospedaje real en Guatemala no cuesta menos.
          Si el monto es < Q100 probablemente es comida del restaurante del
          hotel, parqueo, lavandería, etc., no el hospedaje en sí.
        """
        # Parsear monto como float para el guard de hospedaje. Si falla,
        # deja monto_num=None y saltamos el guard (fallback al comportamiento
        # previo).
        monto_num: Optional[float] = None
        if monto:
            try:
                monto_num = float(
                    str(monto).replace('Q', '').replace(',', '').strip()
                )
            except (ValueError, TypeError):
                monto_num = None
        HOSPEDAJE_MONTO_MINIMO = 100.0
        combined = f"{razon_social} {establecimiento}".lower()

        # ── 1. Gasolina y Hospedaje (no dependen de hora) ──────────────────
        # Algunas palabras son cortas y aparecen como SUBSTRING dentro de otras
        # (ej. "shell" dentro de "Hotel Rahshell", "Marshall", "Russell"). Por
        # eso usamos limites de palabra (\b) con regex.
        keywords_hospedaje = [
            r"\bhotel\b", r"\bhostal\b", r"\bmotel\b", r"\bhospedaje\b",
            r"\bhosped", r"\bposada\b", r"\blodge\b", r"\binn\b",
            r"\bresort\b", r"\balojamiento\b", r"\bsuite\b", r"\bapart hotel\b",
        ]
        # Detectamos primero si hay keyword de hospedaje, para no clasificar como
        # gasolina nombres tipo "Hotel Casa Rashell" donde el OCR puede separar
        # "RASHELL" en "RA SHELL" / "RA-SHELL" y crear un word boundary falso
        # que dispara \bshell\b. Un "Hotel Gasolinera" no existe en la práctica.
        tiene_hospedaje_kw = any(re.search(p, combined) for p in keywords_hospedaje)

        # Keywords cortas/ambiguas (marcas que se pueden confundir con palabras
        # dentro de otros nombres). Si hay keyword de hospedaje, ignoramos las
        # cortas y solo aceptamos las inequívocas (gasolinera, combustible, etc.)
        keywords_gasolina_cortas = [
            r"\btexaco\b", r"\bshell\b", r"\buno gt\b",
        ]
        keywords_gasolina_largas = [
            r"\bgasolinera\b", r"\bservicentro\b", r"\bpuma energy\b",
            r"\bcombustible\b", r"\bpetroleo\b", r"\bgas station\b",
        ]
        keywords_gasolina = keywords_gasolina_largas + (
            [] if tiene_hospedaje_kw else keywords_gasolina_cortas
        )
        if any(re.search(p, combined) for p in keywords_gasolina):
            return "Gasolina"
        # Señales de comida: si el emisor o la factura mencionan comida, NO clasificar
        # como hospedaje aunque el nombre contenga "hotel"/"posada"/"suite"/etc.
        # (caso típico: "Restaurante del Hotel X", "Cafetería Posada Real").
        # Usamos regex con \b en el EMISOR porque palabras cortas como "pollo"
        # podrían aparecer dentro de "REPOLLOS" o similares y dar falso positivo
        # (mismo bug que tuvimos con "shell" dentro de "Rashell"). En el texto FEL
        # de items mantenemos substring porque los items son más explícitos.
        _food_signals_emisor = [
            r"\brestaur", r"\bcafeter", r"\bcomedor\b", r"\bcomida\b",
            r"\bpollo\b", r"\bpizza\b", r"\bhambur", r"\btaco\b",
            r"\bpanaderi", r"\bpasteler", r"\bgrill\b", r"\basado\b",
            r"\bmariscos\b", r"\bbuffet\b", r"\bfood\b", r"\beat\b",
            r"\bdiner\b", r"\bbakery\b", r"\bantoj", r"\bloncher",
            r"\bcocina\b", r"\bsteak\b", r"\bbbq\b", r"\bcarniter",
        ]
        _food_signals_fel = [
            "desayuno", "almuerzo", "cena", "breakfast", "lunch", "dinner",
            "sandwich", "combo", "hamburguesa", "pollo frito", "pizza",
            "taco", "burrito", "tamal", "pupusa", "sopa", "ensalada",
            "huevos", "frijoles", "menu del dia", "plato del dia",
        ]
        _fel_lower = texto_fel.lower() if texto_fel else ""
        has_food_signal = (
            any(re.search(p, combined) for p in _food_signals_emisor)
            or any(k in _fel_lower for k in _food_signals_fel)
        )
        if any(re.search(p, combined) for p in keywords_hospedaje) and not has_food_signal:
            # Guard de monto: un hospedaje no cuesta Q60. Si el monto es muy
            # bajo, el emisor probablemente es un hotel pero el consumo es
            # otra cosa (restaurante, parqueo, lavandería). Dejamos que
            # siga el flujo normal para clasificarlo como alimento por hora.
            if monto_num is not None and monto_num < HOSPEDAJE_MONTO_MINIMO:
                logger.info(
                    f"  Emisor parece hospedaje pero monto Q{monto_num:.2f} "
                    f"< Q{HOSPEDAJE_MONTO_MINIMO:.0f} — no clasificar como "
                    f"hospedaje (probablemente restaurante/otros del hotel)."
                )
            else:
                return "Hospedaje"

        # ── 2. Escaneo del texto FEL (descripción de ítems) ────────────────
        # Primero busca palabras EXPLÍCITAS de tipo de comida (Desayuno/Almuerzo/Cena).
        # Si las encuentra → usa eso directamente, sin depender de la hora.
        # Luego busca ítems de comida genéricos (sandwich, combo, pollo, etc.) →
        # confirma que es comida y clasifica por hora.
        if texto_fel:
            fel_lower = texto_fel.lower()

            # 2a. Palabras explícitas de tipo de comida → prioridad máxima
            keywords_fel_explicitas = [
                (["desayuno", "breakfast", "brunch", "coffe break", "coffee break"], "Desayuno"),
                (["almuerzo", "lunch", "comida del dia", "menu del dia",
                  "menu diario", "plato del dia", "especial del dia"], "Almuerzo"),
                (["cena", "dinner", "cena ejecutiva"], "Cena"),
            ]
            for kws, tipo in keywords_fel_explicitas:
                if any(kw in fel_lower for kw in kws):
                    logger.info(f"[clasificar_comida] Detectado explícito en FEL: {tipo}")
                    return tipo

            # 2b. Ítems de comida genéricos → confirma que es alimento, luego usa hora
            keywords_items_comida = [
                # Sándwiches / wraps
                "sandwich", "sanwuch", "sándwich", "wrap", "torta", "bocadillo",
                # Hamburguesas
                "hamburguesa", "hamburgesa", "burger", "doble queso", "whopper",
                "big mac", "quarter", "mc pollo", "mcpollo",
                # Pollo / carnes
                "pollo frito", "nugget", "alita", "pechuga", "filete de pollo",
                "churrasco", "carne asada", "costilla", "bistec", "milanesa",
                # Combos / órdenes
                "combo", "orden de", "menú ", "menu ", "platillo", "plato",
                "especial", "porcion", "porción",
                # Bebidas de comida
                "gaseosa", "refresco", "agua pura", "jugo natural", "limonada",
                # Otros ítems comunes
                "papas fritas", "ensalada", "sopa", "taco", "burrito", "quesadilla",
                "enchilada", "tamale", "tamal", "pupusa", "pasta", "pizza entera",
                "hotdog", "hot dog", "pancho",
                # Desayunos típicos
                "huevos", "frijoles", "tortilla", "pan tostado", "cereal",
                "avena", "pancake", "waffle", "crepe",
            ]
            if hora and any(kw in fel_lower for kw in keywords_items_comida):
                logger.info(f"[clasificar_comida] Ítem de comida detectado en FEL, usando hora {hora}")
                try:
                    h, m = map(int, hora.split(":"))
                    minutos = h * 60 + m
                    if minutos < 11 * 60 + 20:
                        return "Desayuno"
                    elif minutos < 17 * 60 + 30:
                        return "Almuerzo"
                    else:
                        return "Cena"
                except (ValueError, AttributeError):
                    pass

        # ── 3. Fallback: nombre del negocio + hora de emisión ──────────────
        keywords_alimentos = [
            "restaur", "comida", "pollo", "pizza", "hambur", "taco", "caf",
            "panaderi", "pastel", "campero", "mcdonald", "burger", "subway",
            "wendy", "domino", "little caesar", "papa john", "denny",
            "ihop", "waffle", "sushi", "chino", "mariscos", "ceviche",
            "tortill", "pupus", "antoj", "loncher", "cocina", "grill",
            "steak", "asado", "barbacoa", "bbq", "carnitas", "birria",
            "buffet", "food", "eat", "diner", "bakery", "donut",
            "helad", "ice cream", "juice", "smoothie", "batido",
            "energetico", "pecuari", "lacteo", "huev", "granja", "agropecuari",
            # Panadería / repostería
            "panes", "panad", "baguett", "croissant", "reposter", "dulceri",
            "confiter", "galleta", "torta", "bizcocho", "pan ", "panka",
            # Consumo general
            "consumo", "aliment", "víveres", "viveres", "tienda", "super",
            "market", "despens", "abastec", "abarrot",
            # Cadenas / marcas locales
            "popeye", "cinnabon", "starbucks", "dunkin", "krispy",
            "quiznos", "five guys", "kfc", "taco bell", "pizza hut",
        ]
        if not any(k in combined for k in keywords_alimentos):
            return ""

        if not hora:
            return ""

        try:
            h, m = map(int, hora.split(":"))
            minutos = h * 60 + m
            if minutos < 11 * 60 + 20:      # antes de 11:20 AM → Desayuno
                return "Desayuno"
            elif minutos < 17 * 60 + 30:    # 11:20 AM – 5:30 PM → Almuerzo
                return "Almuerzo"
            else:                            # después de 5:30 PM → Cena
                return "Cena"
        except (ValueError, AttributeError):
            return ""

    async def _encontrar_frame_resultado(self) -> Optional[Frame]:
        """Encuentra el frame que contiene el resultado de la consulta SAT"""
        for frame in self._frames():
            try:
                txt = await frame.evaluate("() => document.body.innerText")
                if "detalle del documento" in txt.lower() or "monto total" in txt.lower():
                    return frame
            except Exception:
                continue
        return None

    async def _descargar_pdf_externo(self, export_data: dict, monto: str, fecha: str,
                                      hora: str = "", establecimiento: str = "",
                                      tipo_comida: str = "") -> Optional[Path]:
        """Descarga PDF via HTTP externo usando cookies del browser.
        NO toca el estado del navegador, así Nueva Consulta sigue funcionando."""
        # Construir nombre: Q200 28.03.2026 12.58 SERVICENTRO KM 12 Almuerzo.pdf
        partes = [monto, fecha]
        if hora:
            partes.append(hora.replace(":", "."))
        if establecimiento:
            # Limpiar caracteres no válidos para nombre de archivo
            est_clean = re.sub(r'[<>:"/\\|?*]', '', establecimiento).strip()
            if est_clean:
                partes.append(est_clean)
        if tipo_comida:
            partes.append(tipo_comida)
        nombre = " ".join(partes) + ".pdf"
        pdf_path = SCREENSHOTS_DIR / nombre
        try:
            # Extraer cookies y User-Agent del browser
            cookies = await self._context.cookies()
            cookie_str = "; ".join(f"{c['name']}={c['value']}" for c in cookies)
            ua = await self._page.evaluate("() => navigator.userAgent")

            logger.info(f"HTTP externo → {export_data['formAction']}")
            logger.info(f"Cookies: {len(cookies)}, exportId: {export_data['exportId']}")

            # Crear cliente HTTP separado (no afecta al browser)
            request_ctx = await self._playwright.request.new_context(
                extra_http_headers={
                    "Cookie": cookie_str,
                    "User-Agent": ua,
                    "Referer": export_data['formAction'],
                    "Origin": "https://farm3.sat.gob.gt",
                }
            )

            # Construir form data
            form = {
                export_data['formId']: export_data['formId'],
                'javax.faces.ViewState': export_data['viewState'],
                export_data['exportId']: export_data['exportId'],
            }
            for k, v in export_data.get('hiddens', {}).items():
                form[k] = v

            try:
                response = await request_ctx.post(export_data['formAction'], form=form)
                status = response.status
                ct = response.headers.get('content-type', '')
                body = await response.body()
            finally:
                await request_ctx.dispose()

            logger.info(f"HTTP externo respuesta: status={status}, content-type={ct}, size={len(body) if body else 0}")

            if body and len(body) > 500 and ('pdf' in ct.lower() or 'octet' in ct.lower() or body[:5] == b'%PDF-'):
                pdf_path.write_bytes(body)
                logger.info(f"PDF descargado (HTTP externo): {nombre} ({len(body)} bytes)")
                return pdf_path
            else:
                # Log primeros bytes para debug
                preview = body[:200].decode('utf-8', errors='replace') if body else '(vacío)'
                logger.warning(f"HTTP externo no retornó PDF. Preview: {preview[:100]}")
                return None

        except Exception as e:
            logger.warning(f"Error descarga PDF externo: {e}")
            return None




# ──────────────────────────────────────────────
# Gestor de sesiones concurrentes
# ──────────────────────────────────────────────
MAX_NAVEGADORES = 5          # máximo navegadores simultáneos
INACTIVIDAD_MAX = 600        # cerrar sesión tras 10 min sin uso

_sesiones: dict = {}         # user_id -> VerificadorSAT
_ultimo_uso: dict = {}       # user_id -> timestamp último uso
_locks_usuario: dict = {}    # user_id -> asyncio.Lock (evita 2 verificaciones del mismo usuario)
_semaforo = None              # se inicializa lazy (necesita event loop)


def _get_semaforo() -> asyncio.Semaphore:
    global _semaforo
    if _semaforo is None:
        _semaforo = asyncio.Semaphore(MAX_NAVEGADORES)
    return _semaforo


def _get_lock_usuario(user_id: str) -> asyncio.Lock:
    if user_id not in _locks_usuario:
        _locks_usuario[user_id] = asyncio.Lock()
    return _locks_usuario[user_id]


async def obtener_verificador(user_id: str) -> VerificadorSAT:
    """Obtiene o crea un verificador para el usuario.
    Limita navegadores simultáneos con semáforo."""
    _ultimo_uso[user_id] = time.time()

    # Si ya hay sesion pero el navegador murio (crash, cerrado por el SO, etc.),
    # cerrarla limpiamente para liberar su slot y crear una nueva. Sin esto, una
    # sesion zombie ocupa un slot del semaforo para siempre.
    if user_id in _sesiones and not _sesiones[user_id].esta_viva():
        logger.warning(f"Sesion de {user_id} tiene navegador muerto — reciclando")
        await cerrar_verificador(user_id)

    if user_id not in _sesiones:
        # Limpiar sesiones inactivas antes de crear una nueva
        await _limpiar_inactivas()
        sem = _get_semaforo()
        await sem.acquire()
        try:
            v = VerificadorSAT()
            await v.iniciar()
            _sesiones[user_id] = v
            logger.info(f"Navegador creado para {user_id} (activos: {len(_sesiones)}/{MAX_NAVEGADORES})")
        except Exception:
            sem.release()
            raise

    return _sesiones[user_id]


async def cerrar_verificador(user_id: str):
    """Cierra la sesion del usuario y libera su slot del semaforo.

    Robusto: saca la sesion del dict ANTES de cerrar el navegador, asi aunque
    cerrar() falle o cuelgue, el slot del semaforo SIEMPRE se libera (finally)
    y no queda una sesion zombie ocupando lugar.
    """
    v = _sesiones.pop(user_id, None)
    _ultimo_uso.pop(user_id, None)
    _locks_usuario.pop(user_id, None)
    if v is None:
        return  # no habia sesion activa: nada que liberar
    try:
        await v.cerrar()
    except Exception as e:
        logger.warning(f"Error cerrando navegador de {user_id} (se libera el slot igual): {e}")
    finally:
        try:
            _get_semaforo().release()
        except ValueError:
            # release() sobre un semaforo ya en su valor maximo lanza ValueError.
            # Significaria un doble-release; lo ignoramos para no romper.
            logger.debug(f"Semaforo ya estaba liberado para {user_id}")
        logger.info(f"Navegador cerrado para {user_id} (activos: {len(_sesiones)})")


async def _limpiar_inactivas():
    """Cierra sesiones que llevan más de INACTIVIDAD_MAX sin usarse."""
    ahora = time.time()
    inactivos = [uid for uid, t in _ultimo_uso.items()
                 if ahora - t > INACTIVIDAD_MAX and uid in _sesiones]
    for uid in inactivos:
        logger.info(f"Cerrando sesión inactiva: {uid}")
        await cerrar_verificador(uid)
