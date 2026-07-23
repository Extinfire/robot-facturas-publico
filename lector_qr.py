"""
Lectura de códigos QR de facturas FEL guatemaltecas.

El QR de una FEL casi siempre codifica (directo o dentro de una URL del
verificador) el "Número de Autorización": un UUID 8-4-4-4-12. De ese UUID:

    - El PRIMER segmento (8 hex) ES la SERIE de la factura.
    - El UUID completo es el número de autorización.

Muchos QR además incluyen, como parámetros de la URL, el NIT del emisor, el
número de DTE, el monto y la fecha.

Leer el QR da estos campos SIN errores de OCR — justo los que hacen que el
SAT rechace (serie/numero mal leídos). Por eso se intenta ANTES/junto al OCR
y, cuando hay QR válido, sus campos pisan a los del OCR.

Dependencias: intenta pyzbar (mejor lector) y cae a cv2.QRCodeDetector
(ya instalado para el deskew). Si ninguno está, degrada silenciosamente.
"""

import re
import logging
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse, parse_qs

logger = logging.getLogger(__name__)

# ── Backends de lectura (ambos opcionales) ──────────────────────────────
# Guardamos el motivo del fallo de import para diagnosticar (sobre todo en
# Windows, donde pyzbar suele instalarse pero fallar al cargar la DLL de zbar
# por falta del "Visual C++ Redistributable").
_PYZBAR_ERROR = ""
try:
    from pyzbar.pyzbar import decode as _pyzbar_decode  # type: ignore
    _PYZBAR = True
except Exception as _e:
    _PYZBAR = False
    _PYZBAR_ERROR = f"{type(_e).__name__}: {_e}"
    logging.getLogger(__name__).warning(
        f"pyzbar no se pudo cargar ({_PYZBAR_ERROR}). "
        f"Se usara cv2 (lector de QR mas debil). En Windows suele faltar el "
        f"'Visual C++ Redistributable 2013': instalalo desde "
        f"https://www.microsoft.com/download/details.aspx?id=40784"
    )

try:
    import cv2  # type: ignore
    import numpy as _np  # type: ignore
    _CV2 = True
except Exception:
    _CV2 = False


# UUID FEL: 8-4-4-4-12 hex. La serie es el primer grupo de 8.
_RE_UUID = re.compile(
    r'([0-9A-Fa-f]{8})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{4})-'
    r'([0-9A-Fa-f]{4})-([0-9A-Fa-f]{12})'
)
# NIT: 5-9 digitos con guion+verificador opcional (digito o K). Los NIT
# guatemaltecos varian: hay de 5 digitos (90494-5) hasta 8+ (12000430-5).
_RE_NIT = re.compile(r'\b(\d{5,9}-?[0-9Kk]?)\b')


def _variantes_cv2(img):
    """Genera versiones preprocesadas de la imagen para mejorar la deteccion
    de QR en fotos malas (sobre-expuestas, oscuras, pequeñas, con sombra).
    Cada variante es un ndarray que se le pasa al detector."""
    variantes = [img]
    try:
        gris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        variantes.append(gris)
        # Ampliar x2: ayuda cuando el QR ocupa pocos pixeles en la foto.
        h, w = gris.shape[:2]
        if max(h, w) < 2500:
            grande = cv2.resize(gris, (w * 2, h * 2), interpolation=cv2.INTER_CUBIC)
            variantes.append(grande)
        # Umbral Otsu: recupera QR lavado por sobre-exposicion o sombra.
        _, otsu = cv2.threshold(gris, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        variantes.append(otsu)
        # Umbral adaptativo: para luz desigual.
        adap = cv2.adaptiveThreshold(
            gris, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 31, 5,
        )
        variantes.append(adap)
    except Exception as e:
        logger.debug(f"_variantes_cv2 fallo: {e}")
    return variantes


def leer_qr_texto(image_path) -> tuple:
    """Devuelve (backend, lista_de_textos) decodificados de los QR en la imagen.

    backend es 'pyzbar', 'cv2' o '' (nada encontrado) — indica que lector
    decodifico el QR, para que el llamador pueda decidir cuanto confiar en
    el resultado (pyzbar aplica correccion Reed-Solomon completa y es mucho
    mas robusto que el detector de cv2 en fotos degradadas).

    Estrategia (de mejor a peor lector), probando varias versiones
    preprocesadas para fotos malas. Nunca lanza: si falla, retorna ('', [])."""
    image_path = str(image_path)
    textos: list = []

    # 1. pyzbar sobre varias versiones preprocesadas (es el lector mas robusto).
    if _PYZBAR:
        try:
            from PIL import Image
            with Image.open(image_path) as _im:
                base = _im.convert('RGB')
            intentos = [base]
            # Escala de grises + ampliada (pyzbar mejora con QR grandes).
            try:
                g = base.convert('L')
                intentos.append(g)
                if max(base.size) < 2500:
                    intentos.append(g.resize((base.size[0] * 2, base.size[1] * 2)))
            except Exception:
                pass
            for im in intentos:
                for obj in _pyzbar_decode(im):
                    try:
                        t = obj.data.decode('utf-8', errors='replace').strip()
                    except Exception:
                        t = ""
                    if t and t not in textos:
                        textos.append(t)
                if textos:
                    return "pyzbar", textos
        except Exception as e:
            logger.debug(f"pyzbar fallo: {e}")

    # 2. cv2 sobre varias variantes preprocesadas.
    if _CV2:
        try:
            img = cv2.imread(image_path)
            if img is not None:
                detector = cv2.QRCodeDetector()
                for var in _variantes_cv2(img):
                    try:
                        ok, decoded, _pts, _ = detector.detectAndDecodeMulti(var)
                        if ok and decoded:
                            for t in decoded:
                                t = (t or "").strip()
                                if t and t not in textos:
                                    textos.append(t)
                        if not textos:
                            t, _pts, _ = detector.detectAndDecode(var)
                            t = (t or "").strip()
                            if t and t not in textos:
                                textos.append(t)
                    except Exception:
                        continue
                    if textos:
                        break
        except Exception as e:
            logger.debug(f"cv2 QR fallo: {e}")

    return ("cv2" if textos else ""), textos


def parsear_qr_fel(texto: str) -> dict:
    """Extrae {autorizacion, serie, nit, numero} de un texto de QR FEL.

    Soporta:
    - URL con parametros (numero/autorizacion, emisor/nit, dte/numeroDte, etc.)
    - Texto plano que contenga el UUID en alguna parte.
    Devuelve solo las claves que pudo determinar con confianza.
    """
    datos: dict = {}
    if not texto:
        return datos

    # 1. UUID de autorizacion (de aqui sale la serie con certeza).
    m = _RE_UUID.search(texto)
    if m:
        datos['autorizacion'] = m.group(0).upper()
        datos['serie'] = m.group(1).upper()  # primer segmento = serie

    # 2. Si es URL, mirar parametros conocidos.
    try:
        parsed = urlparse(texto)
        if parsed.query:
            q = parse_qs(parsed.query)
            # nombres comunes segun certificador (case-insensitive)
            ql = {k.lower(): v for k, v in q.items()}

            def _primero(*keys):
                for k in keys:
                    if k in ql and ql[k]:
                        return ql[k][0].strip()
                return ""

            # NIT del emisor
            nit_q = _primero('emisor', 'nitemisor', 'nit_emisor', 'nit')
            if nit_q:
                mnit = _RE_NIT.search(nit_q.replace(' ', ''))
                if mnit:
                    datos['nit'] = mnit.group(1).upper()
            # Numero de DTE / documento
            num_q = _primero('numero', 'numerodte', 'numero_dte', 'dte',
                             'numdoc', 'numerodocumento')
            # Cuidado: 'numero' a veces ES el UUID de autorizacion. Solo lo
            # tomamos como numero de DTE si es puramente numerico.
            if num_q and num_q.isdigit():
                datos['numero'] = num_q
            # Autorizacion explicita (si no salio del UUID del texto)
            if 'autorizacion' not in datos:
                aut_q = _primero('autorizacion', 'uuid', 'numeroautorizacion')
                ma = _RE_UUID.search(aut_q or "")
                if ma:
                    datos['autorizacion'] = ma.group(0).upper()
                    datos.setdefault('serie', ma.group(1).upper())
    except Exception as e:
        logger.debug(f"parsear_qr_fel: no era URL parseable ({e})")

    return datos


def extraer_datos_qr(image_path) -> dict:
    """Lee el/los QR de la imagen y devuelve el mejor conjunto de datos FEL.

    Retorna {} si no hay QR legible. Si hay varios QR, prioriza el que tenga
    UUID de autorizacion (es el de la FEL)."""
    if not (_PYZBAR or _CV2):
        return {}
    backend, textos = leer_qr_texto(image_path)
    if not textos:
        logger.info(f"📷 No se pudo leer ningun QR (backend: {backend_disponible()})")
        return {}
    mejor: dict = {}
    for t in textos:
        logger.info(f"📷 QR detectado: {t[:120]}")
        d = parsear_qr_fel(t)
        # Preferimos el que tenga autorizacion (UUID FEL).
        if 'autorizacion' in d and 'autorizacion' not in mejor:
            mejor = d
        elif not mejor:
            mejor = d
    if mejor:
        # pyzbar aplica correccion de errores Reed-Solomon completa al decodificar;
        # si paso esa validacion, el contenido es confiable (no es lectura de texto
        # ambigua como el OCR). cv2 es menos robusto en fotos degradadas, asi que
        # el llamador debe seguir tratandolo como pista, no como fuente de verdad.
        mejor['_qr_backend'] = backend
        logger.info(f"✅ Datos del QR: {mejor}")
    return mejor


def backend_disponible() -> str:
    """Describe que backend de QR esta activo (para logs de arranque).
    Si pyzbar fallo al cargar, incluye el motivo entre parentesis."""
    if _PYZBAR:
        return "pyzbar"
    if _CV2:
        if _PYZBAR_ERROR:
            return f"cv2 (pyzbar fallo: {_PYZBAR_ERROR})"
        return "cv2 (pyzbar no instalado)"
    return "ninguno"
