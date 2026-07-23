"""
Sistema inteligente de extracción de facturas SAT con Ollama Vision + Catálogo
Flujo: Foto → OCR Llava-Phi → Parse NIT → Validar
       Si falla: Busca razón social → Catálogo → Confirma usuario
"""

import os
import io
import json
import base64
import re
import threading
import requests
from pathlib import Path
from typing import Optional, Tuple
from dataclasses import dataclass, asdict
import logging
from PIL import Image

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

# Lectura de QR (alta confianza: serie/numero/NIT exactos del codigo de la FEL).
# Es opcional: si el modulo o sus backends no estan, el flujo sigue solo con OCR.
try:
    from lector_qr import extraer_datos_qr as _extraer_datos_qr
    from lector_qr import backend_disponible as _qr_backend
    _QR_DISPONIBLE = True
except Exception:
    _QR_DISPONIBLE = False
    def _extraer_datos_qr(_):
        return {}
    def _qr_backend():
        return "ninguno"

# Logging
logging.basicConfig(
    level=logging.INFO,
    handlers=[logging.StreamHandler()],
)
logging.getLogger().handlers[0].stream = open(
    __import__('sys').stdout.fileno(),
    mode='w', encoding='utf-8', errors='replace', closefd=False
)
logger = logging.getLogger(__name__)

# Rutas
PROJECT_ROOT = Path(__file__).parent
DATA_DIR = PROJECT_ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

CATALOG_FILE = DATA_DIR / "empresas.json"
CACHE_FILE = DATA_DIR / "extraction_cache.json"
APRENDIZAJE_FILE = DATA_DIR / "aprendizaje.json"


def _guardar_json_atomico(ruta: Path, data) -> bool:
    """Guarda datos JSON de forma atómica usando archivo temporal.
    Nunca deja el archivo original corrupto o a medias."""
    tmp = ruta.with_suffix('.tmp')
    try:
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        tmp.replace(ruta)
        return True
    except Exception as e:
        logger.error(f"Error guardando {ruta.name} de forma atomica: {e}")
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass
        return False

# Config Ollama local
OLLAMA_BASE_URL = "http://localhost:11434"
OLLAMA_MODEL = "llama3.2-vision"  # Modelo multimodal local

# Modelos Anthropic: Haiku por defecto (rápido/barato), Sonnet para escalación
# cuando Haiku falla en facturas degradadas (papel arrugado, tinta gastada).
HAIKU_MODEL = "claude-haiku-4-5-20251001"
SONNET_MODEL = "claude-sonnet-4-5-20250929"
OCR_TIMEOUT = 300

# Config SAT
SAT_API_URL = "https://consulta.sat.gob.gt/consultas/api3/Contribuyente"
SAT_TIMEOUT = 10

# NITs que siempre se excluyen como emisor (ej. comprador frecuente, bancos, cadenas grandes)
# Guardar aquí el NÚMERO BASE del NIT (sin dígito verificador ni guión)
NITS_EXCLUIDOS: set = {
    "5074333",   # Banco Azteca         — NIT completo: 5074333-3
    "657697",    # Elektra de Guatemala — NIT completo: 657697-4 (aparece como COMPRADOR en sus recibos)
}


def _es_nit_excluido(nit_raw: str) -> bool:
    """Devuelve True si el número base del NIT está en NITS_EXCLUIDOS.

    Acepta varios formatos: '5074333-3', '50743333', '5074333' (solo base).
    Para entradas de dígitos puros considera ambas interpretaciones:
    el string completo como base, y el string sin el último dígito (que
    podría ser el verificador). Si cualquiera matchea, devuelve True.
    """
    if not nit_raw:
        return False
    limpio = nit_raw.strip().upper().replace(" ", "")
    # Caso 1: tiene separador (guion/punto/slash) → la base es lo de antes
    base_pre_sep = re.split(r'[\.\-/]', limpio)[0]
    if base_pre_sep in NITS_EXCLUIDOS:
        return True
    # Caso 2: número puro → probar AMBAS interpretaciones
    if limpio.isdigit() and len(limpio) >= 5:
        # 2a. Tratar el string como base completa (el último dígito NO es verificador)
        if limpio in NITS_EXCLUIDOS:
            return True
        # 2b. Tratar el último dígito como verificador, base = todo menos el último
        if limpio[:-1] in NITS_EXCLUIDOS:
            return True
    return False

# Resolución máxima para OCR — 1600px captura mejor números pequeños que 1024px
# 1568px es el borde largo óptimo de la API de Anthropic: imágenes más
# grandes las reescala el servidor (perdiendo nitidez en un resample que no
# controlamos). Quedarnos justo en ese límite evita el doble reescalado.
OCR_MAX_SIZE = 1568

# Prompt base para facturas FEL guatemaltecas
PROMPT_FEL_BASE = """You are reading a Guatemalan FEL (Factura Electronica en Linea) tax invoice.
Transcribe ALL text visible in the image, exactly character by character.
Pay close attention to:
- NIT (tax ID): format is digits followed by dash and a digit or letter (e.g. 1234567-K, 90494-5)
- SERIE: alphanumeric code after the word SERIE, often a long UUID (e.g. A78B2394-5678-ABCD-0012).
  Some invoices show SERIE and NUMERO together on one line as "DTE | FACTURA |SERIE-NUMERO" (e.g. "63CAC965-1561870338" means serie=63CAC965, numero=1561870338).
- NUMERO or Correlativo or No. de Documento: the invoice number, usually 7-10 digits (e.g. 1561870338)
- Razon social or company name (usually at the top before NIT)
Do NOT skip any number or code. Preserve hyphens and special characters exactly."""


# ─────────────────────────────────────────────
#  PREPROCESAMIENTO DE IMAGEN PARA OCR
#  Pipeline: EXIF -> RGB -> evaluacion calidad -> (opcional) deskew/binarizar
#           -> autocontraste -> contraste/nitidez
#  cv2 + numpy son opcionales: si estan instalados, se activa deskew y
#  binarizacion adaptativa. Si no, se usa solo PIL (funciona igual, menos potente).
# ─────────────────────────────────────────────
try:
    import cv2  # type: ignore
    import numpy as _np  # type: ignore
    _CV2_DISPONIBLE = True
except ImportError:
    _CV2_DISPONIBLE = False
    logger.info("ℹ cv2/numpy no instalados — deskew y binarizacion adaptativa desactivados.")
    logger.info("   Para activarlos: pip install opencv-python numpy")


def _evaluar_calidad_imagen(img: Image.Image) -> dict:
    """Mide blur (si cv2 disponible) y brillo. Devuelve metricas + banderas.

    blur_score: varianza del Laplaciano (estandar de la industria).
        Solo se calcula con cv2 — el FIND_EDGES de PIL no es fiable.
        <100 = muy borrosa, 100-300 = aceptable, >300 = nitida.
    brillo_medio: 0-255. <60 = muy oscura, >225 = sobre-expuesta.
    """
    from PIL import ImageStat
    metricas = {
        "blur_score": None,            # None = no medido (no cv2)
        "brillo_medio": 128.0,
        "muy_borrosa": False,
        "muy_oscura": False,
        "sobre_expuesta": False,
    }
    try:
        gris = img.convert('L')
        # Brillo siempre con PIL
        metricas["brillo_medio"] = round(ImageStat.Stat(gris).mean[0], 1)
        metricas["muy_oscura"] = metricas["brillo_medio"] < 60
        metricas["sobre_expuesta"] = metricas["brillo_medio"] > 225
        # Blur solo con cv2 (Laplacian variance = estandar fiable)
        if _CV2_DISPONIBLE:
            arr = _np.array(gris)
            lap = cv2.Laplacian(arr, cv2.CV_64F)
            blur = float(lap.var())
            metricas["blur_score"] = round(blur, 1)
            metricas["muy_borrosa"] = blur < 100
    except Exception as e:
        logger.warning(f"_evaluar_calidad_imagen fallo: {e}")
    return metricas


def _corregir_inclinacion(img: Image.Image, max_angulo: float = 15.0) -> Image.Image:
    """Detecta angulo de inclinacion via Hough transform y rota la imagen.
    Solo activo si cv2 esta instalado. max_angulo limita rotaciones absurdas."""
    if not _CV2_DISPONIBLE:
        return img
    try:
        arr = _np.array(img.convert('L'))
        # Binarizar para que Hough vea solo bordes de texto
        _, bw = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        # Detectar lineas largas (bordes de texto/tablas)
        bordes = cv2.Canny(bw, 50, 150, apertureSize=3)
        lineas = cv2.HoughLinesP(bordes, 1, _np.pi / 180,
                                 threshold=100, minLineLength=arr.shape[1] // 4, maxLineGap=20)
        if lineas is None or len(lineas) == 0:
            return img
        # Calcular angulo dominante (cercanos a 0° = horizontal)
        angulos = []
        for linea in lineas[:200]:
            x1, y1, x2, y2 = linea[0]
            if x2 == x1:
                continue
            ang = _np.degrees(_np.arctan2(y2 - y1, x2 - x1))
            # Solo angulos cercanos a horizontal (texto)
            if -max_angulo <= ang <= max_angulo:
                angulos.append(ang)
        if not angulos:
            return img
        angulo_mediana = float(_np.median(angulos))
        # Si la inclinacion es minima, no rotar (rotar gratis empeora calidad)
        if abs(angulo_mediana) < 0.5:
            return img
        # Rotar en sentido contrario para enderezar
        img_rotada = img.rotate(angulo_mediana, expand=True, fillcolor=(255, 255, 255))
        logger.info(f"Imagen enderezada {angulo_mediana:+.2f}° (deskew)")
        return img_rotada
    except Exception as e:
        logger.warning(f"_corregir_inclinacion fallo: {e}")
        return img


def _binarizar_adaptivo(img: Image.Image) -> Image.Image:
    """Binarizacion adaptativa para imagenes con bajo contraste o iluminacion desigual.
    Solo activo si cv2 esta instalado. Devuelve la imagen original si no aplica."""
    if not _CV2_DISPONIBLE:
        return img
    try:
        arr = _np.array(img.convert('L'))
        # Threshold adaptativo: cada pixel se compara con su vecindario local
        bw = cv2.adaptiveThreshold(arr, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY, blockSize=31, C=10)
        # Convertir de vuelta a RGB para que el modelo de OCR lo maneje igual
        bw_rgb = cv2.cvtColor(bw, cv2.COLOR_GRAY2RGB)
        return Image.fromarray(bw_rgb)
    except Exception as e:
        logger.warning(f"_binarizar_adaptivo fallo: {e}")
        return img


def _preprocesar_imagen_ocr(img: Image.Image, calidad_out: Optional[dict] = None) -> Image.Image:
    """Pipeline completo de preprocesamiento para OCR.

    1. Corrige orientacion EXIF (rotaciones de smartphone)
    2. Convierte a RGB
    3. Evalua calidad (blur + brillo) — guarda en calidad_out si se pasa
    4. Corrige inclinacion (cv2)
    5. Si la imagen es muy oscura/desigual, binariza adaptivamente (cv2)
    6. Autocontraste PIL si el contraste global es bajo
    7. Realce de contraste y nitidez (PIL)
    """
    from PIL import ImageEnhance, ImageOps
    # 1. EXIF
    try:
        exif = img._getexif()
        if exif:
            orientation = exif.get(274)
            rotaciones = {3: 180, 6: 270, 8: 90}
            if orientation in rotaciones:
                img = img.rotate(rotaciones[orientation], expand=True)
                logger.info(f"Imagen rotada {rotaciones[orientation]}° segun EXIF")
    except Exception:
        pass
    # 2. RGB
    if img.mode != 'RGB':
        img = img.convert('RGB')
    # 3. Evaluar calidad antes de tocar nada
    metricas = _evaluar_calidad_imagen(img)
    if calidad_out is not None:
        calidad_out.update(metricas)
    if metricas["muy_borrosa"]:
        logger.warning(f"⚠ Imagen MUY BORROSA (blur_score={metricas['blur_score']:.0f}); el OCR puede fallar.")
    if metricas["muy_oscura"]:
        logger.warning(f"⚠ Imagen MUY OSCURA (brillo={metricas['brillo_medio']:.0f}/255).")
    if metricas["sobre_expuesta"]:
        logger.warning(f"⚠ Imagen SOBRE-EXPUESTA (brillo={metricas['brillo_medio']:.0f}/255).")
    # 4. Deskew (cv2)
    img = _corregir_inclinacion(img)
    # 5. Binarizacion adaptativa solo si la imagen tiene problemas de iluminacion
    if metricas["muy_oscura"] or metricas["sobre_expuesta"]:
        img = _binarizar_adaptivo(img)
    # 6. Autocontraste PIL: estira el histograma si el contraste es global bajo
    try:
        img = ImageOps.autocontrast(img, cutoff=1)
    except Exception:
        pass
    # 7. Realce final (mantiene los valores originales que ya funcionaban)
    img = ImageEnhance.Contrast(img).enhance(1.4)
    img = ImageEnhance.Sharpness(img).enhance(1.8)
    return img


@dataclass
class ExtraccionFactura:
    """Datos extraídos de una factura"""
    nit: str
    serie: str
    numero_factura: str
    razon_social: Optional[str] = None
    metodo_extraccion: str = "ocr_directo"  # ocr_directo, catalogo, manual
    confianza: float = 1.0
    texto_ocr: str = ""  # Texto OCR crudo, usado por el sistema de aprendizaje
    # Historial: rellenado por HistorialEmisores cuando hay registro previo del NIT.
    # Estructura: {"vistas": int, "serie_habitual": str, "duplicado_exacto": bool,
    #              "serie_corregida_de": str|None, "razon_social_conocida": str}
    historial: Optional[dict] = None
    # Calidad de la imagen procesada (blur, brillo, banderas accionables).
    # Permite al cliente avisar al usuario "foto muy borrosa, repite la toma".
    calidad_imagen: Optional[dict] = None
    # True si el OCR detectó que la factura fue emitida a CONSUMIDOR FINAL
    # (en vez de a un NIT específico). Estas facturas no sirven para reporte
    # de viáticos porque no son nominativas. El cliente debe rechazarlas
    # antes de gastar una consulta en el SAT.
    receptor_cf: bool = False

    def to_dict(self):
        return asdict(self)


# ─────────────────────────────────────────────
#  HISTORIAL DE EMISORES (caché por usuario)
#  Lee registro.json del usuario para aprender que NITs/series ya conoce
#  y mejorar precision: boost de confianza, correccion de typos OCR en serie,
#  deteccion temprana de duplicado exacto antes de llamar al SAT.
# ─────────────────────────────────────────────

class HistorialEmisores:
    """Indexa el registro.json del usuario para aprender sus emisores frecuentes.
    Lazy: solo lee disco la primera vez que se consulta."""

    def __init__(self, registro_path: Optional[Path] = None):
        self.registro_path = Path(registro_path) if registro_path else None
        self._index: Optional[dict] = None

    def _build(self) -> dict:
        if self._index is not None:
            return self._index
        self._index = {
            "por_nit": {},        # nit -> {"series": {serie: count}, "count": int, "razon_social": str, "last": str}
            "duplicados": set(),  # set de tuplas (nit, serie, numero) ya registradas
        }
        if not self.registro_path or not self.registro_path.exists():
            return self._index
        try:
            with open(self.registro_path, "r", encoding="utf-8") as f:
                registros = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            logger.warning(f"HistorialEmisores: no se pudo leer {self.registro_path}: {e}")
            return self._index
        for r in registros:
            nit = (r.get("nit") or "").strip()
            if not nit:
                continue
            serie = (r.get("serie") or "").strip()
            numero = (r.get("numero") or "").strip()
            entry = self._index["por_nit"].setdefault(nit, {
                "series": {}, "count": 0, "razon_social": "", "last": "",
            })
            entry["count"] += 1
            if serie:
                entry["series"][serie] = entry["series"].get(serie, 0) + 1
            rs = (r.get("razon_social") or "").strip()
            if rs and not entry["razon_social"]:
                entry["razon_social"] = rs
            fp = r.get("fecha_proceso", "")
            if fp > entry["last"]:
                entry["last"] = fp
            if serie and numero:
                self._index["duplicados"].add((nit, serie, numero))
        return self._index

    def vistas_por_nit(self, nit: str) -> int:
        if not nit:
            return 0
        return self._build()["por_nit"].get(nit, {}).get("count", 0)

    def razon_social_conocida(self, nit: str) -> str:
        """NIT → razon_social ya vista para ese NIT (estable por emisor)."""
        if not nit:
            return ""
        return self._build()["por_nit"].get(nit, {}).get("razon_social", "")

    def es_duplicado_exacto(self, nit: str, serie: str, numero: str) -> bool:
        """Detecta si (NIT, serie, numero) ya esta registrado. La serie es
        unica por documento en FEL, asi que esto solo matchea si es la misma
        factura exacta."""
        if not (nit and serie and numero):
            return False
        return (nit.strip(), serie.strip(), numero.strip()) in self._build()["duplicados"]

    def boost_confianza(self, nit: str) -> float:
        """Incremento de confianza segun cuantas veces hemos visto este NIT.
        NOTA: el NIT es lo unico estable por emisor; la serie cambia por
        documento, por eso no se usa aqui."""
        if not nit:
            return 0.0
        info = self._build()["por_nit"].get(nit)
        if not info:
            return 0.0
        boost = 0.0
        if info["count"] >= 1:
            boost += 0.02
        if info["count"] >= 5:
            boost += 0.02
        if info["count"] >= 20:
            boost += 0.01
        return boost


class CatalogoEmpresas:
    """Gestiona el catálogo de empresas conocidas"""
    
    def __init__(self, file_path: Path = CATALOG_FILE):
        self.file_path = file_path
        self.empresas = self._cargar()
        # RLock: la lectura del catálogo (buscar_por_razon_social) corre en
        # hilos worker (OCR via to_thread) mientras la escritura
        # (reconciliar_con_sat/agregar_empresa) corre en el event loop. Sin
        # esto, iterar self.empresas mientras otro hilo lo modifica lanza
        # 'dictionary changed size during iteration'. Reentrante porque
        # reconciliar_con_sat llama internamente a agregar_empresa.
        self._lock = threading.RLock()

    def _cargar(self) -> dict:
        """Carga catálogo desde JSON"""
        if not self.file_path.exists():
            logger.info(f"Creando catálogo vacío en {self.file_path}")
            self._guardar({})
            return {}

        try:
            with open(self.file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            # Migración: el catálogo antes guardaba "serie" por empresa.
            # Serie y número cambian por factura, así que se descartan.
            mutado = False
            for clave, info in data.items():
                if isinstance(info, dict) and ("serie" in info or "numero" in info):
                    info.pop("serie", None)
                    info.pop("numero", None)
                    mutado = True
            # Limpieza: eliminar entradas obviamente basura (OCR de encabezados,
            # asteriscos de markdown, placeholders de formulario, etc.).
            claves_basura = []
            for clave, info in list(data.items()):
                rs = info.get("razon_social", clave) if isinstance(info, dict) else clave
                if not type(self)._razon_social_valida(rs):
                    claves_basura.append(clave)
            for k in claves_basura:
                basura = data.pop(k)
                logger.warning(f"Catálogo: eliminando entrada basura {k!r} -> {basura}")
                mutado = True
            if mutado:
                logger.info("Migrando/limpiando catálogo")
                _guardar_json_atomico(self.file_path, data)
            return data
        except Exception as e:
            logger.error(f"Error cargando catálogo: {e}")
            return {}
    
    def _guardar(self, data: dict = None):
        """Guarda catálogo en JSON"""
        if data is None:
            data = self.empresas
        
        _guardar_json_atomico(self.file_path, data)
        logger.info(f"Catálogo guardado: {len(data)} empresas")
    
    # Palabras/tokens tipicos cuando el OCR captura encabezados de factura en
    # vez del nombre del emisor. Si una "razon social" contiene cualquiera de
    # estos, la rechazamos para no contaminar el catalogo.
    _RAZON_SOCIAL_BLACKLIST = (
        "ESTABLISHMENT INFO", "ESTABLISHMENT", "INFO:",
        "DOCUMENTO TRIBUTARIO", "DOCUMENTO ELECTRONICO",
        "FACTURA ELECTRONICA", "NUMERO DE AUTORIZACION",
        "FECHA EMISION", "SERIE:", "NUMERO:",
        "NIT:", "RAZON SOCIAL:", "NOMBRE:", "DIRECCION:",
        "NIT EMISOR", "COD. ESTABLECIMIENTO",
    )

    @classmethod
    def _razon_social_valida(cls, razon_social: str) -> bool:
        """Detecta razones sociales que son claramente texto crudo de OCR
        (encabezados, placeholders, asteriscos de markdown, etc.)."""
        if not razon_social:
            return False
        limpio = razon_social.strip()
        if len(limpio) < 3 or len(limpio) > 120:
            return False
        # Asteriscos dobles: Haiku a veces devuelve markdown bold
        if "**" in limpio:
            return False
        # Proporcion excesiva de no-letras (e.g. "02-013 ZONA 3")
        letras = sum(1 for c in limpio if c.isalpha())
        if letras < max(3, len(limpio) // 3):
            return False
        upper = limpio.upper()
        for token in cls._RAZON_SOCIAL_BLACKLIST:
            if token in upper:
                return False
        return True

    def agregar_empresa(self, razon_social: str, nit: str, serie: str = "",
                        fuente: str = "ocr"):
        """Agrega o actualiza una empresa en el catálogo.

        - `serie` se acepta por compatibilidad pero se ignora (cambia por factura).
        - `fuente`: 'ocr' (lectura inicial) o 'sat' (confirmado por SAT).
          Los registros marcados como 'sat' tienen prioridad: no se sobrescriben
          con datos 'ocr', pero sí entre sí (última versión SAT manda).

        Regla dura: solo aceptamos fuente='sat'. El OCR es demasiado ruidoso
        y antes metia basura tipo "**ESTABLISHMENT INFO:**" al catalogo.
        """
        if fuente != "sat":
            logger.info(
                f"Ignorando alta de catalogo sin confirmacion SAT: "
                f"'{razon_social}' ({nit}) [{fuente}]"
            )
            return False
        if not self._razon_social_valida(razon_social):
            logger.warning(
                f"Razon social descartada por formato sospechoso: {razon_social!r}"
            )
            return False
        clave = razon_social.strip().lower()
        if not clave or not nit:
            return False

        with self._lock:
            existente = self.empresas.get(clave)
            if existente:
                fuente_previa = existente.get("fuente", "ocr")
                # No degradar un registro confirmado por SAT con uno del OCR
                if fuente_previa == "sat" and fuente != "sat":
                    return False
                # Si el NIT coincide y la fuente no mejora, nada que hacer
                if existente.get("nit") == nit and fuente_previa == fuente:
                    return False
                logger.info(f"Actualizando empresa '{razon_social}': "
                            f"{existente.get('nit')} ({fuente_previa}) → {nit} ({fuente})")

            self.empresas[clave] = {
                "razon_social": razon_social,
                "nit": nit,
                "fuente": fuente,
                "agregada": existente.get("agregada") if existente else (
                    self.file_path.stat().st_ctime if self.file_path.exists() else 0
                ),
            }
            self._guardar()
        if not existente:
            logger.info(f"Empresa agregada: {razon_social} ({nit}) [{fuente}]")
        return True

    def reconciliar_con_sat(self, razon_social_sat: str, nit_sat: str):
        """Cuando el SAT confirma una factura, llama a esto para que el catálogo
        quede con la razón social OFICIAL y se deduplique si había entradas
        previas del mismo NIT con nombres distintos leídos por OCR."""
        if not razon_social_sat or not nit_sat:
            return
        # Buscar entradas previas con el mismo NIT pero razón social distinta
        clave_sat = razon_social_sat.strip().lower()
        with self._lock:
            duplicados = [
                k for k, v in self.empresas.items()
                if v.get("nit") == nit_sat and k != clave_sat
            ]
            for k in duplicados:
                anterior = self.empresas.pop(k)
                logger.info(f"Deduplicando catálogo: '{anterior.get('razon_social')}' "
                            f"→ '{razon_social_sat}' (mismo NIT {nit_sat})")
            self.agregar_empresa(razon_social_sat, nit_sat, fuente="sat")
    
    def buscar_por_razon_social(self, razon_social: str, umbral_similitud: float = 0.85) -> Optional[dict]:
        """
        Busca empresa por razón social con búsqueda fuzzy básica
        Si exacta: retorna immediately
        Si parcial: retorna mejores matches
        """
        clave = razon_social.strip().lower()

        # Snapshot bajo lock: copiamos las entradas para iterar sin riesgo de
        # 'dictionary changed size during iteration' si otro hilo escribe el
        # catálogo. El cálculo de similitud (costoso) se hace fuera del lock.
        with self._lock:
            if clave in self.empresas:
                return self.empresas[clave]
            items = list(self.empresas.items())

        # Búsqueda parcial (contiene)
        coincidencias = []
        for k, v in items:
            # Búsqueda por palabras clave
            if self._similitud_simple(clave, k) > umbral_similitud:
                coincidencias.append((v, self._similitud_simple(clave, k)))

        if coincidencias:
            # Retorna la mejor coincidencia
            mejor = max(coincidencias, key=lambda x: x[1])
            logger.info(f"Encontrado por similitud: {mejor[0]['razon_social']} ({mejor[1]:.2%})")
            return mejor[0]
        
        logger.warning(f"No encontrado en catálogo: {razon_social}")
        return None
    
    @staticmethod
    def _similitud_simple(s1: str, s2: str) -> float:
        """Similitud combinada: Jaccard de palabras + SequenceMatcher de caracteres"""
        from difflib import SequenceMatcher

        if not s1 or not s2:
            return 0.0

        palabras1 = set(s1.split())
        palabras2 = set(s2.split())

        # Jaccard por palabras
        interseccion = len(palabras1 & palabras2)
        union = len(palabras1 | palabras2)
        jaccard = interseccion / union if union > 0 else 0.0

        # Similitud de caracteres entre palabras largas
        char_bonus = 0.0
        for p1 in palabras1:
            if len(p1) < 5:
                continue
            for p2 in palabras2:
                if len(p2) < 5:
                    continue
                ratio = SequenceMatcher(None, p1, p2).ratio()
                char_bonus = max(char_bonus, ratio * 0.6)

        return min(1.0, jaccard + char_bonus)
    
    def listar_empresas(self) -> list:
        """Lista todas las empresas en catálogo"""
        with self._lock:
            return list(self.empresas.values())


class ExtractorOllama:
    """Realiza OCR usando Ollama Vision (Llava-Phi)"""
    
    def __init__(self, base_url: str = OLLAMA_BASE_URL, model: str = OLLAMA_MODEL):
        self.base_url = base_url
        self.model = model
        self.disponible = False
        self._verificar_ollama()
    
    def _verificar_ollama(self):
        """Verifica que Ollama está corriendo"""
        try:
            r = requests.get(f"{self.base_url}/api/tags", timeout=5)
            if r.status_code == 200:
                logger.info("✓ Ollama está corriendo")
            else:
                raise Exception("Ollama no responde")
        except Exception as e:
            logger.warning(f"✗ Ollama no disponible (se usará solo Claude): {e}")
            self.disponible = False
            return
        self.disponible = True
    
    def extraer_texto(self, image_path: str, ejemplos: str = "") -> str:
        """
        Extrae texto de una imagen usando Llava-Phi
        Retorna el texto completo extraído (OCR)
        """
        try:
            # Preprocesar imagen y redimensionar a 1600px para mejor lectura de números
            img = Image.open(image_path)
            img = _preprocesar_imagen_ocr(img)
            img.thumbnail((OCR_MAX_SIZE, OCR_MAX_SIZE), Image.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=95)
            image_data = base64.b64encode(buf.getvalue()).decode("utf-8")

            # Prompt especifico para facturas FEL guatemaltecas
            prompt = PROMPT_FEL_BASE
            if ejemplos:
                prompt += (
                    "\n\nKNOWN ISSUERS (reference only — match by NAME to confirm the issuer NIT; "
                    "does NOT indicate the serie/numero of THIS invoice):\n"
                    f"{ejemplos}\n"
                )
            prompt += "\nNow transcribe ALL text from the invoice image:"
            
            # Llamar a Ollama con streaming para evitar timeout por inactividad
            payload = {
                "model": self.model,
                "prompt": prompt,
                "images": [image_data],
                "stream": True
            }

            logger.info(f"Enviando OCR a Ollama ({self.model})...")
            response = requests.post(
                f"{self.base_url}/api/generate",
                json=payload,
                timeout=OCR_TIMEOUT,
                stream=True
            )

            if response.status_code != 200:
                raise Exception(f"Ollama error: {response.text}")

            # Acumular chunks del stream
            texto = ""
            for line in response.iter_lines(chunk_size=None):
                if line:
                    chunk = json.loads(line)
                    texto += chunk.get("response", "")
                    if chunk.get("done"):
                        break

            logger.info(f"OCR completado: {len(texto)} caracteres")
            return texto
        
        except requests.Timeout:
            logger.error("Timeout en Ollama (modelo muy lento)")
            raise
        except Exception as e:
            logger.error(f"Error OCR: {e}")
            raise
    
    def extraer_json(self, image_path: str) -> dict:
        """
        Extrae información estructurada (JSON) de una factura
        Intenta 2 veces si falla JSON parsing
        """
        for intento in range(2):
            try:
                texto = self.extraer_texto(image_path)
                
                # Intentar parsear JSON
                # A veces Llava envuelve en ```json ... ```
                if "```json" in texto:
                    json_str = texto.split("```json")[1].split("```")[0]
                elif "```" in texto:
                    json_str = texto.split("```")[1].split("```")[0]
                else:
                    json_str = texto
                
                datos = json.loads(json_str)
                logger.info(f"JSON extraído (intento {intento + 1})")
                return datos
            
            except json.JSONDecodeError as e:
                logger.warning(f"JSON inválido (intento {intento + 1}): {e}")
                if intento < 1:
                    logger.info("Reintentando...")
                    continue
                else:
                    logger.error("No se pudo parsear JSON de Ollama")
                    return {}
        
        return {}


class ExtractorHaiku:
    """OCR usando Claude (Anthropic API). Por defecto Haiku; se puede
    instanciar con otro modelo (p.ej. Sonnet) para escalación cuando Haiku
    falla en facturas degradadas."""

    MODEL = HAIKU_MODEL

    def __init__(self, model: str = None):
        api_key = os.getenv("ANTHROPIC_API_KEY")
        if not api_key or not ANTHROPIC_AVAILABLE:
            raise RuntimeError("Anthropic API key no configurada o SDK no instalado")
        self.client = anthropic.Anthropic(api_key=api_key)
        if model:
            self.MODEL = model

    def extraer_texto(self, image_path: str, ejemplos: str = "") -> str:
        """Extrae texto de la factura usando Claude Haiku con prompt especifico FEL."""
        with Image.open(image_path) as _img_src:
            img = _preprocesar_imagen_ocr(_img_src)
            img.thumbnail((OCR_MAX_SIZE, OCR_MAX_SIZE), Image.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=95)
        image_data = base64.b64encode(buf.getvalue()).decode("utf-8")

        prompt = PROMPT_FEL_BASE
        if ejemplos:
            prompt += (
                "\n\nKNOWN ISSUERS (reference only — match by NAME to confirm the issuer NIT; "
                "this does NOT indicate the serie/numero of THIS invoice, which are unique per document):\n"
                f"{ejemplos}\n"
            )
        prompt += "\nTranscribe ALL text from this invoice image:"

        logger.info("Enviando OCR a Claude Haiku...")
        message = self.client.messages.create(
            model=self.MODEL,
            max_tokens=2048,
            temperature=0,  # OCR: determinista, sin "creatividad" en chars ambiguos
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {"type": "base64", "media_type": "image/jpeg", "data": image_data},
                    },
                    {"type": "text", "text": prompt}
                ],
            }]
        )
        texto = message.content[0].text
        logger.info(f"Haiku OCR completado: {len(texto)} caracteres")
        return texto

    def extraer_estructurado(self, image_path: str, ejemplos: str = "") -> dict:
        """Extrae directamente NIT, serie y numero usando Claude Haiku con JSON output.
        Usado como segundo intento cuando el OCR+regex no encuentra todos los campos."""
        with Image.open(image_path) as _img_src:
            img = _preprocesar_imagen_ocr(_img_src)
            img.thumbnail((OCR_MAX_SIZE, OCR_MAX_SIZE), Image.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=95)
        image_data = base64.b64encode(buf.getvalue()).decode("utf-8")

        # Lista de NITs conocidos como RECEPTOR (siempre el comprador, nunca emisor).
        # Construida desde NITS_EXCLUIDOS para no duplicar mantenimiento — agrega
        # aqui el formato canonico con guion-digito para que el modelo lo reconozca.
        _RECEPTORES_HINT = {
            "5074333": "5074333-3 (Banco Azteca)",
            "657697":  "657697-4 (Elektra)",
        }
        _hint_receptores = ", ".join(
            v for k, v in _RECEPTORES_HINT.items() if k in NITS_EXCLUIDOS
        )

        prompt = (
            "This is a Guatemalan FEL (Factura Electronica en Linea) invoice.\n"
            "Extract ONLY the ISSUER (emisor/vendedor) fields — ignore the COMPRADOR/RECEPTOR/CLIENTE section entirely.\n\n"
            "Fields to extract:\n"
            "1. nit: The ISSUER NIT — found at the TOP of the invoice, BEFORE any 'DATOS DEL COMPRADOR' or 'RECEPTOR' section.\n"
            "   Format: digits-digit/letter (e.g. '1234567-K', '90494-5'). NEVER use the NIT from the COMPRADOR/CLIENTE section.\n"
            f"   SPECIFICALLY: these NITs are ALWAYS the receiver, NEVER the issuer — if you see them, the issuer NIT is elsewhere: {_hint_receptores}.\n"
            "2. serie: The SERIE code. Typically 8 HEX characters (0-9, A-F) such as '161273F6', '3F49EC03', '63CAC965'.\n"
            "   Appears AFTER the label 'Serie:' OR matches the first 8 characters of the 'No. Autorizacion' UUID.\n"
            "   MAY contain letters A-F. Always read letters carefully (do not convert F to a digit).\n"
            "3. numero: The invoice number. DIGITS ONLY (no letters), typically 7-12 digits.\n"
            "   Labels: 'No.', 'Numero:', 'No de DTE', 'Correlativo', 'Número de Documento'.\n"
            "   Example: '1620724585', '3019983911', '857885300'.\n"
            "4. razon_social: The ISSUER company name (at the top, before DATOS DEL COMPRADOR)\n\n"
            "CRITICAL RULES:\n"
            "- Read the SERIE and NUMERO ONE CHARACTER AT A TIME, left to right. Do NOT\n"
            "  autocomplete to a value that looks familiar — transcribe exactly what is\n"
            "  printed, even if it seems unusual.\n"
            "- Watch out for these OCR confusions in SERIE (it is hexadecimal): 4<->A, 0<->O/D,\n"
            "  8<->B, 5<->S, 1<->I, 6<->G, F<->E. When unsure between two, pick the one that\n"
            "  keeps the string valid hexadecimal (only 0-9 and A-F).\n"
            "- serie and numero are ALWAYS different values. Never return the same string for both.\n"
            "- serie may contain letters (A-F); numero is digits only.\n"
            "- If serie and numero look similar (e.g. both start with same digits), double-check the image.\n"
            "- If a field is not clearly visible, return null for it rather than guessing.\n\n"
        )
        if ejemplos:
            prompt += (
                "KNOWN ISSUERS (reference only — match by NAME to confirm the issuer NIT and its "
                "check digit; this does NOT indicate the serie/numero of THIS invoice):\n"
                f"{ejemplos}\n\n"
            )
        prompt += (
            "Respond ONLY with a JSON object, no explanation:\n"
            '{"nit": "...", "serie": "...", "numero": "...", "razon_social": "..."}\n'
            "If a field is not visible, use null."
        )

        # El nombre del modelo varia (haiku o sonnet) — loggear el real para
        # que la traza no mienta durante una escalacion.
        _model_short = self.MODEL.split('-')[1] if '-' in self.MODEL else self.MODEL
        logger.info(f"Extraccion estructurada con Claude {_model_short.capitalize()}...")
        message = self.client.messages.create(
            model=self.MODEL,
            max_tokens=512,
            temperature=0,  # OCR: determinista, sin "creatividad" en chars ambiguos
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {"type": "base64", "media_type": "image/jpeg", "data": image_data},
                    },
                    {"type": "text", "text": prompt}
                ],
            }]
        )
        respuesta = message.content[0].text.strip()
        # Limpiar posible markdown
        if "```" in respuesta:
            respuesta = respuesta.split("```")[1].lstrip("json").strip()
            if "```" in respuesta:
                respuesta = respuesta.split("```")[0]
        try:
            datos = json.loads(respuesta)
            logger.info(f"Extraccion estructurada: {datos}")
            return datos
        except Exception:
            logger.warning(f"No se pudo parsear JSON estructurado: {respuesta[:200]}")
            return {}

    def sugerir_campo(self, image_path: str, campo: str, valor_previo: str) -> list:
        """Re-lee la imagen enfocándose SOLO en un campo (nit/serie/numero)
        y devuelve hasta 3 candidatos ordenados por probabilidad.

        Usado bajo demanda cuando el usuario tapea 'Corregir <Campo>': el
        modelo mira la imagen original sabiendo que el valor previo era
        incorrecto y propone alternativas con foco en chars ambiguos.
        """
        descr_campo = {
            'nit': "ISSUER NIT — at the TOP of the invoice, format XXXXXXX-Y (digits, dash, digit or K). NOT the receptor/comprador NIT.",
            'serie': "SERIE — typically 8 HEXADECIMAL characters (0-9 and A-F only). Near the top right, labeled 'Serie:' or as the first segment of the 'No. Autorizacion' UUID.",
            'numero': "INVOICE NUMBER — digits only (no letters), 7-12 digits, labeled 'No.', 'Numero:', 'No de DTE' or 'Correlativo'.",
        }.get(campo, campo)

        with Image.open(image_path) as _img_src:
            img = _preprocesar_imagen_ocr(_img_src)
            img.thumbnail((OCR_MAX_SIZE, OCR_MAX_SIZE), Image.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=95)
        image_data = base64.b64encode(buf.getvalue()).decode("utf-8")

        prompt = (
            f"This is a Guatemalan FEL invoice. Look VERY CAREFULLY at the field below.\n\n"
            f"FIELD: {descr_campo}\n\n"
            f"Previous reading was '{valor_previo}' but it was REJECTED as wrong.\n"
            f"Re-read the image with extreme attention to characters easily confused by OCR:\n"
            f"  0/O, 1/I, 4/A, 5/S, 8/B, F/E, D/0, C/6, Z/2.\n\n"
            f"Give me your TOP 3 most likely candidates for this field, ordered by confidence.\n"
            f"Include the previous reading ONLY if you are now sure it was correct.\n"
            f"If you can only see one or two candidates clearly, return fewer.\n\n"
            f"Respond ONLY with JSON, no explanation:\n"
            f'{{"candidatos": ["valor1", "valor2", "valor3"]}}\n'
        )

        _model_short = self.MODEL.split('-')[1] if '-' in self.MODEL else self.MODEL
        logger.info(f"Sugiriendo campo '{campo}' con {_model_short.capitalize()} (valor previo: {valor_previo!r})")
        try:
            message = self.client.messages.create(
                model=self.MODEL,
                max_tokens=256,
                temperature=0,  # sugerencia de campo: determinista
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": image_data}},
                        {"type": "text", "text": prompt}
                    ],
                }]
            )
            respuesta = message.content[0].text.strip()
            if "```" in respuesta:
                respuesta = re.sub(r'```\w*\n?', '', respuesta).rsplit('```', 1)[0]
            m = re.search(r'\{[\s\S]*"candidatos"[\s\S]*\}', respuesta)
            if not m:
                logger.warning(f"sugerir_campo: respuesta sin JSON: {respuesta[:200]}")
                return []
            datos = json.loads(m.group())
            crudos = datos.get('candidatos', []) or []
            # Limpiar, dedupe, descartar el valor_previo (era el incorrecto)
            previo_norm = (valor_previo or "").strip().upper()
            seen = {previo_norm} if previo_norm else set()
            result = []
            for c in crudos:
                c_clean = str(c).strip()
                if not c_clean:
                    continue
                c_norm = c_clean.upper()
                if c_norm in seen:
                    continue
                seen.add(c_norm)
                result.append(c_clean)
                if len(result) >= 3:
                    break
            logger.info(f"sugerir_campo '{campo}': {result}")
            return result
        except Exception as e:
            logger.error(f"sugerir_campo falló: {e}")
            return []


class ParserFactura:
    """Parsea y valida datos extraídos de facturas"""
    
    # Regex para NIT guatemalteco: 1234567-A (7 dígitos, guión, 1 letra/número)
    REGEX_NIT = r'(\d{7,8})-([A-Z0-9])'
    REGEX_NUMERO_FACTURA = r'(\d{1,2})-(\d{1,2})-(\d{8})'  # O simplemente 00000001

    # Correcciones OCR: letras que se confunden con dígitos
    # O→0, o→0, I→1, i→1, l→1, B→8, S→5, Z→2, G→6, D→0, d→8, q→0, g→8, b→8
    _OCR_DIGIT_MAP = str.maketrans('OoIilBSZGDdqgb', '00111852608088')
    
    @staticmethod
    def _corregir_ocr_numero(texto: str) -> str:
        """Corrige letras que el OCR confunde con dígitos en campos numéricos"""
        return texto.translate(ParserFactura._OCR_DIGIT_MAP)

    # Letras IMPOSIBLES en una serie FEL (hex: 0-9, A-F) y su gemelo hex obvio.
    # Solo incluye caracteres que NO son hex válidos, así que reparar nunca
    # daña una serie correcta (A/B/C/D/E/F jamás se tocan).
    _SERIE_HEX_TWIN = {
        'O': '0', 'Q': '0', 'I': '1', 'L': '1',
        'S': '5', 'G': '6', 'Z': '2', 'T': '7',
    }

    @staticmethod
    def reparar_serie_hex(serie: str) -> str:
        """Repara errores OCR en una serie FEL (siempre hexadecimal).

        Mapea caracteres imposibles-en-hex a su gemelo (O->0, I/L->1, S->5,
        G->6, Z->2, Q->0, T->7) y normaliza a mayúsculas. No modifica
        caracteres que ya son hex válidos, por lo que es seguro aplicarlo
        siempre. Reemplaza al viejo `.replace('O','0')` disperso por el código.
        """
        if not serie:
            return serie
        s = serie.strip().upper()
        return ''.join(ParserFactura._SERIE_HEX_TWIN.get(c, c) for c in s)

    @staticmethod
    def validar_digito_nit(nit_str: str) -> bool:
        """
        Valida el dígito verificador de un NIT guatemalteco.
        Formato entrada: '1234567-K' o '12345678'
        Algoritmo SAT: suma(digito_i * factor_i) % 11
        Si resto=0 → dígito=0, si resto=1 → dígito=K, resto>1 → dígito=11-resto
        """
        nit_str = nit_str.strip().upper().replace(" ", "")
        if "-" in nit_str:
            partes = nit_str.split("-")
            numero = partes[0]
            verificador = partes[1]
        elif len(nit_str) >= 2:
            numero = nit_str[:-1]
            verificador = nit_str[-1]
        else:
            return False

        if not numero.isdigit():
            return False

        # Calcular dígito esperado
        n = len(numero)
        factor_inicial = n + 1
        total = sum(int(d) * (factor_inicial - i) for i, d in enumerate(numero))
        resto = total % 11

        if resto == 0:
            esperado = "0"
        elif resto == 1:
            esperado = "K"
        else:
            esperado = str(11 - resto)

        return verificador == esperado

    @staticmethod
    def limpiar_nit(nit: str) -> Optional[str]:
        """
        Limpia y valida NIT
        Retorna formato: 1234567-A o None si inválido
        Acepta: 1234567-K  o  12345678  (sin guión, último char es dígito verificador)
        """
        if not nit:
            return None

        nit = nit.strip().upper().replace(" ", "")
        # Normalizar separadores OCR: puntos, slashes → guión
        nit = re.sub(r'(\d)[\.\/]([A-Z0-9])$', r'\1-\2', nit)
        nit = nit.replace(".", "")

        # Con guión: 90494-5 o 1234567-K (4-8 dígitos)
        match = re.search(r'(\d{4,8})-([A-Z0-9])', nit)
        if match:
            return f"{match.group(1)}-{match.group(2)}"

        # Sin guión con letra final: 1234567K
        match2 = re.search(r'^(\d{4,8})([A-Z])$', nit)
        if match2:
            return f"{match2.group(1)}-{match2.group(2)}"

        # Solo dígitos: último dígito es verificador
        if re.match(r'^\d{6,9}$', nit):
            numero_part = nit[:-1]
            if len(numero_part) > 7:
                logger.warning(f"  ⚠️ NIT con número inusualmente largo ({len(numero_part)} dígitos): {nit} — posible dígito extra del OCR")
            return f"{numero_part}-{nit[-1]}"

        return None
    
    @staticmethod
    def limpiar_numero_factura(numero: str, serie: str = "") -> Optional[str]:
        """
        Limpia número de factura
        Acepta: 00000001, 1-1-00000001, etc.
        Retorna: 00000001
        """
        if not numero:
            return None
        
        numero = numero.strip().replace(" ", "")

        # Si es serie-número format
        if "-" in numero:
            partes = numero.split("-")
            numero = partes[-1]  # Toma la última parte

        # Reparar letras que el OCR confundió con dígitos (el número es SOLO
        # dígitos). Guard: solo si hay pocos errores (<=2) en un número de
        # longitud plausible, para no "inventar" un número a partir de basura.
        if numero and not numero.isdigit():
            no_digitos = sum(1 for c in numero if not c.isdigit())
            if no_digitos <= 2 and len(numero) >= 5:
                numero = ParserFactura._corregir_ocr_numero(numero)

        # FEL acepta números de hasta 12 dígitos
        return numero if numero.isdigit() and len(numero) >= 1 else None
    
    @staticmethod
    def extraer_campos(texto_ocr: str) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str], bool]:
        """
        Parsea texto OCR bruto y extrae: NIT, serie, número, razón social.
        Retorna: (nit, serie, numero, razon_social, receptor_es_cf).
        receptor_es_cf=True si el bloque del receptor indica CONSUMIDOR FINAL —
        en ese caso la factura no sirve para reporte de viáticos.
        """
        nit = None
        numero_factura = None
        serie = None
        razon_social = None

        texto_upper = texto_ocr.upper()
        lineas = texto_ocr.split("\n")

        # --- Separar bloque EMISOR ---
        # Cortar antes de secciones del comprador/receptor para evitar tomar NITs del comprador
        # NOTA: "FACTURA" y "DOCUMENTO TRIBUTARIO" se omiten aquí porque aparecen
        # en el encabezado del recibo (primeras líneas), no en la sección del comprador.
        # Incluirlos cortaba bloque_emisor antes del NIT y la razón social del emisor.
        SEPARADORES = [
            "COMPRADOR", "DATOS DEL CLIENTE", "DATOS DEL COMPRADOR",
            "RECEPTOR", "DATOS DEL RECEPTOR", "ADQUIRIENTE",
            "CLIENTE:",
        ]
        idx_corte = len(texto_ocr)
        for sep in SEPARADORES:
            idx = texto_upper.find(sep)
            if idx > 20:
                idx_corte = min(idx_corte, idx)
        bloque_emisor = texto_ocr[:idx_corte]
        bloque_receptor = texto_ocr[idx_corte:]

        # --- Detección de "Consumidor Final" en el bloque del receptor ---
        # Si la factura fue emitida a CF, no sirve para reporte de viáticos.
        # Buscamos en bloque_receptor PRIMERO; si está vacío (no había
        # separador COMPRADOR/RECEPTOR explícito) caemos al texto completo
        # buscando solo patrones MUY específicos para no falsear sobre el
        # nombre del emisor.
        _PATRON_CF_FUERTE = (
            r'\bCONSUMIDOR\s*FINAL\b'                                  # "CONSUMIDOR FINAL"
            r'|NIT\s*[:\-]?\s*(?:CF|C/F|C\.F\.|0)(?=\s|$|\n)'           # "NIT: CF" / "NIT: 0"
            r'|NIT\s+RECEPTOR\s*[:\-]?\s*(?:CF|C/F|C\.F\.|0)(?=\s|$|\n)'
        )
        _PATRON_CF_DEBIL = r'|\bC[\.\s\/]+F\b'  # solo en bloque_receptor (ambiguo)

        _receptor_upper = bloque_receptor.upper()
        if _receptor_upper.strip():
            _receptor_es_cf = bool(re.search(
                _PATRON_CF_FUERTE + _PATRON_CF_DEBIL, _receptor_upper
            ))
        else:
            # No hubo split por SEPARADORES — buscar solo patrones fuertes en
            # todo el texto. Como "CONSUMIDOR FINAL" y "NIT: CF/0" son frases
            # muy específicas, no aparecen en el nombre del emisor casi nunca.
            _receptor_es_cf = bool(re.search(_PATRON_CF_FUERTE, texto_upper))

        # --- NIT del EMISOR (en bloque antes del comprador) ---
        # Etiquetas reconocidas para NIT: NIT, N.I.T., RTU (Registro Tributario Unificado), Registro Fiscal
        _ETIQUETA_NIT = r'(?:N\.?I\.?T\.?|R\.?T\.?U\.?|REGISTRO\s+FISCAL)'
        # Patrón ESTRICTO: exige separador explícito (guión, punto o slash) antes del verificador
        # Esto evita capturar NITs del comprador que suelen aparecer sin separador
        patron_nit_estricto = _ETIQUETA_NIT + r'\s*[:.]?\s*(\d{4,8}[\.\-/][A-Z0-9])'
        # Patrón RELAJADO: acepta número puro (se valida matemáticamente más adelante)
        patron_nit_relajado = _ETIQUETA_NIT + r'\s*[:.]?\s*(\d{4,9})'

        def _limpiar_nit_ocr(raw: str) -> str:
            """Corrige errores OCR en el número del NIT: O→0, I→1, S→5, B→8."""
            # Solo aplica correcciones al segmento de dígitos (antes del separador)
            # El verificador (después del guión) puede ser K o un dígito, se deja como está
            partes = re.split(r'([\.\-/])', raw, maxsplit=1)
            num = partes[0].translate(str.maketrans('OoIilBSZ', '00111852'))
            return num + ''.join(partes[1:]) if len(partes) > 1 else num

        # Paso 1: buscar NIT con separador explícito en bloque emisor
        m = re.search(patron_nit_estricto, bloque_emisor, re.IGNORECASE)
        if m and not _es_nit_excluido(m.group(1)):
            nit = _limpiar_nit_ocr(m.group(1))

        if not nit:
            # Paso 2: NIT suelto con formato correcto (digits-letra) en bloque emisor
            m = re.search(r'\b(\d{4,8}-[A-Z0-9])\b', bloque_emisor)
            if m and not _es_nit_excluido(m.group(1)):
                nit = _limpiar_nit_ocr(m.group(1))

        if not nit:
            # Paso 3: patrón relajado (número puro) — SOLO si el dígito verificador es válido
            # Esto evita que NITs del comprador (que suelen no validar) sean aceptados
            for candidato_match in re.finditer(patron_nit_relajado, bloque_emisor, re.IGNORECASE):
                candidato_raw = _limpiar_nit_ocr(candidato_match.group(1))
                if _es_nit_excluido(candidato_raw):
                    continue
                # limpiar_nit convierte dígitos puros a formato XXXXXX-V
                candidato_limpio = ParserFactura.limpiar_nit(candidato_raw)
                if candidato_limpio and ParserFactura.validar_digito_nit(candidato_limpio):
                    nit = candidato_raw
                    logger.info(f"  NIT relajado validado: {candidato_raw} → {candidato_limpio}")
                    break

        if not nit:
            # Paso 4: NIT en sección DATOS DEL VENDEDOR/EMISOR (después del bloque emisor)
            for kw in ["VENDEDOR", "EMISOR"]:
                idx_kw = texto_upper.find(kw)
                if idx_kw >= 0:
                    idx_fin_kw = texto_upper.find("COMPRADOR", idx_kw)
                    fin_bloque = idx_fin_kw if idx_fin_kw > idx_kw else idx_kw + 500
                    bloque_vendedor = texto_ocr[idx_kw:fin_bloque]
                    m = re.search(patron_nit_estricto, bloque_vendedor, re.IGNORECASE)
                    if m and not _es_nit_excluido(m.group(1)):
                        nit = _limpiar_nit_ocr(m.group(1))
                        break
                    m = re.search(r'\b(\d{4,8}-[A-Z0-9])\b', bloque_vendedor)
                    if m and not _es_nit_excluido(m.group(1)):
                        nit = _limpiar_nit_ocr(m.group(1))
                        break

        # --- SERIE (FEL: UUID o código alfanumérico largo después de "SERIE:") ---
        # FEL series son UUIDs tipo: A78B2394-5678-1234-ABCD-001234567890
        # o códigos cortos tipo: A, AB, A1, 3F49EC03
        # Etiquetas reconocidas: SERIE, UUID, COD. SERIE, NUM. SERIE, CÓDIGO DE SERIE, NO. SERIE
        _ETIQUETA_SERIE = r'(?:SERIE|UUID|COD\.?\s*SERIE|NUM\.?\s*SERIE|C[OÓ]DIGO\s+DE\s+SERIE|NO\.?\s+SERIE)'
        # Primero intentar capturar serie larga (UUID o código ≥4 chars)
        m = re.search(_ETIQUETA_SERIE + r'\s*[:.]?\s*([A-Z0-9][A-Z0-9\-]{3,})', texto_upper)
        if m:
            serie_raw = re.sub(r"['\`]", '1', m.group(1)).strip().rstrip('-')
            # La SERIE de una FEL es SOLO el primer segmento (8 hex). Si el OCR
            # capturo el UUID de autorizacion completo (ej.
            # "1C607D9F-97A7-42B0-A64C-08082515...") o "SERIE-NUMERO" juntos,
            # nos quedamos con lo anterior al primer guion: la serie nunca lo
            # lleva.
            if '-' in serie_raw:
                serie_raw = serie_raw.split('-', 1)[0]
            serie = ParserFactura.reparar_serie_hex(serie_raw)  # hex: O->0, I->1, S->5, etc.
        else:
            # Fallback: cualquier token después de la etiqueta (incluye series cortas A, AB, A1)
            m = re.search(_ETIQUETA_SERIE + r'\s*[:.]?\s*(\S+)', texto_upper)
            if m:
                serie = ParserFactura.reparar_serie_hex(re.sub(r"['\`]", '1', m.group(1)).strip())

        # Patrón Shell/UNO y otras estaciones: "DTE | FACTURA |SERIE-NUMERO" sin etiquetas separadas
        # Ej: "DTE | FACTURA |63CAC965-1561870338"
        # Tolerante a variaciones OCR: pipes pueden leerse como "I", "l", "1" o perderse;
        # el guión entre serie y número a veces se lee como espacio o se pierde.
        if not serie or not numero_factura:
            # Separadores aceptados entre FACTURA y SERIE: pipes, espacios, guiones,
            # letras I/l/L (que el OCR suele confundir con pipes). Cuidado: no incluir
            # dígitos porque se comerían el primer carácter de una serie numérica.
            _SEP = r'[\|IlL\s\-:]*'
            patrones_dte = [
                # Con guión claro entre serie y número (caso óptimo)
                r'(?:DTE|FACTURA)' + _SEP + r'([A-Z0-9]{6,})[\-\s]+(\d{7,})',
                # Tolerar que serie y número queden pegados por el OCR: 8 hex + dígitos
                r'(?:DTE|FACTURA)' + _SEP + r'([A-F0-9]{8})(\d{7,})',
            ]
            for pat in patrones_dte:
                m_fel = re.search(pat, texto_upper)
                if m_fel:
                    cand_serie = ParserFactura.reparar_serie_hex(m_fel.group(1).rstrip('-'))
                    cand_numero = m_fel.group(2)
                    # Validar que la serie parezca hex/alfanum y el número sea razonable
                    if len(cand_serie) >= 5 and 7 <= len(cand_numero) <= 12:
                        if not serie:
                            serie = cand_serie
                        if not numero_factura:
                            numero_factura = cand_numero
                        break

        # --- NÚMERO DE FACTURA ---
        # Patrón FEL directo: etiqueta + número en misma línea o en la siguiente
        PATRONES_NUMERO_DIRECTO = [
            r'(?:N[uúU]mero|NRO\.?)\s+de\s+(?:Documento|Factura|DTE|Comprobante)\s*[:\-]?\s*(\d{3,})',
            r'No\.?\s+de\s+DTE\s*[:\-]?\s*(\d{3,})',           # "No de DTE: 3019983911"
            r'(?:DTE|FOLIO)\s+No\.?\s*[:\-]?\s*(\d{3,})',
            r'No\.?\s+(?:de\s+)?(?:DTE|Comprobante|Factura)\s*[:\-]?\s*(\d{3,})',
            r'Correlativo\s+(?:No\.?|N[uúU]mero)\s*[:\-]?\s*[\w\-]*?(\d{3,})',
            r'No\.\s*(?:de\s+)?(?:Documento|Factura)\s*[:\-]?\s*(\d{3,})',
            # Formato FEL corto: "Numero: 1620724585" (sin "de") — común en McDonald's,
            # San Rafael, Burger King y otros. Requiere ≥6 dígitos para no capturar
            # números cortos del cuerpo del ticket.
            r'N[uúU]mero\s*[:\-]\s*(\d{6,})',
            # Formato San Rafael y similares: "No. 800474621" (sin palabra intermedia).
            # \b al inicio evita capturar "Nombre"; ≥6 dígitos descarta números cortos
            # del cuerpo (cantidades, totales) y NITs (8 dígitos pero siempre con guión).
            r'\bNo\.\s+(\d{6,})\b',
            r'Autorizaci[oó]n\s*(?:No\.?|N[uúU]mero)?\s*[:\-]?\s*(\d{5,})',  # algunos usan "Autorización: NNNNN"
        ]
        # Primera pasada: mismo renglón (sin DOTALL)
        for pat in PATRONES_NUMERO_DIRECTO:
            m = re.search(pat, texto_ocr, re.IGNORECASE)
            if m:
                candidato = ParserFactura._corregir_ocr_numero(m.group(1).strip())
                if candidato.isdigit() and len(candidato) >= 3:
                    numero_factura = candidato
                    break

        # Segunda pasada: con DOTALL por si el número está en la línea siguiente a la etiqueta
        if not numero_factura:
            for pat in PATRONES_NUMERO_DIRECTO:
                m = re.search(pat, texto_ocr, re.IGNORECASE | re.DOTALL)
                if m:
                    candidato = ParserFactura._corregir_ocr_numero(m.group(1).strip())
                    if candidato.isdigit() and len(candidato) >= 3:
                        numero_factura = candidato
                        break

        # Si no se encontró con patrones directos, buscar con patrón genérico
        if not numero_factura:
            patron_numero = r'(?:N[ÚUúu]MERO|NUMERO|NUMBER|NRO\.?|No\.)\s*[:]?\s*([\w ,\.]{1,20})'
            for m in re.finditer(patron_numero, texto_ocr, re.IGNORECASE):
                candidato = re.sub(r'[\s,\.]+', '', m.group(1)).strip()
                candidato = ParserFactura._corregir_ocr_numero(candidato)
                if candidato.isdigit() and len(candidato) >= 3:
                    numero_factura = candidato
                    break
                # Buscar hasta 3 líneas adelante si el valor está separado
                pos_fin = m.end()
                resto = texto_ocr[pos_fin:].lstrip()
                for linea_adelante in resto.split('\n')[:3]:
                    linea_adelante = linea_adelante.strip()
                    if not linea_adelante:
                        continue
                    siguiente_limpio = ParserFactura._corregir_ocr_numero(re.sub(r'[\s,\.]+', '', linea_adelante))
                    if siguiente_limpio.isdigit() and len(siguiente_limpio) >= 3:
                        numero_factura = siguiente_limpio
                        break
                if numero_factura:
                    break

        # --- RAZÓN SOCIAL del EMISOR ---
        PALABRAS_EXCLUIR = [
            "SERIE", "NUMERO", "FACTURA", "KILOMETRO", "CARRETERA",
            "TELEFONO", "DIRECCION", "FECHA", "NOMBRE", "CONTROL",
            "DOCUMENTO", "TRIBUT", "AUTORIZACION", "N.I.T", "NIT",
            "REGIMEN", "MONEDA", "ZONA", "CALLE", "ALDEA", "KM.",
            "COLONIA", "HERE ARE", "WORDS", "CIUDAD", "MUNICIPIO",
            "DEPARTAMENTO", "DE MARZO", "DE ENERO", "DE FEBRERO",
            "DE ABRIL", "UNDEFINED", "DE MAYO", "P.M.", "A.M.",
            "NO DE DTE", "DTE:", "FOLIO", "CORRELATIVO", "COMPROBANTE",
            "N/A", "---", "REGISTRO FISCAL", "CONTRIBUYENTE",
        ]
        SUFIJOS_EMPRESA = [
            "S.A.", "S.R.L.", "LTDA", "S.A", "S.R.L", "CIA.", "INC.",
            "SOCIEDAD ANONIMA", "EIRL", "COOPERATIVA",
        ]

        lineas_emisor = bloque_emisor.split("\n")

        def es_linea_nombre(l):
            l = re.sub(r'^[\[\(]?\d+[.\):\-]\s*', '', l.strip())
            return (len(l) > 3
                    and not re.match(r'^\d', l)
                    and str(l).lower() not in ("undefined", "null", "none", "undefinied", "")
                    and not any(x in l.upper() for x in PALABRAS_EXCLUIR))

        def buscar_razon_social_en_lineas(lineas):
            """Busca razón social en lista de líneas; prefiere sufijos empresariales y múltiples palabras"""
            # Paso 1: línea con sufijo empresarial (más confiable)
            # Se verifica que la línea tenga contenido MÁS ALLÁ del sufijo solo,
            # para evitar devolver "SOCIEDAD ANONIMA" como razón social cuando
            # aparece en una línea propia sin el nombre de la empresa.
            for linea in lineas:
                linea = re.sub(r'^[\[\(]?\d+[.\):\-]\s*', '', linea.strip())
                if any(s in linea.upper() for s in SUFIJOS_EMPRESA) and es_linea_nombre(linea):
                    # Quitar todos los sufijos y ver si queda contenido real
                    linea_sin_sufijos = linea.upper()
                    for suf in SUFIJOS_EMPRESA:
                        linea_sin_sufijos = linea_sin_sufijos.replace(suf, '')
                    linea_sin_sufijos = linea_sin_sufijos.strip().strip(',').strip()
                    if linea_sin_sufijos:  # Hay texto además del sufijo → es un nombre completo
                        return linea
            # Paso 2: línea multi-palabra válida más larga
            multi = [re.sub(r'^\d+\.\s*', '', l.strip()) for l in lineas
                     if es_linea_nombre(re.sub(r'^\d+\.\s*', '', l.strip())) and
                     len(re.sub(r'^\d+\.\s*', '', l.strip()).split()) >= 2]
            if multi:
                return max(multi, key=len)
            # Paso 3: cualquier línea válida (fallback)
            for linea in lineas:
                linea = re.sub(r'^[\[\(]?\d+[.\):\-]\s*', '', linea.strip())
                if es_linea_nombre(linea):
                    return linea
            return None

        razon_social = buscar_razon_social_en_lineas(lineas_emisor)

        # Si no encontrado en bloque emisor, buscar en sección VENDEDOR
        if not razon_social:
            for kw in ["VENDEDOR", "EMISOR"]:
                idx_kw = texto_upper.find(kw)
                if idx_kw >= 0:
                    idx_fin_kw = texto_upper.find("COMPRADOR", idx_kw)
                    fin_bloque = idx_fin_kw if idx_fin_kw > idx_kw else idx_kw + 500
                    lineas_vend = texto_ocr[idx_kw:fin_bloque].split("\n")
                    razon_social = buscar_razon_social_en_lineas(lineas_vend)
                    if razon_social:
                        break

        # === VALIDACIONES CRUZADAS (determinísticas) ===============

        # Fix A: serie y número nunca pueden ser iguales. Si vienen iguales,
        # alguno está mal leído — descartar el menos confiable (numero, que
        # suele ser el que se recortó del prefijo de la serie).
        if serie and numero_factura and serie.strip().upper() == numero_factura.strip().upper():
            logger.warning(f"  ⚠️ Serie y número idénticos ({serie}) — descartando número")
            numero_factura = None

        # Fix B: el UUID de autorización SAT empieza con los 8 chars hex de
        # la serie. Es determinístico por diseño de SAT. Si tenemos el UUID,
        # la serie derivada es verdad absoluta — gana sobre cualquier otro.
        # texto_upper ya está en mayúsculas; IGNORECASE además blinda casing mixto.
        m_auth = re.search(
            r'AUTORIZACI[OÓ]N[:\s\n]*([A-F0-9]{8})-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12}',
            texto_upper,
            re.IGNORECASE,
        )
        if m_auth:
            serie_desde_auth = m_auth.group(1)
            if not serie:
                serie = serie_desde_auth
                logger.info(f"  ✅ Serie derivada del UUID autorización: {serie}")
            elif serie.upper() != serie_desde_auth:
                # Solo sobrescribir si la serie actual es sospechosa (≠ hex 8 chars)
                serie_actual_limpia = re.sub(r'[^A-F0-9]', '', serie.upper())
                if len(serie_actual_limpia) != 8 or serie_actual_limpia != serie_desde_auth:
                    logger.warning(
                        f"  ⚠️ Serie OCR ({serie}) ≠ autorización ({serie_desde_auth}) "
                        f"— usando la del autorización (determinística)"
                    )
                    serie = serie_desde_auth

        # Fix C: si el número es muy corto Y coincide con el prefijo de la serie,
        # probablemente es un recorte mal leído — descartar para que el flujo
        # intente Haiku estructurado o catálogo.
        if numero_factura and serie and len(numero_factura) < 7:
            serie_digits = re.sub(r'[^0-9]', '', serie)
            if numero_factura in serie_digits or serie_digits.startswith(numero_factura):
                logger.warning(
                    f"  ⚠️ Número '{numero_factura}' parece recorte de serie '{serie}' "
                    f"— descartando (longitud < 7)"
                )
                numero_factura = None

        return nit, serie, numero_factura, razon_social, _receptor_es_cf


class SistemaExtraccion:
    """Sistema completo: OCR + Catálogo + Validación"""
    
    def __init__(self):
        # Ollama desactivado — el OCR activo es Claude Haiku.
        # La clase ExtractorOllama se mantiene en el modulo por si se reactiva
        # en el futuro (p.ej. para uso offline). Para reactivar: self.ollama = ExtractorOllama()
        self.ollama = None
        self.catalogo = CatalogoEmpresas()
        self.parser = ParserFactura()
        self.cache = self._cargar_cache()
        # Lock para el cache: procesar_factura ahora corre en hilos paralelos
        # (asyncio.to_thread, hasta concurrent_updates simultaneos). Sin esto,
        # un hilo escribiendo self.cache mientras otro lo serializa a JSON
        # lanza 'dictionary changed size during iteration'.
        self._cache_lock = threading.Lock()
        # Haiku es el OCR principal (no fallback)
        try:
            self.haiku = ExtractorHaiku(model=HAIKU_MODEL)
            logger.info("✓ Claude Haiku disponible (OCR principal)")
        except Exception:
            self.haiku = None
            logger.error("✗ Claude Haiku no disponible — configura ANTHROPIC_API_KEY en .env")
        # Sonnet es escalación cuando SAT rechaza lo que leyó Haiku
        try:
            self.sonnet = ExtractorHaiku(model=SONNET_MODEL)
            logger.info("✓ Claude Sonnet disponible (escalación OCR)")
        except Exception:
            self.sonnet = None

        # Lector de QR (opcional). Si hay backend, los datos del QR corrigen
        # serie/numero/NIT antes de consultar al SAT.
        if _QR_DISPONIBLE:
            _backend = _qr_backend()
            if _backend != "ninguno":
                logger.info(f"✓ Lector de QR disponible (backend: {_backend})")
            else:
                logger.info("ℹ Lector de QR sin backend (instala pyzbar o opencv-python "
                            "para leer QR de las facturas)")
        else:
            logger.info("ℹ Modulo lector_qr no disponible — solo OCR")

    def _cargar_cache(self) -> dict:
        """Carga cache de extracciones previas"""
        if CACHE_FILE.exists():
            with open(CACHE_FILE, "r", encoding='utf-8') as f:
                return json.load(f)
        return {}

    def _guardar_cache(self):
        """Guarda cache de extracciones (thread-safe)."""
        with self._cache_lock:
            _guardar_json_atomico(CACHE_FILE, self.cache)
    
    @staticmethod
    def _corregir_nit_ocr(nit_limpio: str) -> Optional[str]:
        """
        Intenta corregir un NIT con dígito verificador incorrecto (error OCR).

        Estrategia A (prioritaria): el OCR unió el verificador real dentro del número.
          Ejemplo: NIT real 536565-1, OCR lee "53656511" → limpiar_nit → "5365651-1"
          Se detecta porque numero[-1] == dígito correcto para numero[:-1].
          Resultado: devuelve "536565-1" en vez de corromper a "5365651-2".

        Estrategia B (fallback): solo el dígito verificador fue leído mal por el OCR.
          Se recalcula el dígito correcto para el número tal como está.
        """
        try:
            partes = nit_limpio.split("-")
            numero = partes[0]
            if not numero.isdigit():
                return None

            def _calcular_digito(num_str: str) -> str:
                n = len(num_str)
                total = sum(int(d) * (n + 1 - i) for i, d in enumerate(num_str))
                resto = total % 11
                return "0" if resto == 0 else ("K" if resto == 1 else str(11 - resto))

            # Estrategia A: OCR metió el verificador dentro del número (dígito extra)
            # Si el último dígito del número ES el verificador correcto para numero[:-1],
            # el NIT real es numero[:-1]-numero[-1]
            if len(numero) > 5:
                numero_corto = numero[:-1]
                digito_esperado = _calcular_digito(numero_corto)
                if numero[-1] == digito_esperado:
                    nit_corregido = f"{numero_corto}-{digito_esperado}"
                    # Guard: si la corrección produce un NIT excluido (caso típico:
                    # OCR leyó mal pero al recalcular llega a Banco Azteca/Elektra),
                    # rechazar. Sin esto un OCR "5074333X" reconstruiría 5074333-3.
                    if _es_nit_excluido(nit_corregido):
                        logger.warning(
                            f"  ⛔ Corrección OCR (Estrategia A) produciría NIT "
                            f"excluido {nit_corregido} ← {nit_limpio}. Descartando."
                        )
                        return None
                    logger.info(f"  OCR unió verificador al número (dígito extra): {nit_limpio} → {nit_corregido}")
                    return nit_corregido

            # Estrategia B: solo corregir el dígito verificador (OCR leyó mal el último char)
            digito = _calcular_digito(numero)
            nit_corregido = f"{numero}-{digito}"
            if _es_nit_excluido(nit_corregido):
                logger.warning(
                    f"  ⛔ Corrección OCR (Estrategia B) produciría NIT "
                    f"excluido {nit_corregido} ← {nit_limpio}. Descartando."
                )
                return None
            logger.info(f"  Dígito verificador corregido: {nit_limpio} → {nit_corregido}")
            return nit_corregido

        except Exception:
            return None

    @staticmethod
    def _limpiar_texto_modelo(texto: str) -> str:
        """Elimina líneas de introducción/comentario del modelo antes del contenido real"""
        frases_modelo = [
            "here are", "the following", "i can see", "the receipt", "the text",
            "printed on", "words and numbers", "i found", "based on", "the image",
            "there is", "there are", "this is", "it appears", "the document",
            "transcription", "certainly", "sure", "below is", "above is"
        ]
        lineas = texto.split("\n")
        lineas_limpias = []
        for linea in lineas:
            linea_strip = linea.strip()
            linea_lower = linea_strip.lower()
            # Quitar encabezados markdown (# Titulo)
            if re.match(r'^#+\s', linea_strip):
                continue
            if any(frase in linea_lower for frase in frases_modelo):
                continue
            lineas_limpias.append(linea)
        return "\n".join(lineas_limpias).strip()

    @staticmethod
    def _catalogo_match_valido(razon_social_ocr: str, empresa_cat: dict) -> bool:
        """Valida que un match del catálogo sea creíble.

        Casos cubiertos:
        1. Empresas vs empresas: exige al menos 2 palabras distintivas en común
           cuando ambos nombres tienen 2+ palabras distintivas. Evita que
           OCR='POLLO' colapse con catálogo='POLLO CAMPERO S.A.' solo porque
           el char_bonus del fuzzy matcher es muy alto cuando hay 1 palabra
           larga idéntica.
        2. Personas físicas: exige al menos un primer nombre en común. Evita
           que 'MARIA ELENA CASTRO CASTRO DE MORENO' colapse con
           'ROBERTO RAFAEL MORENO CASTRO' solo por compartir apellidos.
        """
        if not empresa_cat or not razon_social_ocr:
            return bool(empresa_cat)

        def _es_empresa(n: str) -> bool:
            return any(s in n.upper() for s in
                       ["S.A.", "SOCIEDAD", "LTDA", "CIA.", "S.R.L.", "EIRL",
                        "COOPERATIVA", "INC.", "CORP", "RETAIL"])

        # Stopwords: sufijos legales y conectores que NO aportan identidad.
        _STOP_DISTINTIVAS = {
            "s.a.", "sa", "ltda", "cia", "cía", "srl", "s.r.l.", "eirl",
            "de", "del", "la", "el", "los", "las", "y", "e",
            "sociedad", "anonima", "anónima", "limitada", "inc", "corp",
            "retail", "comercial", "compania", "compañia", "compañía",
            "cooperativa",
        }

        def _distintivas(n: str) -> set:
            palabras = re.split(r'[\s,.\-]+', n.lower().strip())
            return {p for p in palabras if len(p) >= 4 and p not in _STOP_DISTINTIVAS}

        def _primeros_nombres(n: str) -> set:
            palabras = [p for p in re.split(r'[\s,]+', n.strip()) if len(p) > 2]
            palabras = [p for p in palabras if p.upper() not in
                        ("S.A.", "SA", "LTDA", "CIA", "SRL", "DE", "DEL", "LA", "EL")]
            return set(p.lower() for p in palabras[:2]) if palabras else set()

        rs_cat = empresa_cat.get('razon_social', '')

        # Validación de palabras distintivas — aplica a empresas y personas.
        # Si el catálogo tiene 2+ palabras distintivas, el OCR debe traer al
        # menos 2 distintivas en común para considerarse match válido.
        distintivas_ocr = _distintivas(razon_social_ocr)
        distintivas_cat = _distintivas(rs_cat)
        if len(distintivas_cat) >= 2:
            interseccion = distintivas_ocr & distintivas_cat
            if len(interseccion) < 2:
                logger.warning(
                    f"  ⚠️ Catálogo matcheó '{rs_cat}' con '{razon_social_ocr}' "
                    f"pero solo coinciden {len(interseccion)} palabra(s) "
                    f"distintiva(s) ({interseccion or '∅'}); se requieren 2+. "
                    f"Descartando match."
                )
                return False

        if _es_empresa(razon_social_ocr) or _es_empresa(rs_cat):
            return True  # Empresas: confiar en similitud textual + chequeo distintivas previo
        nombres_ocr = _primeros_nombres(razon_social_ocr)
        nombres_cat = _primeros_nombres(rs_cat)
        if nombres_ocr and nombres_cat and not (nombres_ocr & nombres_cat):
            logger.warning(
                f"  ⚠️ Catálogo matcheó '{rs_cat}' con '{razon_social_ocr}' "
                f"pero no comparten primer nombre — probablemente dueños "
                f"distintos. Descartando match."
            )
            return False
        return True

    def procesar_factura(self, image_path: str, user_id: Optional[str] = None,
                         registro_path: Optional[Path] = None) -> ExtraccionFactura:
        """
        Procesa una factura: OCR → Parse → Validar → Fallback

        Flujo:
        1. OCR con Llava-Phi
        2. Parse de texto
        3. Si NIT válido → ✓ Done
        4. Si falla: Buscar razón social en catálogo
        5. Si encuentra → Retorna con metodo='catalogo'
        6. Si no → Requiere confirmación manual

        registro_path: ruta al registro.json del usuario, para activar el
            HistorialEmisores (correccion de serie por typo, boost de confianza,
            deteccion temprana de duplicado).
        """
        image_path = Path(image_path)
        if not image_path.exists():
            raise FileNotFoundError(f"Imagen no encontrada: {image_path}")
        # Historial del usuario (lazy: solo lee si lo consultamos)
        historial = HistorialEmisores(registro_path)
        # Evaluar calidad de la imagen ANTES del OCR para poder alertar al usuario
        calidad_img: dict = {}
        try:
            with Image.open(image_path) as _img_raw:
                calidad_img = _evaluar_calidad_imagen(_img_raw)
            if calidad_img.get("muy_borrosa"):
                logger.warning(f"⚠ Foto borrosa (blur={calidad_img['blur_score']:.0f}); el OCR puede fallar.")
        except Exception as e:
            logger.warning(f"No se pudo evaluar calidad de imagen: {e}")
        
        logger.info(f"\n{'='*60}")
        logger.info(f"PROCESANDO: {image_path.name}")
        logger.info(f"{'='*60}")
        
        # Cargar ejemplos de aprendizaje para enriquecer el prompt
        ejemplos_prompt = self._construir_ejemplos_prompt()

        # Paso 1: OCR con Claude Haiku
        logger.info("📸 Paso 1: OCR con Claude Haiku...")
        texto_ocr = ""
        try:
            texto_ocr = self.haiku.extraer_texto(str(image_path), ejemplos=ejemplos_prompt)
            texto_ocr = self._limpiar_texto_modelo(texto_ocr)
            logger.info(f"📝 Texto Haiku:\n{texto_ocr[:400]}")
        except Exception as e:
            logger.error(f"Haiku falló: {e}")
            raise RuntimeError(f"OCR falló: {e}")

        # Paso 2: Parsear campos del texto con regex
        logger.info("🔍 Paso 2: Parseando campos...")
        nit, serie, numero, razon_social, receptor_es_cf = self.parser.extraer_campos(texto_ocr)
        if receptor_es_cf:
            logger.warning("  ⚠️ Factura emitida a CONSUMIDOR FINAL (detectado en OCR)")

        logger.info(f"  NIT crudo: {nit}")
        logger.info(f"  Serie: {serie}")
        logger.info(f"  Número: {numero}")
        logger.info(f"  Razón social: {razon_social}")

        # Paso 2b: Leer el QR de la FEL — SOLO para RELLENAR campos que el OCR
        # no logro leer. NUNCA pisa un valor que el OCR ya tiene.
        #
        # Por que solo rellenar: el lector cv2 (fallback) puede decodificar mal
        # un QR de foto y devolver un UUID con formato valido pero EQUIVOCADO.
        # Si eso pisara la serie del OCR, mandariamos datos erroneos al SAT.
        # Para PDFs ya usamos el texto exacto del PDF (en telegram_bot), asi que
        # el QR aqui es solo una ayuda conservadora para fotos con campos vacios.
        try:
            qr = _extraer_datos_qr(image_path)
        except Exception as e:
            logger.debug(f"Lectura de QR fallo: {e}")
            qr = {}
        if qr:
            if qr.get('serie') and not serie:
                logger.info(f"  ✓ Serie tomada del QR (OCR no la leyo): {qr['serie']}")
                serie = qr['serie']
            if qr.get('numero') and not numero:
                logger.info(f"  ✓ Numero tomado del QR (OCR no lo leyo): {qr['numero']}")
                numero = qr['numero']
            if qr.get('nit') and not nit:
                logger.info(f"  ✓ NIT tomado del QR (OCR no lo leyo): {qr['nit']}")
                nit = qr['nit']

        # Paso 3: Limpiar y validar NIT con dígito verificador
        if nit and _es_nit_excluido(nit):
            logger.warning(f"  ⛔ NIT excluido (regex): {nit} — ignorando")
            nit = None
        nit_limpio = self.parser.limpiar_nit(nit) if nit else None
        numero_limpio = self.parser.limpiar_numero_factura(numero, serie) if numero else None

        # Verificación cruzada con catálogo ANTES de corregir matemáticamente.
        # Umbral alto (0.85) para evitar falsos positivos en negocios con nombres
        # parecidos pero dueños distintos (ej. "Restaurante San Rafael" de dos
        # personas diferentes en la misma ciudad — mismo rótulo, distinto NIT).
        # Para personas individuales verificamos que al menos los primeros
        # nombres coincidan, no solo los apellidos.
        empresa_cat = None
        if razon_social:
            empresa_cat = self.catalogo.buscar_por_razon_social(
                razon_social, umbral_similitud=0.85
            )
        if empresa_cat and not self._catalogo_match_valido(razon_social, empresa_cat):
            empresa_cat = None
        if empresa_cat:
                nit_catalogo = empresa_cat['nit']
                if nit_limpio and nit_limpio != nit_catalogo:
                    logger.warning(
                        f"  ⚠️ NIT OCR ({nit_limpio}) no coincide con catálogo "
                        f"({nit_catalogo}) para '{empresa_cat['razon_social']}'. "
                        f"Usando NIT del catálogo."
                    )
                nit_limpio = nit_catalogo
                razon_social = empresa_cat['razon_social']
                logger.info(f"  ✅ NIT desde catálogo: {nit_limpio} ({razon_social})")
        elif nit_limpio:
            if self.parser.validar_digito_nit(nit_limpio):
                logger.info(f"  ✅ NIT válido: {nit_limpio}")
            else:
                logger.warning(f"  ⚠️ Dígito verificador incorrecto: {nit_limpio} — intentando corregir...")
                nit_limpio = self._corregir_nit_ocr(nit_limpio) or nit_limpio
                logger.info(f"  NIT corregido: {nit_limpio}")

        if nit_limpio and numero_limpio:
            logger.info(f"\n✅ ÉXITO - OCR directo")
            logger.info(f"   NIT: {nit_limpio}")
            logger.info(f"   Número: {numero_limpio}")

            # Aplicar historial del usuario: NIT conocido → boost de confianza
            # y razon social ya cacheada. La serie NO se toca (varia por documento).
            info_hist = self._aplicar_historial(historial, nit_limpio, serie or "", numero_limpio)

            resultado = ExtraccionFactura(
                nit=nit_limpio,
                serie=serie or "",
                numero_factura=numero_limpio,
                razon_social=razon_social or info_hist.get("razon_social_conocida") or None,
                metodo_extraccion="ocr_directo+historial" if info_hist.get("vistas", 0) > 0 else "ocr_directo",
                confianza=min(0.99, 0.95 + info_hist.get("boost", 0.0)),
                historial=info_hist or None,
                calidad_imagen=calidad_img or None,
                receptor_cf=receptor_es_cf,
            )
            resultado.texto_ocr = texto_ocr  # guardamos para aprendizaje posterior
            self._guardar_extraccion(image_path, resultado)
            return resultado
        
        # Paso 3b: Si falta NIT o numero, intentar extraccion estructurada con Haiku
        if (not nit_limpio or not numero_limpio) and self.haiku:
            logger.info("🔍 Paso 2b: Extraccion estructurada con Haiku (regex no encontro todos los campos)...")
            try:
                datos_struct = self.haiku.extraer_estructurado(str(image_path), ejemplos=ejemplos_prompt)
                if datos_struct:
                    if not nit_limpio and datos_struct.get("nit"):
                        nit_raw_struct = datos_struct["nit"]
                        if _es_nit_excluido(nit_raw_struct):
                            logger.warning(f"  ⛔ NIT excluido (Haiku): {nit_raw_struct}")
                        else:
                            nit_limpio = self.parser.limpiar_nit(nit_raw_struct)
                            if nit_limpio and not self.parser.validar_digito_nit(nit_limpio):
                                nit_limpio = self._corregir_nit_ocr(nit_limpio) or nit_limpio
                            logger.info(f"  NIT estructurado: {nit_limpio}")
                    # Validar ANTES de aceptar: serie y numero no pueden ser iguales
                    cand_serie = (ParserFactura.reparar_serie_hex(str(datos_struct.get("serie") or "").strip())
                                  if datos_struct.get("serie") else None)
                    cand_numero = (str(datos_struct.get("numero") or "").strip()
                                   if datos_struct.get("numero") else None)
                    if (cand_serie and cand_numero
                            and cand_serie.upper() == cand_numero.upper()):
                        logger.warning(
                            f"  ⚠️ Haiku devolvió serie=numero=({cand_serie}) "
                            f"— descartando ambos (valores sospechosos)"
                        )
                        cand_serie = None
                        cand_numero = None

                    if not numero_limpio and cand_numero:
                        numero_limpio = self.parser.limpiar_numero_factura(cand_numero)
                        logger.info(f"  Numero estructurado: {numero_limpio}")
                    if not serie and cand_serie:
                        serie = cand_serie
                        logger.info(f"  Serie estructurada: {serie}")
                    if not razon_social and datos_struct.get("razon_social"):
                        razon_social = datos_struct["razon_social"]
            except Exception as e:
                logger.warning(f"Extraccion estructurada fallo: {e}")

        # Con datos estructurados, si ahora tenemos NIT y numero, reportar exito
        if nit_limpio and numero_limpio:
            logger.info(f"\n✅ ÉXITO - OCR + extraccion estructurada")
            info_hist = self._aplicar_historial(historial, nit_limpio, serie or "", numero_limpio)
            resultado = ExtraccionFactura(
                nit=nit_limpio,
                serie=serie or "",
                numero_factura=numero_limpio,
                razon_social=razon_social or info_hist.get("razon_social_conocida") or None,
                metodo_extraccion="ocr_estructurado+historial" if info_hist.get("vistas", 0) > 0 else "ocr_estructurado",
                confianza=min(0.98, 0.90 + info_hist.get("boost", 0.0)),
                historial=info_hist or None,
                calidad_imagen=calidad_img or None,
                receptor_cf=receptor_es_cf,
            )
            self._guardar_extraccion(image_path, resultado)
            return resultado

        # Paso 3c: Escalacion preventiva a Sonnet cuando Haiku:
        #   (a) dejo algun campo en blanco — quemariamos consulta SAT con datos parciales
        #   (b) la serie tiene caracteres alfanumericos sospechosos (mezcla A/4, 0/O,
        #       1/I, 8/B, 5/S) Y el NIT no esta en el historial — alto riesgo de
        #       confusion OCR en serie (caso real: Haiku leyo "A3D7447F" en lugar
        #       de "A3D7A47F" y SAT lo rechazo)
        # Sonnet lee mejor caracteres ambiguos en fotos malas.
        falta_algo = not nit_limpio or not numero_limpio or not serie
        # Serie sospechosa: tiene digitos+letras de confusion comun OCR
        # Conjunto de chars que Haiku confunde frecuentemente entre si:
        # A↔4, O↔0, I↔1, B↔8, S↔5, Z↔2.
        _CHARS_CONFUSION = set('A4O01I8BS5Z2')
        serie_sospechosa = False
        if serie and len(serie) >= 4 and not falta_algo and nit_limpio:
            # Solo aplica si el NIT no es conocido (historial.boost == 0) —
            # NITs vistos antes ya quedaron validados en consultas previas
            # y no vale la pena el Sonnet preventivo.
            _hist_boost = historial.boost_confianza(nit_limpio)
            if _hist_boost == 0.0:
                chars_conf = sum(1 for c in serie.upper() if c in _CHARS_CONFUSION)
                # Umbral: 3+ chars de confusion o ≥1/3 del total. Caso real
                # 'A3D7447F' tiene 3 (A,4,4) → activa Sonnet preventivo.
                if chars_conf >= 3 or chars_conf >= len(serie) / 3:
                    serie_sospechosa = True

        if (falta_algo or serie_sospechosa) and self.sonnet:
            motivo = "faltan campos" if falta_algo else "serie sospechosa OCR"
            logger.info(
                f"🔁 Pre-SAT: {motivo} (nit={bool(nit_limpio)}, "
                f"serie={bool(serie)}, num={bool(numero_limpio)}) — escalando a Sonnet"
            )
            pista = ExtraccionFactura(
                nit=nit_limpio or "",
                serie=serie or "",
                numero_factura=numero_limpio or "",
                razon_social=razon_social or "",
                metodo_extraccion="sonnet_preview",
                confianza=0.0,
            )
            sonnet_res = self.reextraer_con_sonnet(str(image_path), resultado_previo=pista)
            if sonnet_res and sonnet_res.nit and sonnet_res.numero_factura:
                logger.info(
                    f"✅ Sonnet completo los campos faltantes: "
                    f"NIT={sonnet_res.nit} Serie={sonnet_res.serie} Num={sonnet_res.numero_factura}"
                )
                info_hist = self._aplicar_historial(
                    historial, sonnet_res.nit, sonnet_res.serie or "", sonnet_res.numero_factura
                )
                sonnet_res.metodo_extraccion = "sonnet_preview"
                sonnet_res.historial = info_hist or None
                sonnet_res.calidad_imagen = calidad_img or None
                sonnet_res.receptor_cf = receptor_es_cf
                self._guardar_extraccion(image_path, sonnet_res)
                return sonnet_res
            else:
                logger.info("Sonnet preventivo no completo los datos — cae a catalogo/manual")

        # Paso 4: Fallback - Buscar por razón social en catálogo
        logger.info(f"\n⚠️  Paso 3: NIT no valido. Intentando catálogo...")
        
        if razon_social:
            logger.info(f"   Buscando: '{razon_social}'")
            empresa = self.catalogo.buscar_por_razon_social(
                razon_social, umbral_similitud=0.85
            )
            if empresa and not self._catalogo_match_valido(razon_social, empresa):
                empresa = None

            if empresa:
                logger.info(f"\n✅ ENCONTRADO en catálogo")
                logger.info(f"   Razón social: {empresa['razon_social']}")
                logger.info(f"   NIT: {empresa['nit']}")
                
                # El catálogo solo aporta NIT + razón social. Serie y número
                # se respetan si el OCR los extrajo; nunca se heredan del catálogo
                # porque son distintos en cada factura.
                resultado = ExtraccionFactura(
                    nit=empresa['nit'],
                    serie=serie or "",
                    numero_factura=numero_limpio or "",
                    razon_social=empresa['razon_social'],
                    metodo_extraccion="catalogo",
                    confianza=0.85 if (serie and numero_limpio) else 0.6,
                    receptor_cf=receptor_es_cf,
                )

                self._guardar_extraccion(image_path, resultado)
                return resultado

        # Paso 5: Manual (requiere confirmación del usuario)
        logger.warning(f"\n❌ No se pudo extraer automáticamente")
        logger.warning(f"   Razón: NIT inválido y empresa no en catálogo")
        logger.info(f"   Requiere confirmación manual del usuario")

        resultado = ExtraccionFactura(
            nit="",
            serie=serie or "A",
            numero_factura=numero_limpio or "",
            razon_social=razon_social or "",
            metodo_extraccion="manual_pending",
            confianza=0.0,
            receptor_cf=receptor_es_cf,
        )

        return resultado
    
    @staticmethod
    def _aplicar_historial(historial: HistorialEmisores, nit: str, serie: str, numero: str) -> dict:
        """Consulta el historial del usuario y devuelve metadata accionable.
        Solo se basa en cosas estables por emisor: el NIT y la razon social.
        La serie NO se usa para correccion (cambia por documento en FEL)."""
        info: dict = {
            "vistas": 0,
            "boost": 0.0,
            "razon_social_conocida": "",
            "duplicado_exacto": False,
        }
        if not nit:
            return info
        vistas = historial.vistas_por_nit(nit)
        info["vistas"] = vistas
        info["boost"] = historial.boost_confianza(nit)
        info["razon_social_conocida"] = historial.razon_social_conocida(nit)
        info["duplicado_exacto"] = historial.es_duplicado_exacto(nit, serie, numero)
        if vistas > 0:
            logger.info(f"  📚 Historial: emisor visto {vistas}x, boost +{info['boost']:.2f}"
                        + (f", razon='{info['razon_social_conocida']}'" if info['razon_social_conocida'] else "")
                        + (" ⚠️ DUPLICADO EXACTO" if info['duplicado_exacto'] else ""))
        return info

    def _guardar_extraccion(self, image_path: Path, extraccion: ExtraccionFactura):
        """Guarda extracción en caché para estadísticas.

        Protegido por _cache_lock porque corre en hilos paralelos: el
        read-modify-write del dict + su serialización a JSON deben ser
        atómicos entre sí para no chocar con otro hilo.
        """
        clave = image_path.stem
        with self._cache_lock:
            self.cache[clave] = {
                **extraccion.to_dict(),
                "timestamp": image_path.stat().st_ctime
            }
            _guardar_json_atomico(CACHE_FILE, self.cache)

    # ── Sistema de aprendizaje ──────────────────────────────────────────────

    def _cargar_aprendizaje(self) -> list:
        if APRENDIZAJE_FILE.exists():
            try:
                with open(APRENDIZAJE_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return []

    def _guardar_aprendizaje(self, datos: list):
        _guardar_json_atomico(APRENDIZAJE_FILE, datos)

    def registrar_confirmacion_sat(self, nit: str, serie: str, numero: str, razon_social: str, texto_ocr: str = ""):
        """Llamado cuando SAT confirma una factura como valida.

        Aprende EMISORES frecuentes (NIT ↔ razon social), no documentos.
        Se cuenta por NIT — la clave estable — porque serie y numero son
        unicos por documento y nunca se repiten; contar por documento dejaba
        'confirmaciones' siempre en 1 y llenaba el archivo de entradas sueltas.
        """
        from datetime import datetime
        nit = (nit or "").strip()
        razon_social = (razon_social or "").strip()
        if not nit:
            return
        ejemplos = self._cargar_aprendizaje()
        ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Si el emisor (NIT) ya esta, sumamos una confirmacion y refrescamos.
        for e in ejemplos:
            if (e.get("nit") or "").strip() == nit:
                e["confirmaciones"] = e.get("confirmaciones", 1) + 1
                e["ultima"] = ahora
                if razon_social:
                    e["razon_social"] = razon_social  # ultima oficial del SAT
                self._guardar_aprendizaje(ejemplos)
                logger.info(f"Aprendizaje: emisor {nit} visto {e['confirmaciones']}x")
                return
        # Emisor nuevo.
        ejemplos.append({
            "timestamp": ahora,
            "ultima": ahora,
            "nit": nit,
            "razon_social": razon_social,
            "confirmaciones": 1,
        })
        # Cap por cantidad de emisores; conservar los mas frecuentes.
        if len(ejemplos) > 80:
            ejemplos = sorted(
                ejemplos, key=lambda e: e.get("confirmaciones", 1), reverse=True
            )[:80]
        self._guardar_aprendizaje(ejemplos)
        logger.info(f"Aprendizaje: emisor nuevo {nit} ({razon_social})")

    def _construir_ejemplos_prompt(self, max_ejemplos: int = 6) -> str:
        """Devuelve una tabla de referencia EMISOR → NIT de los emisores mas
        frecuentes confirmados por SAT.

        Importante: NO incluye serie ni numero. Esos cambian en CADA factura
        (son unicos por documento), asi que mostrarlos como "ejemplo" no ayuda
        a leer una factura nueva y hasta puede sesgar al modelo a copiarlos.
        Lo unico estable por emisor es NIT ↔ razon social — eso si sirve: si
        el modelo reconoce el nombre del emisor, sabe que NIT le corresponde
        (y con que formato/digito verificador).
        """
        ejemplos = self._cargar_aprendizaje()
        if not ejemplos:
            return ""
        # Deduplicar por NIT, quedandonos con el mas confirmado de cada emisor.
        por_nit: dict = {}
        for e in ejemplos:
            nit = (e.get("nit") or "").strip()
            rs = (e.get("razon_social") or "").strip()
            if not nit or not rs:
                continue
            prev = por_nit.get(nit)
            if prev is None or e.get("confirmaciones", 1) > prev.get("confirmaciones", 1):
                por_nit[nit] = e
        if not por_nit:
            return ""
        seleccion = sorted(
            por_nit.values(), key=lambda e: e.get("confirmaciones", 1), reverse=True
        )[:max_ejemplos]
        lineas = [f'- "{e["razon_social"]}" → NIT {e["nit"]}' for e in seleccion]
        return "\n".join(lineas)

    def reextraer_con_sonnet(self, image_path: str,
                              resultado_previo: Optional[ExtraccionFactura] = None) -> Optional[ExtraccionFactura]:
        """Re-extrae la factura usando Claude Sonnet en vez de Haiku.

        Se dispara como escalación cuando SAT rechaza los datos que extrajo
        Haiku (facturas degradadas donde Haiku lee mal caracteres ambiguos
        como F→5, 5→1). Sonnet lee mejor en papel arrugado/borroso.

        Usa solo el prompt estructurado (JSON) para evitar gastar tokens en
        transcripción completa. Pasa el resultado previo como pista — Sonnet
        puede corregir específicamente los campos que Haiku probablemente
        leyó mal.
        """
        if not self.sonnet:
            logger.warning("Sonnet no disponible — no se puede escalar")
            return None

        logger.info("🔁 Escalando a Claude Sonnet (Haiku falló en SAT)...")

        # Construir hint con lo que Haiku leyó, para que Sonnet sepa qué revisar
        hint = ""
        if resultado_previo:
            hint = (
                f"\n\nPREVIOUS ATTEMPT (likely has errors in 1-2 characters):\n"
                f"  NIT: {resultado_previo.nit or 'unknown'}\n"
                f"  serie: {resultado_previo.serie or 'unknown'}\n"
                f"  numero: {resultado_previo.numero_factura or 'unknown'}\n"
                f"  razon_social: {resultado_previo.razon_social or 'unknown'}\n"
                f"SAT rejected these values. Re-read the image CAREFULLY, paying\n"
                f"special attention to ambiguous characters: F vs 5, 5 vs 1/6, 0 vs O.\n"
                f"The errors are usually 1-2 digits, not the whole field.\n"
            )

        try:
            datos = self.sonnet.extraer_estructurado(image_path, ejemplos=hint)
        except Exception as e:
            logger.error(f"Sonnet falló: {e}")
            return None

        if not datos:
            return None

        # Aplicar misma validación que Haiku: serie != numero
        s_raw = ParserFactura.reparar_serie_hex(str(datos.get("serie") or "").strip()) or None
        n_raw = str(datos.get("numero") or "").strip() or None
        if s_raw and n_raw and s_raw.upper() == n_raw.upper():
            logger.warning(f"  ⚠️ Sonnet también devolvió serie=numero=({s_raw}) — descartando")
            s_raw = None
            n_raw = None

        nit_raw = datos.get("nit")
        # Filtro NITs excluidos (Banco Azteca, Elektra, etc.) — Sonnet a veces
        # confunde la sección del receptor con la del emisor, especialmente en
        # facturas degradadas donde la separación visual no es clara.
        if nit_raw and _es_nit_excluido(nit_raw):
            logger.warning(
                f"  ⛔ NIT excluido devuelto por Sonnet: {nit_raw} — "
                f"probablemente leyó al COMPRADOR. Descartando resultado."
            )
            return None
        nit_limpio = self.parser.limpiar_nit(nit_raw) if nit_raw else None
        if nit_limpio and not self.parser.validar_digito_nit(nit_limpio):
            nit_limpio = self._corregir_nit_ocr(nit_limpio) or nit_limpio
        # Re-chequear tras el limpiar/corregir por si el formato cambió y ahora
        # coincide con un NIT excluido conocido.
        if nit_limpio and _es_nit_excluido(nit_limpio):
            logger.warning(
                f"  ⛔ NIT excluido tras normalizar: {nit_limpio}. Descartando."
            )
            return None
        numero_limpio = self.parser.limpiar_numero_factura(n_raw) if n_raw else None

        if not (nit_limpio and numero_limpio):
            logger.warning(f"Sonnet no produjo datos completos: nit={nit_limpio}, num={numero_limpio}")
            return None

        resultado = ExtraccionFactura(
            nit=nit_limpio,
            serie=s_raw or (resultado_previo.serie if resultado_previo else "") or "",
            numero_factura=numero_limpio,
            razon_social=datos.get("razon_social") or (resultado_previo.razon_social if resultado_previo else None),
            metodo_extraccion="sonnet_escalacion",
            confianza=0.92,
            # Sonnet re-lee la imagen pero NO el bloque receptor; heredamos el
            # flag CF del resultado previo. Si la imagen era CF, sigue siendo CF.
            receptor_cf=bool(resultado_previo and resultado_previo.receptor_cf),
        )
        logger.info(f"✅ Sonnet: NIT={resultado.nit} Serie={resultado.serie} Num={resultado.numero_factura}")
        return resultado


# Ejemplo de uso
if __name__ == "__main__":
    # Inicializar catálogo con ejemplos
    catalogo = CatalogoEmpresas()
    
    # Agregar empresas de ejemplo (solo si está vacío)
    if len(catalogo.empresas) == 0:
        logger.info("Inicializando catálogo con ejemplos...")
        catalogo.agregar_empresa("Pollo Campero S.A.", "4342321-K")
        catalogo.agregar_empresa("Ánfora Retail", "5432123-L")
        catalogo.agregar_empresa("Farmacias Similares", "3456789-M")
        catalogo.agregar_empresa("Supermercado La Paiz", "2234567-N")
    
    # Crear sistema
    sistema = SistemaExtraccion()
    
    # Ejemplo: procesar una imagen
    # imagen_test = Path("factura_prueba.jpg")
    # if imagen_test.exists():
    #     resultado = sistema.procesar_factura(str(imagen_test))
    #     print(f"\n{'='*60}")
    #     print("RESULTADO FINAL:")
    #     print(f"{'='*60}")
    #     print(json.dumps(resultado.to_dict(), ensure_ascii=False, indent=2))
