"""
Utilidades base compartidas del bot (sin dependencias de Telegram ni Playwright).

Aquí viven los cimientos que usan tanto telegram_bot.py como reportes.py:
- Rutas de datos.
- Locks por archivo y guardado JSON atómico (escrituras seguras).
- Carpeta/perfil por usuario.
- Identidad y deduplicación de facturas (clave NIT|SERIE|NUMERO).
- Helpers de fecha (orden y formato de rango para nombres de reporte).

Separado para romper el import circular: reportes.py necesita estos helpers
y no puede importarlos de telegram_bot (que a su vez importa reportes).
"""

import json
import logging
import threading
from pathlib import Path
from typing import Optional
from datetime import datetime

logger = logging.getLogger(__name__)

# Carpeta raíz de datos (relativa al directorio de ejecución del bot).
DATOS_DIR = Path("datos")
DATOS_DIR.mkdir(exist_ok=True)

# Nombres de meses en español (índice 1 = Enero) para nombres de reporte.
_MESES_ES = [
    '', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre',
]


# ─────────────────────────────────────────────
#  LOCKS POR ARCHIVO (thread-safe read-modify-write)
# ─────────────────────────────────────────────
_FILE_LOCKS: dict = {}
_FILE_LOCKS_GUARD = threading.Lock()


def _lock_archivo(ruta) -> threading.RLock:
    """Devuelve un RLock unico por ruta de archivo. Reentrant para anidamiento seguro."""
    key = str(ruta)
    with _FILE_LOCKS_GUARD:
        lock = _FILE_LOCKS.get(key)
        if lock is None:
            lock = threading.RLock()
            _FILE_LOCKS[key] = lock
        return lock


def _safe_int(valor, default=None):
    """Convierte a int de forma segura. Retorna default si falla."""
    try:
        return int(str(valor).strip())
    except (ValueError, TypeError, AttributeError):
        return default


def _guardar_json_atomico(ruta: Path, data) -> bool:
    """Guarda datos JSON de forma atómica usando un archivo temporal.
    El archivo original nunca queda corrupto: o existe el viejo completo
    o existe el nuevo completo, nunca un archivo a medias.
    Retorna True si tuvo éxito, False si falló."""
    tmp = ruta.with_suffix('.tmp')
    try:
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        tmp.replace(ruta)  # Operación atómica en todos los SO
        return True
    except Exception as e:
        logger.error(f"Error guardando {ruta.name} de forma atomica: {e}")
        try:
            tmp.unlink(missing_ok=True)  # Limpiar .tmp si algo falló
        except Exception:
            pass
        return False


# ─────────────────────────────────────────────
#  CARPETA Y PERFIL POR USUARIO
# ─────────────────────────────────────────────

def carpeta_usuario(user_id: int, username: str = "") -> Path:
    """Crea y retorna la carpeta del usuario: datos/usuario_ID_nombre/"""
    nombre = f"usuario_{user_id}"
    if username:
        nombre += f"_{username}"
    carpeta = DATOS_DIR / nombre
    (carpeta / "facturas").mkdir(parents=True, exist_ok=True)
    (carpeta / "pdf_sat").mkdir(parents=True, exist_ok=True)
    return carpeta


def subcarpeta_reportes(carpeta: Path) -> Path:
    """Devuelve (creando si hace falta) datos/usuario_X/reportes/ — donde se
    guardan TODOS los archivos generados del reporte (xlsx, txt, zip, pdf
    combinado) para no ensuciar la raiz de la carpeta del usuario."""
    destino = carpeta / "reportes"
    destino.mkdir(parents=True, exist_ok=True)
    return destino


def guardar_perfil_usuario(user_id: int, username: str = "", **campos):
    """Guarda campos en el perfil persistente del usuario (perfil.json) — atomico + lock."""
    carpeta = carpeta_usuario(user_id, username)
    perfil_path = carpeta / "perfil.json"
    with _lock_archivo(perfil_path):
        perfil = {}
        if perfil_path.exists():
            try:
                with open(perfil_path, 'r', encoding='utf-8') as f:
                    perfil = json.load(f)
            except (json.JSONDecodeError, OSError) as e:
                logger.warning(f"perfil.json ilegible para user {user_id} ({e}); recreando.")
                perfil = {}
        perfil.update(campos)
        _guardar_json_atomico(perfil_path, perfil)


def leer_perfil_usuario(user_id: int, username: str = "") -> dict:
    """Lee el perfil persistente del usuario."""
    carpeta = carpeta_usuario(user_id, username)
    perfil_path = carpeta / "perfil.json"
    if not perfil_path.exists():
        return {}
    with _lock_archivo(perfil_path):
        try:
            with open(perfil_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            logger.error(f"perfil.json ilegible para user {user_id}: {e}")
            return {}


# ─────────────────────────────────────────────
#  IDENTIDAD Y DEDUPLICACIÓN DE FACTURAS
# ─────────────────────────────────────────────
# Toda la lógica de "¿es la misma factura?" pasa por estos helpers para que
# las rutas que la usan no diverjan.

def _clave_factura(nit, serie, numero) -> str:
    """Clave canónica de identidad de una factura: NIT|SERIE|NUMERO.
    Normaliza a mayúsculas sin espacios para que diferencias triviales de
    formato no hagan pasar la misma factura como distinta."""
    def _n(x):
        return str(x if x is not None else "").strip().upper()
    return f"{_n(nit)}|{_n(serie)}|{_n(numero)}"


def _clave_de_resultado(resultado) -> str:
    """Clave de un ExtraccionFactura (campos .nit/.serie/.numero_factura)."""
    return _clave_factura(
        getattr(resultado, 'nit', ''),
        getattr(resultado, 'serie', ''),
        getattr(resultado, 'numero_factura', ''),
    )


def _clave_de_registro(reg: dict) -> str:
    """Clave de una entrada del registro.json (campos nit/serie/numero)."""
    return _clave_factura(reg.get('nit'), reg.get('serie'), reg.get('numero'))


def _cargar_registros_usuario(user_id: int, username: str = "") -> list:
    """Lee registro.json del usuario de forma segura. [] si no existe o falla."""
    registro_path = carpeta_usuario(user_id, username) / "registro.json"
    if not registro_path.exists():
        return []
    try:
        with open(registro_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []


# ─────────────────────────────────────────────
#  HELPERS DE FECHA
# ─────────────────────────────────────────────

def _formatear_rango_fechas(fecha_inicio: str, fecha_fin: str) -> str:
    """Formatea un rango de fechas en español para el nombre del reporte.
    Ejemplos:
      ('14/04/2026','17/04/2026') -> '14-17 Abril'
      ('30/04/2026','03/05/2026') -> '30 Abr - 03 May'
      ('14/04/2026','14/04/2026') -> '14 Abril'
      ('', '')                    -> 'Todas'
    """
    def _parse(f):
        if not f:
            return None
        f = f.replace('.', '/').replace('-', '/')
        try:
            return datetime.strptime(f, '%d/%m/%Y')
        except ValueError:
            return None
    fi = _parse(fecha_inicio)
    ff = _parse(fecha_fin)
    if not fi and not ff:
        return "Todas"
    if fi and ff:
        if fi == ff:
            return f"{fi.day} {_MESES_ES[fi.month]}"
        if fi.year == ff.year and fi.month == ff.month:
            return f"{fi.day}-{ff.day} {_MESES_ES[fi.month]}"
        # meses/años distintos → abreviar mes
        return f"{fi.day} {_MESES_ES[fi.month][:3]} - {ff.day} {_MESES_ES[ff.month][:3]}"
    d = fi or ff
    return f"{d.day} {_MESES_ES[d.month]}"


def _fecha_para_orden(f: str):
    """Convierte fecha dd.mm.yyyy / dd/mm/yyyy / dd-mm-yyyy / yyyy-mm-dd a datetime
    para poder ordenar facturas. Devuelve None si es inválida."""
    if not f:
        return None
    f = f.strip().replace('.', '/').replace('-', '/')
    for fmt in ("%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(f, fmt)
        except ValueError:
            continue
    return None
