"""
Bot de Telegram para procesamiento de facturas SAT
Flujo: Usuario envía foto → Bot procesa → Responde con datos extraídos o pide confirmación
"""

import os
import re
import uuid
import asyncio
import threading
from pathlib import Path
from typing import Optional
import json
import logging
from datetime import datetime, timedelta
from dotenv import load_dotenv
import shutil
import zipfile
import openpyxl

# Utilidades base compartidas (rutas, locks, JSON atomico, carpeta/perfil de
# usuario, identidad de facturas, helpers de fecha) viven en comun.py — sin
# dependencias de Telegram, asi reportes.py tambien las usa sin import circular.
from comun import (
    DATOS_DIR,
    _MESES_ES,
    _lock_archivo,
    _safe_int,
    _guardar_json_atomico,
    carpeta_usuario,
    leer_perfil_usuario,
    guardar_perfil_usuario,
    _clave_factura,
    _clave_de_resultado,
    _clave_de_registro,
    _cargar_registros_usuario,
    _formatear_rango_fechas,
    _fecha_para_orden,
)
from openpyxl.styles import Font as XLFont
from openpyxl.utils.datetime import to_excel as _date_to_excel

load_dotenv(Path(__file__).parent / ".env", override=True)

try:
    from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
    from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
    import telegram.error as tg_error
    TELEGRAM_AVAILABLE = True
except ImportError:
    TELEGRAM_AVAILABLE = False
    tg_error = None
    logging.warning("python-telegram-bot no instalado. Install: pip install python-telegram-bot")

from sat_invoice_bot import SistemaExtraccion, ExtraccionFactura
from sat_verificador import obtener_verificador, cerrar_verificador, _get_lock_usuario

import sys as _sys
from logging.handlers import RotatingFileHandler as _RotatingFileHandler

# Formato con timestamp: indispensable para diagnosticar "que paso a las 3am"
# (antes los logs no tenian hora). Ej: 2026-06-08 14:03:21 INFO sat_invoice_bot: ...
_LOG_FMT = logging.Formatter(
    "%(asctime)s %(levelname)s %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

# Handler a consola (stdout, utf-8 tolerante a emojis en Windows).
_consola = logging.StreamHandler(
    open(_sys.stdout.fileno(), mode='w', encoding='utf-8',
         errors='replace', closefd=False)
)
_consola.setFormatter(_LOG_FMT)

_handlers = [_consola]

# Handler a archivo con rotacion: logs/bot.log, 5 MB x 5 backups (~25 MB max).
# Sobrevive reinicios y deja rastro para revisar fallos pasados. Si por
# permisos no se puede crear, el bot sigue solo con consola.
try:
    _LOG_DIR = Path(__file__).parent / "logs"
    _LOG_DIR.mkdir(exist_ok=True)
    _archivo = _RotatingFileHandler(
        _LOG_DIR / "bot.log", maxBytes=5 * 1024 * 1024,
        backupCount=5, encoding='utf-8',
    )
    _archivo.setFormatter(_LOG_FMT)
    _handlers.append(_archivo)
except Exception as _e:
    logging.warning(f"No se pudo crear el archivo de log (se usa solo consola): {_e}")

# force=True reemplaza cualquier config previa del root logger. Es necesario
# porque sat_invoice_bot (importado mas arriba) ya llamo a basicConfig; sin
# force, este basicConfig seria no-op y el archivo de log no se adjuntaria.
logging.basicConfig(level=logging.INFO, handlers=_handlers, force=True)

# Silenciar logs ruidosos de librerias de red. httpx logueaba una linea INFO
# por CADA getUpdates de Telegram (varias por segundo), ahogando los logs
# propios del bot y haciendo imposible diagnosticar. Solo nos interesan sus
# WARNING/ERROR. Idem para httpcore, telegram.ext y apscheduler.
for _ruidoso in ("httpx", "httpcore", "telegram.ext.Updater", "apscheduler"):
    logging.getLogger(_ruidoso).setLevel(logging.WARNING)

logger = logging.getLogger(__name__)

# Config
TELEGRAM_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "YOUR_TOKEN_HERE")
ADMIN_CHAT_ID = os.getenv("ADMIN_CHAT_ID", "")  # Chat ID del admin para alertas
LIMITE_FACTURAS_MES = 100  # Alerta si un usuario supera este límite
# DATOS_DIR se importa de comun.py (arriba).

# Estado del bot
sistema = SistemaExtraccion()

# ─────────────────────────────────────────────
#  GESTIÓN DE USUARIOS (users_db.json)
# ─────────────────────────────────────────────
USERS_DB_PATH = DATOS_DIR / "users_db.json"

def _cargar_users_db() -> dict:
    if not USERS_DB_PATH.exists():
        return {}
    with _lock_archivo(USERS_DB_PATH):
        try:
            return json.loads(USERS_DB_PATH.read_text(encoding='utf-8'))
        except (json.JSONDecodeError, OSError) as e:
            logger.error(f"users_db.json ilegible ({e}); devolviendo dict vacio.")
            return {}

def _guardar_users_db(db: dict):
    """Guardado atomico + bajo lock para evitar corrupcion en escrituras concurrentes."""
    with _lock_archivo(USERS_DB_PATH):
        _guardar_json_atomico(USERS_DB_PATH, db)

def _registrar_usuario(user_id: int, username: str = "", first_name: str = "") -> dict:
    """Registra un usuario nuevo si no existe. Retorna su entrada."""
    with _lock_archivo(USERS_DB_PATH):
        db = _cargar_users_db()
        key = str(user_id)
        if key not in db:
            db[key] = {
                "user_id": user_id,
                "username": username,
                "first_name": first_name,
                "aprobado": True,
                "bloqueado": False,
                "limite_mensual": LIMITE_FACTURAS_MES,
                "fecha_registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "ultima_actividad": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "notas": "",
            }
            _guardar_users_db(db)
        else:
            # Actualizar última actividad y username por si cambió
            db[key]["ultima_actividad"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if username:
                db[key]["username"] = username
            if first_name:
                db[key]["first_name"] = first_name
            _guardar_users_db(db)
        return db[key]

def _es_admin(user_id: int) -> bool:
    return ADMIN_CHAT_ID and str(user_id) == str(ADMIN_CHAT_ID)

def _usuario_permitido(user_id: int) -> tuple[bool, str]:
    """Retorna (permitido, motivo). El admin siempre tiene acceso."""
    if _es_admin(user_id):
        return True, "admin"
    db = _cargar_users_db()
    entrada = db.get(str(user_id))
    if not entrada:
        return False, "no_registrado"
    if entrada.get("bloqueado"):
        return False, "bloqueado"
    if not entrada.get("aprobado"):
        return False, "pendiente"
    return True, "ok"

def _aprobar_usuario(user_id: int) -> bool:
    with _lock_archivo(USERS_DB_PATH):
        db = _cargar_users_db()
        key = str(user_id)
        if key not in db:
            return False
        db[key]["aprobado"] = True
        db[key]["bloqueado"] = False
        _guardar_users_db(db)
        return True

def _bloquear_usuario(user_id: int, motivo: str = "") -> bool:
    with _lock_archivo(USERS_DB_PATH):
        db = _cargar_users_db()
        key = str(user_id)
        if key not in db:
            return False
        db[key]["bloqueado"] = True
        if motivo:
            db[key]["notas"] = motivo
        _guardar_users_db(db)
        return True

def _desbloquear_usuario(user_id: int) -> bool:
    with _lock_archivo(USERS_DB_PATH):
        db = _cargar_users_db()
        key = str(user_id)
        if key not in db:
            return False
        db[key]["bloqueado"] = False
        db[key]["aprobado"] = True
        _guardar_users_db(db)
        return True

def _set_limite_usuario(user_id: int, limite: int) -> bool:
    with _lock_archivo(USERS_DB_PATH):
        db = _cargar_users_db()
        key = str(user_id)
        if key not in db:
            return False
        db[key]["limite_mensual"] = limite
        _guardar_users_db(db)
        return True

def _set_nota_usuario(user_id: int, nota: str) -> bool:
    with _lock_archivo(USERS_DB_PATH):
        db = _cargar_users_db()
        key = str(user_id)
        if key not in db:
            return False
        db[key]["notas"] = nota
        _guardar_users_db(db)
        return True


async def notificar_admin(bot, mensaje: str):
    """Envía un mensaje de alerta al admin si ADMIN_CHAT_ID está configurado."""
    if not ADMIN_CHAT_ID:
        logger.warning("⚠️  ADMIN_CHAT_ID no configurado en .env — ninguna alerta sera enviada.")
        logger.warning("    Agrega tu chat_id de Telegram en el archivo .env: ADMIN_CHAT_ID=123456789")
        return
    admin_id = _safe_int(ADMIN_CHAT_ID)
    if admin_id is None:
        logger.error(f"ADMIN_CHAT_ID invalido en .env (no es entero): {ADMIN_CHAT_ID!r}")
        return
    try:
        await bot.send_message(chat_id=admin_id, text=f"🚨 ALERTA BOT:\n\n{mensaje}")
    except Exception as e:
        logger.error(f"Error notificando admin (chat_id={admin_id}): {e}")


def contar_facturas_mes(user_id: int, username: str = "") -> int:
    """Cuenta las facturas procesadas en el mes actual por un usuario."""
    carpeta = carpeta_usuario(user_id, username)
    registro_path = carpeta / "registro.json"
    if not registro_path.exists():
        return 0
    try:
        with open(registro_path, 'r', encoding='utf-8') as f:
            registros = json.load(f)
    except Exception:
        return 0
    mes_actual = datetime.now().strftime("%Y-%m")
    return sum(1 for r in registros if r.get('fecha_proceso', '').startswith(mes_actual))


# guardar_perfil_usuario, leer_perfil_usuario, carpeta_usuario y los helpers
# de identidad de facturas (_clave_*, _cargar_registros_usuario) se importan
# de comun.py (arriba).

# ───── Identidad y deduplicación de facturas (fuente única) ─────
# La lógica de "¿es la misma factura?" vive en comun.py (_clave_factura,
# _clave_de_resultado, _clave_de_registro). Aquí solo se usa.


def _buscar_duplicado_en_registros(registros: list, carpeta: Path, clave_nueva: str,
                                   verificar_pdf: bool = True) -> Optional[dict]:
    """Busca en una lista YA cargada de registros una factura duplicada:
    misma clave canónica + monto > 0.

    verificar_pdf=True: además exige que el PDF declarado exista en disco.
      Si el registro no declara PDF, igual cuenta como duplicado (es la
      semántica histórica de buscar_duplicado/guardar_registro).
    verificar_pdf=False: solo clave + monto (semántica de _validar_pre_sat).

    Retorna el registro existente o None.
    """
    for reg in registros:
        if _clave_de_registro(reg) != clave_nueva:
            continue
        if _monto_efectivo(reg.get('monto', '')) <= 0:
            continue
        if not verificar_pdf:
            return reg
        pdf_guardado = reg.get('pdf_sat', '')
        if pdf_guardado:
            pdf_nombre = Path(pdf_guardado).name
            if (carpeta / "pdf_sat" / pdf_nombre).exists():
                return reg
            # PDF declarado pero ausente → no es duplicado confirmado; seguir
        else:
            return reg
    return None


def buscar_duplicado(user_id: int, resultado, username: str = "") -> tuple:
    """Solo verifica si una factura ya existe en el registro del usuario (sin guardar nada).
    Retorna (es_duplicado: bool, registro_existente: dict o None)."""
    carpeta = carpeta_usuario(user_id, username)
    registros = _cargar_registros_usuario(user_id, username)
    reg = _buscar_duplicado_en_registros(
        registros, carpeta, _clave_de_resultado(resultado), verificar_pdf=True
    )
    return (True, reg) if reg else (False, None)


def guardar_registro(user_id: int, resultado, pdf_path=None, username: str = "", res_sat=None):
    """Guarda los datos de una factura procesada en el registro del usuario.
    Retorna (duplicada: bool, entrada: dict). Si es duplicada, no guarda y retorna la existente.
    Read-modify-write protegido por lock por archivo."""
    carpeta = carpeta_usuario(user_id, username)
    registro_path = carpeta / "registro.json"

    with _lock_archivo(registro_path):
        # Leer registro existente
        registros = []
        if registro_path.exists():
            try:
                with open(registro_path, 'r', encoding='utf-8') as f:
                    registros = json.load(f)
            except (json.JSONDecodeError, OSError) as e:
                logger.error(f"registro.json ilegible para user {user_id} ({e}); inicializando vacio.")
                registros = []

        # Verificar duplicados (segunda barrera de seguridad)
        dup = _buscar_duplicado_en_registros(
            registros, carpeta, _clave_de_resultado(resultado), verificar_pdf=True
        )
        if dup:
            return True, dup

        entrada = {
            "fecha_proceso": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "nit": resultado.nit,
            "serie": resultado.serie,
            "numero": resultado.numero_factura,
            "razon_social": resultado.razon_social or "",
            "metodo": resultado.metodo_extraccion,
            "confianza": resultado.confianza,
            "pdf_sat": str(pdf_path) if pdf_path else "",
        }
        # Agregar datos del SAT si están disponibles
        if res_sat:
            entrada["monto"] = getattr(res_sat, 'monto', '') or ""
            entrada["fecha_emision"] = getattr(res_sat, 'fecha_emision', '') or ""
            entrada["hora_emision"] = getattr(res_sat, 'hora_emision', '') or ""
            entrada["tipo_comida"] = getattr(res_sat, 'tipo_comida', '') or ""
            entrada["establecimiento"] = getattr(res_sat, 'establecimiento', '') or ""
            entrada["departamento"] = normalizar_departamento(getattr(res_sat, 'departamento', '') or "")
            entrada["valida_sat"] = getattr(res_sat, 'valida', None)

        registros.append(entrada)

        _guardar_json_atomico(registro_path, registros)

        return False, entrada


# Tipos de comida que se verifican para duplicados diarios
_COMIDAS_DIARIAS = {"desayuno", "almuerzo", "cena"}

def _verificar_comida_duplicada(user_id: int, username: str, entrada_nueva: dict) -> Optional[str]:
    """
    Verifica si ya existe otra factura del mismo tipo de comida en la misma fecha.
    Solo aplica a Desayuno / Almuerzo / Cena.
    Retorna un mensaje de alerta si hay duplicado, o None si todo está bien.
    """
    tipo = (entrada_nueva.get("tipo_comida") or "").strip().lower()
    if tipo not in _COMIDAS_DIARIAS:
        return None  # Gasolina, Hospedaje, etc. no tienen límite diario

    fecha_nueva = (entrada_nueva.get("fecha_emision") or "").strip()
    if not fecha_nueva:
        return None  # Sin fecha no podemos comparar

    # Normalizar fecha a YYYY-MM-DD para comparación segura
    def _fecha_ymd(f: str) -> str:
        """Convierte dd.mm.yyyy, dd-mm-yyyy o yyyy-mm-dd a yyyy-mm-dd."""
        f = f.strip()
        m = re.match(r'(\d{2})[.\-/](\d{2})[.\-/](\d{4})', f)
        if m:
            return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
        return f[:10]  # Asumir que ya es yyyy-mm-dd

    fecha_nueva_ymd = _fecha_ymd(fecha_nueva)

    # Leer todos los registros guardados
    carpeta = carpeta_usuario(user_id, username)
    registro_path = carpeta / "registro.json"
    if not registro_path.exists():
        return None

    try:
        with open(registro_path, 'r', encoding='utf-8') as f:
            registros = json.load(f)
    except Exception:
        return None

    # Buscar registros con misma fecha y mismo tipo de comida
    duplicados = []
    clave_nueva = _clave_de_registro(entrada_nueva)
    for reg in registros:
        # Excluir la entrada recién guardada (misma clave)
        if _clave_de_registro(reg) == clave_nueva:
            continue
        tipo_reg = (reg.get("tipo_comida") or "").strip().lower()
        if tipo_reg != tipo:
            continue
        fecha_reg_ymd = _fecha_ymd(reg.get("fecha_emision") or "")
        if fecha_reg_ymd == fecha_nueva_ymd:
            duplicados.append(reg)

    if not duplicados:
        return None

    # Formatear fecha legible para el mensaje
    try:
        from datetime import datetime as _dt
        fecha_display = _dt.strptime(fecha_nueva_ymd, "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        fecha_display = fecha_nueva_ymd

    total = len(duplicados) + 1  # +1 por la recién guardada
    tipo_display = tipo.capitalize()
    razones = ", ".join(
        r.get("razon_social") or r.get("establecimiento") or "desconocido"
        for r in duplicados
    )
    return (
        f"⚠️ *Alerta: {total} {tipo_display}s el {fecha_display}*\n"
        f"Ya tenías {len(duplicados)} {tipo_display} registrado(s) ese día ({razones}).\n"
        f"La factura fue guardada igualmente — verifica si es correcto."
    )


# ───── Kilometraje / Odómetro (solo para facturas de gasolina) ─────
# _MESES_ES, _formatear_rango_fechas y _fecha_para_orden se importan de comun.py.


def guardar_odometro(user_id: int, username: str, nit: str, serie: str, numero: str, odometro: int) -> bool:
    """Guarda/actualiza el odómetro de una factura de gasolina ya registrada.
    Retorna True si se actualizó, False si no se encontró la factura."""
    carpeta = carpeta_usuario(user_id, username)
    registro_path = carpeta / "registro.json"
    if not registro_path.exists():
        return False
    with _lock_archivo(registro_path):
        try:
            with open(registro_path, 'r', encoding='utf-8') as f:
                registros = json.load(f)
        except Exception as e:
            logger.error(f"registro.json ilegible al guardar odometro: {e}")
            return False
        clave = f"{nit}|{serie}|{numero}"
        encontrado = False
        for r in registros:
            if _clave_de_registro(r) == clave:
                r['odometro'] = int(odometro)
                encontrado = True
        if encontrado:
            _guardar_json_atomico(registro_path, registros)
        return encontrado


def obtener_odometro_anterior(user_id: int, username: str, fecha_referencia: str, clave_excluir: str = "") -> Optional[dict]:
    """Busca el registro de gasolina con odómetro más reciente ANTES de fecha_referencia.
    Retorna dict con {odometro, fecha_emision, nit, serie, numero} o None."""
    carpeta = carpeta_usuario(user_id, username)
    registro_path = carpeta / "registro.json"
    if not registro_path.exists():
        return None
    try:
        with open(registro_path, 'r', encoding='utf-8') as f:
            registros = json.load(f)
    except Exception:
        return None
    ref = _fecha_para_orden(fecha_referencia)
    candidatos = []
    for r in registros:
        if (r.get('tipo_comida') or '').lower() != 'gasolina':
            continue
        if r.get('odometro') is None:
            continue
        clave_r = _clave_de_registro(r)
        if clave_r == clave_excluir:
            continue
        f_r = _fecha_para_orden(r.get('fecha_emision') or '')
        if ref and f_r and f_r > ref:
            continue
        candidatos.append((f_r or datetime.min, r))
    if not candidatos:
        return None
    candidatos.sort(key=lambda x: x[0])
    return candidatos[-1][1]


def listar_facturas_gasolina(user_id: int, username: str) -> list:
    """Devuelve lista de facturas de gasolina ordenadas por fecha (más reciente primero)."""
    carpeta = carpeta_usuario(user_id, username)
    registro_path = carpeta / "registro.json"
    if not registro_path.exists():
        return []
    try:
        with open(registro_path, 'r', encoding='utf-8') as f:
            registros = json.load(f)
    except Exception:
        return []
    # Enriquecer con clasificación para no perder gasolina sin tipo_comida explícito
    gas = []
    for r in registros:
        _enriquecer_desde_pdf(r)
        _clasificar_por_razon_social(r)
        if (r.get('tipo_comida') or '').lower() == 'gasolina':
            gas.append(r)
    gas.sort(key=lambda r: _fecha_para_orden(r.get('fecha_emision') or '') or datetime.min, reverse=True)
    return gas


def validar_odometro_nuevo(odometro: int, anterior: Optional[dict]) -> Optional[str]:
    """Retorna un mensaje de advertencia si el odómetro es sospechoso, None si está OK."""
    if odometro < 0:
        return "⚠️ El odómetro no puede ser negativo."
    if anterior is None:
        return None
    prev = anterior.get('odometro')
    if prev is None:
        return None
    if odometro < prev:
        return (f"⚠️ El nuevo odómetro ({odometro:,} km) es MENOR que el anterior "
                f"({prev:,} km del {anterior.get('fecha_emision', '?')}). Verifica.")
    if odometro - prev > 2000:
        return (f"⚠️ Salto grande: {odometro - prev:,} km desde la última carga "
                f"({prev:,} km el {anterior.get('fecha_emision', '?')}). ¿Está correcto?")
    return None


# _guardar_json_atomico se importa de comun.py (arriba).


# Clasificación, política de viáticos y normalización de departamentos viven
# en clasificacion.py (modulo autocontenido, sin dependencias de Telegram).
# Se importan aqui para mantener el resto del codigo sin cambios.
from clasificacion import (
    DEPARTAMENTOS_BD,
    DEPTOS_TABLA1_BD,
    POLITICA_VIATICOS,
    HOSPEDAJE_MONTO_MINIMO,
    KEYWORDS_GASOLINA_LARGAS,
    KEYWORDS_GASOLINA_CORTAS,
    KEYWORDS_GASOLINA,
    KEYWORDS_HOSPEDAJE,
    KEYWORDS_FOOD_SIGNAL,
    KEYWORDS_PARQUEO,
    _reclasificar_tipo_si_necesario,
    _normalizar,
    normalizar_departamento,
    _monto_efectivo,
    _clasificar_por_razon_social,
    _enriquecer_desde_pdf,
    _extraer_depto_de_pdf_sat,
    _resolver_departamento,
)

# Generación de reportes (Excel + TXT) vive en reportes.py. PLANTILLA_REPORTE
# se reexporta porque telegram_bot la usa en chequeos de arranque.
from reportes import (
    PLANTILLA_REPORTE,
    generar_reporte_excel,
    _filtrar_facturas_para_reporte,
)


def _extraer_fel_de_pdf(pdf_path: Path) -> dict:
    """Extrae serie/numero/NIT del TEXTO embebido de un PDF oficial de FEL.

    El PDF de la SAT trae el texto exacto (sin errores de OCR ni de lectura de
    QR). Para un PDF de FEL esto es la fuente mas confiable: evita que un OCR
    de imagen sobre-expuesta o un QR mal decodificado manden datos erroneos al
    SAT. Devuelve {} si el PDF no parece una FEL o no se pudo leer.
    """
    datos: dict = {}
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(str(pdf_path))
        txt = "".join(p.get_text() for p in doc)
        doc.close()
    except Exception as e:
        logger.debug(f"_extraer_fel_de_pdf no pudo leer {pdf_path}: {e}")
        return {}
    # Solo serie y numero: son los que el OCR de imagen lee mal y hacen que el
    # SAT rechace. El NIT NO se extrae del PDF porque en el texto viene pegado
    # al UUID de autorizacion (ej. "Nit Emisor: 708759366BD48E0E...") y no se
    # puede separar con confianza; ademas el NIT ya viene bien por catalogo/QR.
    # Ej. de texto: "Serie: 6BD48E0E Número de DTE: 7522390939"
    m = re.search(r'Serie:\s*([0-9A-Za-z]{6,12})', txt)
    if m:
        datos['serie'] = m.group(1).upper()
    m = re.search(r'N[uú]mero\s+de\s+DTE:\s*(\d{6,})', txt, re.IGNORECASE)
    if m:
        datos['numero'] = m.group(1)
    return datos


def _guardar_factura_fallida(user_id: int, username: str, resultado, error_msg: str, pend_id: str):
    """Guarda una factura que agotó todos los reintentos en datos/facturas_fallidas.json."""
    fallidas_path = DATOS_DIR / "facturas_fallidas.json"
    fallidas = []
    if fallidas_path.exists():
        try:
            with open(fallidas_path, 'r', encoding='utf-8') as f:
                fallidas = json.load(f)
        except Exception:
            fallidas = []
    entrada = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "user_id": user_id,
        "username": username,
        "pend_id": pend_id,
        "nit": resultado.nit,
        "serie": resultado.serie,
        "numero": resultado.numero_factura,
        "ultimo_error": error_msg,
        "resuelta": False,
    }
    # Actualizar si ya existe el mismo pend_id
    for i, falla in enumerate(fallidas):
        if falla.get('pend_id') == pend_id:
            fallidas[i] = entrada
            break
    else:
        fallidas.append(entrada)
    _guardar_json_atomico(fallidas_path, fallidas)
    logger.info(f"Factura fallida guardada en facturas_fallidas.json: NIT={resultado.nit} Serie={resultado.serie} Num={resultado.numero_factura}")


def _copiar_pdf_a_usuario(res, user_id: int, username: str) -> "Optional[Path]":
    """Copia el PDF descargado de SAT (en downloads/) a la carpeta propia del usuario.
    Retorna la ruta destino si tuvo éxito, o la ruta original si falla."""
    if not res.pdf_path or not res.pdf_path.exists():
        return res.pdf_path
    try:
        carpeta = carpeta_usuario(user_id, username)
        nombre_pdf = res.pdf_path.name
        destino = carpeta / "pdf_sat" / nombre_pdf
        n = 1
        while destino.exists():
            destino = carpeta / "pdf_sat" / f"{nombre_pdf.rsplit('.', 1)[0]} ({n}).pdf"
            n += 1
        destino.write_bytes(res.pdf_path.read_bytes())
        logger.info(f"PDF copiado a carpeta usuario: {destino}")
        return destino
    except Exception as e:
        logger.warning(f"No se pudo copiar PDF a carpeta usuario: {e}")
        return res.pdf_path


# Almacenar datos pendientes de confirmación — cada factura tiene su propio ID
pendientes = {}  # {pend_id: {resultado, user_id, archivo, username, timestamp}}
PENDIENTES_FILE = DATOS_DIR / "pendientes.json"

# Cola de fotos por usuario para detección de batch (no depende de media_group_id)
_cola_fotos: dict = {}  # user_id -> {archivos, username, chat_id, progress_msg, task}
BATCH_TIMEOUT = 3.0  # segundos de espera tras la última foto para procesar el batch

# Sesiones temporales para flujo de reporte editable y eliminación
reporte_sessions = {}  # {session_id: {user_id, username, ceco, fi, ff, facturas, excluidas, timestamp}}
eliminar_sessions = {}  # {session_id: {user_id, username, facturas_display, all_registros}}
clasificar_sessions = {}  # {session_id: {user_id, username, ceco, fi, ff, codigo_empleado, pendientes, idx}}
odometro_sessions = {}  # {session_id: {user_id, username, claves, facturas, timestamp}}
tipo_sessions = {}  # {session_id: {user_id, username, facturas, timestamp}}


def _guardar_pendientes():
    """Persiste el dict pendientes en disco para sobrevivir reinicios del bot."""
    try:
        data = {}
        for pid, info in pendientes.items():
            resultado = info.get('resultado')
            ts = info.get('timestamp', datetime.now())
            data[pid] = {
                'resultado': resultado.to_dict() if resultado else {},
                'user_id': info.get('user_id'),
                'archivo': info.get('archivo', ''),
                'username': info.get('username', ''),
                'timestamp': ts.isoformat() if isinstance(ts, datetime) else str(ts),
            }
        _guardar_json_atomico(PENDIENTES_FILE, data)
    except Exception as e:
        logger.error(f"Error guardando pendientes en disco: {e}")


def _cargar_pendientes():
    """Carga pendientes guardados desde disco al arrancar el bot."""
    if not PENDIENTES_FILE.exists():
        return
    try:
        with open(PENDIENTES_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        count = 0
        descartados = 0
        for pid, info in data.items():
            res_dict = info.get('resultado', {})
            if not res_dict.get('nit'):
                continue  # Saltar entradas sin datos
            resultado = ExtraccionFactura(
                nit=res_dict.get('nit', ''),
                serie=res_dict.get('serie', ''),
                numero_factura=res_dict.get('numero_factura', ''),
                razon_social=res_dict.get('razon_social'),
                metodo_extraccion=res_dict.get('metodo_extraccion', 'ocr_directo'),
                confianza=res_dict.get('confianza', 0.5),
                texto_ocr=res_dict.get('texto_ocr', ''),
            )
            ts_str = info.get('timestamp', '')
            try:
                ts = datetime.fromisoformat(ts_str) if ts_str else datetime.now()
            except ValueError:
                ts = datetime.now()
            # Solo restaurar si no tiene mas de 7 dias (suficiente para que el
            # usuario corrija facturas que quedaron pendientes despues de un
            # reinicio o de un /pendientes diferido).
            if (datetime.now() - ts).total_seconds() < 7 * 86400:
                pendientes[pid] = {
                    'resultado': resultado,
                    'user_id': info.get('user_id'),
                    'archivo': info.get('archivo', ''),
                    'username': info.get('username', ''),
                    'timestamp': ts,
                }
                count += 1
            else:
                descartados += 1
                logger.warning(
                    f"Pendiente descartado por edad (>7d): pid={pid} "
                    f"user_id={info.get('user_id')} nit={res_dict.get('nit', '?')} "
                    f"numero={res_dict.get('numero_factura', '?')} creado={ts_str}"
                )
        if count:
            logger.info(f"Pendientes restaurados desde disco: {count} facturas")
        if descartados:
            logger.warning(f"Pendientes descartados por edad (>7d): {descartados}")
    except Exception as e:
        logger.error(f"Error cargando pendientes desde disco: {e}")


class BotFacturasGT:
    """Gestor del bot de Telegram para facturas"""
    
    def __init__(self, token: str):
        if not TELEGRAM_AVAILABLE:
            raise ImportError("Telegram no disponible")
        
        self.token = token
        # concurrent_updates=4: procesa hasta 4 mensajes en paralelo, compatible con JobQueue
        self.app = Application.builder().token(token).concurrent_updates(4).build()
        self._configurar_handlers()
    
    def _configurar_handlers(self):
        """Registra todos los handlers del bot"""
        # Comandos
        self.app.add_handler(CommandHandler("start", self.cmd_start))
        self.app.add_handler(CommandHandler("catalogo", self.cmd_catalogo))

        self.app.add_handler(CommandHandler("reporte", self.cmd_reporte))
        self.app.add_handler(CommandHandler("perfil", self.cmd_perfil))
        self.app.add_handler(CommandHandler("buscar", self.cmd_buscar))
        self.app.add_handler(CommandHandler("eliminar", self.cmd_eliminar))
        self.app.add_handler(CommandHandler("fallidas", self.cmd_fallidas))
        self.app.add_handler(CommandHandler("estado", self.cmd_estado))
        self.app.add_handler(CommandHandler("pendientes", self.cmd_pendientes))
        self.app.add_handler(CommandHandler("cancelar_todas", self.cmd_cancelar_todas))
        self.app.add_handler(CommandHandler("odometro", self.cmd_odometro))
        self.app.add_handler(CommandHandler("tipo", self.cmd_tipo))
        self.app.add_handler(CommandHandler("reload", self.cmd_reload))
        self.app.add_handler(CommandHandler("help", self.cmd_help))
        self.app.add_handler(CommandHandler("admin", self.cmd_admin))
        
        # Mensajes
        self.app.add_handler(MessageHandler(filters.PHOTO, self.msg_procesar_foto))
        self.app.add_handler(MessageHandler(filters.Document.PDF, self.msg_procesar_pdf))
        self.app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, self.msg_texto))
        
        # Callbacks (botones)
        self.app.add_handler(CallbackQueryHandler(self.btn_confirmación))

        # Error handler global — captura errores no manejados para que no rompan el bot
        self.app.add_error_handler(self._error_handler)
    
    async def _guard_acceso(self, update: Update) -> bool:
        """Verifica si el usuario tiene acceso. Retorna True si puede continuar."""
        user = update.effective_user
        user_id = user.id
        username = user.username or user.first_name or ""
        first_name = user.first_name or ""

        # Registrar usuario si es nuevo
        entrada = _registrar_usuario(user_id, username, first_name)

        # Admin siempre tiene acceso
        if _es_admin(user_id):
            return True

        permitido, motivo = _usuario_permitido(user_id)

        if motivo == "bloqueado":
            await update.effective_message.reply_text(
                "🚫 Tu acceso ha sido suspendido. Contacta al administrador."
            )
            return False

        return True

    async def cmd_start(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Comando /start"""
        user = update.effective_user
        user_id = user.id
        username = user.username or user.first_name or ""
        _registrar_usuario(user_id, username, user.first_name or "")

        if not await self._guard_acceso(update):
            return

        mensaje = """
🤖 *Bot Extractor de Facturas SAT*

Hola! Soy tu asistente para extraer datos de facturas guatemaltecas.

*Como usar:*
1. Enviame una foto o PDF de una factura
2. Extraere: NIT, numero, serie automaticamente
3. Si algo falla, buscare en mi catalogo de empresas
4. Si todo ok, confirmas y listo!

*Comandos:*
/catalogo - Ver empresas conocidas
/agregar - Agregar nueva empresa
/reporte - Generar Excel de gastos (editable)
/buscar - Buscar facturas por empresa, NIT, etc.
/eliminar - Eliminar una factura del registro
/fallidas - Ver facturas fallidas y reintentar
/pendientes - Ver facturas sin confirmar
/odometro - Kilometraje de facturas de gasolina
/tipo - Cambiar tipo (Desayuno/Almuerzo/Cena/Gasolina/Hospedaje/Parqueo) de una factura
/reload - Reiniciar sesion SAT (cuando se queda trabada o no llega captcha)
/estado - Estado actual del bot
/help - Esta ayuda

*Tips:*
- Foto clara y bien enfocada o PDF directo
- Si SAT falla, reintentare automaticamente
- Si una factura falla, queda guardada en /fallidas
- Las facturas pendientes se restauran si el bot se reinicia
"""
        await update.message.reply_text(mensaje, parse_mode="Markdown")
    
    async def cmd_help(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Comando /help"""
        await self.cmd_start(update, context)

    async def cmd_pendientes(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Lista las facturas pendientes de confirmar del usuario
        (o de todos, si es admin). Son las facturas que ya pasaron OCR/SAT
        pero estan esperando que el usuario confirme/corrija/rechace."""
        if not await self._guard_acceso(update):
            return
        user_id = update.message.from_user.id
        es_admin = _es_admin(user_id)

        if es_admin:
            items = list(pendientes.items())
            encabezado = "📋 *Pendientes (todos los usuarios)*"
        else:
            items = [(pid, info) for pid, info in pendientes.items()
                     if info.get('user_id') == user_id]
            encabezado = "📋 *Tus facturas pendientes*"

        if not items:
            await update.message.reply_text(
                "✅ No tienes facturas pendientes de confirmar."
                if not es_admin else
                "✅ No hay facturas pendientes en memoria."
            )
            return

        # Ordenar por mas reciente primero
        items.sort(key=lambda x: x[1].get('timestamp', datetime.now()), reverse=True)

        lineas = [f"{encabezado}  (total: {len(items)})\n"]
        for pid, info in items[:20]:  # limitar a 20 para no pasar 4096 chars
            resultado = info.get('resultado')
            if not resultado:
                continue
            ts = info.get('timestamp', datetime.now())
            if not isinstance(ts, datetime):
                try:
                    ts = datetime.fromisoformat(str(ts))
                except (ValueError, TypeError):
                    ts = datetime.now()
            edad = datetime.now() - ts
            horas = int(edad.total_seconds() // 3600)
            minutos = int((edad.total_seconds() % 3600) // 60)
            edad_str = f"{horas}h{minutos}m" if horas else f"{minutos}m"
            nit = getattr(resultado, 'nit', '?') or '?'
            serie = getattr(resultado, 'serie', '?') or '?'
            numero = getattr(resultado, 'numero_factura', '?') or '?'
            lineas.append(
                f"• `{pid[:8]}` — NIT {nit} · Serie {serie} · Num {numero}  ({edad_str})"
            )
            if es_admin:
                uname = info.get('username', '') or '?'
                lineas[-1] += f"\n   👤 @{uname} (id {info.get('user_id','?')})"

        if len(items) > 20:
            lineas.append(f"\n… y {len(items) - 20} mas.")

        lineas.append(
            "\n💡 Para confirmar/corregir, usa los botones del mensaje original "
            "de cada factura (busca mas arriba en el chat)."
        )
        teclado_acciones = None
        if len(items) >= 1:
            teclado_acciones = InlineKeyboardMarkup([
                [InlineKeyboardButton(
                    f"✅ Validar TODAS en SAT ({len(items)})",
                    callback_data="verificar_todas",
                )],
                [InlineKeyboardButton(
                    f"🔧 Revisar 1 a 1 ({len(items)})",
                    callback_data="revisar_pendientes",
                )],
                [InlineKeyboardButton(
                    f"🗑️ Cancelar TODAS ({len(items)})",
                    callback_data="cancelar_todas_confirm",
                )],
            ])
        await update.message.reply_text(
            "\n".join(lineas),
            parse_mode="Markdown",
            reply_markup=teclado_acciones,
        )

    async def cmd_cancelar_todas(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Cancela y elimina todas las facturas pendientes del usuario."""
        if not await self._guard_acceso(update):
            return
        user_id = update.message.from_user.id
        mis_ids = [pid for pid, info in pendientes.items() if info.get('user_id') == user_id]
        if not mis_ids:
            await update.message.reply_text("✅ No tienes facturas pendientes.")
            return
        teclado = InlineKeyboardMarkup([[
            InlineKeyboardButton(f"🗑️ Sí, cancelar las {len(mis_ids)}", callback_data="cancelar_todas_confirm"),
            InlineKeyboardButton("No", callback_data="cancelar_todas_abort"),
        ]])
        await update.message.reply_text(
            f"⚠️ Tienes {len(mis_ids)} facturas pendientes.\n¿Cancelar y descartar todas?",
            reply_markup=teclado
        )

    # ─────────────────────────────────────────────
    #  PANEL DE ADMIN
    # ─────────────────────────────────────────────
    async def cmd_admin(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Panel de administración — solo accesible por el admin."""
        user_id = update.message.from_user.id
        if not _es_admin(user_id):
            await update.message.reply_text("⛔ No tienes permisos para este comando.")
            return

        args = context.args or []
        subcomando = args[0].lower() if args else ""

        # /admin  →  menú principal
        if not subcomando:
            db = _cargar_users_db()
            total = len(db)
            aprobados = sum(1 for u in db.values() if u.get('aprobado') and not u.get('bloqueado'))
            pendientes = sum(1 for u in db.values() if not u.get('aprobado') and not u.get('bloqueado'))
            bloqueados = sum(1 for u in db.values() if u.get('bloqueado'))
            texto = (
                f"🛠 *Panel de Administración*\n\n"
                f"👥 Usuarios totales: {total}\n"
                f"✅ Activos: {aprobados}\n"
                f"⏳ Pendientes: {pendientes}\n"
                f"🚫 Bloqueados: {bloqueados}\n\n"
                f"*Comandos disponibles:*\n"
                f"`/admin usuarios` — lista todos\n"
                f"`/admin info <id>` — detalle de un usuario\n"
                f"`/admin aprobar <id>` — aprobar acceso\n"
                f"`/admin bloquear <id> [motivo]` — bloquear\n"
                f"`/admin desbloquear <id>` — desbloquear\n"
                f"`/admin limite <id> <n>` — cambiar límite mensual\n"
                f"`/admin nota <id> <texto>` — agregar nota\n"
                f"`/admin pendientes` — ver solo pendientes\n"
            )
            await update.message.reply_text(texto, parse_mode="Markdown")
            return

        # /admin usuarios
        if subcomando == "usuarios":
            db = _cargar_users_db()
            if not db:
                await update.message.reply_text("No hay usuarios registrados aún.")
                return
            lineas = ["👥 *Todos los usuarios:*\n"]
            for uid, u in sorted(db.items(), key=lambda x: x[1].get('fecha_registro', '')):
                estado = "✅" if u.get('aprobado') and not u.get('bloqueado') else ("🚫" if u.get('bloqueado') else "⏳")
                nombre = u.get('username') or u.get('first_name') or '?'
                facturas = contar_facturas_mes(int(uid), u.get('username', ''))
                limite = u.get('limite_mensual', LIMITE_FACTURAS_MES)
                lineas.append(f"{estado} *{nombre}* (ID: `{uid}`)\n   Facturas este mes: {facturas}/{limite}")
            await update.message.reply_text("\n".join(lineas), parse_mode="Markdown")
            return

        # /admin pendientes
        if subcomando == "pendientes":
            db = _cargar_users_db()
            pendientes = {uid: u for uid, u in db.items() if not u.get('aprobado') and not u.get('bloqueado')}
            if not pendientes:
                await update.message.reply_text("✅ No hay usuarios pendientes.")
                return
            for uid, u in pendientes.items():
                nombre = u.get('username') or u.get('first_name') or '?'
                teclado = InlineKeyboardMarkup([[
                    InlineKeyboardButton("✅ Aprobar", callback_data=f"admin_aprobar:{uid}"),
                    InlineKeyboardButton("🚫 Bloquear", callback_data=f"admin_bloquear:{uid}"),
                ]])
                await update.message.reply_text(
                    f"⏳ *{nombre}*\nID: `{uid}`\nRegistro: {u.get('fecha_registro','')}",
                    parse_mode="Markdown", reply_markup=teclado
                )
            return

        # /admin info <id>
        if subcomando == "info" and len(args) >= 2:
            uid = args[1]
            db = _cargar_users_db()
            u = db.get(uid)
            if not u:
                await update.message.reply_text(f"Usuario `{uid}` no encontrado.", parse_mode="Markdown")
                return
            estado = "✅ Activo" if u.get('aprobado') and not u.get('bloqueado') else ("🚫 Bloqueado" if u.get('bloqueado') else "⏳ Pendiente")
            facturas = contar_facturas_mes(int(uid), u.get('username', ''))
            texto = (
                f"👤 *Información de usuario*\n\n"
                f"ID: `{uid}`\n"
                f"Usuario: @{u.get('username','')}\n"
                f"Nombre: {u.get('first_name','')}\n"
                f"Estado: {estado}\n"
                f"Facturas este mes: {facturas}/{u.get('limite_mensual', LIMITE_FACTURAS_MES)}\n"
                f"Registro: {u.get('fecha_registro','')}\n"
                f"Última actividad: {u.get('ultima_actividad','')}\n"
                f"Notas: {u.get('notas','—')}\n"
            )
            teclado = InlineKeyboardMarkup([[
                InlineKeyboardButton("✅ Aprobar", callback_data=f"admin_aprobar:{uid}"),
                InlineKeyboardButton("🚫 Bloquear", callback_data=f"admin_bloquear:{uid}"),
                InlineKeyboardButton("🔓 Desbloquear", callback_data=f"admin_desbloquear:{uid}"),
            ]])
            await update.message.reply_text(texto, parse_mode="Markdown", reply_markup=teclado)
            return

        # /admin aprobar <id>
        if subcomando == "aprobar" and len(args) >= 2:
            uid = int(args[1])
            if _aprobar_usuario(uid):
                await update.message.reply_text(f"✅ Usuario `{uid}` aprobado.", parse_mode="Markdown")
                try:
                    await self.app.bot.send_message(chat_id=uid, text="✅ Tu acceso ha sido aprobado. Ya puedes usar el bot.")
                except Exception:
                    pass
            else:
                await update.message.reply_text(f"❌ Usuario `{uid}` no encontrado.", parse_mode="Markdown")
            return

        # /admin bloquear <id> [motivo]
        if subcomando == "bloquear" and len(args) >= 2:
            uid = int(args[1])
            motivo = " ".join(args[2:]) if len(args) > 2 else ""
            if _bloquear_usuario(uid, motivo):
                await update.message.reply_text(f"🚫 Usuario `{uid}` bloqueado.", parse_mode="Markdown")
                try:
                    await self.app.bot.send_message(chat_id=uid, text="🚫 Tu acceso ha sido suspendido.")
                except Exception:
                    pass
            else:
                await update.message.reply_text(f"❌ Usuario `{uid}` no encontrado.", parse_mode="Markdown")
            return

        # /admin desbloquear <id>
        if subcomando == "desbloquear" and len(args) >= 2:
            uid = int(args[1])
            if _desbloquear_usuario(uid):
                await update.message.reply_text(f"🔓 Usuario `{uid}` desbloqueado y aprobado.", parse_mode="Markdown")
                try:
                    await self.app.bot.send_message(chat_id=uid, text="✅ Tu acceso ha sido restaurado. Ya puedes usar el bot.")
                except Exception:
                    pass
            else:
                await update.message.reply_text(f"❌ Usuario `{uid}` no encontrado.", parse_mode="Markdown")
            return

        # /admin limite <id> <n>
        if subcomando == "limite" and len(args) >= 3:
            uid = int(args[1])
            try:
                limite = int(args[2])
            except ValueError:
                await update.message.reply_text("Uso: /admin limite <id> <numero>")
                return
            if _set_limite_usuario(uid, limite):
                await update.message.reply_text(f"✅ Límite de `{uid}` actualizado a *{limite}* facturas/mes.", parse_mode="Markdown")
            else:
                await update.message.reply_text(f"❌ Usuario `{uid}` no encontrado.", parse_mode="Markdown")
            return

        # /admin nota <id> <texto>
        if subcomando == "nota" and len(args) >= 3:
            uid = int(args[1])
            nota = " ".join(args[2:])
            if _set_nota_usuario(uid, nota):
                await update.message.reply_text(f"✅ Nota guardada para `{uid}`.", parse_mode="Markdown")
            else:
                await update.message.reply_text(f"❌ Usuario `{uid}` no encontrado.", parse_mode="Markdown")
            return

        await update.message.reply_text("Subcomando no reconocido. Usa `/admin` para ver opciones.", parse_mode="Markdown")
    
    async def cmd_catalogo(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Muestra catálogo de empresas"""
        empresas = sistema.catalogo.listar_empresas()
        
        if not empresas:
            await update.message.reply_text("📭 Catálogo vacío. Agrega empresas con /agregar")
            return
        
        mensaje = "📋 *Empresas en catálogo:*\n\n"
        for i, emp in enumerate(empresas, 1):
            mensaje += f"{i}. *{emp['razon_social']}*\n"
            mensaje += f"   NIT: `{emp['nit']}`\n"
            if emp.get('serie'):
                mensaje += f"   Serie: {emp['serie']}\n"
            mensaje += "\n"
        
        await update.message.reply_text(mensaje, parse_mode="Markdown")
    
    async def cmd_reporte(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Comando /reporte — pide formato, CECO y muestra selector de fechas con botones."""
        user_id = update.message.from_user.id
        username = update.message.from_user.username or update.message.from_user.first_name or ""
        args = context.args or []

        # Sin args: primero preguntamos el formato del reporte.
        if len(args) == 0:
            teclado = InlineKeyboardMarkup([
                [InlineKeyboardButton("📊 Completo (Excel + TXT con política)",
                                      callback_data="reporte_formato:completo")],
                [InlineKeyboardButton("📄 Simple (solo TXT sin política)",
                                      callback_data="reporte_formato:simple")],
            ])
            await update.message.reply_text(
                "¿Qué tipo de reporte deseas generar?\n\n"
                "• *Completo*: Excel con fórmulas de política + TXT con columna "
                "'A COBRAR'. Ideal para liquidación interna.\n"
                "• *Simple*: solo TXT sin columna 'A COBRAR'. Útil para "
                "compartir con externos que no aplican tu política.",
                parse_mode="Markdown",
                reply_markup=teclado,
            )
            return

        # Con args: flujo CLI legado, formato completo por defecto.
        context.user_data['formato_reporte'] = 'completo'
        ceco = args[0]
        if len(args) >= 3:
            await self._generar_y_enviar_reporte(
                update.message, user_id, username, ceco, args[1], args[2],
                formato='completo',
            )
        elif len(args) == 2:
            await self._generar_y_enviar_reporte(
                update.message, user_id, username, ceco, args[1], args[1],
                formato='completo',
            )
        else:
            context.user_data['ceco'] = ceco
            context.user_data['estado'] = 'esperando_fechas_reporte'
            await self._pedir_fechas_reporte(update.message, ceco)

    async def _continuar_reporte_tras_formato(self, message, user_id: int, username: str, context):
        """Paso 2 de /reporte: ya se eligió formato, ahora pedir CECO o fechas.

        Se llama tanto desde el callback del menú de formato como desde la
        CLI (que asume formato=completo).
        """
        perfil = leer_perfil_usuario(user_id, username)
        ceco_guardado = perfil.get('ceco', '')
        cod_guardado = perfil.get('codigo_empleado', '')
        if ceco_guardado and cod_guardado:
            context.user_data['ceco'] = ceco_guardado
            context.user_data['codigo_empleado'] = cod_guardado
            context.user_data['estado'] = 'esperando_fechas_reporte'
            await message.reply_text(
                f"Usando CECO: *{ceco_guardado}* | Codigo: *{cod_guardado}*\n\n"
                f"Para cambiarlos usa /perfil",
                parse_mode="Markdown",
            )
            await self._pedir_fechas_reporte(message, ceco_guardado)
            return
        context.user_data['estado'] = 'esperando_ceco_reporte'
        msg = "Ingresa tu numero de CECO para el reporte:\n(Ej: 349065)"
        if ceco_guardado:
            msg += f"\n\nUltimo CECO usado: *{ceco_guardado}* — escribe 'ok' para usarlo"
        await message.reply_text(msg, parse_mode="Markdown")

    async def _pedir_fechas_reporte(self, message, ceco: str):
        """Muestra selector de fechas con botones rapidos."""
        hoy = datetime.now()
        # Semana actual: lunes a hoy
        lunes = hoy - timedelta(days=hoy.weekday())
        fi_semana = lunes.strftime("%d/%m/%Y")
        ff_semana = hoy.strftime("%d/%m/%Y")
        # Semana pasada
        lunes_pasado = lunes - timedelta(days=7)
        domingo_pasado = lunes - timedelta(days=1)
        fi_pasada = lunes_pasado.strftime("%d/%m/%Y")
        ff_pasada = domingo_pasado.strftime("%d/%m/%Y")
        # Mes actual
        fi_mes = hoy.replace(day=1).strftime("%d/%m/%Y")
        ff_mes = hoy.strftime("%d/%m/%Y")

        teclado = InlineKeyboardMarkup([
            [InlineKeyboardButton(f"📅 Semana actual ({fi_semana} – {ff_semana})",
                                  callback_data=f"reporte_preview:{ceco}|{fi_semana}|{ff_semana}")],
            [InlineKeyboardButton(f"📅 Semana pasada ({fi_pasada} – {ff_pasada})",
                                  callback_data=f"reporte_preview:{ceco}|{fi_pasada}|{ff_pasada}")],
            [InlineKeyboardButton(f"📆 Mes actual ({fi_mes} – {ff_mes})",
                                  callback_data=f"reporte_preview:{ceco}|{fi_mes}|{ff_mes}")],
            [InlineKeyboardButton("📋 Todas las facturas",
                                  callback_data=f"reporte_preview:{ceco}||")],
        ])
        await message.reply_text(
            f"CECO: {ceco}\n\nElige el periodo del reporte:",
            reply_markup=teclado
        )

    async def _crear_zip_pdfs(self, facturas: list, carpeta: Path, nombre_zip: str) -> Optional[Path]:
        """Crea un ZIP con los PDFs de las facturas del reporte.
        Lee los PDFs de carpeta/pdf_sat y escribe el ZIP en carpeta/reportes."""
        from comun import subcarpeta_reportes
        zip_path = subcarpeta_reportes(carpeta) / nombre_zip
        pdfs_incluidos = 0
        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for fac in facturas:
                    pdf_str = fac.get('pdf_sat', '')
                    if not pdf_str:
                        continue
                    pdf_nombre = Path(pdf_str).name
                    pdf_candidato = carpeta / "pdf_sat" / pdf_nombre
                    if pdf_candidato.exists():
                        zf.write(pdf_candidato, pdf_nombre)
                        pdfs_incluidos += 1
            if pdfs_incluidos == 0:
                zip_path.unlink(missing_ok=True)
                return None
            return zip_path
        except Exception as e:
            logger.warning(f"No se pudo crear ZIP de PDFs: {e}")
            return None

    async def _crear_pdf_combinado(self, facturas: list, carpeta: Path, nombre_pdf: str) -> Optional[Path]:
        """Combina los PDFs SAT de todas las facturas en un solo archivo.

        Mismo orden que el ZIP (orden del reporte). Útil para imprimir todas
        las facturas de un tirón en vez de abrir una por una.
        """
        import fitz  # PyMuPDF
        from comun import subcarpeta_reportes
        pdf_path = subcarpeta_reportes(carpeta) / nombre_pdf
        try:
            combinado = fitz.open()
            paginas_agregadas = 0
            for fac in facturas:
                pdf_str = fac.get('pdf_sat', '')
                if not pdf_str:
                    continue
                pdf_nombre = Path(pdf_str).name
                pdf_candidato = carpeta / "pdf_sat" / pdf_nombre
                if not pdf_candidato.exists():
                    continue
                try:
                    src = fitz.open(str(pdf_candidato))
                    combinado.insert_pdf(src)
                    paginas_agregadas += src.page_count
                    src.close()
                except Exception as ex:
                    logger.warning(f"No se pudo insertar {pdf_nombre} al PDF combinado: {ex}")
                    continue
            if paginas_agregadas == 0:
                combinado.close()
                return None
            # garbage=4 + deflate=True reduce el tamaño del archivo final
            # (PDFs SAT suelen tener objetos redundantes que comprimen bien).
            combinado.save(str(pdf_path), garbage=4, deflate=True)
            combinado.close()
            return pdf_path
        except Exception as e:
            logger.warning(f"No se pudo crear PDF combinado: {e}")
            return None

    async def _mostrar_preview_reporte(self, message, user_id: int, username: str, ceco: str, fi: str, ff: str, codigo_empleado: str = "", formato: str = "completo"):
        """Muestra el contenido del reporte como texto con botones Generar / Editar."""
        facturas = _filtrar_facturas_para_reporte(user_id, username, fi, ff)
        if not facturas:
            await message.reply_text(
                "No hay facturas verificadas para ese rango.\n"
                "Usa /reporte para intentar con otro periodo."
            )
            return

        # ── Paso previo: clasificar facturas sin tipo ──────────────────────
        sin_tipo = [f for f in facturas if not f.get('tipo_comida')]
        if sin_tipo:
            sid = uuid.uuid4().hex[:10]
            clasificar_sessions[sid] = {
                'user_id': user_id,
                'username': username,
                'ceco': ceco,
                'fi': fi,
                'ff': ff,
                'codigo_empleado': codigo_empleado,
                'formato': formato,
                'pendientes': sin_tipo,
                'idx': 0,
                'timestamp': datetime.now(),
            }
            await self._mostrar_clasificacion_pendiente(message, sid)
            return
        # ──────────────────────────────────────────────────────────────────
        # Guardar sesion
        sid = uuid.uuid4().hex[:10]
        reporte_sessions[sid] = {
            'user_id': user_id,
            'username': username,
            'ceco': ceco,
            'codigo_empleado': codigo_empleado,
            'fi': fi,
            'ff': ff,
            'facturas': facturas,
            'fac_ids': [id(f) for f in facturas],
            'excluidas': set(),
            'formato': formato,
            'timestamp': datetime.now(),
        }
        # Construir preview de texto
        TIPO_CATEGORIA = {'Gasolina': 'GASOLINA', 'Desayuno': 'ALIMENTACION', 'Almuerzo': 'ALIMENTACION', 'Cena': 'ALIMENTACION', 'Hospedaje': 'HOSPEDAJE', 'Parqueo': 'PARQUEO'}
        lineas = [f"📋 *Preview del reporte — {len(facturas)} facturas*"]
        if fi or ff:
            lineas.append(f"Periodo: {fi or '...'} – {ff or '...'}")
        lineas.append("")
        total = 0.0
        for i, fac in enumerate(facturas, 1):
            empresa = fac.get('razon_social') or fac.get('establecimiento') or fac.get('nit', '?')
            fecha_raw = fac.get('fecha_emision', '') or ''
            # Tratar "00.00.0000" o "00/00/0000" como fecha sin datos
            fecha = fecha_raw if (fecha_raw and not fecha_raw.startswith('00')) else 'sin fecha'
            monto_val = _monto_efectivo(fac.get('monto', ''))
            total += monto_val
            tipo = TIPO_CATEGORIA.get(fac.get('tipo_comida', ''), 'OTROS')
            lineas.append(f"{i}. {empresa}\n   📅 {fecha} | 💰 Q{monto_val:.2f} | {tipo}")
        lineas.append(f"\n💰 *Total: Q{total:.2f}*")
        preview_txt = "\n".join(lineas)
        # Limitar a 4000 chars (limite Telegram)
        if len(preview_txt) > 3900:
            preview_txt = preview_txt[:3900] + "\n...(lista truncada)"
        teclado = InlineKeyboardMarkup([
            [InlineKeyboardButton("📊 Generar reporte", callback_data=f"reporte_generar:{sid}"),
             InlineKeyboardButton("✂️ Quitar facturas", callback_data=f"reporte_editar:{sid}")],
            [InlineKeyboardButton("❌ Cancelar", callback_data=f"reporte_cancelar:{sid}")],
        ])
        await message.reply_text(preview_txt, parse_mode="Markdown", reply_markup=teclado)

    async def _mostrar_clasificacion_pendiente(self, message, sid: str):
        """Muestra una factura sin clasificar y pide al usuario que la identifique."""
        session = clasificar_sessions.get(sid)
        if not session:
            return
        pendientes = session['pendientes']
        idx = session['idx']

        if idx >= len(pendientes):
            # Todas clasificadas → continuar al preview
            await message.reply_text("✅ Listo, todas clasificadas. Generando preview...")
            await self._mostrar_preview_reporte(
                message,
                session['user_id'], session['username'],
                session['ceco'], session['fi'], session['ff'],
                codigo_empleado=session['codigo_empleado'],
                formato=session.get('formato', 'completo'),
            )
            clasificar_sessions.pop(sid, None)
            return

        fac = pendientes[idx]
        empresa = fac.get('razon_social') or fac.get('establecimiento') or fac.get('nit', '?')
        fecha = fac.get('fecha_emision', '').replace('.', '/') or 'sin fecha'
        hora = fac.get('hora_emision', '') or 'sin hora'
        monto = fac.get('monto', 'Q?')

        texto = (
            f"❓ *Factura sin clasificar* ({idx + 1}/{len(pendientes)})\n\n"
            f"🏪 {empresa}\n"
            f"📅 {fecha}  🕐 {hora}\n"
            f"💰 {monto}\n\n"
            f"¿Qué tipo de gasto es?"
        )
        teclado = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("🌅 Desayuno",  callback_data=f"clasificar:{sid}|Desayuno"),
                InlineKeyboardButton("☀️ Almuerzo",  callback_data=f"clasificar:{sid}|Almuerzo"),
                InlineKeyboardButton("🌙 Cena",       callback_data=f"clasificar:{sid}|Cena"),
            ],
            [
                InlineKeyboardButton("⛽ Gasolina",   callback_data=f"clasificar:{sid}|Gasolina"),
                InlineKeyboardButton("🏨 Hospedaje",  callback_data=f"clasificar:{sid}|Hospedaje"),
                InlineKeyboardButton("🅿️ Parqueo",    callback_data=f"clasificar:{sid}|Parqueo"),
            ],
            [
                InlineKeyboardButton("⏭️ Saltar",     callback_data=f"clasificar:{sid}|"),
            ],
        ])
        await message.reply_text(texto, parse_mode="Markdown", reply_markup=teclado)

    async def _mostrar_editor_facturas(self, message, sid: str, session: dict):
        """Muestra lista de facturas con toggle para incluir/excluir del reporte."""
        facturas = session['facturas']
        excluidas = session.get('excluidas', set())
        fac_ids = session['fac_ids']
        botones = []
        for i, fac in enumerate(facturas):
            empresa = fac.get('razon_social') or fac.get('establecimiento') or fac.get('nit', '?')
            empresa = empresa[:25]
            monto = fac.get('monto', 'Q?')
            excluida = fac_ids[i] in excluidas
            icono = "❌" if excluida else "✅"
            botones.append([InlineKeyboardButton(
                f"{icono} {empresa} | {monto}",
                callback_data=f"reporte_excluir:{sid}|{i}"
            )])
        botones.append([InlineKeyboardButton("✔️ Confirmar y generar", callback_data=f"reporte_confirmar_edicion:{sid}")])
        botones.append([InlineKeyboardButton("❌ Cancelar", callback_data=f"reporte_cancelar:{sid}")])
        teclado = InlineKeyboardMarkup(botones)
        incluidas = len(facturas) - len(excluidas)
        texto = f"✂️ *Editar reporte*\nToca para incluir ✅ / quitar ❌ cada factura.\nActualmente: {incluidas}/{len(facturas)} facturas incluidas."
        try:
            await message.reply_text(texto, parse_mode="Markdown", reply_markup=teclado)
        except Exception:
            pass

    async def _generar_y_enviar_reporte(self, message, user_id, username, ceco, fecha_inicio="", fecha_fin="", facturas_override=None, codigo_empleado="", formato="completo"):
        """Genera el reporte según `formato` y lo envía.

        - formato="completo": Excel + TXT (con A COBRAR) + ZIP de PDFs.
        - formato="simple":   solo TXT (sin A COBRAR) + ZIP de PDFs.
        """
        etiqueta = "simple" if formato == "simple" else "completo"
        await message.reply_text(f"⏳ Generando reporte {etiqueta}...")
        try:
            resultado = generar_reporte_excel(
                user_id, username, fecha_inicio, fecha_fin,
                ceco=ceco, codigo_empleado=codigo_empleado,
                facturas_override=facturas_override, formato=formato,
            )
            if not resultado:
                await message.reply_text(
                    "No hay facturas verificadas para ese rango.\n\n"
                    "Uso: /reporte CECO (todas las facturas)\n"
                    "/reporte CECO 23/03/2026 26/03/2026 (rango de fechas)"
                )
                return
            ruta, ruta_txt, facturas, carpeta = resultado
            # En modo simple, ruta == None (sin Excel).
            if ruta is not None:
                with open(ruta, 'rb') as f:
                    await message.reply_document(
                        document=f, filename=ruta.name,
                        caption=f"✅ Reporte generado: {len(facturas)} facturas.",
                        write_timeout=120, read_timeout=60,
                    )
                # Avisar si hay facturas de gasolina sin odómetro (solo aplica al Excel)
                gas_sin_odo = [f for f in facturas if f.get('tipo_comida') == 'Gasolina' and f.get('odometro') is None]
                if gas_sin_odo:
                    lineas_aviso = [f"⚠️ *{len(gas_sin_odo)} factura(s) de gasolina sin odómetro:*"]
                    for f in gas_sin_odo[:5]:
                        lineas_aviso.append(f"   - {f.get('fecha_emision', '?')} {f.get('razon_social', '')[:25]} {f.get('monto', '')}")
                    lineas_aviso.append("\nUsa `/odometro` para agregarlos y regenera el reporte.")
                    await message.reply_text("\n".join(lineas_aviso), parse_mode="Markdown")
            caption_txt = (
                "📄 Detalle de facturas (sin política interna)."
                if formato == "simple"
                else "📄 Detalle de facturas para copiar/pegar."
            )
            with open(ruta_txt, 'rb') as f:
                await message.reply_document(document=f, filename=ruta_txt.name,
                                             caption=caption_txt,
                                             write_timeout=120, read_timeout=60)
            # Enviar ZIP con PDFs individuales (aplica a ambos formatos)
            base_zip = (ruta.stem if ruta is not None else ruta_txt.stem)
            nombre_zip = base_zip + "_PDFs.zip"
            zip_path = await self._crear_zip_pdfs(facturas, carpeta, nombre_zip)
            if zip_path:
                with open(zip_path, 'rb') as f:
                    await message.reply_document(document=f, filename=zip_path.name,
                                                 caption=f"📦 PDFs individuales ({zip_path.stat().st_size // 1024} KB).",
                                                 write_timeout=120, read_timeout=60)
            else:
                await message.reply_text("ℹ️ No hay PDFs guardados en tu carpeta para comprimir.")

            # PDF combinado (mismo orden del reporte) para impresión masiva.
            # Solo tiene sentido si hay 2+ PDFs — con uno solo basta el del ZIP.
            pdfs_disponibles = sum(
                1 for f in facturas
                if f.get('pdf_sat') and (carpeta / "pdf_sat" / Path(f['pdf_sat']).name).exists()
            )
            if pdfs_disponibles >= 2:
                nombre_pdf_combo = base_zip + "_PDFs_combinados.pdf"
                pdf_combo = await self._crear_pdf_combinado(facturas, carpeta, nombre_pdf_combo)
                if pdf_combo:
                    with open(pdf_combo, 'rb') as f:
                        await message.reply_document(
                            document=f, filename=pdf_combo.name,
                            caption=(
                                f"🖨️ PDF combinado ({pdfs_disponibles} facturas, "
                                f"{pdf_combo.stat().st_size // 1024} KB).\n"
                                f"Imprime todas de una sola vez."
                            ),
                            write_timeout=180, read_timeout=120,
                        )
        except Exception as e:
            logger.error(f"Error generando reporte: {e}", exc_info=True)
            await message.reply_text(f"❌ Error generando reporte: {e}")

    async def cmd_perfil(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Muestra y permite actualizar CECO y código de empleado del usuario."""
        user_id = update.message.from_user.id
        username = update.message.from_user.username or update.message.from_user.first_name or ""
        args = context.args or []
        perfil = leer_perfil_usuario(user_id, username)

        # /perfil ceco 349065  o  /perfil codigo 12345  o  /perfil nombre Juan  o  /perfil apellido Perez
        if len(args) >= 2:
            campo = args[0].lower()
            valor = " ".join(args[1:]).strip()
            if campo in ('ceco',):
                if not valor.isdigit():
                    await update.message.reply_text("El CECO debe ser un numero.")
                    return
                guardar_perfil_usuario(user_id, username, ceco=valor)
                await update.message.reply_text(f"CECO actualizado: *{valor}*", parse_mode="Markdown")
            elif campo in ('codigo', 'cod', 'empleado'):
                guardar_perfil_usuario(user_id, username, codigo_empleado=valor)
                await update.message.reply_text(f"Codigo de empleado actualizado: *{valor}*", parse_mode="Markdown")
            elif campo in ('nombre',):
                guardar_perfil_usuario(user_id, username, nombre=valor)
                await update.message.reply_text(f"Nombre actualizado: *{valor}*", parse_mode="Markdown")
            elif campo in ('apellido',):
                guardar_perfil_usuario(user_id, username, apellido=valor)
                await update.message.reply_text(f"Apellido actualizado: *{valor}*", parse_mode="Markdown")
            else:
                await update.message.reply_text(
                    "Uso:\n/perfil ceco 349065\n/perfil codigo 12345\n/perfil nombre Juan\n/perfil apellido Perez"
                )
            return

        # Sin args: mostrar perfil actual
        ceco = perfil.get('ceco', 'No configurado')
        cod = perfil.get('codigo_empleado', 'No configurado')
        nombre = perfil.get('nombre', 'No configurado')
        apellido = perfil.get('apellido', 'No configurado')
        await update.message.reply_text(
            f"Tu perfil:\n\n"
            f"CECO: *{ceco}*\n"
            f"Codigo de empleado: *{cod}*\n"
            f"Nombre: *{nombre}*\n"
            f"Apellido: *{apellido}*\n\n"
            f"Para cambiar:\n"
            f"/perfil ceco 349065\n"
            f"/perfil codigo 12345\n"
            f"/perfil nombre Juan\n"
            f"/perfil apellido Perez",
            parse_mode="Markdown"
        )

    async def msg_procesar_foto(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Procesa fotos de facturas. 1 foto → flujo inmediato. 2+ fotos seguidas → flujo batch."""
        if not await self._guard_acceso(update):
            return
        user_id = update.message.from_user.id
        message_id = update.message.message_id
        username = update.message.from_user.username or update.message.from_user.first_name or ""
        logger.info(f"📷 Foto recibida — msg_id={message_id} user={user_id}")

        # Descargar foto inmediatamente
        carpeta = carpeta_usuario(user_id, username)
        foto = update.message.photo[-1]
        archivo_path = carpeta / "facturas" / f"{message_id}.jpg"

        try:
            file = await context.bot.get_file(foto.file_id)
            await file.download_to_drive(str(archivo_path))
            logger.info(f"Foto descargada: {archivo_path}")
        except Exception as e:
            logger.error(f"Error descargando foto: {e}")
            await update.message.reply_text("❌ Error descargando la foto. Intenta de nuevo.")
            return

        # ── Detección de batch por timer de usuario ──
        # IMPORTANTE: el dict se crea ANTES del primer await para evitar
        # condición de carrera cuando varias fotos llegan simultáneamente.
        primera_foto = user_id not in _cola_fotos
        if primera_foto:
            _cola_fotos[user_id] = {
                'username': username,
                'chat_id': update.message.chat_id,
                'archivos': [],
                'progress_msg': None,  # se asigna tras el sendMessage
                'task': None,
            }

        _cola_fotos[user_id]['archivos'].append(archivo_path)
        n = len(_cola_fotos[user_id]['archivos'])

        if primera_foto:
            # Solo la primera foto crea el mensaje de progreso
            msg = await update.message.reply_text("📥 Recibiendo foto...")
            if user_id in _cola_fotos:  # puede haber sido procesado ya si llegó tarde
                _cola_fotos[user_id]['progress_msg'] = msg
        else:
            # Las siguientes actualizan el contador
            try:
                pm = _cola_fotos[user_id].get('progress_msg')
                if pm:
                    txt = f"📥 Recibiendo fotos... ({n})"
                    await pm.edit_text(txt)
            except Exception:
                pass

        # Cancelar timer anterior y reiniciar (ventana deslizante)
        task_anterior = _cola_fotos[user_id].get('task')
        if task_anterior and not task_anterior.done():
            task_anterior.cancel()
        task = asyncio.create_task(self._procesar_cola_fotos(user_id, context))
        _cola_fotos[user_id]['task'] = task

    @staticmethod
    def _generar_variantes_ocr(valor: str, campo: str) -> list:
        """Genera variantes 1-char-swap del valor para chars que el OCR
        confunde frecuentemente. Costo cero (puro Python), complementa al
        Sonnet bajo demanda cuando éste no sugirió el valor correcto.

        Las confusiones son específicas por tipo de campo: la serie es hex,
        el número es solo dígitos, el NIT tiene un formato mixto.
        """
        if not valor:
            return []
        v = valor.upper()
        if campo == 'serie':
            # Series FEL son hex (0-9 y A-F). Solo swaps a chars hex válidos.
            mapping = {
                '0': ['O', 'D', '8'], 'O': ['0'],
                '4': ['A'], 'A': ['4'],
                '8': ['B', '0'], 'B': ['8'],
                'F': ['E', 'P'], 'E': ['F'],
                'D': ['0'], 'C': ['6'], '6': ['C', 'G'],
                '1': ['I', '7'], '7': ['1'],
                '5': ['S'], '2': ['Z'],
            }
            valid = set('0123456789ABCDEF')
        elif campo == 'numero':
            # Solo dígitos.
            mapping = {
                '0': ['8', '6'], '8': ['0', '3'], '5': ['6', '8'],
                '6': ['5', '8'], '1': ['7'], '7': ['1'],
                '3': ['8'], '9': ['0'], '4': ['9'],
            }
            valid = set('0123456789')
        elif campo == 'nit':
            # NIT: digits + '-' + (digit o K). El char tras el guión es el
            # verificador, lo dejamos como va.
            mapping = {
                '0': ['8', '6'], '8': ['0', '3'],
                '5': ['6', '8'], '6': ['5', '8'],
                '1': ['7'], '7': ['1'], '3': ['8'], '9': ['0'],
            }
            valid = set('0123456789-K')
        else:
            return []

        variantes = []
        seen = {v}
        for i, c in enumerate(v):
            for nuevo in mapping.get(c, []):
                if nuevo not in valid:
                    continue
                var = v[:i] + nuevo + v[i+1:]
                if var in seen:
                    continue
                seen.add(var)
                variantes.append(var)
                if len(variantes) >= 6:
                    return variantes
        return variantes

    async def _pantalla_correccion(self, query, pend_id, texto, teclado, parse_mode=None):
        """Muestra una pantalla del flujo de corrección manteniendo la foto
        visible. Si el mensaje actual YA es la foto → edita su caption.
        Si es un mensaje de texto → adjunta la foto como mensaje nuevo y
        quita los botones del texto previo para no confundir al usuario.
        """
        info = pendientes.get(pend_id, {}) or {}
        archivo = info.get('archivo', '')
        foto_ok = bool(archivo and Path(archivo).exists())

        es_foto_msg = bool(query.message.photo or query.message.document)
        if es_foto_msg:
            try:
                await query.edit_message_caption(
                    caption=texto, reply_markup=teclado, parse_mode=parse_mode
                )
                return
            except Exception as e:
                logger.debug(f"edit_caption falló, reintentando: {e}")

        if foto_ok:
            # Quitar botones del mensaje viejo (texto) para evitar interacción doble
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            try:
                with open(archivo, 'rb') as f:
                    await query.message.reply_photo(
                        photo=f, caption=texto, reply_markup=teclado, parse_mode=parse_mode
                    )
                return
            except Exception as e:
                logger.warning(f"No se pudo enviar foto en correccion: {e}")

        # Fallback: sin foto disponible, edita o responde texto
        try:
            await query.edit_message_text(texto, reply_markup=teclado, parse_mode=parse_mode)
        except Exception:
            await query.message.reply_text(texto, reply_markup=teclado, parse_mode=parse_mode)

    async def _handler_corregir_campo(self, query, context, pend_id, resultado, campo):
        """Punto de entrada cuando el usuario tapea Corregir NIT/Serie/Numero.

        Flujo:
        1. Sonnet bajo demanda re-lee el campo → top 3 candidatos como botones.
        2. Si ninguna sirve → 'Ver variantes OCR' calcula swaps locales.
        3. Si nada sirve → 'Escribir manualmente' (flujo viejo).
        """
        nombres = {"nit": "NIT", "serie": "Serie", "numero": "Número"}
        nombre_campo = nombres.get(campo, campo)
        valor_prev = {
            "nit": resultado.nit or "",
            "serie": resultado.serie or "",
            "numero": resultado.numero_factura or "",
        }[campo]

        info = pendientes.get(pend_id, {}) or {}
        archivo = info.get('archivo', '')
        puede_sonnet = bool(archivo and Path(archivo).exists() and sistema.sonnet)

        if not puede_sonnet:
            # Sin imagen o sin Sonnet → ir directo a escritura manual
            await self._mostrar_correccion_manual(query, context, pend_id, campo, resultado)
            return

        await query.answer(f"🔍 Re-leyendo {nombre_campo}...")
        # Pantalla "cargando" con foto visible
        await self._pantalla_correccion(
            query, pend_id,
            f"🔍 Re-leyendo {nombre_campo} con Sonnet...\n"
            f"Valor actual: {valor_prev or '(vacío)'}\n\n"
            f"Espera unos segundos.",
            None,
        )

        # Llamar Sonnet en hilo separado para no bloquear el event loop
        try:
            candidatos = await asyncio.to_thread(
                sistema.sonnet.sugerir_campo, archivo, campo, valor_prev
            )
        except Exception as e:
            logger.error(f"Error Sonnet sugerir_campo: {e}")
            candidatos = []

        # Guardar candidatos para los callbacks de aplicación
        context.user_data[f'corr_sug_{pend_id}_{campo}'] = candidatos

        if candidatos:
            filas = []
            for i, c in enumerate(candidatos):
                filas.append([InlineKeyboardButton(
                    f"✅ {c}", callback_data=f"csug:{pend_id}:{campo}:{i}"
                )])
            filas.append([InlineKeyboardButton("↻ Ver variantes OCR", callback_data=f"cver:{pend_id}:{campo}")])
            filas.append([InlineKeyboardButton("✏️ Escribir manualmente", callback_data=f"cman:{pend_id}:{campo}")])
            filas.append([InlineKeyboardButton("⬅ Atrás", callback_data=f"corregir:{pend_id}")])
            texto = (
                f"🤖 Sonnet sugiere para *{nombre_campo}*:\n"
                f"Valor actual: `{valor_prev or '(vacío)'}`\n\n"
                f"Elige uno o ve otras opciones:"
            )
        else:
            filas = [
                [InlineKeyboardButton("↻ Ver variantes OCR", callback_data=f"cver:{pend_id}:{campo}")],
                [InlineKeyboardButton("✏️ Escribir manualmente", callback_data=f"cman:{pend_id}:{campo}")],
                [InlineKeyboardButton("⬅ Atrás", callback_data=f"corregir:{pend_id}")],
            ]
            texto = (
                f"⚠️ Sonnet no pudo sugerir un valor confiable para *{nombre_campo}*.\n"
                f"Valor actual: `{valor_prev or '(vacío)'}`\n\n"
                f"Probá las otras opciones:"
            )
        await self._pantalla_correccion(
            query, pend_id, texto, InlineKeyboardMarkup(filas), parse_mode="Markdown"
        )

    async def _handler_corregir_sub(self, query, context, accion, data):
        """Maneja callbacks del sub-flujo de corrección: csug/cvar/cver/cman."""
        # data trae formato '{pend_id}:{campo}[:{idx}]'
        partes = data.split(":")
        if len(partes) < 2:
            await query.answer("Datos invalidos.", show_alert=True)
            return
        pend_id = partes[0]
        campo = partes[1]
        idx = int(partes[2]) if len(partes) >= 3 and partes[2].isdigit() else -1

        info = pendientes.get(pend_id)
        if not info:
            await query.answer("Esta factura ya no esta pendiente.", show_alert=True)
            return
        resultado = info['resultado']

        if accion == "cman":
            # Fallback al flujo de escritura manual (set estado + pedir input)
            await self._mostrar_correccion_manual(query, context, pend_id, campo, resultado)
            return

        if accion == "cver":
            # Calcular variantes OCR del valor actual y mostrarlas como botones
            valor_prev = {
                "nit": resultado.nit or "",
                "serie": resultado.serie or "",
                "numero": resultado.numero_factura or "",
            }.get(campo, "")
            variantes = self._generar_variantes_ocr(valor_prev, campo)
            context.user_data[f'corr_var_{pend_id}_{campo}'] = variantes
            nombres = {"nit": "NIT", "serie": "Serie", "numero": "Número"}
            nombre_campo = nombres.get(campo, campo)
            if variantes:
                filas = []
                for i, v in enumerate(variantes):
                    filas.append([InlineKeyboardButton(
                        f"🔄 {v}", callback_data=f"cvar:{pend_id}:{campo}:{i}"
                    )])
                filas.append([InlineKeyboardButton("✏️ Escribir manualmente", callback_data=f"cman:{pend_id}:{campo}")])
                filas.append([InlineKeyboardButton("⬅ Atrás", callback_data=f"corregir_{campo}:{pend_id}")])
                texto = (
                    f"🔄 Variantes OCR para *{nombre_campo}*:\n"
                    f"Valor actual: `{valor_prev or '(vacío)'}`\n\n"
                    f"(1 carácter cambiado: A↔4, 0↔O, 8↔B, etc.)"
                )
            else:
                filas = [
                    [InlineKeyboardButton("✏️ Escribir manualmente", callback_data=f"cman:{pend_id}:{campo}")],
                    [InlineKeyboardButton("⬅ Atrás", callback_data=f"corregir_{campo}:{pend_id}")],
                ]
                texto = (
                    f"No hay variantes OCR razonables para `{valor_prev}`.\n"
                    "Escribe el valor manualmente."
                )
            await query.answer()
            await self._pantalla_correccion(
                query, pend_id, texto, InlineKeyboardMarkup(filas), parse_mode="Markdown"
            )
            return

        if accion in ("csug", "cvar"):
            # Aplicar la sugerencia/variante seleccionada
            key = f'corr_sug_{pend_id}_{campo}' if accion == "csug" else f'corr_var_{pend_id}_{campo}'
            candidatos = context.user_data.get(key, [])
            if idx < 0 or idx >= len(candidatos):
                await query.answer("Sugerencia no válida.", show_alert=True)
                return
            nuevo_valor = candidatos[idx]
            await self._aplicar_correccion_campo(query, context, pend_id, campo, nuevo_valor, resultado)
            return

    async def _aplicar_correccion_campo(self, query, context, pend_id, campo, nuevo_valor, resultado):
        """Aplica un nuevo valor al campo y muestra el menú con datos actualizados."""
        # Limpiar formato según campo
        if campo == 'nit':
            nit_limpio = sistema.parser.limpiar_nit(nuevo_valor.upper())
            if not nit_limpio:
                await query.answer("NIT inválido. Formato: 1234567-A", show_alert=True)
                return
            resultado.nit = nit_limpio
        elif campo == 'serie':
            resultado.serie = sistema.parser.reparar_serie_hex(nuevo_valor)
        elif campo == 'numero':
            resultado.numero_factura = nuevo_valor.strip()

        # Persistir en pendientes
        if pend_id in pendientes:
            pendientes[pend_id]['resultado'] = resultado
            _guardar_pendientes()

        # Limpiar sugerencias guardadas para este campo
        context.user_data.pop(f'corr_sug_{pend_id}_{campo}', None)
        context.user_data.pop(f'corr_var_{pend_id}_{campo}', None)

        await query.answer(f"✅ {campo.upper()} actualizado")
        # Re-render del menú con el valor nuevo (mismo que tras corregir manual)
        user_id = query.from_user.id
        filas_corr = [[
            InlineKeyboardButton("✅ Validar en SAT", callback_data=f"verificar_sat:{pend_id}"),
            InlineKeyboardButton("✏️ Corregir", callback_data=f"corregir:{pend_id}"),
            InlineKeyboardButton("❌ Rechazar", callback_data=f"rechazar:{pend_id}"),
        ]]
        mis_pend_corr = [pid for pid, info in pendientes.items() if info.get('user_id') == user_id]
        if len(mis_pend_corr) >= 2:
            filas_corr.append([InlineKeyboardButton(
                f"Validar TODAS ({len(mis_pend_corr)})",
                callback_data=f"verificar_todas:{pend_id}",
            )])
        teclado = InlineKeyboardMarkup(filas_corr)
        texto = (
            f"Datos actualizados:\n"
            f"NIT: {resultado.nit}\n"
            f"Serie: {resultado.serie}\n"
            f"Numero: {resultado.numero_factura}\n"
            f"Empresa: {resultado.razon_social or '(no detectada)'}"
        )
        await self._pantalla_correccion(query, pend_id, texto, teclado)

    async def _mostrar_correccion_manual(self, query, context, pend_id, campo, resultado):
        """Flujo de escritura manual del valor (comportamiento original)."""
        nombres = {"nit": "NIT (formato: 1234567-A)", "serie": "Serie", "numero": "Numero de factura"}
        await query.answer()
        texto_corr = (
            f"Datos actuales:\n"
            f"NIT: {resultado.nit}\n"
            f"Serie: {resultado.serie}\n"
            f"Numero: {resultado.numero_factura}\n\n"
            f"Enviame el nuevo valor para {nombres[campo]}:"
        )
        await self._pantalla_correccion(query, pend_id, texto_corr, None)
        context.user_data['estado'] = f'corrigiendo_{campo}'
        context.user_data['resultado_corrigiendo'] = resultado
        context.user_data['pend_id'] = pend_id

    @staticmethod
    def _validar_pre_sat(resultado, user_id: int, username: str):
        """Valida la factura ANTES de SAT/UI: detecta CF y duplicados.

        Devuelve (ok: bool, motivo_texto: str, dup_entry: dict|None).
        - ok=False: la factura debe rechazarse sin consumir captcha
        - dup_entry: si es duplicado, la entrada anterior (para mostrar info)
        """
        # 1. Consumidor Final: factura no nominativa, no sirve para reporte.
        if getattr(resultado, 'receptor_cf', False):
            return False, "consumidor_final", None
        # 2. Duplicado: ya existe en el registro.json del usuario.
        if not (resultado and resultado.nit and resultado.serie and resultado.numero_factura):
            return True, "", None
        try:
            carpeta = carpeta_usuario(user_id, username)
            registros = _cargar_registros_usuario(user_id, username)
            dup = _buscar_duplicado_en_registros(
                registros, carpeta, _clave_de_resultado(resultado), verificar_pdf=False
            )
            if dup:
                return False, "duplicado", dup
        except Exception as e:
            logger.warning(f"Pre-validacion no pudo leer registro {user_id}: {e}")
        return True, "", None

    @staticmethod
    async def _ocr_factura(archivo, user_id: int, username: str):
        """Ejecuta procesar_factura en un hilo aparte para NO congelar el
        event loop.

        procesar_factura es síncrona y hace I/O de red bloqueante (OCR con
        Claude vía el cliente síncrono de Anthropic, que tarda 3-10s) además
        de preprocesamiento de imagen con OpenCV (CPU). Si se llamara directo
        en un handler async, el bot quedaría congelado para TODOS los usuarios
        durante esos segundos. asyncio.to_thread lo corre fuera del loop.
        """
        registro_path = carpeta_usuario(user_id, username) / "registro.json"
        return await asyncio.to_thread(
            sistema.procesar_factura, str(archivo),
            user_id=str(user_id), registro_path=registro_path,
        )

    async def _procesar_cola_fotos(self, user_id: int, context):
        """Espera BATCH_TIMEOUT tras la última foto y luego procesa. 1 foto→flujo rápido, 2+→batch."""
        await asyncio.sleep(BATCH_TIMEOUT)

        cola = _cola_fotos.pop(user_id, None)
        if not cola:
            return

        archivos = cola['archivos']
        username = cola['username']
        chat_id = cola['chat_id']
        progress_msg = cola['progress_msg']
        total = len(archivos)
        bot = context.bot

        # ── 1 sola foto: flujo inmediato (sin numeración) ──
        if total == 1:
            archivo_path = archivos[0]
            try:
                await progress_msg.edit_text("⏳ Analizando factura...")
            except Exception:
                pass
            try:
                resultado = await self._ocr_factura(archivo_path, user_id, username)
            except Exception as e:
                logger.error(f"Error procesando factura: {e}")
                await bot.send_message(chat_id=chat_id, text=f"❌ Error procesando: {e}")
                archivo_path.unlink(missing_ok=True)
                return

            # Pre-validación: CF y duplicados → rechazar antes de SAT
            pre_ok, motivo, dup_entry = self._validar_pre_sat(resultado, user_id, username)
            if not pre_ok:
                if motivo == "consumidor_final":
                    await bot.send_message(
                        chat_id=chat_id,
                        text=(
                            "🚫 *Factura emitida a CONSUMIDOR FINAL*\n\n"
                            "Esta factura no es nominativa, así que no sirve para el reporte de viáticos. "
                            "Pídela emitida con tu NIT y reenvíala."
                        ),
                        parse_mode="Markdown",
                    )
                elif motivo == "duplicado" and dup_entry:
                    fecha_prev = dup_entry.get('fecha_emision') or dup_entry.get('fecha_proceso', '?')
                    monto_prev = dup_entry.get('monto', '?')
                    await bot.send_message(
                        chat_id=chat_id,
                        text=(
                            "♻️ *Factura ya registrada anteriormente*\n\n"
                            f"NIT: `{resultado.nit}`\n"
                            f"Serie: `{resultado.serie}`\n"
                            f"Número: `{resultado.numero_factura}`\n"
                            f"Fecha: {fecha_prev}\n"
                            f"Monto: {monto_prev}\n\n"
                            "No se procesará para evitar duplicarla en el reporte."
                        ),
                        parse_mode="Markdown",
                    )
                archivo_path.unlink(missing_ok=True)
                return

            pend_id = uuid.uuid4().hex[:8]
            pendientes[pend_id] = {
                'resultado': resultado,
                'user_id': user_id,
                'archivo': str(archivo_path),
                'username': username,
                'timestamp': datetime.now()
            }
            _guardar_pendientes()

            cal = getattr(resultado, 'calidad_imagen', None) or {}
            if cal.get('muy_borrosa') or cal.get('muy_oscura') or cal.get('sobre_expuesta'):
                avisos = []
                if cal.get('muy_borrosa'): avisos.append("📸 La foto está borrosa")
                if cal.get('muy_oscura'): avisos.append("🌑 Muy oscura")
                if cal.get('sobre_expuesta'): avisos.append("☀️ Sobre-expuesta")
                await bot.send_message(
                    chat_id=chat_id,
                    text="⚠️ " + " · ".join(avisos) + "\nSi la extracción falla, intenta otra foto con mejor luz/enfoque."
                )

            # Construir teclado igual que _mostrar_extraccion
            mis_ids = [pid for pid, info in pendientes.items() if info.get('user_id') == user_id]
            fila = [
                InlineKeyboardButton("✅ Validar en SAT", callback_data=f"verificar_sat:{pend_id}"),
                InlineKeyboardButton("✏️ Corregir", callback_data=f"corregir:{pend_id}"),
                InlineKeyboardButton("❌ Rechazar", callback_data=f"rechazar:{pend_id}"),
            ]
            filas = [fila]
            if len(mis_ids) >= 2:
                filas.append([InlineKeyboardButton(f"Validar TODAS ({len(mis_ids)})", callback_data=f"verificar_todas:{pend_id}")])

            if resultado.metodo_extraccion == "manual_pending":
                teclado = InlineKeyboardMarkup([
                    [InlineKeyboardButton("NIT", callback_data=f"corregir_nit:{pend_id}"),
                     InlineKeyboardButton("Serie", callback_data=f"corregir_serie:{pend_id}"),
                     InlineKeyboardButton("Numero", callback_data=f"corregir_numero:{pend_id}")],
                    [InlineKeyboardButton("✅ Validar en SAT", callback_data=f"verificar_sat:{pend_id}")],
                    [InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar_factura:{pend_id}")],
                ])
                texto = (
                    f"Datos parciales encontrados:\n"
                    f"NIT: {resultado.nit or '(no detectado)'}\n"
                    f"Serie: {resultado.serie or '(no detectada)'}\n"
                    f"Numero: {resultado.numero_factura or '(no detectado)'}\n"
                    f"Razon social: {resultado.razon_social or '(no detectada)'}\n\n"
                    f"Corrige lo que falte y luego valida en SAT:"
                )
            elif resultado.metodo_extraccion == "catalogo":
                teclado = InlineKeyboardMarkup(filas)
                texto = (
                    f"Encontrado en catalogo\n\n"
                    f"Empresa: {resultado.razon_social}\n"
                    f"NIT: {resultado.nit}\n"
                    f"Numero: {resultado.numero_factura}\n"
                    f"Serie: {resultado.serie}\n\n"
                    f"Es correcto?"
                )
            else:
                teclado = InlineKeyboardMarkup(filas)
                texto = (
                    f"Extraccion exitosa (OCR directo)\n\n"
                    f"NIT: {resultado.nit}\n"
                    f"Numero: {resultado.numero_factura}\n"
                    f"Serie: {resultado.serie}\n\n"
                    f"Razon social: {resultado.razon_social or '(no detectada)'}\n"
                    f"Confianza: {resultado.confianza:.0%}"
                )
            await bot.send_message(chat_id=chat_id, text=texto, reply_markup=teclado)
            return

        # ── 2+ fotos: flujo batch con progreso ──

        # Procesar cada foto secuencialmente mostrando progreso
        # tupla: (resultado | None, pend_id | None, error_msg | None, dup_entry | None)
        # error_msg con prefijo "__SKIP__:" indica que la factura se rechazó
        # antes de SAT (consumidor_final o duplicado).
        resultados = []
        for i, archivo_path in enumerate(archivos, 1):
            try:
                await progress_msg.edit_text(f"⏳ Procesando {i} de {total}...")
            except Exception:
                pass

            try:
                resultado = await self._ocr_factura(archivo_path, user_id, username)
                # Pre-validación: CF y duplicados → no añadir a pendientes
                pre_ok, motivo, dup_entry = self._validar_pre_sat(resultado, user_id, username)
                if not pre_ok:
                    archivo_path.unlink(missing_ok=True)
                    resultados.append((None, None, f"__SKIP__:{motivo}", dup_entry))
                    continue
                pend_id = uuid.uuid4().hex[:8]
                pendientes[pend_id] = {
                    'resultado': resultado,
                    'user_id': user_id,
                    'archivo': str(archivo_path),
                    'username': username,
                    'timestamp': datetime.now()
                }
                _guardar_pendientes()
                resultados.append((resultado, pend_id, None, None))
            except Exception as e:
                logger.error(f"Batch: error procesando {archivo_path}: {e}")
                resultados.append((None, None, str(e), None))

        # Actualizar mensaje de progreso con resumen
        exitosas = sum(1 for r, _, _, _ in resultados if r is not None)
        try:
            await progress_msg.edit_text(
                f"✅ {exitosas} de {total} facturas leídas. Mostrando resultados..."
            )
        except Exception:
            pass

        # Enviar resultados uno por uno
        for idx, (resultado, pend_id, error, dup_entry) in enumerate(resultados, 1):
            if resultado is None:
                # Caso skip (consumidor_final o duplicado) — mensaje específico
                if error and error.startswith("__SKIP__:"):
                    motivo = error.split(":", 1)[1]
                    if motivo == "consumidor_final":
                        await bot.send_message(
                            chat_id=chat_id,
                            text=(
                                f"🚫 Factura {idx}/{total}: emitida a *CONSUMIDOR FINAL*. "
                                "No sirve para reporte — pídela con tu NIT."
                            ),
                            parse_mode="Markdown",
                        )
                    elif motivo == "duplicado":
                        fecha_prev = (dup_entry or {}).get('fecha_emision') or (dup_entry or {}).get('fecha_proceso', '?')
                        monto_prev = (dup_entry or {}).get('monto', '?')
                        nit_prev = (dup_entry or {}).get('nit', '?')
                        await bot.send_message(
                            chat_id=chat_id,
                            text=(
                                f"♻️ Factura {idx}/{total}: ya registrada anteriormente "
                                f"(NIT {nit_prev}, {fecha_prev}, {monto_prev}). Saltada."
                            ),
                        )
                    continue
                await bot.send_message(
                    chat_id=chat_id,
                    text=f"❌ Factura {idx}/{total}: no se pudo procesar.\n{error}"
                )
                continue

            # Avisos de calidad de imagen
            cal = getattr(resultado, 'calidad_imagen', None) or {}
            avisos = []
            if cal.get('muy_borrosa'): avisos.append("📸 Borrosa")
            if cal.get('muy_oscura'): avisos.append("🌑 Muy oscura")
            if cal.get('sobre_expuesta'): avisos.append("☀️ Sobre-expuesta")
            if avisos:
                await bot.send_message(
                    chat_id=chat_id,
                    text=f"⚠️ Factura {idx}/{total}: " + " · ".join(avisos)
                )

            # Construir teclado
            mis_ids = [pid for pid, info in pendientes.items() if info.get('user_id') == user_id]
            if resultado.metodo_extraccion == "manual_pending":
                teclado = InlineKeyboardMarkup([
                    [
                        InlineKeyboardButton("NIT", callback_data=f"corregir_nit:{pend_id}"),
                        InlineKeyboardButton("Serie", callback_data=f"corregir_serie:{pend_id}"),
                        InlineKeyboardButton("Numero", callback_data=f"corregir_numero:{pend_id}"),
                    ],
                    [InlineKeyboardButton("✅ Validar en SAT", callback_data=f"verificar_sat:{pend_id}")],
                    [InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar_factura:{pend_id}")],
                ])
                texto = (
                    f"📋 {idx}/{total} — Datos parciales\n\n"
                    f"NIT: {resultado.nit or '(no detectado)'}\n"
                    f"Serie: {resultado.serie or '(no detectada)'}\n"
                    f"Numero: {resultado.numero_factura or '(no detectado)'}\n"
                    f"Razon social: {resultado.razon_social or '(no detectada)'}\n\n"
                    f"Corrige lo que falte y luego valida en SAT:"
                )
            else:
                fila = [
                    InlineKeyboardButton("✅ Validar en SAT", callback_data=f"verificar_sat:{pend_id}"),
                    InlineKeyboardButton("✏️ Corregir", callback_data=f"corregir:{pend_id}"),
                    InlineKeyboardButton("❌ Rechazar", callback_data=f"rechazar:{pend_id}"),
                ]
                filas = [fila]
                if len(mis_ids) >= 2:
                    filas.append([InlineKeyboardButton(
                        f"Validar TODAS ({len(mis_ids)})", callback_data=f"verificar_todas:{pend_id}"
                    )])
                teclado = InlineKeyboardMarkup(filas)

                if resultado.metodo_extraccion == "catalogo":
                    texto = (
                        f"📋 {idx}/{total} — Catálogo\n\n"
                        f"Empresa: {resultado.razon_social}\n"
                        f"NIT: {resultado.nit}\n"
                        f"Numero: {resultado.numero_factura}\n"
                        f"Serie: {resultado.serie}\n\n"
                        f"¿Es correcto?"
                    )
                else:
                    texto = (
                        f"📋 {idx}/{total}\n\n"
                        f"NIT: {resultado.nit}\n"
                        f"Numero: {resultado.numero_factura}\n"
                        f"Serie: {resultado.serie}\n\n"
                        f"Razon social: {resultado.razon_social or '(no detectada)'}\n"
                        f"Confianza: {resultado.confianza:.0%}"
                    )

            await bot.send_message(chat_id=chat_id, text=texto, reply_markup=teclado)

        # Botón para cancelar todo el batch de un golpe
        total_pendientes = sum(1 for info in pendientes.values() if info.get('user_id') == user_id)
        if total_pendientes >= 2:
            teclado_cancelar = InlineKeyboardMarkup([[
                InlineKeyboardButton(f"🗑️ Cancelar TODAS ({total_pendientes})", callback_data="cancelar_todas_confirm"),
            ]])
            await bot.send_message(
                chat_id=chat_id,
                text=f"👆 {exitosas} facturas listas para validar.\nSi alguna está mal, puedes cancelar todo:",
                reply_markup=teclado_cancelar
            )

    async def msg_procesar_pdf(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Procesa documentos PDF de facturas — convierte a imagen y usa el mismo flujo de OCR"""
        if not await self._guard_acceso(update):
            return
        user_id = update.message.from_user.id
        message_id = update.message.message_id
        username = update.message.from_user.username or update.message.from_user.first_name or ""

        carpeta = carpeta_usuario(user_id, username)
        doc = update.message.document
        archivo_pdf = carpeta / "facturas" / f"{message_id}.pdf"
        archivo_img = carpeta / "facturas" / f"{message_id}.jpg"

        try:
            file = await context.bot.get_file(doc.file_id)
            await file.download_to_drive(str(archivo_pdf))
            logger.info(f"PDF descargado: {archivo_pdf}")
        except Exception as e:
            logger.error(f"Error descargando PDF: {e}")
            await update.message.reply_text("❌ Error descargando el PDF. Intenta de nuevo.")
            return

        # Convertir PDF a imagen usando PyMuPDF (no requiere instalar nada externo)
        await update.message.reply_text("⏳ Procesando PDF de factura...")
        try:
            import fitz  # PyMuPDF
            doc = fitz.open(str(archivo_pdf))
            pagina = doc[0]  # Primera página
            # Renderizar a 200 DPI
            mat = fitz.Matrix(200 / 72, 200 / 72)
            pix = pagina.get_pixmap(matrix=mat)
            pix.save(str(archivo_img))
            doc.close()
            logger.info(f"PDF convertido a imagen: {archivo_img}")
        except ImportError:
            await update.message.reply_text(
                "❌ Falta instalar PyMuPDF.\n"
                "Ejecuta: pip install PyMuPDF"
            )
            return
        except Exception as e:
            logger.error(f"Error convirtiendo PDF: {e}")
            await update.message.reply_text(f"❌ Error procesando PDF: {e}")
            return

        # Procesar la imagen con el mismo flujo de OCR
        try:
            resultado = await self._ocr_factura(archivo_img, user_id, username)
        except Exception as e:
            logger.error(f"Error procesando factura PDF: {e}")
            await update.message.reply_text(f"❌ Error procesando: {e}")
            return

        # Override con el TEXTO del PDF (exacto). Para un PDF de FEL esto manda
        # sobre el OCR de imagen y sobre el QR: el texto embebido no tiene
        # errores. Evita rechazos del SAT por serie/numero mal leidos.
        fel = _extraer_fel_de_pdf(archivo_pdf)
        if fel:
            logger.info(f"📄 FEL desde texto del PDF (exacto): {fel}")
            if fel.get('serie') and fel['serie'] != resultado.serie:
                logger.info(f"  🔁 Serie del PDF pisa: {resultado.serie!r} → {fel['serie']!r}")
                resultado.serie = fel['serie']
            if fel.get('numero') and fel['numero'] != resultado.numero_factura:
                logger.info(f"  🔁 Numero del PDF pisa: {resultado.numero_factura!r} → {fel['numero']!r}")
                resultado.numero_factura = fel['numero']

        # Pre-validación: CF y duplicados → no añadir a pendientes
        pre_ok, motivo, dup_entry = self._validar_pre_sat(resultado, user_id, username)
        if not pre_ok:
            if motivo == "consumidor_final":
                await update.message.reply_text(
                    "🚫 *Factura emitida a CONSUMIDOR FINAL*\n\n"
                    "No es nominativa, así que no sirve para el reporte de viáticos. "
                    "Pídela emitida con tu NIT y reenvíala.",
                    parse_mode="Markdown",
                )
            elif motivo == "duplicado" and dup_entry:
                fecha_prev = dup_entry.get('fecha_emision') or dup_entry.get('fecha_proceso', '?')
                monto_prev = dup_entry.get('monto', '?')
                await update.message.reply_text(
                    "♻️ *Factura ya registrada anteriormente*\n\n"
                    f"NIT: `{resultado.nit}`\n"
                    f"Serie: `{resultado.serie}`\n"
                    f"Número: `{resultado.numero_factura}`\n"
                    f"Fecha: {fecha_prev}\n"
                    f"Monto: {monto_prev}\n\n"
                    "No se procesará para evitar duplicarla.",
                    parse_mode="Markdown",
                )
            return

        pend_id = uuid.uuid4().hex[:8]
        pendientes[pend_id] = {
            'resultado': resultado,
            'user_id': user_id,
            'archivo': str(archivo_img),
            'username': username,
            'timestamp': datetime.now()
        }
        _guardar_pendientes()

        # NOTA: NO se revisa duplicado por OCR/historial. Duplicados SOLO se
        # detectan tras SAT validar y descargar el PDF (ver buscar_duplicado
        # post-verificacion).

        # Cubre todos los metodos exitosos (incluye los nuevos +historial)
        if resultado.metodo_extraccion == "manual_pending":
            await self._pedir_confirmacion_manual(update, resultado, archivo_img, pend_id)
        else:
            await self._mostrar_extraccion(update, resultado, pend_id)

    async def msg_texto(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Procesa mensajes de texto"""
        if not await self._guard_acceso(update):
            return
        user_id = update.message.from_user.id
        texto = update.message.text.strip()

        # Esperando odómetro tras guardar factura de gasolina
        pend_odo = context.user_data.get('esperando_odometro')
        if pend_odo:
            username = update.message.from_user.username or update.message.from_user.first_name or ""
            if texto.lower() in ('saltar', 'luego', 'despues', 'después', 'cancelar'):
                context.user_data.pop('esperando_odometro', None)
                await update.message.reply_text(
                    "Ok. Puedes agregarlo después con `/odometro` o `/odometro <km>`.",
                    parse_mode="Markdown"
                )
                return
            valor_txt = texto.replace(',', '').replace('.', '').replace(' ', '')
            if valor_txt.isdigit():
                valor = int(valor_txt)
                anterior = pend_odo.get('anterior')
                aviso = validar_odometro_nuevo(valor, anterior)
                ok = guardar_odometro(
                    user_id, username,
                    pend_odo['nit'], pend_odo['serie'], pend_odo['numero'], valor
                )
                context.user_data.pop('esperando_odometro', None)
                if ok:
                    msg = f"✅ Odómetro guardado: *{valor:,} km*"
                    if anterior and anterior.get('odometro') is not None:
                        km = valor - anterior['odometro']
                        if km >= 0:
                            msg += f"\n🛣️ Km recorridos desde la carga anterior: *{km:,} km*"
                    if aviso:
                        msg += f"\n\n{aviso}"
                    await update.message.reply_text(msg, parse_mode="Markdown")
                else:
                    await update.message.reply_text("No se pudo actualizar la factura (puede haber sido eliminada).")
                return
            # Texto que no es número ni comando: avisar pero no bloquear otras funciones
            context.user_data.pop('esperando_odometro', None)
            await update.message.reply_text(
                "No entendí el odómetro. Usa `/odometro` más tarde para agregarlo.",
                parse_mode="Markdown"
            )
            # Continuar con el flujo normal en caso de que sea otra acción

        # Si está esperando CECO para reporte
        if context.user_data.get('estado') == 'esperando_ceco_reporte':
            if texto.lower() in ('cancelar', 'no', 'salir'):
                context.user_data.clear()
                await update.message.reply_text("Reporte cancelado.")
                return
            username = update.message.from_user.username or update.message.from_user.first_name or ""
            perfil = leer_perfil_usuario(user_id, username)
            # Permitir "ok" para reusar CECO guardado
            if texto.lower() == 'ok' and perfil.get('ceco'):
                ceco = perfil['ceco']
            elif texto.strip().isdigit():
                ceco = texto.strip()
            else:
                await update.message.reply_text("El CECO debe ser un numero.\nIntenta de nuevo o escribe 'cancelar':")
                return
            context.user_data['ceco'] = ceco
            # Pedir código de empleado (o usar el guardado)
            cod_guardado = perfil.get('codigo_empleado', '')
            context.user_data['estado'] = 'esperando_codigo_empleado_reporte'
            msg = "Ingresa tu codigo de empleado:\n(Ej: 12345)"
            if cod_guardado:
                msg += f"\n\nUltimo codigo usado: *{cod_guardado}* — escribe 'ok' para usarlo"
            await update.message.reply_text(msg, parse_mode="Markdown")
            return

        if context.user_data.get('estado') == 'esperando_codigo_empleado_reporte':
            if texto.lower() in ('cancelar', 'salir'):
                context.user_data.clear()
                await update.message.reply_text("Reporte cancelado.")
                return
            username = update.message.from_user.username or update.message.from_user.first_name or ""
            perfil = leer_perfil_usuario(user_id, username)
            if texto.lower() == 'ok' and perfil.get('codigo_empleado'):
                codigo_empleado = perfil['codigo_empleado']
            elif texto.strip():
                codigo_empleado = texto.strip()
            else:
                await update.message.reply_text("Ingresa un codigo de empleado valido o escribe 'cancelar':")
                return
            # Guardar en perfil para próxima vez
            guardar_perfil_usuario(user_id, username, ceco=context.user_data.get('ceco', ''), codigo_empleado=codigo_empleado)
            context.user_data['codigo_empleado'] = codigo_empleado
            context.user_data['estado'] = 'esperando_fechas_reporte'
            await self._pedir_fechas_reporte(update.message, context.user_data.get('ceco', ''))
            return

        if context.user_data.get('estado') == 'esperando_fechas_reporte':
            if texto.lower() in ('cancelar', 'salir'):
                context.user_data.clear()
                await update.message.reply_text("Reporte cancelado.")
                return
            ceco = context.user_data.get('ceco', '')
            username = update.message.from_user.username or update.message.from_user.first_name or ""
            fecha_inicio = ""
            fecha_fin = ""
            if texto.lower() not in ('todas', 'todo', 'all', ''):
                partes = texto.split()
                def _fecha_valida(f):
                    try:
                        datetime.strptime(f.replace('.', '/').replace('-', '/'), "%d/%m/%Y")
                        return True
                    except ValueError:
                        return False
                if len(partes) >= 2:
                    if not _fecha_valida(partes[0]) or not _fecha_valida(partes[1]):
                        await update.message.reply_text(
                            "Formato de fecha invalido. Usa DD/MM/YYYY\nEj: 07/04/2026 10/04/2026\nO escribe 'todas' para ver todas las facturas."
                        )
                        return
                    fecha_inicio = partes[0]
                    fecha_fin = partes[1]
                elif len(partes) == 1:
                    if not _fecha_valida(partes[0]):
                        await update.message.reply_text(
                            "Formato de fecha invalido. Usa DD/MM/YYYY\nEj: 07/04/2026"
                        )
                        return
                    fecha_inicio = partes[0]
                    fecha_fin = partes[0]
            codigo_empleado = context.user_data.get('codigo_empleado', '')
            formato = context.user_data.get('formato_reporte', 'completo')
            context.user_data.clear()
            await self._generar_y_enviar_reporte(
                update.message, user_id, username, ceco, fecha_inicio, fecha_fin,
                codigo_empleado=codigo_empleado, formato=formato,
            )
            return

        if context.user_data.get('estado', '').startswith('corrigiendo_'):
            campo = context.user_data['estado'].replace('corrigiendo_', '')
            resultado = context.user_data.get('resultado_corrigiendo')
            # Si el bot se reinicio en medio de la correccion, context.user_data
            # se vacia. Pero `pendientes` SI se persiste en disco, asi que
            # podemos recuperar el resultado por pend_id sin perder la sesion.
            if not resultado:
                pend_id = context.user_data.get('pend_id', '')
                if pend_id and pend_id in pendientes:
                    resultado = pendientes[pend_id].get('resultado')
                    context.user_data['resultado_corrigiendo'] = resultado
                    logger.info(f"Resultado recuperado desde pendientes (pend_id={pend_id[:8]}) tras reinicio")
            if not resultado:
                await update.message.reply_text("❌ Sesión expirada. Envía otra foto.")
                context.user_data.clear()
                return

            if campo == 'nit':
                nit_limpio = sistema.parser.limpiar_nit(texto.upper())
                if not nit_limpio:
                    await update.message.reply_text("❌ NIT inválido. Formato: 1234567-A\nIntenta de nuevo:")
                    return
                resultado.nit = nit_limpio
            elif campo == 'serie':
                resultado.serie = sistema.parser.reparar_serie_hex(texto)
            elif campo == 'numero':
                resultado.numero_factura = texto.strip()

            # Actualizar en pendientes usando pend_id
            pend_id = context.user_data.get('pend_id', '')
            if pend_id and pend_id in pendientes:
                pendientes[pend_id]['resultado'] = resultado

            # Limpiar estado del flujo de correccion, PERO preservar el estado
            # del asistente de revision (revision_cola/total/resueltas). Si lo
            # borraramos, tras validar una factura corregida el hook veria
            # cola=0 y no avanzaria a las demas fallidas.
            p = pend_id
            _rev = {k: context.user_data.get(k) for k in
                    ('revision_cola', 'revision_total', 'revision_resueltas')
                    if context.user_data.get(k) is not None}
            context.user_data.clear()
            context.user_data.update(_rev)
            filas_corr = [[
                InlineKeyboardButton("✅ Validar en SAT", callback_data=f"verificar_sat:{p}"),
                InlineKeyboardButton("✏️ Corregir", callback_data=f"corregir:{p}"),
                InlineKeyboardButton("❌ Rechazar", callback_data=f"rechazar:{p}"),
            ]]
            # Si quedan 2+ pendientes (otras fallidas del lote), agregar
            # "Validar TODAS" para no obligar al usuario a hacer /pendientes
            # de nuevo solo porque corrigio una.
            mis_pend_corr = [pid for pid, info in pendientes.items()
                             if info.get('user_id') == user_id]
            if len(mis_pend_corr) >= 2:
                filas_corr.append([InlineKeyboardButton(
                    f"Validar TODAS ({len(mis_pend_corr)})",
                    callback_data=f"verificar_todas:{p}",
                )])
            teclado = InlineKeyboardMarkup(filas_corr)
            texto_actualizado = (
                f"Datos actualizados:\n"
                f"NIT: {resultado.nit}\n"
                f"Serie: {resultado.serie}\n"
                f"Numero: {resultado.numero_factura}\n"
                f"Empresa: {resultado.razon_social or '(no detectada)'}"
            )
            # Enviar foto + caption para que el usuario siga viendo la imagen
            # al validar/corregir de nuevo, sin tener que buscarla arriba.
            info_pend = pendientes.get(pend_id, {}) or {}
            archivo_pend = info_pend.get('archivo', '')
            if archivo_pend and Path(archivo_pend).exists():
                try:
                    with open(archivo_pend, 'rb') as f:
                        await update.message.reply_photo(
                            photo=f, caption=texto_actualizado, reply_markup=teclado
                        )
                except Exception as e:
                    logger.warning(f"No se pudo enviar foto tras corregir: {e}")
                    await update.message.reply_text(texto_actualizado, reply_markup=teclado)
            else:
                await update.message.reply_text(texto_actualizado, reply_markup=teclado)

        elif context.user_data.get('estado') == 'esperando_captcha_sat':
            # Usuario envió el texto del captcha — cancelar timeout
            timeout_task = context.user_data.pop('captcha_timeout_task', None)
            if timeout_task and not timeout_task.done():
                timeout_task.cancel()

            resultado = context.user_data.get('resultado_sat')
            if not resultado:
                await update.message.reply_text("❌ Sesión expirada. Envía otra foto.")
                context.user_data.clear()
                return

            await update.message.reply_text("⏳ Enviando captcha al SAT...")
            verificador = await obtener_verificador(str(user_id))
            res = await verificador.verificar_con_captcha(texto, str(user_id))

            # Si hay otro captcha (incorrecto), volver a pedir
            if res.esperando_captcha and res.captcha_screenshot:
                with open(res.captcha_screenshot, 'rb') as f:
                    await update.message.reply_photo(photo=f, caption="🔒 Captcha incorrecto. Intenta de nuevo:")
                return

            uname = update.message.from_user.username or update.message.from_user.first_name or ""
            resultado_sat = context.user_data.get('resultado_sat')
            captcha_pend_id = context.user_data.get('pend_id', '')
            # Actualizar razón social con la oficial del SAT
            if res.razon_social_sat and resultado_sat and resultado_sat.nit:
                resultado_sat.razon_social = res.razon_social_sat
                sistema.catalogo.reconciliar_con_sat(res.razon_social_sat, resultado_sat.nit)
                logger.info(f"Catálogo reconciliado con SAT: {resultado_sat.nit} = {res.razon_social_sat}")
            # Detectar si estamos en medio de un lote (hay cola pendiente) o en
            # el asistente de revision. En ambos casos NO mostramos UI manual
            # cuando SAT rechaza — la factura se queda pendiente y el flujo
            # correspondiente (lote o asistente) la vuelve a presentar.
            en_lote = bool(context.user_data.get('verificar_cola')
                           or context.user_data.get('revision_cola'))

            if res.valida is False:
                # Intentar escalacion a Sonnet primero — el flujo principal de
                # _intentar_verificacion_sat lo hace, pero aqui no pasaba.
                info = pendientes.get(captcha_pend_id, {})
                archivo = info.get('archivo', '')
                ya_escalado = context.user_data.get(f'sonnet_escalado_{captcha_pend_id}', False)
                sonnet_ok = False
                if (not ya_escalado and archivo and Path(archivo).exists()
                        and sistema.sonnet and resultado_sat):
                    context.user_data[f'sonnet_escalado_{captcha_pend_id}'] = True
                    logger.info(
                        f"SAT (captcha) rechazó {resultado_sat.nit}/"
                        f"{resultado_sat.serie}/{resultado_sat.numero_factura} "
                        f"— escalando a Sonnet"
                    )
                    try:
                        await update.message.reply_text(
                            "🔁 SAT no encontró la factura. Reintentando con lectura más precisa..."
                        )
                        nuevo = sistema.reextraer_con_sonnet(archivo, resultado_previo=resultado_sat)
                        if nuevo and (
                            nuevo.nit != resultado_sat.nit
                            or nuevo.serie != resultado_sat.serie
                            or nuevo.numero_factura != resultado_sat.numero_factura
                        ):
                            logger.info(
                                f"Sonnet (post-captcha) corrigió: {nuevo.nit}/"
                                f"{nuevo.serie}/{nuevo.numero_factura}"
                            )
                            if captcha_pend_id in pendientes:
                                pendientes[captcha_pend_id]['resultado'] = nuevo
                            # Reintentar SAT con los datos corregidos.
                            # Usamos _verificar_en_sat con modo_batch=en_lote
                            # asi participa correctamente del flujo del lote.
                            ok_sonnet = await self._verificar_en_sat(
                                update.message, context, nuevo, user_id,
                                pend_id=captcha_pend_id, modo_batch=en_lote,
                            )
                            sonnet_ok = (ok_sonnet is True)
                            if not sonnet_ok and en_lote:
                                # Acumular en fallidas_lote del context
                                fl = context.user_data.get('fallidas_lote', [])
                                fl.append((captcha_pend_id, nuevo))
                                context.user_data['fallidas_lote'] = fl
                    except Exception as ex:
                        logger.warning(f"Escalacion Sonnet post-captcha falló: {ex}")

                # Si Sonnet no resolvió: UI manual solo fuera de lote.
                if not sonnet_ok and not en_lote:
                    teclado = InlineKeyboardMarkup([
                        [
                            InlineKeyboardButton("NIT", callback_data=f"corregir_nit:{captcha_pend_id}"),
                            InlineKeyboardButton("Serie", callback_data=f"corregir_serie:{captcha_pend_id}"),
                            InlineKeyboardButton("Numero", callback_data=f"corregir_numero:{captcha_pend_id}"),
                        ],
                        [InlineKeyboardButton("Reintentar en SAT", callback_data=f"verificar_sat:{captcha_pend_id}")],
                        [InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar_factura:{captcha_pend_id}")],
                    ])
                    texto_error = (
                        f"Factura NO encontrada en SAT\n\n"
                        f"NIT: {resultado_sat.nit}\n"
                        f"Serie: {resultado_sat.serie}\n"
                        f"Numero: {resultado_sat.numero_factura}\n\n"
                        f"Corrige el dato incorrecto o reintenta:"
                    )
                    if archivo and Path(archivo).exists():
                        with open(archivo, 'rb') as f:
                            await update.message.reply_photo(photo=f, caption=texto_error, reply_markup=teclado)
                    else:
                        await update.message.reply_text(texto_error, reply_markup=teclado)
                elif not sonnet_ok and en_lote:
                    # En lote: acumular como fallida para el resumen final.
                    fl = context.user_data.get('fallidas_lote', [])
                    fl.append((captcha_pend_id, resultado_sat))
                    context.user_data['fallidas_lote'] = fl
                    logger.info(
                        f"[batch] Factura tras-captcha rechazada por SAT "
                        f"({captcha_pend_id[:8]}) — diferida al resumen final"
                    )
            else:
                if resultado_sat:
                    # Verificar duplicado ANTES de copiar el PDF
                    duplicada, reg_existente = buscar_duplicado(user_id, resultado_sat, uname)
                    if duplicada:
                        await self._enviar_duplicado(update.message, resultado_sat, reg_existente, res, user_id, uname)
                        pendientes.pop(captcha_pend_id, None)
                        _guardar_pendientes()
                    else:
                        pdf_usuario = _copiar_pdf_a_usuario(res, user_id, uname)
                        _, entrada_guardada = guardar_registro(user_id, resultado_sat, pdf_path=pdf_usuario, username=uname, res_sat=res)
                        await self._enviar_resultado_sat(update.message, res, user_id=user_id, username=uname, pend_id=captcha_pend_id)
                        alerta = _verificar_comida_duplicada(user_id, uname, entrada_guardada)
                        if alerta:
                            await update.message.reply_text(alerta, parse_mode="Markdown")
                        # Clasificar si aún no tiene tipo (por razón social) y pedir odómetro si es gasolina
                        _clasificar_por_razon_social(entrada_guardada)
                        await self._prompt_odometro(update.message, context, user_id, uname, entrada_guardada)
                else:
                    await self._enviar_resultado_sat(update.message, res, user_id=user_id, username=uname, pend_id=captcha_pend_id)
                pendientes.pop(captcha_pend_id, None)
                _guardar_pendientes()

            # Continuar con la cola de verificación si existe
            cola = context.user_data.get('verificar_cola', [])
            fallidas_lote = context.user_data.get('fallidas_lote', [])
            context.user_data.pop('verificar_cola', None)
            context.user_data.pop('fallidas_lote', None)
            # CRITICO: limpiar el 'estado' = 'esperando_captcha_sat'. Si no, el
            # loop de continuacion lo lee como "hay captcha pendiente" y rompe
            # despues de la primera factura, atascando el lote.
            context.user_data.pop('estado', None)
            context.user_data.pop('resultado_sat', None)
            context.user_data.pop('pend_id', None)

            if cola:
                logger.info(f"[batch-cola] Continuando con {len(cola)} facturas tras captcha")
                await update.message.reply_text(f"Continuando con {len(cola)} facturas restantes...")
                completado = False
                for i, (pid, info) in enumerate(cola, 1):
                    if pid not in pendientes:
                        logger.info(f"[batch-cola] Factura {i}/{len(cola)} ya no esta en pendientes (pid={pid[:8]})")
                        continue
                    res_factura = info['resultado']
                    logger.info(
                        f"[batch-cola] >>> Iniciando {i}/{len(cola)} "
                        f"(pid={pid[:8]}, NIT={res_factura.nit}, "
                        f"serie={res_factura.serie}, num={res_factura.numero_factura})"
                    )
                    try:
                        # Pausa anti-bot SAT
                        if i > 1:
                            await asyncio.sleep(4.0)
                        await update.message.reply_text(
                            f"[{i}/{len(cola)}] Verificando...\n"
                            f"NIT: {res_factura.nit} | Num: {res_factura.numero_factura}"
                        )
                        ok = await self._verificar_en_sat(
                            update.message, context, res_factura, user_id,
                            pend_id=pid, modo_batch=True,
                        )
                        logger.info(
                            f"[batch-cola] <<< Termino {i}/{len(cola)}: ok={ok} "
                            f"estado={context.user_data.get('estado', '<none>')}"
                        )
                        if ok is False:
                            fallidas_lote.append((pid, res_factura))

                        if context.user_data.get('estado') == 'esperando_captcha_sat':
                            restantes = len(cola) - i
                            if restantes > 0:
                                context.user_data['verificar_cola'] = [
                                    (pid2, info2) for pid2, info2 in cola[i:]
                                ]
                                context.user_data['fallidas_lote'] = fallidas_lote
                                await update.message.reply_text(
                                    f"Otro captcha. Resuelve y las {restantes} restantes continuaran."
                                )
                            break
                    except Exception as e:
                        logger.exception(f"[batch-cola] Excepcion en {i}/{len(cola)}: {e}")
                        try:
                            await update.message.reply_text(
                                f"⚠️ Error en factura {i}: {e}\nContinuando..."
                            )
                        except Exception:
                            pass
                        fallidas_lote.append((pid, res_factura))
                else:
                    completado = True
                logger.info(
                    f"[batch-cola] Cola completada={completado} "
                    f"Fallidas={len(fallidas_lote)} Estado={context.user_data.get('estado', '<none>')}"
                )
                if completado:
                    total_lote = len(cola) + len(fallidas_lote)
                    await self._mostrar_resumen_lote_sat(update.message, total_lote, fallidas_lote, context=context)
            elif context.user_data.get('revision_cola'):
                # El captcha se resolvio dentro del asistente de revision (no
                # era un lote). Avanzar a la siguiente factura con problema.
                # Si la actual sigue en pendientes (SAT la rechazo de nuevo),
                # _mostrar_siguiente_revision la vuelve a mostrar.
                await self._mostrar_siguiente_revision(update.message, context, user_id)

        else:
            # Mensaje genérico
            await update.message.reply_text(
                "👋 Envíame una foto de una factura para analizarla.\n"
                "Usa /help para más comandos.",
                parse_mode="Markdown"
            )
    
    @staticmethod
    async def _edit_query_message(query, texto: str, reply_markup=None):
        """Edita el mensaje de un callback: usa caption si es foto, text si es texto."""
        kwargs = {"reply_markup": reply_markup} if reply_markup else {}
        if query.message.photo or query.message.document:
            await query.edit_message_caption(caption=texto, **kwargs)
        else:
            await query.edit_message_text(texto, **kwargs)

    async def btn_confirmación(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Procesa clics en botones de confirmación"""
        query = update.callback_query
        user_id = query.from_user.id

        # Extraer acción y pend_id del callback_data (formato: "accion:pend_id")
        parts = query.data.split(":", 1)
        accion = parts[0]
        pend_id = parts[1] if len(parts) > 1 else ""

        # Router por categoría: cada grupo maneja un set disjunto de acciones
        # y hace su propio return. Lo que no cae en ninguno usa el flujo de
        # factura pendiente (requiere info/resultado) más abajo.
        if accion in {"admin_aprobar", "admin_bloquear", "admin_desbloquear", "dup_cancel", "dup_pdf", "reintentar_fallida", "resolver_fallida"}:
            return await self._router_cb_sistema(query, context, accion, pend_id, user_id)
        if accion in {"clasificar", "odometro_cancelar", "odometro_cerrar", "odometro_edit", "tipo_cerrar", "tipo_pick", "tipo_set"}:
            return await self._router_cb_tipo_odometro(query, context, accion, pend_id, user_id)
        if accion in {"eliminar_factura", "recordatorio_reporte", "recordatorio_sin_facturas", "reporte_cancelar", "reporte_confirmar_edicion", "reporte_editar", "reporte_excluir", "reporte_formato", "reporte_generar", "reporte_preview"}:
            return await self._router_cb_reportes(query, context, accion, pend_id, user_id)
        if accion in {"cancelar_todas_abort", "cancelar_todas_confirm", "cman", "csug", "cvar", "cver", "revisar_pendientes"} or (accion == "verificar_todas" and not pend_id):
            return await self._router_cb_lote(query, context, accion, pend_id, user_id)

        info = pendientes.get(pend_id)
        if not info:
            await query.answer(
                "Esta factura ya no esta pendiente (procesada, cancelada o "
                "expirada). Envía la foto otra vez si necesitas reintentar.",
                show_alert=True,
            )
            return
        await self._router_callbacks_pendiente(
            query, context, accion, pend_id, user_id, info
        )
        # Asistente de revision: si esta activo y la factura se resolvio o
        # cancelo, avanzar a la siguiente automaticamente.
        await self._avanzar_revision_si_aplica(
            query.message, context, user_id, pend_id, accion
        )


    async def _router_cb_sistema(self, query, context, accion, pend_id, user_id):
        """Admin, duplicados-PDF y facturas fallidas."""
        if accion == "dup_cancel":
            await query.answer()
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            return

        if accion == "dup_pdf":
            # pend_id aquí es "nit|serie|numero"
            try:
                nit, serie, numero = pend_id.split("|", 2)
            except ValueError:
                await query.answer("Datos invalidos.", show_alert=True)
                return
            username = query.from_user.username or query.from_user.first_name or ""
            carpeta = carpeta_usuario(user_id, username)
            registro_path = carpeta / "registro.json"
            if not registro_path.exists():
                await query.answer("No hay registro.", show_alert=True)
                return
            try:
                with open(registro_path, 'r', encoding='utf-8') as f:
                    registros = json.load(f)
            except Exception:
                await query.answer("Error leyendo registro.", show_alert=True)
                return
            reg = next(
                (r for r in registros
                 if r.get('nit') == nit and r.get('serie') == serie and r.get('numero') == numero),
                None,
            )
            if not reg:
                await query.answer("Factura no encontrada en registro.", show_alert=True)
                return
            pdf_guardado = reg.get('pdf_sat', '')
            if not pdf_guardado:
                await query.answer("PDF no disponible.", show_alert=True)
                return
            pdf_path = carpeta / "pdf_sat" / Path(pdf_guardado).name
            if not pdf_path.exists():
                await query.answer("PDF ya no esta en disco.", show_alert=True)
                return
            await query.answer()
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            with open(pdf_path, 'rb') as f:
                await query.message.reply_document(document=f, filename=pdf_path.name)
            return

        # Callbacks de admin (aprobar/bloquear desde botones inline)
        if accion in ("admin_aprobar", "admin_bloquear", "admin_desbloquear"):
            if not _es_admin(user_id):
                await query.answer("⛔ Sin permisos.", show_alert=True)
                return
            uid = _safe_int(pend_id)
            if uid is None:
                await query.answer("ID de usuario invalido.", show_alert=True)
                logger.warning(f"Callback admin con pend_id no numerico: {pend_id!r}")
                return
            if accion == "admin_aprobar":
                _aprobar_usuario(uid)
                await query.answer("✅ Usuario aprobado.")
                try:
                    await self.app.bot.send_message(chat_id=uid, text="✅ Tu acceso ha sido aprobado. Ya puedes usar el bot.")
                except Exception:
                    pass
                try:
                    await query.edit_message_reply_markup(reply_markup=None)
                    await query.edit_message_text(query.message.text + "\n\n✅ *Aprobado*", parse_mode="Markdown")
                except Exception:
                    pass
            elif accion == "admin_bloquear":
                _bloquear_usuario(uid)
                await query.answer("🚫 Usuario bloqueado.")
                try:
                    await query.edit_message_reply_markup(reply_markup=None)
                    await query.edit_message_text(query.message.text + "\n\n🚫 *Bloqueado*", parse_mode="Markdown")
                except Exception:
                    pass
            elif accion == "admin_desbloquear":
                _desbloquear_usuario(uid)
                await query.answer("🔓 Usuario desbloqueado.")
                try:
                    await self.app.bot.send_message(chat_id=uid, text="✅ Tu acceso ha sido restaurado.")
                except Exception:
                    pass
                try:
                    await query.edit_message_reply_markup(reply_markup=None)
                    await query.edit_message_text(query.message.text + "\n\n🔓 *Desbloqueado*", parse_mode="Markdown")
                except Exception:
                    pass
            return

        # Reintentar factura fallida — re-encolar en pendientes y verificar en SAT
        if accion == "reintentar_fallida":
            fallidas_path = DATOS_DIR / "facturas_fallidas.json"
            if not fallidas_path.exists():
                await query.answer("No hay facturas fallidas.", show_alert=True)
                return
            try:
                with open(fallidas_path, 'r', encoding='utf-8') as f:
                    fallidas = json.load(f)
            except Exception:
                await query.answer("Error leyendo fallidas.", show_alert=True)
                return
            # pend_id aqui es el pend_id original de la factura fallida
            falla = next((f for f in fallidas if f.get('pend_id') == pend_id and not f.get('resuelta')), None)
            if not falla:
                await query.answer("Factura no encontrada o ya resuelta.", show_alert=True)
                return
            resultado = ExtraccionFactura(
                nit=falla.get('nit', ''),
                serie=falla.get('serie', ''),
                numero_factura=falla.get('numero', ''),
                razon_social=None,
                metodo_extraccion='ocr_directo',
                confianza=0.9,
            )
            username = falla.get('username', '')
            uid = falla.get('user_id', user_id)
            nuevo_id = uuid.uuid4().hex[:8]
            pendientes[nuevo_id] = {
                'resultado': resultado,
                'user_id': uid,
                'archivo': '',
                'username': username,
                'timestamp': datetime.now(),
            }
            _guardar_pendientes()
            await query.answer("Reintentando en SAT...")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await query.message.reply_text(
                f"Reintentando verificacion en SAT...\n"
                f"NIT: {resultado.nit} | Serie: {resultado.serie} | Num: {resultado.numero_factura}"
            )
            await self._verificar_en_sat(query.message, context, resultado, uid, pend_id=nuevo_id)
            return

        # Marcar factura fallida como resuelta manualmente
        if accion == "resolver_fallida":
            fallidas_path = DATOS_DIR / "facturas_fallidas.json"
            try:
                # Lock: evita que dos admins clickeando "resolver" en paralelo
                # pisoteen el archivo (RMW concurrente).
                with _lock_archivo(fallidas_path):
                    fallidas = []
                    if fallidas_path.exists():
                        with open(fallidas_path, 'r', encoding='utf-8') as f:
                            fallidas = json.load(f)
                    for falla in fallidas:
                        if falla.get('pend_id') == pend_id:
                            falla['resuelta'] = True
                            break
                    _guardar_json_atomico(fallidas_path, fallidas)
                await query.answer("Marcada como resuelta.")
                try:
                    await query.edit_message_reply_markup(reply_markup=None)
                    await query.edit_message_text(query.message.text + "\n\n✅ Marcada como resuelta.")
                except Exception:
                    pass
            except Exception as e:
                await query.answer(f"Error: {e}", show_alert=True)
            return

        # ── Callback: cerrar listado de /tipo o /odometro ──────────────────

    async def _router_cb_tipo_odometro(self, query, context, accion, pend_id, user_id):
        """Clasificación de tipo de gasto y odómetro."""
        if accion in ("tipo_cerrar", "odometro_cerrar"):
            await query.answer("Menú cerrado.")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            sid = pend_id
            if accion == "tipo_cerrar":
                tipo_sessions.pop(sid, None)
            else:
                odometro_sessions.pop(sid, None)
            return

        # ── Callback: elegir factura a reclasificar (mostrar selector) ─────
        if accion == "tipo_pick":
            partes = pend_id.split("|", 1)
            sid = partes[0]
            idx_txt = partes[1] if len(partes) > 1 else ""
            sess = tipo_sessions.get(sid)
            if not sess or sess.get('user_id') != user_id:
                await query.answer("Sesión expirada. Usa /tipo de nuevo.", show_alert=True)
                return
            try:
                idx = int(idx_txt)
                fac = sess['facturas'][idx]
            except (ValueError, IndexError):
                await query.answer("Factura no encontrada.", show_alert=True)
                return
            await query.answer()
            est = (fac.get('razon_social') or fac.get('establecimiento') or '?')
            tipo_actual = fac.get('tipo_comida') or 'sin tipo'
            teclado = InlineKeyboardMarkup([
                [
                    InlineKeyboardButton("🌅 Desayuno",  callback_data=f"tipo_set:{sid}|{idx}|Desayuno"),
                    InlineKeyboardButton("☀️ Almuerzo",  callback_data=f"tipo_set:{sid}|{idx}|Almuerzo"),
                    InlineKeyboardButton("🌙 Cena",       callback_data=f"tipo_set:{sid}|{idx}|Cena"),
                ],
                [
                    InlineKeyboardButton("⛽ Gasolina",   callback_data=f"tipo_set:{sid}|{idx}|Gasolina"),
                    InlineKeyboardButton("🏨 Hospedaje",  callback_data=f"tipo_set:{sid}|{idx}|Hospedaje"),
                    InlineKeyboardButton("🅿️ Parqueo",    callback_data=f"tipo_set:{sid}|{idx}|Parqueo"),
                ],
                [
                    InlineKeyboardButton("❌ Cancelar",   callback_data=f"tipo_set:{sid}|{idx}|"),
                ],
            ])
            await query.message.reply_text(
                f"🏷 *Cambiar tipo*\n\n"
                f"📅 {fac.get('fecha_emision', '?')}\n"
                f"🏪 {est[:40]}\n"
                f"💰 {fac.get('monto', 'Q?')}\n"
                f"Tipo actual: *{tipo_actual}*\n\n"
                f"¿A qué tipo cambiar?",
                parse_mode="Markdown",
                reply_markup=teclado
            )
            return

        # ── Callback: aplicar reclasificacion de tipo ──────────────────────
        if accion == "tipo_set":
            partes = pend_id.split("|", 2)
            if len(partes) < 3:
                await query.answer("Datos incompletos.", show_alert=True)
                return
            sid, idx_txt, nuevo_tipo = partes[0], partes[1], partes[2]
            sess = tipo_sessions.get(sid)
            if not sess or sess.get('user_id') != user_id:
                await query.answer("Sesión expirada. Usa /tipo de nuevo.", show_alert=True)
                return
            try:
                idx = int(idx_txt)
                fac = sess['facturas'][idx]
            except (ValueError, IndexError):
                await query.answer("Factura no encontrada.", show_alert=True)
                return
            if not nuevo_tipo:
                await query.answer("Cancelado.")
                try:
                    await query.edit_message_reply_markup(reply_markup=None)
                except Exception:
                    pass
                return
            # Persistir en registro.json con lock y escritura atómica
            carpeta = carpeta_usuario(user_id, sess['username'])
            registro_path = carpeta / "registro.json"
            try:
                with _lock_archivo(registro_path):
                    with open(registro_path, 'r', encoding='utf-8') as f:
                        registros = json.load(f)
                    clave_fac = _clave_de_registro(fac)
                    actualizadas = 0
                    for r in registros:
                        clave_r = _clave_de_registro(r)
                        if clave_r == clave_fac:
                            r['tipo_comida'] = nuevo_tipo
                            actualizadas += 1
                    _guardar_json_atomico(registro_path, registros)
                fac['tipo_comida'] = nuevo_tipo  # reflejar en la sesión
            except Exception as e:
                logger.error(f"Error guardando reclasificacion: {e}")
                await query.answer(f"Error: {e}", show_alert=True)
                return
            await query.answer(f"✅ Reclasificado como {nuevo_tipo}")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
                await query.edit_message_text(
                    query.message.text + f"\n\n✅ *Cambiado a {nuevo_tipo}*",
                    parse_mode="Markdown"
                )
            except Exception:
                pass
            return

        # ── Callback: cancelar prompt de odómetro ──────────────────────────
        if accion == "odometro_cancelar":
            context.user_data.pop('esperando_odometro', None)
            await query.answer("Saltado.")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await query.message.reply_text(
                "Ok. Puedes agregarlo después con `/odometro`.",
                parse_mode="Markdown"
            )
            return

        # ── Callback: editar odómetro desde listado /odometro ──────────────
        if accion == "odometro_edit":
            partes = pend_id.split("|", 1)
            sid = partes[0]
            idx_txt = partes[1] if len(partes) > 1 else ""
            sess = odometro_sessions.get(sid)
            if not sess or sess.get('user_id') != user_id:
                await query.answer("Sesión expirada. Usa /odometro de nuevo.", show_alert=True)
                return
            try:
                idx = int(idx_txt)
                fac = sess['facturas'][idx]
            except (ValueError, IndexError):
                await query.answer("Factura no encontrada.", show_alert=True)
                return
            await query.answer()
            context.user_data['esperando_odometro'] = {
                'clave': _clave_de_registro(fac),
                'nit': fac.get('nit'),
                'serie': fac.get('serie'),
                'numero': fac.get('numero'),
                'fecha': fac.get('fecha_emision', ''),
                'anterior': obtener_odometro_anterior(
                    user_id, sess['username'], fac.get('fecha_emision', ''),
                    clave_excluir=_clave_de_registro(fac)
                ),
            }
            actual = fac.get('odometro')
            extra = f"\nOdómetro actual: *{actual:,} km*" if actual is not None else ""
            teclado = InlineKeyboardMarkup([[
                InlineKeyboardButton("❌ Cancelar", callback_data="odometro_cancelar:")
            ]])
            await query.message.reply_text(
                f"✏️ Editar odómetro de la factura del *{fac.get('fecha_emision', '?')}*.{extra}\n\n"
                f"Envíame el nuevo kilometraje (solo el número).",
                parse_mode="Markdown", reply_markup=teclado
            )
            return

        # ── Callback: clasificar factura sin tipo ──────────────────────────
        if accion == "clasificar":
            partes = pend_id.split("|", 1)
            sid = partes[0]
            tipo_elegido = partes[1] if len(partes) > 1 else ""
            session_cl = clasificar_sessions.get(sid)
            if not session_cl:
                await query.answer("Sesión expirada.", show_alert=True)
                return
            idx = session_cl['idx']
            pendientes_cl = session_cl['pendientes']
            if idx < len(pendientes_cl) and tipo_elegido:
                fac = pendientes_cl[idx]
                # Actualizar tipo_comida en memoria y en registro.json
                fac['tipo_comida'] = tipo_elegido
                carpeta = carpeta_usuario(session_cl['user_id'], session_cl['username'])
                registro_path = carpeta / "registro.json"
                try:
                    with _lock_archivo(registro_path):
                        with open(registro_path, 'r', encoding='utf-8') as f:
                            registros = json.load(f)
                        clave_fac = _clave_de_registro(fac)
                        for r in registros:
                            clave_r = _clave_de_registro(r)
                            if clave_r == clave_fac:
                                r['tipo_comida'] = tipo_elegido
                        # Escritura atomica para evitar corrupcion si crashea
                        _guardar_json_atomico(registro_path, registros)
                except (json.JSONDecodeError, OSError) as e:
                    logger.warning(f"No se pudo guardar clasificación: {e}")
            session_cl['idx'] += 1
            await query.answer("✅ Guardado")
            try:
                await query.message.delete()
            except Exception:
                pass
            # Si el usuario clasificó como Gasolina y aún no tiene odómetro, pedirlo
            if tipo_elegido == 'Gasolina' and fac.get('odometro') is None:
                await self._prompt_odometro(
                    query.message, context,
                    session_cl['user_id'], session_cl['username'],
                    fac,
                )
                # El loop de clasificación continuará cuando el usuario responda o salte;
                # pero para simplicidad seguimos mostrando la siguiente igual (la sesión de
                # odómetro vive en user_data y no bloquea).
            await self._mostrar_clasificacion_pendiente(query.message, sid)
            return

        # Callbacks de reporte (no usan pendientes)

    async def _router_cb_reportes(self, query, context, accion, pend_id, user_id):
        """Generación/edición de reportes, recordatorios, eliminar."""
        if accion == "reporte_formato":
            # pend_id = "completo" | "simple"
            formato = pend_id if pend_id in ("completo", "simple") else "completo"
            context.user_data['formato_reporte'] = formato
            etiqueta = "simple (solo TXT sin política)" if formato == "simple" else "completo (Excel + TXT)"
            await query.answer(f"Formato: {etiqueta}")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            username = query.from_user.username or query.from_user.first_name or ""
            await self._continuar_reporte_tras_formato(query.message, user_id, username, context)
            return

        if accion == "reporte_preview":
            # Mostrar preview editable antes de generar el reporte
            partes_rango = pend_id.split("|")
            ceco = partes_rango[0] if len(partes_rango) > 0 else ""
            fi = partes_rango[1] if len(partes_rango) > 1 else ""
            ff = partes_rango[2] if len(partes_rango) > 2 else ""
            username = query.from_user.username or query.from_user.first_name or ""
            await query.answer("Preparando preview...")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            codigo_empleado = context.user_data.get('codigo_empleado', '')
            formato = context.user_data.get('formato_reporte', 'completo')
            context.user_data.clear()
            await self._mostrar_preview_reporte(
                query.message, user_id, username, ceco, fi, ff,
                codigo_empleado=codigo_empleado, formato=formato,
            )
            return

        if accion == "reporte_generar":
            # pend_id es el session_id del reporte
            session = reporte_sessions.get(pend_id)
            if not session or session.get('user_id') != user_id:
                await query.answer("Sesion expirada. Usa /reporte de nuevo.", show_alert=True)
                return
            await query.answer("Generando reporte...")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            excluidas = session.get('excluidas', set())
            facturas_final = [f for f in session['facturas'] if id(f) not in excluidas]
            await self._generar_y_enviar_reporte(
                query.message, user_id, session['username'],
                session['ceco'], session['fi'], session['ff'],
                facturas_override=facturas_final,
                codigo_empleado=session.get('codigo_empleado', ''),
                formato=session.get('formato', 'completo'),
            )
            reporte_sessions.pop(pend_id, None)
            return

        if accion == "reporte_editar":
            session = reporte_sessions.get(pend_id)
            if not session or session.get('user_id') != user_id:
                await query.answer("Sesion expirada.", show_alert=True)
                return
            await query.answer("Elige las facturas a quitar:")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await self._mostrar_editor_facturas(query.message, pend_id, session)
            return

        if accion == "reporte_excluir":
            # pend_id = "session_id|fac_index"
            partes = pend_id.split("|", 1)
            sid = partes[0]
            try:
                fac_idx = int(partes[1]) if len(partes) > 1 else -1
            except ValueError:
                await query.answer("Error en indice de factura.", show_alert=True)
                return
            session = reporte_sessions.get(sid)
            if not session or session.get('user_id') != user_id:
                await query.answer("Sesion expirada.", show_alert=True)
                return
            if fac_idx < 0 or fac_idx >= len(session['fac_ids']):
                await query.answer("Factura no encontrada.", show_alert=True)
                return
            excluidas = session.setdefault('excluidas', set())
            fac_id = session['fac_ids'][fac_idx]
            if fac_id in excluidas:
                excluidas.discard(fac_id)
                await query.answer("Factura incluida de nuevo ✅")
            else:
                excluidas.add(fac_id)
                await query.answer("Factura quitada del reporte ❌")
            await self._mostrar_editor_facturas(query.message, sid, session)
            return

        if accion == "reporte_confirmar_edicion":
            session = reporte_sessions.get(pend_id)
            if not session or session.get('user_id') != user_id:
                await query.answer("Sesion expirada.", show_alert=True)
                return
            excluidas = session.get('excluidas', set())
            facturas_final = [f for i, f in enumerate(session['facturas']) if session['fac_ids'][i] not in excluidas]
            if not facturas_final:
                await query.answer("No puedes quitar todas las facturas.", show_alert=True)
                return
            await query.answer("Generando reporte editado...")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await self._generar_y_enviar_reporte(
                query.message, user_id, session['username'],
                session['ceco'], session['fi'], session['ff'],
                facturas_override=facturas_final,
                codigo_empleado=session.get('codigo_empleado', ''),
                formato=session.get('formato', 'completo'),
            )
            reporte_sessions.pop(pend_id, None)
            return

        if accion == "reporte_cancelar":
            reporte_sessions.pop(pend_id, None)
            await query.answer("Reporte cancelado")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await query.message.reply_text("❌ Reporte cancelado. Usa /reporte para generar uno nuevo.")
            return

        if accion == "eliminar_factura":
            # pend_id = "session_id|fac_index"
            partes = pend_id.split("|", 1)
            sid = partes[0]
            try:
                fac_idx = int(partes[1]) if len(partes) > 1 else -1
            except ValueError:
                await query.answer("Error en indice de factura.", show_alert=True)
                return
            session = eliminar_sessions.get(sid)
            if not session or session.get('user_id') != user_id:
                await query.answer("Sesion expirada.", show_alert=True)
                return
            if fac_idx < 0 or fac_idx >= len(session['facturas_display']):
                await query.answer("Factura no encontrada.", show_alert=True)
                return
            fac = session['facturas_display'][fac_idx]
            username = session['username']
            carpeta = carpeta_usuario(user_id, username)
            registro_path = carpeta / "registro.json"
            try:
                with _lock_archivo(registro_path):
                    with open(registro_path, 'r', encoding='utf-8') as f:
                        registros = json.load(f)
                    clave = _clave_de_registro(fac)
                    registros = [r for r in registros if _clave_de_registro(r) != clave]
                    _guardar_json_atomico(registro_path, registros)
                await query.answer("✅ Factura eliminada del registro.")
                try:
                    await query.edit_message_reply_markup(reply_markup=None)
                    empresa = fac.get('razon_social') or fac.get('establecimiento') or fac.get('nit', '?')
                    await query.edit_message_text(
                        query.message.text + f"\n\n❌ Eliminada: {empresa} | {fac.get('fecha_emision','?')} | {fac.get('monto','?')}"
                    )
                except Exception:
                    pass
                eliminar_sessions.pop(sid, None)
                logger.info(f"Factura eliminada por user {user_id}: NIT={fac.get('nit')} Serie={fac.get('serie')} Num={fac.get('numero')}")
            except Exception as e:
                await query.answer(f"Error eliminando: {e}", show_alert=True)
            return

        if accion == "recordatorio_reporte":
            # pend_id = "ceco" o vacío
            username = query.from_user.username or query.from_user.first_name or ""
            ceco = pend_id if pend_id else ""
            if not ceco:
                context.user_data['estado'] = 'esperando_ceco_reporte'
                await query.answer()
                await query.message.reply_text("Ingresa tu numero de CECO para el reporte:\n(Ej: 349065)")
                return
            await query.answer("Preparando reporte de la semana...")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            hoy = datetime.now()
            lunes = hoy - timedelta(days=hoy.weekday())
            fi = lunes.strftime("%d/%m/%Y")
            ff = hoy.strftime("%d/%m/%Y")
            await self._mostrar_preview_reporte(query.message, user_id, username, ceco, fi, ff)
            return

        if accion == "recordatorio_sin_facturas":
            await query.answer("Ok, entendido.")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
                await query.edit_message_text(query.message.text + "\n\n✅ Confirmado — sin facturas esta semana.")
            except Exception:
                pass
            return

        # ── Cancelar TODAS las pendientes del usuario ──────────────────────

    async def _router_cb_lote(self, query, context, accion, pend_id, user_id):
        """Cancelar todas, validar todas (sin pend), corrección Sonnet/variantes."""
        if accion == "cancelar_todas_abort":
            await query.answer("Operación cancelada.")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            return

        if accion == "revisar_pendientes":
            # Arranca el asistente de revision 1-a-1 sobre las facturas que ya
            # estan en pendientes (las que fallaron), SIN re-procesarlas por OCR.
            mis_ids = [pid for pid, info in pendientes.items() if info.get('user_id') == user_id]
            if not mis_ids:
                await query.answer("No tenes facturas pendientes.", show_alert=True)
                return
            await query.answer(f"Revisando {len(mis_ids)} facturas una por una...")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            context.user_data['revision_cola'] = list(mis_ids)
            context.user_data['revision_total'] = len(mis_ids)
            context.user_data['revision_resueltas'] = 0
            await self._mostrar_siguiente_revision(query.message, context, user_id)
            return

        if accion == "cancelar_todas_confirm":
            mis_ids = [pid for pid, info in list(pendientes.items()) if info.get('user_id') == user_id]
            for pid in mis_ids:
                pendientes.pop(pid, None)
            _guardar_pendientes()
            context.user_data.clear()
            await query.answer(f"🗑️ {len(mis_ids)} facturas canceladas.")
            try:
                await query.edit_message_text(f"🗑️ {len(mis_ids)} facturas pendientes canceladas y descartadas.")
            except Exception:
                pass
            return

        # `verificar_todas` opera sobre todas las pendientes del usuario, no
        # sobre un pend_id especifico. Saltamos directamente a su handler
        # mas abajo en este mismo bloque, sin pasar por el check de pend_id.
        if accion == "verificar_todas" and not pend_id:
            return await self._verificar_todas_handler(query, context, user_id)

        # Callbacks del flujo de corrección con Sonnet/variantes:
        # csug:{pend_id}:{campo}:{idx} → aplicar sugerencia de Sonnet
        # cvar:{pend_id}:{campo}:{idx} → aplicar variante OCR
        # cver:{pend_id}:{campo}      → calcular y mostrar variantes OCR
        # cman:{pend_id}:{campo}      → fallback a escritura manual
        if accion in ("csug", "cvar", "cver", "cman"):
            return await self._handler_corregir_sub(query, context, accion, pend_id)



    async def _router_callbacks_pendiente(self, query, context, accion, pend_id, user_id, info):
        """Acciones que operan sobre UNA factura pendiente (requieren info/resultado).
        Se llama desde btn_confirmación tras validar que pend_id sigue vivo."""
        resultado = info['resultado']
        username = info.get('username', '')

        if accion == "confirmar":
            duplicada, reg_existente = buscar_duplicado(user_id, resultado, username)
            if duplicada:
                await query.answer("Factura duplicada!", show_alert=True)
                await query.edit_message_text(
                    f"⚠️ Factura duplicada\n\n"
                    f"NIT: {resultado.nit}\n"
                    f"Serie: {resultado.serie}\n"
                    f"Numero: {resultado.numero_factura}\n\n"
                    f"Esta factura ya fue registrada."
                )
                # Reenviar PDF si existe en carpeta del usuario
                await self._enviar_duplicado(query.message, resultado, reg_existente, None, user_id, username)
                pendientes.pop(pend_id, None)
                return
            await query.answer("Datos confirmados!")
            await query.edit_message_text(
                f"Extraccion completada\n\n"
                f"NIT: {resultado.nit}\n"
                f"Numero: {resultado.numero_factura}\n"
                f"Serie: {resultado.serie}\n\n"
                f"Metodo: {resultado.metodo_extraccion}\n"
                f"Confianza: {resultado.confianza:.0%}"
            )
            pendientes.pop(pend_id, None)

        elif accion == "rechazar":
            await query.answer("Descartado")
            await query.edit_message_text("Datos rechazados. Envia otra foto.")
            pendientes.pop(pend_id, None)
            _guardar_pendientes()

        elif accion == "corregir":
            await query.answer("Modo correccion")
            teclado = InlineKeyboardMarkup([
                [
                    InlineKeyboardButton("NIT", callback_data=f"corregir_nit:{pend_id}"),
                    InlineKeyboardButton("Serie", callback_data=f"corregir_serie:{pend_id}"),
                    InlineKeyboardButton("Numero", callback_data=f"corregir_numero:{pend_id}"),
                ],
                [InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar_factura:{pend_id}")],
            ])
            texto_corr = (
                f"Datos actuales:\n"
                f"NIT: {resultado.nit}\n"
                f"Serie: {resultado.serie}\n"
                f"Numero: {resultado.numero_factura}\n\n"
                f"Que campo quieres corregir?"
            )
            # Si hay 2+ pendientes (resumen de lote o /pendientes con varias
            # fallidas), enviar mensaje NUEVO con la foto para no borrar la
            # lista con los botones de las otras fallidas. Si es una sola,
            # mostramos foto + caption con la UI de corrección.
            otras_pend = [pid for pid, info in pendientes.items()
                          if info.get('user_id') == user_id and pid != pend_id]
            archivo = info.get('archivo', '')
            foto_ok = bool(archivo and Path(archivo).exists())
            if otras_pend:
                # Mensaje nuevo (preserva el resumen del lote)
                if foto_ok:
                    with open(archivo, 'rb') as f:
                        await query.message.reply_photo(
                            photo=f, caption=texto_corr, reply_markup=teclado
                        )
                else:
                    await query.message.reply_text(texto_corr, reply_markup=teclado)
            else:
                await self._pantalla_correccion(query, pend_id, texto_corr, teclado)

        elif accion in ("corregir_nit", "corregir_serie", "corregir_numero"):
            campo = accion.replace("corregir_", "")
            await self._handler_corregir_campo(query, context, pend_id, resultado, campo)

        elif accion == "reintentar_ahora":
            await query.answer("Reintentando ahora...")
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await query.message.reply_text(
                f"⚡ Reintentando en SAT ahora...\n"
                f"NIT: {resultado.nit} | Serie: {resultado.serie} | Num: {resultado.numero_factura}"
            )
            await self._verificar_en_sat(query.message, context, resultado, user_id, pend_id=pend_id)

        elif accion == "verificar_sat":
            await query.answer("Verificando en SAT...")
            texto_ver = (
                f"Verificando en portal SAT...\n\n"
                f"NIT: {resultado.nit}\n"
                f"Serie: {resultado.serie}\n"
                f"Numero: {resultado.numero_factura}\n\n"
                f"Espera un momento..."
            )
            await self._edit_query_message(query, texto_ver)
            await self._verificar_en_sat(query.message, context, resultado, user_id, pend_id=pend_id)

        elif accion == "cancelar_factura":
            await query.answer("Factura cancelada")
            # Eliminar de pendientes
            if pendientes.pop(pend_id, None):
                _guardar_pendientes()
            # Limpiar estado de corrección si estaba en ese flujo
            if context.user_data.get('pend_id') == pend_id:
                context.user_data.clear()
            try:
                await self._edit_query_message(query, "❌ Factura cancelada y descartada.")
            except Exception:
                pass

        elif accion == "verificar_todas":
            await self._verificar_todas_handler(query, context, user_id)

    async def _verificar_todas_handler(self, query, context, user_id: int):
        """Valida en SAT todas las facturas pendientes del usuario en modo
        batch (los fallos no interrumpen el lote, se acumulan para el resumen
        final). Reutilizable desde /pendientes y desde el menu de extraccion."""
        mis_pendientes = [(pid, info) for pid, info in pendientes.items() if info.get('user_id') == user_id]
        total = len(mis_pendientes)
        if total == 0:
            await query.answer("No hay facturas pendientes", show_alert=True)
            return

        await query.answer(f"Verificando {total} facturas...")
        try:
            await query.edit_message_text(f"Verificando {total} facturas en SAT...\nEspera, esto puede tomar un momento.")
        except Exception:
            # editMessageText puede fallar si el mensaje no es del bot o ya
            # se edito; no es critico, el flujo sigue con el reply normal.
            pass

        # Pausa entre facturas para no disparar el anti-bot del SAT (que pide
        # captcha tras 2-3 consultas rapidas). Una pausa corta entre facturas
        # mantiene la sesion mas tiempo activa y reduce los captchas nuevos.
        PAUSA_ENTRE_FACTURAS = 4.0
        # Acumulamos los fallos para presentarlos juntos al final del lote.
        # Las exitosas ya emiten su mensaje (PDF + datos), las fallidas no
        # mostraran UI manual durante el lote (modo_batch=True).
        fallidas_lote: list = []  # [(pend_id, resultado), ...]
        logger.info(f"[batch] Iniciando lote de {total} facturas para user_id={user_id}")
        for i, (pid, info) in enumerate(mis_pendientes, 1):
            res_factura = info['resultado']
            logger.info(
                f"[batch] >>> Iniciando factura {i}/{total} "
                f"(pend_id={pid[:8]}, NIT={res_factura.nit}, "
                f"serie={res_factura.serie}, num={res_factura.numero_factura})"
            )
            # Pausa anti-bot SAT a partir de la 2da factura.
            if i > 1:
                await asyncio.sleep(PAUSA_ENTRE_FACTURAS)
            try:
                await query.message.reply_text(
                    f"[{i}/{total}] Verificando...\n"
                    f"NIT: {res_factura.nit} | Num: {res_factura.numero_factura}"
                )
                ok = await self._verificar_en_sat(
                    query.message, context, res_factura, user_id,
                    pend_id=pid, modo_batch=True,
                )
                logger.info(
                    f"[batch] <<< Termino factura {i}/{total}: ok={ok} "
                    f"estado={context.user_data.get('estado', '<none>')}"
                )
                if ok is False:
                    fallidas_lote.append((pid, res_factura))

                # Si pide captcha, detenerse — el usuario debe resolverlo
                if context.user_data.get('estado') == 'esperando_captcha_sat':
                    # Guardar las restantes para después
                    restantes = total - i
                    if restantes > 0:
                        await query.message.reply_text(
                            f"Captcha requerido. Resuelve el captcha y luego "
                            f"las {restantes} facturas restantes continuaran automaticamente."
                        )
                        # Guardar lista de pendientes restantes
                        context.user_data['verificar_cola'] = [
                            (pid2, info2) for pid2, info2 in mis_pendientes[i:]
                        ]
                        context.user_data['fallidas_lote'] = fallidas_lote
                    break
            except Exception as e:
                logger.exception(f"[batch] Excepcion en factura {i}/{total}: {e}")
                try:
                    await query.message.reply_text(
                        f"⚠️ Error procesando factura {i}/{total}: {e}\n"
                        f"Continuando con la siguiente..."
                    )
                except Exception:
                    pass
                # Tratar como fallida y seguir
                fallidas_lote.append((pid, res_factura))

        logger.info(
            f"[batch] Lote completado. Exitosas={total - len(fallidas_lote)} "
            f"Fallidas={len(fallidas_lote)} Estado={context.user_data.get('estado', '<none>')}"
        )
        if not context.user_data.get('estado') == 'esperando_captcha_sat':
            await self._mostrar_resumen_lote_sat(query.message, total, fallidas_lote, context=context)

    async def _mostrar_extraccion(self, update: Update, resultado: ExtraccionFactura, pend_id: str = ""):
        """Muestra resultado de extracción (OCR o catálogo) con botones de acción"""
        p = pend_id
        user_id = update.message.from_user.id
        mis_pendientes = [pid for pid, info in pendientes.items() if info.get('user_id') == user_id]
        fila_principal = [
            InlineKeyboardButton("✅ Validar en SAT", callback_data=f"verificar_sat:{p}"),
            InlineKeyboardButton("✏️ Corregir", callback_data=f"corregir:{p}"),
            InlineKeyboardButton("❌ Rechazar", callback_data=f"rechazar:{p}"),
        ]
        filas = [fila_principal]
        if len(mis_pendientes) >= 2:
            filas.append([InlineKeyboardButton(f"Validar TODAS ({len(mis_pendientes)})", callback_data=f"verificar_todas:{p}")])
        teclado = InlineKeyboardMarkup(filas)

        if resultado.metodo_extraccion == "catalogo":
            mensaje = (
                f"Encontrado en catalogo\n\n"
                f"Empresa: {resultado.razon_social}\n"
                f"NIT: {resultado.nit}\n"
                f"Numero: {resultado.numero_factura}\n"
                f"Serie: {resultado.serie}\n\n"
                f"Es correcto?"
            )
        else:
            mensaje = (
                f"Extraccion exitosa (OCR directo)\n\n"
                f"NIT: {resultado.nit}\n"
                f"Numero: {resultado.numero_factura}\n"
                f"Serie: {resultado.serie}\n\n"
                f"Razon social: {resultado.razon_social or '(no detectada)'}\n"
                f"Confianza: {resultado.confianza:.0%}"
            )

        await update.message.reply_text(mensaje, reply_markup=teclado)
    
    async def _pedir_confirmacion_manual(self, update: Update, resultado: ExtraccionFactura, archivo_path: Path, pend_id: str = ""):
        """Muestra datos parciales con botones para corregir lo que falte"""
        p = pend_id
        teclado = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("NIT", callback_data=f"corregir_nit:{p}"),
                InlineKeyboardButton("Serie", callback_data=f"corregir_serie:{p}"),
                InlineKeyboardButton("Numero", callback_data=f"corregir_numero:{p}"),
            ],
            [InlineKeyboardButton("✅ Validar en SAT", callback_data=f"verificar_sat:{p}")],
            [InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar_factura:{p}")],
        ])
        mensaje = (
            f"Datos parciales encontrados:\n"
            f"NIT: {resultado.nit or '(no detectado)'}\n"
            f"Serie: {resultado.serie or '(no detectada)'}\n"
            f"Numero: {resultado.numero_factura or '(no detectado)'}\n"
            f"Razon social: {resultado.razon_social or '(no detectada)'}\n\n"
            f"Corrige lo que falte y luego valida en SAT:"
        )
        # Reintentar el envio en caso de bache de red (ConnectError/getaddrinfo).
        # Si no se envia, el usuario no ve la UI para corregir y la factura queda
        # huerfana en `pendientes`.
        ultimo_error: Optional[Exception] = None
        for intento in range(1, 5):
            try:
                await update.message.reply_text(mensaje, reply_markup=teclado)
                return
            except Exception as e:
                ultimo_error = e
                logger.warning(
                    f"No se pudo enviar confirmacion manual (intento {intento}/4): {e}"
                )
                await asyncio.sleep(2 * intento)
        logger.error(
            f"Mensaje de confirmacion manual perdido tras 4 intentos. "
            f"pend_id={pend_id} user={update.message.from_user.id} "
            f"ultimo_error={ultimo_error}"
        )
    
    async def _mostrar_resumen_lote_sat(self, message, total: int, fallidas: list, context=None):
        """Resumen final de un lote de validacion SAT.

        `fallidas` es una lista de tuplas (pend_id, ExtraccionFactura). Las
        exitosas ya emitieron su mensaje (datos + PDF) durante el lote.

        En vez de listar todas las fallidas con un boton cada una, arranca un
        ASISTENTE secuencial: muestra la primera fallida (foto + datos +
        botones), y al resolverla (validar OK) o cancelarla, avanza sola a la
        siguiente, hasta que no queden problemas. Asi el usuario no tiene que
        ir tocando una por una de una lista.
        """
        exitosas = total - len(fallidas)
        if not fallidas:
            await message.reply_text(
                f"✅ Verificacion completada: {exitosas}/{total} facturas validadas.\n"
                f"Sin fallos."
            )
            return
        await message.reply_text(
            f"📋 Lote terminado: ✅ {exitosas} OK · ❌ {len(fallidas)} con problemas.\n\n"
            f"Vamos a resolver las {len(fallidas)} con problema una por una. "
            f"Cada una se valida o se cancela, y al terminar paso a la siguiente."
        )
        if context is not None:
            # Iniciar el asistente de revision secuencial.
            cola = [pid for pid, _ in fallidas if pid in pendientes]
            context.user_data['revision_cola'] = cola
            context.user_data['revision_total'] = len(cola)
            context.user_data['revision_resueltas'] = 0
            await self._mostrar_siguiente_revision(message, context, message.chat.id)
            return
        # Fallback (sin context): listado simple con botones (comportamiento viejo).
        botones = [
            [InlineKeyboardButton(
                f"{i}. Corregir N°{r.numero_factura}",
                callback_data=f"corregir:{pid}",
            )]
            for i, (pid, r) in enumerate(fallidas, 1)
        ]
        await message.reply_text(
            "Tocá una para corregirla:",
            reply_markup=InlineKeyboardMarkup(botones),
        )

    async def _mostrar_siguiente_revision(self, message, context, user_id: int):
        """Muestra la siguiente factura con problema del asistente de revision.

        Limpia de la cola las que ya no estan pendientes (resueltas/canceladas)
        y presenta la primera que quede como una tarjeta: foto + datos +
        botones Validar / Corregir / Cancelar. Si no queda ninguna, cierra el
        asistente con un resumen.
        """
        cola = context.user_data.get('revision_cola', []) or []
        # Podar las que ya se resolvieron o cancelaron (ya no estan en pendientes).
        cola = [pid for pid in cola if pid in pendientes]
        context.user_data['revision_cola'] = cola

        total = context.user_data.get('revision_total', 0)
        if not cola:
            # Terminamos: limpiar estado del asistente.
            context.user_data.pop('revision_cola', None)
            context.user_data.pop('revision_total', None)
            context.user_data.pop('revision_resueltas', None)
            await message.reply_text(
                "🎉 Listo. Ya no quedan facturas con problema.\n"
                "Todas fueron validadas o canceladas."
            )
            return

        pend_id = cola[0]
        info = pendientes.get(pend_id, {})
        resultado = info.get('resultado')
        if not resultado:
            # Defensa: si por alguna razon no hay resultado, sacarla y seguir.
            context.user_data['revision_cola'] = cola[1:]
            await self._mostrar_siguiente_revision(message, context, user_id)
            return

        restantes = len(cola)
        posicion = total - restantes + 1 if total else 1
        teclado = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("✅ Validar en SAT", callback_data=f"verificar_sat:{pend_id}"),
                InlineKeyboardButton("✏️ Corregir", callback_data=f"corregir:{pend_id}"),
            ],
            [InlineKeyboardButton("❌ Cancelar esta", callback_data=f"cancelar_factura:{pend_id}")],
        ])
        texto = (
            f"🔧 Factura con problema {posicion} de {total} "
            f"({restantes} restante{'s' if restantes != 1 else ''})\n\n"
            f"NIT: {resultado.nit}\n"
            f"Serie: {resultado.serie}\n"
            f"Numero: {resultado.numero_factura}\n"
            f"Empresa: {resultado.razon_social or '(no detectada)'}\n\n"
            f"Corregí el dato mal leído y validá, o cancelá esta factura."
        )
        archivo = info.get('archivo', '')
        if archivo and Path(archivo).exists():
            try:
                with open(archivo, 'rb') as f:
                    await message.reply_photo(photo=f, caption=texto, reply_markup=teclado)
                return
            except Exception as e:
                logger.warning(f"No se pudo enviar foto en revision: {e}")
        await message.reply_text(texto, reply_markup=teclado)

    async def _avanzar_revision_si_aplica(self, message, context, user_id: int, pend_id: str, accion: str):
        """Hook de auto-avance del asistente de revision.

        Se llama tras manejar un callback de factura. Si hay un asistente activo
        y la factura recien tocada se resolvio (ya no esta en pendientes) o se
        cancelo, pasamos a la siguiente. No avanza si quedo esperando captcha
        (la factura sigue pendiente) ni si la validacion fallo (sigue pendiente
        mostrando su UI de correccion)."""
        cola = context.user_data.get('revision_cola')
        if not cola:
            return
        if accion not in {"verificar_sat", "cancelar_factura", "rechazar"}:
            return
        # Si sigue esperando captcha, no avanzamos (se reanuda tras el captcha).
        if context.user_data.get('estado') == 'esperando_captcha_sat':
            return
        # Solo avanzar si la factura ya salio de pendientes (resuelta/cancelada).
        if pend_id in pendientes:
            return
        await self._mostrar_siguiente_revision(message, context, user_id)

    async def _verificar_en_sat(self, message, context, resultado: ExtraccionFactura, user_id: int, pend_id: str = "", modo_batch: bool = False):
        """Lanza la verificación SAT con reintentos automáticos.
        Flujo: intento 1 → reintento inmediato → reintento 30min → reintento 30min → avisa admin.

        modo_batch=True: cuando viene de una cola (verificar_todas o cola post-captcha).
        En ese caso NO mostramos UI manual al fallar — la factura queda en pendientes
        y el caller acumula los fallos para mostrarlos al final del lote.
        """
        # Limpieza oportunista: sweep de flags 'sonnet_escalado_*' cuyas
        # pend_ids ya no estan en `pendientes`. Sin esto crecen indefinidamente
        # en user_data (un bool por cada factura procesada en la sesion).
        try:
            keys_huerfanas = [
                k for k in list(context.user_data.keys())
                if k.startswith('sonnet_escalado_')
                and k.replace('sonnet_escalado_', '') not in pendientes
            ]
            for k in keys_huerfanas:
                context.user_data.pop(k, None)
            if keys_huerfanas:
                logger.debug(f"Limpiadas {len(keys_huerfanas)} flags sonnet huerfanas")
        except Exception:
            pass  # nunca bloquear la verificacion por la limpieza

        lock = _get_lock_usuario(str(user_id))
        async with lock:
            res = await self._intentar_verificacion_sat(message, context, resultado, user_id, pend_id, modo_batch=modo_batch)
            if res is None:
                return None  # Captcha o factura inválida ya manejada

            # Si fue exitoso (res == True) o ya manejado, verificar alerta de abuso
            if res is True:
                username = pendientes.get(pend_id, {}).get('username', '')
                total_mes = contar_facturas_mes(user_id, username)
                if total_mes >= LIMITE_FACTURAS_MES:
                    await notificar_admin(
                        context.bot,
                        f"Usuario {username} (ID: {user_id}) ha procesado {total_mes} facturas este mes. Limite: {LIMITE_FACTURAS_MES}"
                    )
            return res

    async def _intentar_verificacion_sat(self, message, context, resultado, user_id, pend_id, intento: int = 1, modo_batch: bool = False) -> Optional[bool]:
        """Intenta verificar en SAT. Retorna True si exitoso, None si ya manejado (captcha/invalida)."""
        MAX_REINTENTOS = 3
        ESPERA_REINTENTOS = [0, 180, 300]  # seg: inmediato, 3min, 5min

        try:
            verificador = await obtener_verificador(str(user_id))
            res = await asyncio.wait_for(
                verificador.verificar(
                    nit=resultado.nit,
                    serie=resultado.serie,
                    numero=resultado.numero_factura,
                    user_id=str(user_id)
                ),
                timeout=60  # 60 segundos de timeout
            )

            # Defensa: si SAT pide captcha pero no tenemos screenshot, no podemos
            # mostrarlo al usuario. En batch tratamos esto como fallo (la
            # factura queda pendiente) para no bloquear el lote.
            if res.esperando_captcha and not res.captcha_screenshot:
                logger.warning(
                    f"SAT pidio captcha sin screenshot — sesion probablemente "
                    f"corrupta. Cerrando verificador. (pend_id={pend_id})"
                )
                await cerrar_verificador(str(user_id))
                if modo_batch:
                    return False
                await message.reply_text(
                    "⚠️ SAT pidio un captcha pero no se pudo capturar la imagen. "
                    "Intenta de nuevo en unos segundos."
                )
                return None

            if res.esperando_captcha and res.captcha_screenshot:
                context.user_data['estado'] = 'esperando_captcha_sat'
                context.user_data['resultado_sat'] = resultado
                context.user_data['pend_id'] = pend_id
                with open(res.captcha_screenshot, 'rb') as f:
                    await message.reply_photo(photo=f, caption=res.mensaje, parse_mode="Markdown")
                # Programar timeout de 5 minutos via asyncio (siempre funciona)
                bot_ref = context.bot
                chat_id = message.chat_id
                uid = user_id
                # Cancelar timeout anterior si existia
                old_task = context.user_data.pop('captcha_timeout_task', None)
                if old_task and not old_task.done():
                    old_task.cancel()
                async def _timeout_asyncio():
                    await asyncio.sleep(300)
                    try:
                        await bot_ref.send_message(
                            chat_id=chat_id,
                            text=(
                                "⏰ *Tiempo agotado*\n\n"
                                "No recibimos el captcha en 5 minutos.\n"
                                "La sesion SAT fue cerrada automaticamente.\n"
                                "Puedes reintentar desde /fallidas o enviando la foto de nuevo."
                            ),
                            parse_mode="Markdown"
                        )
                        await cerrar_verificador(str(uid))
                    except asyncio.CancelledError:
                        pass  # Timeout cancelado — usuario respondio a tiempo
                    except Exception as ex:
                        logger.error(f"Error en timeout captcha: {ex}")
                task = asyncio.create_task(_timeout_asyncio())
                task.add_done_callback(
                    lambda t: logger.error(f"Captcha timeout error: {t.exception()}") if not t.cancelled() and t.exception() else None
                )
                context.user_data['captcha_timeout_task'] = task
                return None

            elif res.valida is False:
                # Escalación: si SAT rechazó los datos de Haiku y no hemos
                # probado Sonnet aún, re-extraer con Sonnet y reintentar SAT.
                # Solo un intento de escalación para no loopear infinito.
                info = pendientes.get(pend_id, {})
                archivo = info.get('archivo', '')
                ya_escalado = context.user_data.get(f'sonnet_escalado_{pend_id}', False)
                if (not ya_escalado and archivo and Path(archivo).exists()
                        and sistema.sonnet):
                    context.user_data[f'sonnet_escalado_{pend_id}'] = True
                    logger.info(f"SAT rechazó {resultado.nit}/{resultado.serie}/{resultado.numero_factura} — escalando a Sonnet")
                    try:
                        await message.reply_text("🔁 SAT no encontró la factura. Reintentando con lectura más precisa...")
                        nuevo = sistema.reextraer_con_sonnet(archivo, resultado_previo=resultado)
                        if nuevo and (
                            nuevo.nit != resultado.nit
                            or nuevo.serie != resultado.serie
                            or nuevo.numero_factura != resultado.numero_factura
                        ):
                            logger.info(f"Sonnet corrigió datos: {nuevo.nit}/{nuevo.serie}/{nuevo.numero_factura}")
                            # Actualizar resultado en pendientes y reintentar SAT
                            if pend_id in pendientes:
                                pendientes[pend_id]['extraccion'] = nuevo
                            return await self._intentar_verificacion_sat(
                                message, context, nuevo, user_id, pend_id, intento,
                                modo_batch=modo_batch,
                            )
                        else:
                            logger.info("Sonnet devolvió los mismos datos o nada — cae a UI manual")
                    except Exception as ex:
                        logger.warning(f"Escalación a Sonnet falló: {ex}")

                # Factura no encontrada (escalación no ayudó o no disponible).
                # En modo batch NO mostramos UI manual aqui — la factura se queda
                # en pendientes y el caller acumula los fallos para presentarlos
                # todos juntos al final, evitando que el usuario tenga que
                # ir alternando entre confirmacion y fallo durante el lote.
                if modo_batch:
                    logger.info(
                        f"[batch] SAT rechazo {resultado.nit}/{resultado.serie}/"
                        f"{resultado.numero_factura} — diferido al resumen final"
                    )
                    return False
                teclado = InlineKeyboardMarkup([
                    [
                        InlineKeyboardButton("NIT", callback_data=f"corregir_nit:{pend_id}"),
                        InlineKeyboardButton("Serie", callback_data=f"corregir_serie:{pend_id}"),
                        InlineKeyboardButton("Numero", callback_data=f"corregir_numero:{pend_id}"),
                    ],
                    [InlineKeyboardButton("Reintentar en SAT", callback_data=f"verificar_sat:{pend_id}")],
                    [InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar_factura:{pend_id}")],
                ])
                texto_error = (
                    f"Factura NO encontrada en SAT\n\n"
                    f"NIT: {resultado.nit}\n"
                    f"Serie: {resultado.serie}\n"
                    f"Numero: {resultado.numero_factura}\n\n"
                    f"Corrige el dato incorrecto o reintenta:"
                )
                if archivo and Path(archivo).exists():
                    with open(archivo, 'rb') as f:
                        await message.reply_photo(photo=f, caption=texto_error, reply_markup=teclado)
                else:
                    await message.reply_text(texto_error, reply_markup=teclado)
                return None

            elif res.valida is None and "tardando" in res.mensaje:
                # SAT tardando — aplicar reintentos automáticos
                raise asyncio.TimeoutError("SAT tardando en responder")

            elif res.valida is None:
                # Caso ambiguo: el verificador no pudo determinar si era valida
                # (formulario no cargo, frame perdido, sesion rota). La causa #1
                # es que la sesion del navegador quedo en mal estado y reusarla
                # vuelve a fallar — por eso antes habia que dar /reload a mano.
                #
                # Recuperacion automatica: cerramos el navegador (libera su slot)
                # y reintentamos DE INMEDIATO con una sesion fresca, sin esperar
                # los 180s del backoff normal. Es justo lo que hacia /reload.
                logger.warning(
                    f"SAT respondio con valida=None y mensaje='{res.mensaje}'. "
                    f"Reciclando navegador (sesion en mal estado)."
                )
                await cerrar_verificador(str(user_id))
                if intento < MAX_REINTENTOS:
                    logger.info(
                        f"Navegador reciclado — reintento inmediato {intento + 1}/{MAX_REINTENTOS} con sesion fresca"
                    )
                    if not modo_batch:
                        try:
                            await message.reply_text(
                                "⚠️ El SAT no cargó el formulario. Reintentando "
                                "automáticamente con un navegador nuevo..."
                            )
                        except Exception:
                            pass
                    # Pequena pausa para no martillar al SAT, pero sin el backoff largo.
                    await asyncio.sleep(3)
                    return await self._intentar_verificacion_sat(
                        message, context, resultado, user_id, pend_id,
                        intento + 1, modo_batch=modo_batch,
                    )
                # Agotados los reintentos de reciclaje.
                if modo_batch:
                    return False
                raise asyncio.TimeoutError(res.mensaje or "SAT respuesta ambigua")

            elif res.valida is True:
                # Factura válida
                username = pendientes.get(pend_id, {}).get('username', '')
                if res.razon_social_sat and resultado.nit:
                    resultado.razon_social = res.razon_social_sat
                    sistema.catalogo.reconciliar_con_sat(res.razon_social_sat, resultado.nit)
                    logger.info(f"Catálogo reconciliado con SAT: {resultado.nit} = {res.razon_social_sat}")

                # Verificar duplicado ANTES de copiar el PDF
                duplicada, reg_existente = buscar_duplicado(user_id, resultado, username)
                if duplicada:
                    await self._enviar_duplicado(message, resultado, reg_existente, res, user_id, username)
                else:
                    # No es duplicado: copiar PDF a carpeta del usuario y guardar registro
                    pdf_usuario = _copiar_pdf_a_usuario(res, user_id, username)
                    _, entrada_guardada = guardar_registro(user_id, resultado, pdf_path=pdf_usuario, username=username, res_sat=res)
                    # Registrar en sistema de aprendizaje
                    texto_ocr = getattr(resultado, 'texto_ocr', '')
                    sistema.registrar_confirmacion_sat(
                        nit=resultado.nit, serie=resultado.serie,
                        numero=resultado.numero_factura,
                        razon_social=resultado.razon_social or "",
                        texto_ocr=texto_ocr
                    )
                    await self._enviar_resultado_sat(message, res, user_id=user_id, username=username, pend_id=pend_id)
                    alerta = _verificar_comida_duplicada(user_id, username, entrada_guardada)
                    if alerta:
                        await message.reply_text(alerta, parse_mode="Markdown")
                    _clasificar_por_razon_social(entrada_guardada)
                    await self._prompt_odometro(message, context, user_id, username, entrada_guardada)
                if pendientes.pop(pend_id, None):
                    _guardar_pendientes()
                # Si esta factura venía de una fallida, marcarla como resuelta
                self._marcar_fallida_resuelta(pend_id)
                return True

        except (asyncio.TimeoutError, Exception) as e:
            es_timeout = isinstance(e, asyncio.TimeoutError)
            error_msg = "Timeout: SAT no respondio en 60s" if es_timeout else str(e)
            logger.error(f"Error verificando en SAT (intento {intento}/{MAX_REINTENTOS}): {error_msg}")

            if intento < MAX_REINTENTOS:
                espera = ESPERA_REINTENTOS[intento]  # 0, 1800, 1800

                # En modo batch NO programamos reintentos diferidos: cerrarian
                # el verificador y romperian el procesamiento del resto del
                # lote. Solo aceptamos reintento inmediato; si tampoco funciona,
                # marcamos como fallida y seguimos con la siguiente factura.
                if modo_batch and espera > 0:
                    logger.info(
                        f"[batch] Error tras intento {intento} ({error_msg}). "
                        f"No diferimos reintento en batch — marcando fallida."
                    )
                    return False

                if espera == 0:
                    await message.reply_text(
                        f"⚠️ Error en SAT: {error_msg}\n"
                        f"Reintentando inmediatamente... (intento {intento + 1}/{MAX_REINTENTOS})"
                    )
                else:
                    minutos = espera // 60
                    teclado_reintento = InlineKeyboardMarkup([
                        [InlineKeyboardButton("⚡ Reintentar ahora", callback_data=f"reintentar_ahora:{pend_id}")],
                        [InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar_factura:{pend_id}")],
                    ])
                    await message.reply_text(
                        f"⚠️ Error en SAT: {error_msg}\n"
                        f"Reintentando en {minutos} minutos... (intento {intento + 1}/{MAX_REINTENTOS})\n"
                        f"Te avisare cuando este listo.\n\n"
                        f"O puedes reintentar ahora mismo:",
                        reply_markup=teclado_reintento
                    )

                # Cerrar verificador para limpiar estado
                await cerrar_verificador(str(user_id))

                if espera > 0:
                    # Programar reintento en background (pasar bot directamente, context puede expirar)
                    task = asyncio.create_task(
                        self._reintento_programado(message, context, resultado, user_id, pend_id, intento + 1, espera, context.bot)
                    )
                    task.add_done_callback(
                        lambda t: logger.error(f"Reintento programado fallo: {t.exception()}") if not t.cancelled() and t.exception() else None
                    )
                    return None
                else:
                    # Reintento inmediato
                    return await self._intentar_verificacion_sat(
                        message, context, resultado, user_id, pend_id, intento + 1,
                        modo_batch=modo_batch,
                    )
            else:
                # Se agotaron los reintentos
                await cerrar_verificador(str(user_id))
                username = pendientes.get(pend_id, {}).get('username', '')
                # Guardar en cola persistente para no perder la factura
                _guardar_factura_fallida(user_id, username, resultado, error_msg, pend_id)
                # En batch: factura fallida silenciosa, el resumen la lista
                if modo_batch:
                    logger.info(
                        f"[batch] Se agotaron reintentos para "
                        f"{resultado.nit}/{resultado.serie}/{resultado.numero_factura}: {error_msg}"
                    )
                    return False
                teclado = InlineKeyboardMarkup([
                    [InlineKeyboardButton("Reintentar manualmente", callback_data=f"verificar_sat:{pend_id}")],
                    [InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar_factura:{pend_id}")],
                ])
                await message.reply_text(
                    f"❌ SAT no responde despues de {MAX_REINTENTOS} intentos (~8 minutos).\n"
                    f"El portal puede estar caido.\n\n"
                    f"La factura quedo guardada. Puedes reintentar manualmente mas tarde:\n"
                    f"NIT: {resultado.nit} | Serie: {resultado.serie} | Num: {resultado.numero_factura}",
                    reply_markup=teclado
                )
                # Notificar admin
                await notificar_admin(
                    context.bot,
                    f"SAT no responde despues de {MAX_REINTENTOS} intentos.\n"
                    f"Usuario: {username} (ID: {user_id})\n"
                    f"Factura: NIT={resultado.nit} Serie={resultado.serie} Num={resultado.numero_factura}\n"
                    f"Ultimo error: {error_msg}\n"
                    f"Guardada en facturas_fallidas.json para reintento manual."
                )
                return None

    async def _reintento_programado(self, message, context, resultado, user_id, pend_id, intento, espera_seg, bot=None):
        """Ejecuta un reintento de verificación SAT después de esperar."""
        await asyncio.sleep(espera_seg)
        if pend_id not in pendientes:  # Check seguro — solo lectura, no modifica
            logger.info(f"Reintento cancelado: factura {pend_id} ya no esta pendiente")
            return
        logger.info(f"Ejecutando reintento {intento} para factura {pend_id}")
        lock = _get_lock_usuario(str(user_id))
        async with lock:
            res = await self._intentar_verificacion_sat(message, context, resultado, user_id, pend_id, intento)
            if res is True:
                username = pendientes.get(pend_id, {}).get('username', '')
                total_mes = contar_facturas_mes(user_id, username)
                if total_mes >= LIMITE_FACTURAS_MES:
                    bot_inst = bot or context.bot
                    await notificar_admin(
                        bot_inst,
                        f"Usuario {username} (ID: {user_id}) ha procesado {total_mes} facturas este mes. Limite: {LIMITE_FACTURAS_MES}"
                    )

    async def _prompt_odometro(self, message, context, user_id: int, username: str, entrada: dict):
        """Si la factura es de gasolina, pide al usuario el kilometraje del odómetro."""
        if (entrada.get('tipo_comida') or '').lower() != 'gasolina':
            return
        if entrada.get('odometro') is not None:
            return  # ya tiene odómetro (ej. corregido manualmente)
        clave = _clave_de_registro(entrada)
        fecha = entrada.get('fecha_emision', '')
        anterior = obtener_odometro_anterior(user_id, username, fecha, clave_excluir=clave)
        perfil = leer_perfil_usuario(user_id, username)
        odo_ini = perfil.get('odometro_inicial')
        texto = (
            "⛽ *Factura de gasolina guardada.*\n"
            f"📅 {fecha or 'sin fecha'}\n\n"
            "Envíame el *kilometraje del odómetro* (solo el número, ej: 125430)."
        )
        if anterior and anterior.get('odometro') is not None:
            texto += f"\n\nÚltimo odómetro registrado: *{anterior['odometro']:,} km* ({anterior.get('fecha_emision', '?')})."
        elif odo_ini is not None:
            texto += f"\n\nOdómetro inicial configurado: *{int(odo_ini):,} km*."
        else:
            texto += (
                "\n\n💡 Es tu primera factura de gasolina con odómetro. Esta quedará como "
                "punto de partida y las siguientes calcularán km automáticamente."
            )
        context.user_data['esperando_odometro'] = {
            'clave': clave,
            'nit': entrada.get('nit'),
            'serie': entrada.get('serie'),
            'numero': entrada.get('numero'),
            'fecha': fecha,
            'anterior': anterior,
        }
        teclado = InlineKeyboardMarkup([[
            InlineKeyboardButton("❌ Saltar", callback_data="odometro_cancelar:")
        ]])
        await message.reply_text(texto, parse_mode="Markdown", reply_markup=teclado)

    async def cmd_tipo(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Muestra las últimas facturas para cambiarles el tipo de gasto."""
        if not await self._guard_acceso(update):
            return
        user_id = update.message.from_user.id
        username = update.message.from_user.username or update.message.from_user.first_name or ""
        carpeta = carpeta_usuario(user_id, username)
        registro_path = carpeta / "registro.json"
        if not registro_path.exists():
            await update.message.reply_text("No tienes facturas registradas.")
            return
        try:
            with open(registro_path, 'r', encoding='utf-8') as f:
                registros = json.load(f)
        except Exception:
            await update.message.reply_text("No se pudo leer el registro.")
            return
        # Enriquecer y filtrar facturas con monto > 0
        for r in registros:
            _enriquecer_desde_pdf(r)
            _clasificar_por_razon_social(r)
        facturas = [r for r in registros if _monto_efectivo(r.get('monto', '')) > 0]
        facturas.sort(
            key=lambda r: _fecha_para_orden(r.get('fecha_emision') or '') or datetime.min,
            reverse=True
        )
        facturas = facturas[:15]
        if not facturas:
            await update.message.reply_text("No tienes facturas con monto válido.")
            return
        sid = str(int(datetime.now().timestamp() * 1000))
        tipo_sessions[sid] = {
            'user_id': user_id,
            'username': username,
            'facturas': facturas,
            'timestamp': datetime.now(),
        }
        lineas = ["🏷 *Reclasificar tipo de gasto*\n"]
        botones = []
        for i, r in enumerate(facturas):
            fecha = r.get('fecha_emision', '?')
            monto = r.get('monto', 'Q?')
            tipo = r.get('tipo_comida') or '_sin tipo_'
            est = (r.get('razon_social') or r.get('establecimiento') or '?')[:22]
            lineas.append(f"{i + 1}. {fecha} — {est} — {monto} — *{tipo}*")
            botones.append([InlineKeyboardButton(
                f"✏️ {i + 1}. {fecha} · {tipo}",
                callback_data=f"tipo_pick:{sid}|{i}"
            )])
        botones.append([InlineKeyboardButton("❌ Cerrar", callback_data=f"tipo_cerrar:{sid}")])
        await update.message.reply_text(
            "\n".join(lineas) + "\n\nToca una factura para cambiarle el tipo.",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(botones)
        )

    async def cmd_odometro(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Gestiona kilometraje de facturas de gasolina.
        /odometro                  → lista con estado actual
        /odometro inicial <N>      → fija odómetro inicial de partida
        /odometro <N>              → fija odómetro de la última factura de gasolina sin registrar
        """
        if not await self._guard_acceso(update):
            return
        user_id = update.message.from_user.id
        username = update.message.from_user.username or update.message.from_user.first_name or ""
        args = context.args or []

        # /odometro inicial <N>
        if len(args) >= 2 and args[0].lower() in ('inicial', 'inicio', 'start'):
            valor = args[1].replace(',', '').replace('.', '').strip()
            if not valor.isdigit():
                await update.message.reply_text("Debe ser un número entero. Ej: `/odometro inicial 120000`", parse_mode="Markdown")
                return
            guardar_perfil_usuario(user_id, username, odometro_inicial=int(valor))
            await update.message.reply_text(
                f"✅ Odómetro inicial fijado: *{int(valor):,} km*\n"
                f"Desde esta base se calcularán los km recorridos en el reporte.",
                parse_mode="Markdown"
            )
            return

        # /odometro <N>  → asigna a la última factura de gasolina sin odómetro
        if len(args) == 1 and args[0].replace(',', '').replace('.', '').isdigit():
            valor = int(args[0].replace(',', '').replace('.', ''))
            gas = listar_facturas_gasolina(user_id, username)
            pendiente = next((g for g in gas if g.get('odometro') is None), None)
            if not pendiente:
                await update.message.reply_text(
                    "No hay facturas de gasolina sin odómetro. Usa `/odometro` para ver el listado.",
                    parse_mode="Markdown"
                )
                return
            clave = _clave_de_registro(pendiente)
            anterior = obtener_odometro_anterior(user_id, username, pendiente.get('fecha_emision', ''), clave_excluir=clave)
            aviso = validar_odometro_nuevo(valor, anterior)
            if guardar_odometro(user_id, username, pendiente['nit'], pendiente['serie'], pendiente['numero'], valor):
                msg = f"✅ Odómetro guardado: *{valor:,} km* en la factura del {pendiente.get('fecha_emision', '?')}."
                if aviso:
                    msg += f"\n\n{aviso}"
                await update.message.reply_text(msg, parse_mode="Markdown")
            else:
                await update.message.reply_text("No se pudo actualizar la factura.")
            return

        # Sin args (o args no reconocidos) → mostrar listado con botones
        gas = listar_facturas_gasolina(user_id, username)
        if not gas:
            await update.message.reply_text(
                "No tienes facturas de gasolina registradas aún.\n\n"
                "Puedes fijar un odómetro inicial con: `/odometro inicial 120000`",
                parse_mode="Markdown"
            )
            return
        perfil = leer_perfil_usuario(user_id, username)
        odo_ini = perfil.get('odometro_inicial')
        lineas = ["⛽ *Facturas de gasolina — kilometraje*\n"]
        if odo_ini is not None:
            lineas.append(f"🏁 Odómetro inicial: *{int(odo_ini):,} km*\n")
        else:
            # Si ya hay odómetros registrados, el más antiguo funciona como inicial implícito
            gas_con_odo = [g for g in gas if g.get('odometro') is not None]
            if gas_con_odo:
                gas_con_odo_ord = sorted(
                    gas_con_odo,
                    key=lambda r: _fecha_para_orden(r.get('fecha_emision') or '') or datetime.min
                )
                base = gas_con_odo_ord[0]
                lineas.append(
                    f"🏁 Inicial implícito: *{int(base['odometro']):,} km* "
                    f"({base.get('fecha_emision', '?')}, factura más antigua con odómetro)\n"
                )
            else:
                lineas.append("🏁 Odómetro inicial: _no configurado_ (opcional: `/odometro inicial <N>`)\n")

        # Guardar sesión para callbacks de edición
        sid = str(int(datetime.now().timestamp() * 1000))
        claves = []
        botones = []
        for i, r in enumerate(gas[:15]):  # limitar a 15 más recientes
            fecha = r.get('fecha_emision', '?')
            monto = r.get('monto', 'Q?')
            est = (r.get('razon_social') or r.get('establecimiento') or '?')[:22]
            odo = r.get('odometro')
            odo_txt = f"{odo:,} km" if odo is not None else "_sin odómetro_"
            lineas.append(f"{i + 1}. {fecha} — {est} — {monto} — {odo_txt}")
            claves.append(_clave_de_registro(r))
            etiqueta = f"{i + 1}. {fecha} · {odo_txt.replace('_','')}"
            botones.append([InlineKeyboardButton(f"✏️ {etiqueta}", callback_data=f"odometro_edit:{sid}|{i}")])

        odometro_sessions[sid] = {
            'user_id': user_id,
            'username': username,
            'claves': claves,
            'facturas': gas[:15],
            'timestamp': datetime.now(),
        }
        botones.append([InlineKeyboardButton("❌ Cerrar", callback_data=f"odometro_cerrar:{sid}")])
        texto = "\n".join(lineas)
        texto += "\n\nToca una factura para editar o eliminar su odómetro."
        await update.message.reply_text(texto, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(botones))

    async def _enviar_duplicado(self, message, resultado, reg_existente, res, user_id: int, username: str):
        """Notifica que la factura es duplicada y ofrece botones para ver PDF o cancelar."""
        carpeta = carpeta_usuario(user_id, username)

        # Verificar que haya un PDF guardado (para saber si mostrar botón o no)
        pdf_disponible = False
        if reg_existente:
            pdf_guardado = reg_existente.get('pdf_sat', '')
            if pdf_guardado and (carpeta / "pdf_sat" / Path(pdf_guardado).name).exists():
                pdf_disponible = True

        texto = (
            f"⚠️ Factura duplicada\n\n"
            f"NIT: {resultado.nit}\n"
            f"Serie: {resultado.serie}\n"
            f"Numero: {resultado.numero_factura}\n\n"
            f"Esta factura ya fue registrada anteriormente."
        )

        if pdf_disponible:
            clave = _clave_de_resultado(resultado)
            teclado = InlineKeyboardMarkup([[
                InlineKeyboardButton("📄 Ver PDF", callback_data=f"dup_pdf:{clave}"),
                InlineKeyboardButton("❌ Cancelar", callback_data="dup_cancel:"),
            ]])
            await message.reply_text(texto, reply_markup=teclado)
        else:
            await message.reply_text(
                texto + "\n(PDF no disponible — ya fue eliminado de descargas)"
            )

    async def _enviar_resultado_sat(self, message, res, user_id: int = None, username: str = "", pend_id: str = ""):
        """Envía el resultado de la verificación SAT al usuario"""
        # Nota: el PDF ya fue copiado a la carpeta del usuario antes de llamar a esta función

        # Si es timeout del SAT, ofrecer reintentar
        if res.valida is None and "tardando" in res.mensaje:
            teclado = InlineKeyboardMarkup([
                [InlineKeyboardButton("Reintentar", callback_data=f"verificar_sat:{pend_id}")]
            ])
            if res.captcha_screenshot and res.captcha_screenshot.exists():
                with open(res.captcha_screenshot, 'rb') as f:
                    await message.reply_photo(photo=f, caption=res.mensaje, reply_markup=teclado)
            else:
                await message.reply_text(res.mensaje, reply_markup=teclado)
            return

        # Enviar PDF si fue descargado
        if res.pdf_path and res.pdf_path.exists():
            with open(res.pdf_path, 'rb') as f:
                await message.reply_document(document=f, filename=res.pdf_path.name, caption=res.mensaje)
        elif res.captcha_screenshot and res.captcha_screenshot.exists():
            with open(res.captcha_screenshot, 'rb') as f:
                await message.reply_photo(photo=f, caption=res.mensaje[:1024])
        else:
            await message.reply_text(res.mensaje)

        if res.detalles and not res.pdf_path:
            await message.reply_text(res.detalles[:800])

    async def _limpiar_pendientes_viejos(self, context):
        """Elimina entradas viejas de pendientes, reporte_sessions y eliminar_sessions."""
        ahora = datetime.now()

        # Pendientes: expirar tras 24h
        expirados = [
            pid for pid, info in list(pendientes.items())
            if (ahora - info.get('timestamp', ahora)).total_seconds() > 86400
        ]
        for pid in expirados:
            pendientes.pop(pid, None)
        if expirados:
            logger.info(f"Limpieza pendientes: {len(expirados)} entradas expiradas eliminadas")
            _guardar_pendientes()

        # Sesiones de reporte: expirar tras 2h
        rep_expirados = [
            sid for sid, s in list(reporte_sessions.items())
            if (ahora - s.get('timestamp', ahora)).total_seconds() > 7200
        ]
        for sid in rep_expirados:
            reporte_sessions.pop(sid, None)
        if rep_expirados:
            logger.info(f"Limpieza reporte_sessions: {len(rep_expirados)} sesiones expiradas")

        # Sesiones de eliminar: expirar tras 30 min
        el_expirados = [
            sid for sid, s in list(eliminar_sessions.items())
            if (ahora - s.get('timestamp', ahora)).total_seconds() > 1800
        ]
        for sid in el_expirados:
            eliminar_sessions.pop(sid, None)
        if el_expirados:
            logger.info(f"Limpieza eliminar_sessions: {len(el_expirados)} sesiones expiradas")

        # Sesiones de clasificacion: expirar tras 2h
        # (si el usuario abandona /reporte a mitad de clasificar tipos de comida)
        cl_expirados = [
            sid for sid, s in list(clasificar_sessions.items())
            if (ahora - s.get('timestamp', ahora)).total_seconds() > 7200
        ]
        for sid in cl_expirados:
            clasificar_sessions.pop(sid, None)
        if cl_expirados:
            logger.info(f"Limpieza clasificar_sessions: {len(cl_expirados)} sesiones expiradas")

        # Sesiones de odometro (listado /odometro): expirar tras 1h
        odo_expirados = [
            sid for sid, s in list(odometro_sessions.items())
            if (ahora - s.get('timestamp', ahora)).total_seconds() > 3600
        ]
        for sid in odo_expirados:
            odometro_sessions.pop(sid, None)
        if odo_expirados:
            logger.info(f"Limpieza odometro_sessions: {len(odo_expirados)} sesiones expiradas")

        # Sesiones de /tipo: expirar tras 1h
        tipo_expirados = [
            sid for sid, s in list(tipo_sessions.items())
            if (ahora - s.get('timestamp', ahora)).total_seconds() > 3600
        ]
        for sid in tipo_expirados:
            tipo_sessions.pop(sid, None)
        if tipo_expirados:
            logger.info(f"Limpieza tipo_sessions: {len(tipo_expirados)} sesiones expiradas")

        # ── Limpieza de archivos temporales ──────────────────────────────────
        # Las fotos/JPG en facturas/ ya no sirven después de procesar la factura.
        # Los PDFs en downloads/ son temporales (se copian a pdf_sat/ del usuario).
        # Se conservan los PDFs en pdf_sat/ — esos sí son permanentes.
        archivos_borrados = 0

        # 1. Fotos y PDFs temporales en cada carpeta facturas/ (> 7 días)
        DIAS_FACTURAS = 7
        limite_facturas = ahora.timestamp() - DIAS_FACTURAS * 86400
        for carpeta_usr in DATOS_DIR.iterdir():
            if not carpeta_usr.is_dir():
                continue
            facturas_dir = carpeta_usr / "facturas"
            if not facturas_dir.exists():
                continue
            for archivo in facturas_dir.iterdir():
                if archivo.suffix.lower() in ('.jpg', '.jpeg', '.png', '.pdf'):
                    try:
                        if archivo.stat().st_mtime < limite_facturas:
                            archivo.unlink()
                            archivos_borrados += 1
                    except Exception:
                        pass

        # 2. PDFs temporales en downloads/ (> 24 horas)
        downloads_dir = Path("downloads")
        if downloads_dir.exists():
            limite_downloads = ahora.timestamp() - 86400
            for archivo in downloads_dir.iterdir():
                if archivo.is_file():
                    try:
                        if archivo.stat().st_mtime < limite_downloads:
                            archivo.unlink()
                            archivos_borrados += 1
                    except Exception:
                        pass

        if archivos_borrados:
            logger.info(f"Limpieza archivos: {archivos_borrados} archivo(s) temporal(es) eliminado(s)")

    async def cmd_buscar(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Busca facturas del usuario por empresa, NIT, serie, número o monto."""
        user_id = update.message.from_user.id
        username = update.message.from_user.username or update.message.from_user.first_name or ""
        termino = " ".join(context.args or "").strip().lower()
        if not termino:
            await update.message.reply_text(
                "🔍 *Buscar facturas*\n\nUso: /buscar <término>\n\nEjemplos:\n"
                "/buscar campero\n/buscar 1234567\n/buscar Q150",
                parse_mode="Markdown"
            )
            return
        carpeta = carpeta_usuario(user_id, username)
        registro_path = carpeta / "registro.json"
        if not registro_path.exists():
            await update.message.reply_text("📭 No tienes facturas registradas aún.")
            return
        try:
            with open(registro_path, 'r', encoding='utf-8') as f:
                registros = json.load(f)
        except Exception:
            await update.message.reply_text("❌ Error leyendo tus registros.")
            return
        for r in registros:
            _enriquecer_desde_pdf(r)
            _clasificar_por_razon_social(r)
        # Buscar en todos los campos relevantes
        campos = ['nit', 'serie', 'numero', 'razon_social', 'establecimiento', 'monto', 'fecha_emision', 'tipo_comida']
        resultados = []
        for r in registros:
            texto_busqueda = " ".join(str(r.get(c, '') or '') for c in campos).lower()
            if termino in texto_busqueda:
                resultados.append(r)
        if not resultados:
            await update.message.reply_text(f"🔍 No se encontraron facturas para: *{termino}*", parse_mode="Markdown")
            return
        lineas = [f"🔍 *{len(resultados)} resultado(s) para \"{termino}\"*\n"]
        for i, r in enumerate(resultados[:15], 1):
            empresa = r.get('razon_social') or r.get('establecimiento') or r.get('nit', '?')
            fecha = r.get('fecha_emision', '?') or '?'
            monto = r.get('monto', '?') or '?'
            tipo = r.get('tipo_comida', '') or ''
            nit = r.get('nit', '?')
            serie = r.get('serie', '?')
            lineas.append(f"{i}. *{empresa}*\n   NIT: `{nit}` | Serie: `{serie}`\n   📅 {fecha} | 💰 {monto}" + (f" | {tipo}" if tipo else ""))
        if len(resultados) > 15:
            lineas.append(f"\n_...y {len(resultados) - 15} más. Usa un término más específico._")
        await update.message.reply_text("\n".join(lineas), parse_mode="Markdown")

    async def cmd_eliminar(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Muestra las últimas facturas del usuario con botón para eliminar cada una."""
        user_id = update.message.from_user.id
        username = update.message.from_user.username or update.message.from_user.first_name or ""
        carpeta = carpeta_usuario(user_id, username)
        registro_path = carpeta / "registro.json"
        if not registro_path.exists():
            await update.message.reply_text("📭 No tienes facturas registradas.")
            return
        try:
            with open(registro_path, 'r', encoding='utf-8') as f:
                registros = json.load(f)
        except Exception:
            await update.message.reply_text("❌ Error leyendo tus registros.")
            return
        for r in registros:
            _enriquecer_desde_pdf(r)
            _clasificar_por_razon_social(r)
        # Mostrar las últimas 10 con monto (las más recientes primero)
        con_monto = [r for r in registros if r.get('monto')]
        ultimas = list(reversed(con_monto[-10:]))
        if not ultimas:
            await update.message.reply_text("No tienes facturas verificadas para eliminar.")
            return
        # Guardar sesion
        sid = uuid.uuid4().hex[:10]
        eliminar_sessions[sid] = {
            'user_id': user_id,
            'username': username,
            'facturas_display': ultimas,
            'timestamp': datetime.now(),
        }
        await update.message.reply_text(
            "🗑 *Eliminar factura*\n\nElige la factura a eliminar del registro.\n⚠️ Esta acción no se puede deshacer.",
            parse_mode="Markdown"
        )
        for i, fac in enumerate(ultimas):
            empresa = fac.get('razon_social') or fac.get('establecimiento') or fac.get('nit', '?')
            fecha = fac.get('fecha_emision', '?') or '?'
            monto = fac.get('monto', '?') or '?'
            tipo = fac.get('tipo_comida', '') or ''
            nit = fac.get('nit', '?')
            texto_fac = f"*{empresa}*\nNIT: `{nit}` | 📅 {fecha} | 💰 {monto}" + (f" | {tipo}" if tipo else "")
            teclado = InlineKeyboardMarkup([
                [InlineKeyboardButton(f"🗑 Eliminar esta factura", callback_data=f"eliminar_factura:{sid}|{i}")]
            ])
            await update.message.reply_text(texto_fac, parse_mode="Markdown", reply_markup=teclado)

    async def _recordatorio_viernes(self, context: ContextTypes.DEFAULT_TYPE):
        """Job de viernes: avisar a usuarios con facturas de la semana o recordar que no tienen."""
        hoy = datetime.now()
        lunes = hoy - timedelta(days=hoy.weekday())
        fi = lunes.strftime("%d/%m/%Y")
        ff = hoy.strftime("%d/%m/%Y")
        # Buscar todos los usuarios que tengan carpeta en datos/
        for carpeta in DATOS_DIR.iterdir():
            if not carpeta.is_dir() or not carpeta.name.startswith("usuario_"):
                continue
            try:
                partes = carpeta.name.split("_", 2)
                uid = int(partes[1])
                uname = partes[2] if len(partes) > 2 else ""
            except (IndexError, ValueError):
                continue
            registro_path = carpeta / "registro.json"
            if not registro_path.exists():
                continue
            facturas_semana = _filtrar_facturas_para_reporte(uid, uname, fi, ff)
            try:
                if facturas_semana:
                    total = sum(
                        float(f.get('monto', 'Q0').replace('Q', '').replace(',', '')) if f.get('monto') else 0
                        for f in facturas_semana
                    )
                    teclado = InlineKeyboardMarkup([
                        [InlineKeyboardButton("📊 Generar reporte de la semana", callback_data=f"recordatorio_reporte:")],
                    ])
                    await context.bot.send_message(
                        chat_id=uid,
                        text=(
                            f"📅 *Recordatorio semanal*\n\n"
                            f"Tienes *{len(facturas_semana)} factura(s)* esta semana por un total de *Q{total:.2f}*.\n\n"
                            f"¿Deseas generar el reporte?"
                        ),
                        reply_markup=teclado,
                        parse_mode="Markdown"
                    )
                else:
                    teclado = InlineKeyboardMarkup([
                        [InlineKeyboardButton("✅ Confirmar — no hay facturas", callback_data="recordatorio_sin_facturas:ok")],
                    ])
                    await context.bot.send_message(
                        chat_id=uid,
                        text=(
                            "📅 *Recordatorio semanal*\n\n"
                            "No tienes facturas registradas esta semana.\n\n"
                            "¿Confirmas que no hay facturas para reportar?"
                        ),
                        reply_markup=teclado,
                        parse_mode="Markdown"
                    )
            except Exception as e:
                logger.warning(f"No se pudo enviar recordatorio a uid={uid}: {e}")

    async def _notificar_inicio(self, app):
        """Notifica al admin cuando el bot arranca correctamente."""
        # Restaurar facturas pendientes del disco
        _cargar_pendientes()
        # Programar limpieza periodica de pendientes cada 6 horas (solo si JobQueue esta disponible)
        if app.job_queue:
            try:
                app.job_queue.run_repeating(
                    self._limpiar_pendientes_viejos,
                    interval=21600,
                    first=3600,
                    name="limpiar_pendientes",
                )
                # Recordatorio viernes 6pm — calcular segundos hasta próximo viernes 18:00
                ahora = datetime.now()
                dias_hasta_viernes = (4 - ahora.weekday()) % 7  # 4 = viernes
                proximo_viernes = ahora.replace(hour=18, minute=0, second=0, microsecond=0) + timedelta(days=dias_hasta_viernes)
                if proximo_viernes <= ahora:
                    proximo_viernes += timedelta(days=7)
                segundos_hasta_viernes = (proximo_viernes - ahora).total_seconds()
                app.job_queue.run_repeating(
                    self._recordatorio_viernes,
                    interval=7 * 24 * 3600,  # cada 7 días
                    first=segundos_hasta_viernes,
                    name="recordatorio_viernes",
                )
                logger.info(f"✅ JobQueue activo: limpieza cada 6h, recordatorio viernes a las 18:00 (en {segundos_hasta_viernes/3600:.1f}h).")
            except Exception as e:
                logger.error(f"⚠️  No se pudo programar jobs con JobQueue: {e}")
                logger.warning("    El bot sigue funcionando pero sin limpieza/recordatorio automatico.")
        else:
            logger.warning("⚠️  JobQueue no disponible. Limpieza periodica desactivada.")
            logger.warning("    Para activarla: pip install 'python-telegram-bot[job-queue]'")
        if ADMIN_CHAT_ID:
            admin_id_int = _safe_int(ADMIN_CHAT_ID)
            if admin_id_int is None:
                logger.error(f"ADMIN_CHAT_ID invalido: {ADMIN_CHAT_ID!r} — no se enviara notificacion de inicio.")
            else:
                try:
                    pendientes_restaurados = len(pendientes)
                    msg_inicio = "✅ Bot de facturas iniciado correctamente.\nLas notificaciones de errores llegaran aqui."
                    if pendientes_restaurados:
                        msg_inicio += f"\n\n🔄 {pendientes_restaurados} factura(s) pendiente(s) restaurada(s) desde disco."
                    await app.bot.send_message(
                        chat_id=admin_id_int,
                        text=msg_inicio
                    )
                    logger.info(f"Notificacion de inicio enviada a ADMIN_CHAT_ID={admin_id_int}")
                except Exception as e:
                    logger.error(f"No se pudo enviar notificacion de inicio: {e}")
        else:
            logger.warning("⚠️  ADMIN_CHAT_ID no configurado. No se enviaran alertas de errores.")
            logger.warning("    Para configurar: agrega ADMIN_CHAT_ID=tu_chat_id en el archivo .env")
            logger.warning("    Obtén tu chat_id enviando un mensaje a @userinfobot en Telegram")

    async def cmd_fallidas(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Muestra las facturas que fallaron en SAT y permite reintentarlas."""
        fallidas_path = DATOS_DIR / "facturas_fallidas.json"
        if not fallidas_path.exists():
            await update.message.reply_text("✅ No hay facturas fallidas guardadas.")
            return
        try:
            with open(fallidas_path, 'r', encoding='utf-8') as f:
                fallidas = json.load(f)
        except Exception:
            await update.message.reply_text("❌ Error leyendo lista de fallidas.")
            return

        pendientes_fallidas = [f for f in fallidas if not f.get('resuelta')]
        if not pendientes_fallidas:
            await update.message.reply_text("✅ No hay facturas fallidas pendientes de reintentar.")
            return

        await update.message.reply_text(
            f"❌ *Facturas fallidas ({len(pendientes_fallidas)})*\n\n"
            f"Estas facturas no pudieron verificarse en SAT tras 3 intentos.\n"
            f"Puedes reintentar cada una individualmente:",
            parse_mode="Markdown"
        )

        for falla in pendientes_fallidas[:10]:  # Máximo 10 para no saturar
            pid = falla.get('pend_id', '')
            ts = falla.get('timestamp', 'fecha desconocida')
            nit = falla.get('nit', '?')
            serie = falla.get('serie', '?')
            numero = falla.get('numero', '?')
            username = falla.get('username', '?')
            error = falla.get('ultimo_error', 'desconocido')[:80]

            teclado = InlineKeyboardMarkup([
                [InlineKeyboardButton("🔄 Reintentar en SAT", callback_data=f"reintentar_fallida:{pid}")],
                [InlineKeyboardButton("✅ Marcar como resuelta", callback_data=f"resolver_fallida:{pid}")],
            ])
            await update.message.reply_text(
                f"📄 *Factura*\n"
                f"Usuario: @{username}\n"
                f"NIT: `{nit}` | Serie: `{serie}` | Num: `{numero}`\n"
                f"Fecha: {ts}\n"
                f"Error: _{error}_",
                reply_markup=teclado,
                parse_mode="Markdown"
            )

        if len(pendientes_fallidas) > 10:
            await update.message.reply_text(
                f"...y {len(pendientes_fallidas) - 10} facturas mas. Resuelve estas primero."
            )

    async def cmd_reload(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Cierra la sesion SAT del usuario (navegador + cookies + lock) para
        forzar una sesion fresca en la proxima verificacion. Util cuando la
        sesion quedo trabada (captcha no llega, formulario corrupto, etc.)."""
        if not await self._guard_acceso(update):
            return
        user_id = update.message.from_user.id
        # Tambien limpiamos el estado en memoria del flujo de captcha por si
        # quedo a medias — asi el siguiente intento arranca limpio.
        context.user_data.pop('estado', None)
        context.user_data.pop('resultado_sat', None)
        context.user_data.pop('captcha_timeout_task', None)
        context.user_data.pop('verificar_cola', None)
        context.user_data.pop('fallidas_lote', None)
        context.user_data.pop('revision_cola', None)
        context.user_data.pop('revision_total', None)
        context.user_data.pop('revision_resueltas', None)
        try:
            await cerrar_verificador(str(user_id))
            await update.message.reply_text(
                "🔄 Sesion SAT reiniciada.\n\n"
                "La proxima validacion abrira un navegador nuevo y "
                "pedira un captcha fresco. Usa /pendientes o reenvia "
                "una foto para continuar."
            )
        except Exception as e:
            logger.warning(f"Error en /reload para user {user_id}: {e}")
            await update.message.reply_text(
                f"⚠️ No se pudo cerrar limpiamente: {e}\n\n"
                "Aun asi, el estado en memoria quedo limpio. Probá enviar "
                "una factura."
            )

    async def cmd_estado(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Muestra el estado actual del bot."""
        user_id = update.message.from_user.id
        username = update.message.from_user.username or update.message.from_user.first_name or ""

        # Pendientes en memoria
        total_pendientes = len(pendientes)
        mis_pendientes = sum(1 for info in pendientes.values() if info.get('user_id') == user_id)

        # Fallidas
        fallidas_path = DATOS_DIR / "facturas_fallidas.json"
        total_fallidas = 0
        if fallidas_path.exists():
            try:
                with open(fallidas_path, 'r', encoding='utf-8') as f:
                    fallidas = json.load(f)
                total_fallidas = sum(1 for f in fallidas if not f.get('resuelta'))
            except Exception as e:
                logger.warning(f"No se pudo leer facturas_fallidas.json para /estado: {e}")

        # Facturas del usuario este mes
        facturas_mes = contar_facturas_mes(user_id, username)

        # Ping SAT — conexion TCP al puerto 443 (mas fiable que HTTP GET que el SAT bloquea)
        sat_estado = "⏳ verificando..."
        await update.message.reply_text("🔍 Verificando estado del SAT...")
        try:
            reader, writer = await asyncio.wait_for(
                asyncio.open_connection("felcons.c.sat.gob.gt", 443),
                timeout=10
            )
            writer.close()
            try:
                await writer.wait_closed()
            except Exception:
                pass
            sat_estado = "✅ Portal SAT en linea"
        except asyncio.TimeoutError:
            sat_estado = "⚠️ SAT no responde (timeout 10s) — puede estar caido"
        except OSError as e:
            sat_estado = f"❌ SAT no accesible — {e.strerror}"
        except Exception as e:
            sat_estado = f"❌ SAT no accesible ({type(e).__name__})"

        await update.message.reply_text(
            f"📊 *Estado del Bot*\n\n"
            f"*Facturas pendientes:*\n"
            f"  • Total en cola: {total_pendientes}\n"
            f"  • Tuyas: {mis_pendientes}\n\n"
            f"*Facturas fallidas:* {total_fallidas} (usa /fallidas para reintentar)\n\n"
            f"*Tus facturas este mes:* {facturas_mes}\n\n"
            f"*Portal SAT:* {sat_estado}",
            parse_mode="Markdown"
        )

    def _marcar_fallida_resuelta(self, pend_id: str):
        """Marca como resuelta una factura fallida si el pend_id coincide."""
        fallidas_path = DATOS_DIR / "facturas_fallidas.json"
        if not fallidas_path.exists():
            return
        try:
            with open(fallidas_path, 'r', encoding='utf-8') as f:
                fallidas = json.load(f)
            cambiado = False
            for falla in fallidas:
                if falla.get('pend_id') == pend_id and not falla.get('resuelta'):
                    falla['resuelta'] = True
                    cambiado = True
            if cambiado:
                _guardar_json_atomico(fallidas_path, fallidas)
        except Exception as e:
            logger.warning(f"No se pudo marcar fallida como resuelta: {e}")

    async def _error_handler(self, update, context: ContextTypes.DEFAULT_TYPE):
        """Captura todos los errores no manejados del bot."""
        error = context.error
        if tg_error and isinstance(error, tg_error.Conflict):
            logger.critical(
                "❌ ERROR 409 CONFLICT: Hay OTRA instancia del bot corriendo al mismo tiempo.\n"
                "   Detén el otro proceso (servidor Linux, otra ventana, etc.) antes de\n"
                "   correr el bot aqui. Ambas instancias no pueden coexistir."
            )
            return  # No reintentar — el usuario debe resolverlo manualmente
        if tg_error and isinstance(error, tg_error.NetworkError):
            logger.warning(f"⚠️  Error de red (se reintentara automaticamente): {error}")
            return
        if tg_error and isinstance(error, tg_error.TimedOut):
            logger.warning(f"⚠️  Timeout de red (se reintentara): {error}")
            return
        # Error inesperado — loggear y notificar admin
        logger.error(f"Error no manejado en update {update}: {error}", exc_info=error)
        try:
            await notificar_admin(
                context.bot,
                f"Error inesperado en el bot:\n{type(error).__name__}: {error}"
            )
        except Exception:
            pass

    def iniciar(self):
        """Inicia el bot"""
        logger.info(f"🤖 Iniciando bot con token: {self.token[:20]}...")
        # Validacion temprana: la plantilla del reporte debe existir o nada funcionara
        if not PLANTILLA_REPORTE.exists():
            logger.error(f"❌ Plantilla del reporte no encontrada: {PLANTILLA_REPORTE}")
            logger.error("   Coloca 'plantilla_reporte.xlsx' junto a telegram_bot.py antes de iniciar.")
            raise FileNotFoundError(f"Falta plantilla del reporte: {PLANTILLA_REPORTE}")
        # Asegurar que el admin siempre esté aprobado en la base de usuarios
        if ADMIN_CHAT_ID:
            admin_id_int = _safe_int(ADMIN_CHAT_ID)
            if admin_id_int is None:
                logger.error(f"⚠️  ADMIN_CHAT_ID invalido en .env: {ADMIN_CHAT_ID!r} (debe ser entero).")
            else:
                try:
                    _registrar_usuario(admin_id_int, "admin", "Admin")
                    _aprobar_usuario(admin_id_int)
                except Exception as e:
                    logger.warning(f"No se pudo registrar admin: {e}")
        jq = self.app.job_queue
        if jq:
            logger.info("✅ JobQueue disponible: limpieza periodica y timeouts activos.")
        else:
            logger.warning("⚠️  JobQueue no disponible (instala: pip install 'python-telegram-bot[job-queue]').")
        self.app.post_init = self._notificar_inicio
        self.app.run_polling()


# Script principal
if __name__ == "__main__":
    if not TELEGRAM_AVAILABLE:
        logger.error("❌ python-telegram-bot no instalado")
        logger.error("Instala con: pip install python-telegram-bot --upgrade")
        exit(1)
    
    token = TELEGRAM_TOKEN
    if token == "YOUR_TOKEN_HERE":
        logger.error("❌ TELEGRAM_BOT_TOKEN no configurado")
        logger.error("Establece: export TELEGRAM_BOT_TOKEN='tu_token_aqui'")
        exit(1)
    
    try:
        logger.info("✓ Inicializando SAT Invoice Bot...")
        bot = BotFacturasGT(token)
        logger.info("✓ Bot listo. Esperando mensajes...")
        bot.iniciar()
    except KeyboardInterrupt:
        logger.info("\n✓ Bot detenido")
    except Exception as e:
        logger.error(f"❌ Error fatal: {e}")
        exit(1)
