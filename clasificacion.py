"""
Clasificación de facturas y política de viáticos.

Módulo autocontenido (sin dependencias de Telegram) extraído de telegram_bot.py
para reducir el tamaño de ese archivo y tener una única fuente de verdad para:

- Departamentos válidos del dropdown BD del Excel y su normalización.
- Política de viáticos T1/T2 (montos por tipo de gasto).
- Clasificación por razón social/establecimiento (Gasolina/Hospedaje/Parqueo/comida).
- Reclasificación retroactiva de Hospedaje con monto bajo.
- Extracción/normalización de departamento desde el PDF oficial de la SAT.

Lo importa telegram_bot.py. También puede importarlo sat_verificador para no
duplicar las listas de keywords (evita que se desincronicen).
"""

import re
import logging
import unicodedata
from pathlib import Path

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────
#  DEPARTAMENTOS Y POLÍTICA DE VIÁTICOS
# ─────────────────────────────────────────────

# Valores exactos del dropdown BD del Excel (tal como aparecen en la hoja).
DEPARTAMENTOS_BD = [
    'ALTA VERAPAZ', 'BAJA VERAPAZ', 'CHIMALTENAGO', 'CHIQUIMULA',
    'GUATEMALA', 'EL PROGRESO', 'ESCUINTLA', 'HUEHUETENANGO', 'IZABAL',
    'JALAPA', 'JUTIAPA', 'PETÉN', 'QUETZALTENANGO', 'QUICHÉ',
    'RETALHULEU', 'SACATEPEQUEZ', 'SAN MARCOS', 'SANTA ROSA',
    'SOLOLÁ', 'SUCHITEPEQUEZ', 'TOTONICAPÁN', 'ZACAPA',
]

# Departamentos con tarifa T1 (monto mayor de viáticos).
# ÚNICA fuente de verdad, usada tanto por la fórmula Excel (celda M) como
# por el TXT. Los strings DEBEN coincidir exactamente con DEPARTAMENTOS_BD
# porque la fórmula Excel compara literal contra la celda H.
DEPTOS_TABLA1_BD = [
    'GUATEMALA', 'ALTA VERAPAZ', 'HUEHUETENANGO', 'PETÉN',
    'QUETZALTENANGO', 'RETALHULEU', 'SACATEPEQUEZ', 'SOLOLÁ',
]

# Montos de política por tipo y tabla. Mismo orden/valores que la fórmula
# IF de la celda M del Excel — mantener sincronizado.
POLITICA_VIATICOS = {
    'Desayuno':  {'T1': 45,  'T2': 40},
    'Almuerzo':  {'T1': 90,  'T2': 80},
    'Cena':      {'T1': 45,  'T2': 40},
    'Hospedaje': {'T1': 300, 'T2': 250},
}

# Un hospedaje real en Guatemala no cuesta menos que esto. Debajo de este
# monto, probablemente es comida del restaurante del hotel, parqueo,
# lavandería, etc. — no el hospedaje en sí.
HOSPEDAJE_MONTO_MINIMO = 100.0


def _reclasificar_tipo_si_necesario(fac: dict) -> str:
    """Devuelve el tipo_comida efectivo para generar el reporte.

    Si el tipo guardado es 'Hospedaje' pero el monto está por debajo de
    HOSPEDAJE_MONTO_MINIMO, reclasifica por hora (Desayuno/Almuerzo/Cena)
    usando los mismos umbrales que el verificador SAT. Esto arregla
    retroactivamente facturas viejas guardadas antes del fix del
    clasificador, sin tocar el registros.json (el fix se aplica solo al
    momento de escribir el reporte).
    """
    tipo = fac.get('tipo_comida', '') or ''
    if tipo != 'Hospedaje':
        return tipo
    try:
        monto_val = float(
            str(fac.get('monto', '')).replace('Q', '').replace(',', '').strip()
        )
    except (ValueError, TypeError):
        return tipo
    if monto_val >= HOSPEDAJE_MONTO_MINIMO:
        return tipo
    hora = (fac.get('hora_emision') or '').strip()
    nuevo = 'Almuerzo'  # fallback si no hay hora confiable
    try:
        partes = hora.split(':')
        h = int(partes[0])
        m = int(partes[1]) if len(partes) > 1 else 0
        minutos = h * 60 + m
        if minutos < 11 * 60 + 20:
            nuevo = 'Desayuno'
        elif minutos < 17 * 60 + 30:
            nuevo = 'Almuerzo'
        else:
            nuevo = 'Cena'
    except (ValueError, IndexError):
        pass
    logger.info(
        f"Reclasificando factura NIT {fac.get('nit')} "
        f"serie {fac.get('serie')} num {fac.get('numero')}: "
        f"Hospedaje (Q{monto_val:.2f}) → {nuevo} "
        f"(monto < Q{HOSPEDAJE_MONTO_MINIMO:.0f}, hora '{hora}')"
    )
    return nuevo


def _normalizar(texto: str) -> str:
    """Quita tildes y convierte a mayúsculas para comparación."""
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto.upper())
        if unicodedata.category(c) != 'Mn'
    )


# Mapa normalizado → valor exacto del BD
_DEPTO_NORM = {_normalizar(d): d for d in DEPARTAMENTOS_BD}


def normalizar_departamento(depto: str) -> str:
    """Convierte un departamento o municipio extraído al valor exacto del dropdown BD."""
    if not depto:
        return ""
    norm = _normalizar(depto.strip())
    # Alias para grafías que NO coinciden con el valor (a veces con typo) del
    # dropdown del Excel. El SAT devuelve "CHIMALTENANGO" bien escrito, pero la
    # plantilla del Excel tiene "CHIMALTENAGO" (sin una N). Para que la celda
    # quede con el valor que el dropdown reconoce, traducimos al valor del BD.
    # NO se toca el Excel; solo se mapea el nombre real al que espera la hoja.
    _ALIAS_BD = {
        'CHIMALTENANGO': 'CHIMALTENAGO',  # dropdown del Excel sin la 2a N
    }
    if norm in _ALIAS_BD:
        return _ALIAS_BD[norm]
    # Coincidencia exacta
    if norm in _DEPTO_NORM:
        return _DEPTO_NORM[norm]
    # Coincidencia parcial (ej. "CHIQUIMULA, CHIQUIMULA" → "CHIQUIMULA")
    for key, val in _DEPTO_NORM.items():
        if key in norm or norm in key:
            return val
    # Municipios comunes que no matchean directo → mapear a su departamento BD
    _MUNICIPIO_A_DEPTO = {
        'SANTA CRUZ':     'ALTA VERAPAZ',
        'VERAPAZ':        'ALTA VERAPAZ',
        'COBAN':          'ALTA VERAPAZ',
        'CAHABON':        'ALTA VERAPAZ',
        'CHISEC':         'ALTA VERAPAZ',
        'ANTIGUA':        'SACATEPEQUEZ',
        'JOCOTENANGO':    'SACATEPEQUEZ',
        'PASTORES':       'SACATEPEQUEZ',
        'MIXCO':          'GUATEMALA',
        'VILLA NUEVA':    'GUATEMALA',
        'VILLA CANALES':  'GUATEMALA',
        'AMATITLAN':      'GUATEMALA',
        'CHINAUTLA':      'GUATEMALA',
        'SAN JOSE PINULA':'GUATEMALA',
        'PALENCIA':       'GUATEMALA',
        'XELA':           'QUETZALTENANGO',
        'QUEZALTENANGO':  'QUETZALTENANGO',
        'MAZATENANGO':    'SUCHITEPEQUEZ',
        'ZACUALPA':       'QUICHE',
        'NEBAJ':          'QUICHE',
        'CHICHI':         'QUICHE',
        'LIVINGSTON':     'IZABAL',
        'PUERTO BARRIOS': 'IZABAL',
        'MORALES':        'IZABAL',
        'FLORES':         'PETEN',
        'POPTUN':         'PETEN',
        'SAYAXCHE':       'PETEN',
        'HUEHUE':         'HUEHUETENANGO',
        'JACALTENANGO':   'HUEHUETENANGO',
        'BARILLAS':       'HUEHUETENANGO',
        'SANTA LUCIA':    'ESCUINTLA',
        'TIQUISATE':      'ESCUINTLA',
        'ESQUIPULAS':     'CHIQUIMULA',
        'GUASTATOYA':     'EL PROGRESO',
        'MORAZAN':        'EL PROGRESO',
    }
    for keyword, depto_bd in _MUNICIPIO_A_DEPTO.items():
        if keyword in norm:
            return depto_bd
    return depto.upper()


# ─────────────────────────────────────────────
#  KEYWORDS DE CLASIFICACIÓN
# ─────────────────────────────────────────────
# Usamos patrones regex con limites de palabra (\b) porque palabras cortas
# como "shell", "uno", "total" pueden aparecer dentro de otros nombres
# legitimos (ej. "Hotel Rahshell" o "Marshall" matchearian "shell"). Sin \b
# clasificariamos hoteles como gasolineras por accidente.
#
# Separamos en LARGAS (inequívocas) y CORTAS (marcas con riesgo de falso
# positivo si OCR separa "RASHELL" en "RA SHELL" y crea un \b ficticio).
# Las cortas SOLO aplican si NO hay keyword de hospedaje en el nombre.
KEYWORDS_GASOLINA_LARGAS = [
    r"\bgasolinera\b", r"\bservicentro\b", r"\bpuma energy\b",
    r"\bcombustible\b", r"\bpetroleo\b", r"\bgas station\b",
    r"\btexgaso\b", r"\bestacion de servicio\b", r"\bcombustibles\b",
    r"\blubricantes\b", r"\bdiesel\b", r"\bgasolina\b", r"\bpetrol\b",
    r"\benergetico\b",
]
KEYWORDS_GASOLINA_CORTAS = [
    r"\btexaco\b", r"\bshell\b", r"\buno gt\b", r"\bgulf\b",
    r"\besso\b", r"\bmobil\b", r"\btotal\b",
]
KEYWORDS_GASOLINA = KEYWORDS_GASOLINA_LARGAS + KEYWORDS_GASOLINA_CORTAS

KEYWORDS_HOSPEDAJE = [
    r"\bhotel\b", r"\bhostal\b", r"\bmotel\b", r"\bhospedaje\b",
    r"\bhosped", r"\bposada\b", r"\blodge\b", r"\bresort\b",
    r"\balojamiento\b", r"\bsuite\b", r"\bapart hotel\b",
]

KEYWORDS_FOOD_SIGNAL = [
    r"\brestaur", r"\bcafeter", r"\bcomedor\b", r"\bcomida\b", r"\bpollo\b",
    r"\bpizza\b", r"\bhambur", r"\btaco\b", r"\bpanaderi", r"\bpasteler",
    r"\bgrill\b", r"\basado\b", r"\bmariscos\b", r"\bbuffet\b",
    r"\bfood\b", r"\beat\b", r"\bdiner\b", r"\bbakery\b",
    r"\bantoj", r"\bloncher", r"\bcocina\b", r"\bsteak\b", r"\bbbq\b",
]

KEYWORDS_PARQUEO = [
    r"\bparqueo\b", r"\bparqueos\b", r"\bestacionamiento\b", r"\bparking\b",
    r"\bparquimetro\b", r"\bparquímetro\b", r"\baparcamiento\b",
    r"\bparqueadero\b",
]


def _monto_efectivo(monto_str: str) -> float:
    """Convierte string de monto (ej: 'Q64.00', '64.00') a float.
    Retorna 0.0 si el string está vacío, es inválido o es cero."""
    if not monto_str:
        return 0.0
    try:
        return float(monto_str.replace('Q', '').replace(',', '').strip())
    except ValueError:
        return 0.0


def _clasificar_por_razon_social(r: dict):
    """Si no tiene tipo_comida, intenta clasificar por razon_social/establecimiento."""
    if r.get('tipo_comida'):
        return
    rs = (r.get('razon_social') or '').lower()
    est = (r.get('establecimiento') or '').lower()
    combined = f"{rs} {est}"
    # Si hay hospedaje en el nombre, NO usamos las keywords cortas de gasolina
    # (shell/uno/total/gulf/esso/mobil/texaco) porque el OCR puede separar
    # "RASHELL" en "RA SHELL" y disparar un \bshell\b ficticio. "Hotel
    # Gasolinera" no existe; las largas (gasolinera, combustible, etc.) sí
    # aplican siempre porque son inequívocas.
    tiene_hospedaje_kw = any(re.search(p, combined) for p in KEYWORDS_HOSPEDAJE)
    gas_keywords = KEYWORDS_GASOLINA_LARGAS + (
        [] if tiene_hospedaje_kw else KEYWORDS_GASOLINA_CORTAS
    )
    if any(re.search(p, combined) for p in gas_keywords):
        r['tipo_comida'] = 'Gasolina'
        return
    # Parqueo — también inequívoco y no suele mezclarse con otros tipos
    if any(re.search(p, combined) for p in KEYWORDS_PARQUEO):
        r['tipo_comida'] = 'Parqueo'
        return
    # Hospedaje SOLO si no hay señales de comida en el nombre del emisor.
    # Evita falsos positivos como "Restaurante del Hotel X" o "Cafetería Posada".
    has_food_signal = any(re.search(p, combined) for p in KEYWORDS_FOOD_SIGNAL)
    if tiene_hospedaje_kw and not has_food_signal:
        r['tipo_comida'] = 'Hospedaje'


def _enriquecer_desde_pdf(r: dict):
    """Extrae datos del nombre del PDF para registros antiguos que no tienen campos SAT."""
    if r.get('monto') or not r.get('pdf_sat'):
        return
    nombre_pdf = Path(r['pdf_sat']).stem
    m_monto = re.match(r'Q(\d+)', nombre_pdf)
    if m_monto:
        r['monto'] = f"Q{m_monto.group(1)}"
    m_fecha = re.search(r'(\d{2}\.\d{2}\.\d{4})', nombre_pdf)
    if m_fecha:
        r['fecha_emision'] = m_fecha.group(1)
    m_hora = re.search(r'\d{2}\.\d{2}\.\d{4}\s+(\d{1,2})\.(\d{2})', nombre_pdf)
    if m_hora:
        r['hora_emision'] = f"{m_hora.group(1).zfill(2)}:{m_hora.group(2)}"
    for tipo in ['Desayuno', 'Almuerzo', 'Cena', 'Gasolina', 'Parqueo']:
        if tipo in nombre_pdf:
            r['tipo_comida'] = tipo
            break
    m_est = re.search(
        r'\d{2}\.\d{2}\.\d{4}\s+\d{1,2}\.\d{2}\s+(.+?)(?:\s+(?:Desayuno|Almuerzo|Cena|Gasolina|Parqueo))?$',
        nombre_pdf)
    if m_est:
        r['establecimiento'] = m_est.group(1).strip()


def _extraer_depto_de_pdf_sat(pdf_path: Path) -> str:
    """Extrae el departamento del PDF oficial de la SAT.

    El PDF FEL tiene un bloque tipo:
        Serie: ABCD1234   Numero de DTE: 9876543210
        AVENIDA REFORMA 7-62 ZONA 5
        QUETZALTENANGO, QUETZALTENANGO            <- direccion + departamento
        Numero Acceso ...

    Tomamos la ultima coma de la direccion como nombre del departamento.
    Devuelve "" si no encuentra nada o el PDF no existe.
    """
    if not pdf_path or not pdf_path.exists():
        return ""
    try:
        import fitz as _fitz
        _pdf = _fitz.open(str(pdf_path))
        _texto = _pdf[0].get_text()
        _pdf.close()
        _lineas = [l.strip() for l in _texto.split('\n') if l.strip()]
        for _i, _l in enumerate(_lineas):
            if re.search(r'Serie:.*N[úu]mero de DTE:', _l, re.IGNORECASE):
                _dir_lineas = []
                for _dl in _lineas[_i + 1:]:
                    if re.search(r'Numero Acceso|NIT Receptor|Fecha y hora', _dl, re.IGNORECASE):
                        break
                    _dir_lineas.append(_dl)
                _dir = ' '.join(_dir_lineas)
                _partes = [p.strip() for p in _dir.split(',') if p.strip()]
                if _partes:
                    return _partes[-1].upper()
                break
    except Exception as e:
        logger.debug(f"No se pudo extraer depto desde PDF ({pdf_path.name}): {e}")
    return ""


def _resolver_departamento(fac: dict, proyecto_dir: Path, pdf_sat_dir: Path) -> str:
    """Devuelve el departamento de la factura. Estrategia:

    1. Usa fac['departamento'] si ya esta guardado en el registro.
    2. Si no, abre el PDF SAT y lo extrae del texto (mismo metodo que el Excel).

    Resultado: string en MAYUSCULAS o "" si no se pudo determinar.
    Esta funcion es la fuente unica para el departamento — TXT y Excel deben
    usarla para que el calculo de A COBRAR coincida en ambos.
    """
    depto = (fac.get('departamento') or '').strip()
    if depto:
        return depto
    pdf_path_str = fac.get('pdf_sat') or ''
    if not pdf_path_str:
        return ""
    pdf_path_obj = (proyecto_dir / pdf_path_str) if pdf_path_str else None
    if not pdf_path_obj or not pdf_path_obj.exists():
        nombre = Path(pdf_path_str).name if pdf_path_str else ''
        pdf_path_obj = pdf_sat_dir / nombre if nombre else None
    return _extraer_depto_de_pdf_sat(pdf_path_obj) if pdf_path_obj else ""
