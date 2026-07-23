"""
Generación del reporte de gastos: Excel (con fórmulas de política T1/T2) + TXT.

Extraído de telegram_bot.py para aislar toda la lógica del reporte en un solo
lugar. No depende de Telegram ni Playwright: se puede probar de forma aislada.

Funciones públicas:
- generar_reporte_excel(...)  -> genera el .xlsx (+ .txt) o solo .txt (formato simple)
- _filtrar_facturas_para_reporte(...) -> lista de facturas del rango, deduplicada

Comparte con telegram_bot la base de comun.py y la política de clasificacion.py
para que el TXT y el Excel calculen el mismo "A COBRAR".
"""

import re
import json
import shutil
import zipfile
import logging
from pathlib import Path
from typing import Optional
from datetime import datetime

import openpyxl
from openpyxl.styles import Font as XLFont
from openpyxl.utils.datetime import to_excel as _date_to_excel

from comun import (
    carpeta_usuario,
    subcarpeta_reportes,
    leer_perfil_usuario,
    _fecha_para_orden,
    _formatear_rango_fechas,
    _clave_de_registro,
)
from clasificacion import (
    DEPTOS_TABLA1_BD,
    POLITICA_VIATICOS,
    _normalizar,
    normalizar_departamento,
    _monto_efectivo,
    _reclasificar_tipo_si_necesario,
    _clasificar_por_razon_social,
    _enriquecer_desde_pdf,
    _resolver_departamento,
)

logger = logging.getLogger(__name__)

# Plantilla del Excel y fila donde empiezan los datos en la hoja "Masivo".
PLANTILLA_REPORTE = Path(__file__).parent / "plantilla_reporte.xlsx"
MASIVO_DATA_START = 10


def _extraer_xml_extensions(xlsx_path: Path) -> dict:
    """Extrae extLst de cada hoja y los namespaces del <worksheet> raíz."""
    extensions = {}
    try:
        with zipfile.ZipFile(str(xlsx_path), 'r') as z:
            sheets = [n for n in z.namelist()
                      if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')]
            for sheet_file in sheets:
                xml = z.read(sheet_file).decode('utf-8')
                # Capturar el ÚLTIMO bloque extLst (nivel hoja, no celda)
                last_start = xml.rfind('<extLst>')
                last_end   = xml.rfind('</extLst>')
                if last_start == -1 or last_end == -1:
                    continue
                ext_block = xml[last_start: last_end + len('</extLst>')]
                # Extraer namespaces extra del <worksheet> raíz (x14, mc, etc.)
                root_m = re.search(r'<worksheet([^>]*)>', xml)
                extra_ns = ''
                if root_m:
                    attrs = root_m.group(1)
                    # Quedarse solo con xmlns:x14, xmlns:mc, etc. (no xmlns sin prefijo)
                    extra_ns = ' '.join(
                        a for a in re.findall(r'xmlns:[A-Za-z0-9_]+="[^"]*"', attrs)
                    )
                extensions[sheet_file] = (ext_block, extra_ns)
    except Exception as e:
        logger.warning(f"No se pudieron extraer extensiones XML: {e}")
    return extensions


def _reinyectar_xml_extensions(xlsx_path: Path, extensions: dict):
    """Reinyecta extLst preservado y declara los namespaces necesarios."""
    if not extensions:
        return
    tmp = xlsx_path.with_suffix('.tmp.xlsx')
    try:
        with zipfile.ZipFile(str(xlsx_path), 'r') as zin, \
             zipfile.ZipFile(str(tmp), 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in extensions:
                    xml = data.decode('utf-8')
                    ext_block, extra_ns = extensions[item.filename]
                    # Inyectar namespaces extra en el tag <worksheet ...> de openpyxl
                    # Solo los que openpyxl no declaró ya (evitar duplicados que Excel rechaza)
                    if extra_ns:
                        ws_m = re.search(r'<worksheet\b[^>]*>', xml)
                        if ws_m:
                            existing = set(re.findall(r'xmlns:[A-Za-z0-9_]+=', ws_m.group(0)))
                            needed = [a for a in extra_ns.split()
                                      if re.match(r'xmlns:[A-Za-z0-9_]+=', a)
                                      and a.split('=')[0] + '=' not in existing]
                            if needed:
                                xml = xml[:ws_m.start()] + \
                                      ws_m.group(0).replace('>', ' ' + ' '.join(needed) + '>', 1) + \
                                      xml[ws_m.end():]
                    # Reemplazar el último extLst o insertar antes de </worksheet>
                    last_start = xml.rfind('<extLst>')
                    last_end   = xml.rfind('</extLst>')
                    if last_start != -1 and last_end != -1:
                        xml = xml[:last_start] + ext_block + xml[last_end + len('</extLst>'):]
                    else:
                        xml = xml.replace('</worksheet>', ext_block + '</worksheet>')
                    data = xml.encode('utf-8')
                # item.filename (no item) para evitar arrastrar flag_bits corruptos
                zout.writestr(item.filename, data)
        shutil.move(str(tmp), str(xlsx_path))
    except Exception as e:
        logger.warning(f"No se pudieron reinyectar extensiones XML: {e}")
        if tmp.exists():
            tmp.unlink()




def generar_reporte_excel(user_id: int, username: str = "", fecha_inicio: str = "", fecha_fin: str = "", ceco: str = "", codigo_empleado: str = "", facturas_override: list = None, formato: str = "completo") -> Optional[Path]:
    """Genera reporte copiando la plantilla original y llenando la hoja Masivo.

    formato:
      - "completo" (default): Excel + TXT con columna 'A COBRAR' (política interna).
      - "simple": solo TXT sin columna 'A COBRAR' — para externos que no usan la política.
    """
    carpeta = carpeta_usuario(user_id, username)
    proyecto_dir = Path(__file__).parent
    pdf_sat_dir = carpeta / "pdf_sat"

    # Si viene lista externa (reporte editado), aplicar enriquecimiento y filtro igual que el path normal
    if facturas_override is not None:
        if not facturas_override:
            return None
        # Enriquecer por si alguna factura no tiene tipo_comida o datos del PDF
        for r in facturas_override:
            _enriquecer_desde_pdf(r)
            _clasificar_por_razon_social(r)
        # Filtrar Q0 también en el override (seguridad extra)
        facturas = [r for r in facturas_override if _monto_efectivo(r.get('monto', '')) > 0]
        if not facturas:
            return None
    else:
        registro_path = carpeta / "registro.json"
        if not registro_path.exists():
            return None
        try:
            with open(registro_path, 'r', encoding='utf-8') as f:
                registros = json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            logger.error(f"Error leyendo registro.json de usuario {user_id}: {e}")
            return None
        # Enriquecer registros antiguos desde nombre del PDF y razon social
        for r in registros:
            _enriquecer_desde_pdf(r)
            _clasificar_por_razon_social(r)
        # Filtrar facturas con datos suficientes (monto verificado por SAT)
        facturas = [r for r in registros if _monto_efectivo(r.get('monto', '')) > 0]
        if not facturas:
            return None

    # Deduplicar: quedarse con el registro más completo por NIT+Serie+Numero
    vistos = {}
    for fac in facturas:
        clave = _clave_de_registro(fac)
        existente = vistos.get(clave)
        if not existente:
            vistos[clave] = fac
        else:
            # Preferir el que tenga más datos SAT (tipo_comida, hora, etc.)
            campos_nuevos = sum(1 for k in ['tipo_comida', 'hora_emision', 'establecimiento', 'valida_sat'] if fac.get(k))
            campos_exist = sum(1 for k in ['tipo_comida', 'hora_emision', 'establecimiento', 'valida_sat'] if existente.get(k))
            if campos_nuevos > campos_exist:
                vistos[clave] = fac
    facturas = list(vistos.values())
    if not facturas:
        return None

    # Filtrar por rango de fechas
    if fecha_inicio or fecha_fin:
        def parse_fecha(f_str):
            f_str = f_str.replace('.', '/').replace('-', '/')
            try:
                return datetime.strptime(f_str, "%d/%m/%Y")
            except ValueError:
                return None

        fi = parse_fecha(fecha_inicio) if fecha_inicio else None
        ff = parse_fecha(fecha_fin) if fecha_fin else None
        filtradas = []
        for fac in facturas:
            fecha_fac = parse_fecha(fac.get('fecha_emision', '').replace('.', '/'))
            if not fecha_fac:
                # Fecha no parseable: incluir siempre para no perder facturas validas
                filtradas.append(fac)
                continue
            if fi and fecha_fac < fi:
                continue
            if ff and fecha_fac > ff:
                continue
            filtradas.append(fac)
        facturas = filtradas

    if not facturas:
        return None

    # Copiar plantilla original (con fórmulas intactas)
    # Formato de nombre: "{codigo} {nombre} {apellido} {rango_fechas}.xlsx"
    # Ej: "12345 Juan Perez 14-17 Abril.xlsx"
    perfil_nom = leer_perfil_usuario(user_id, username)
    cod_str = str(codigo_empleado or perfil_nom.get('codigo_empleado', '') or '').strip()
    nombre_str = (perfil_nom.get('nombre', '') or '').strip()
    apellido_str = (perfil_nom.get('apellido', '') or '').strip()
    rango_str = _formatear_rango_fechas(fecha_inicio, fecha_fin)
    partes_nombre = [p for p in (cod_str, nombre_str, apellido_str, rango_str) if p]
    if partes_nombre:
        nombre_base = " ".join(partes_nombre)
    else:
        # Fallback si no hay nada configurado
        nombre_base = f"Reporte_{user_id}"
    # Limpiar caracteres no válidos en nombres de archivo
    nombre_base = re.sub(r'[\\/:*?"<>|]', '', nombre_base)

    # Todos los archivos del reporte se guardan en datos/usuario_X/reportes/
    # para no ensuciar la raiz de la carpeta del usuario.
    dir_reportes = subcarpeta_reportes(carpeta)

    # ─────────────── Rama SIMPLE: solo TXT sin 'A COBRAR' ───────────────
    if formato == "simple":
        ruta_txt = dir_reportes / f"{nombre_base}.txt"
        with open(ruta_txt, 'w', encoding='utf-8') as txt:
            txt.write("# Reporte simple (sin política interna)\n")
            txt.write(
                f"{'#':<4} {'NIT EMISOR':<15} {'SERIE':<12} {'NUMERO':<15} "
                f"{'FECHA':<12} {'MONTO':<12} {'TIPO':<15} {'CONCEPTO'}\n"
            )
            txt.write("-" * 100 + "\n")
            for idx, fac in enumerate(facturas, 1):
                nit = fac.get('nit', '').replace('-', '')
                serie = fac.get('serie', '')
                numero = fac.get('numero', '')
                fecha = fac.get('fecha_emision', '').replace('.', '/')
                monto_raw = fac.get('monto', 'Q0').replace('Q', '').strip()
                tipo_comida = _reclasificar_tipo_si_necesario(fac)
                TIPO_CATEGORIA_SIMPLE = {
                    'Gasolina':  'GASOLINA',
                    'Desayuno':  'ALIMENTACION',
                    'Almuerzo':  'ALIMENTACION',
                    'Cena':      'ALIMENTACION',
                    'Hospedaje': 'HOSPEDAJE',
                    'Parqueo':   'PARQUEO',
                }
                categoria = TIPO_CATEGORIA_SIMPLE.get(tipo_comida, 'OTROS')
                concepto = tipo_comida.upper() if tipo_comida else 'OTROS'
                txt.write(
                    f"{idx:<4} {nit:<15} {serie:<12} {numero:<15} "
                    f"{fecha:<12} {monto_raw:<12} {categoria:<15} {concepto}\n"
                )
            txt.write("-" * 100 + "\n")
            txt.write(f"Total facturas: {len(facturas)}\n")
        # Devolver ruta=None para que el caller sepa que no hay Excel
        return None, ruta_txt, facturas, carpeta
    # ────────────────────────────────────────────────────────────────────

    nombre = f"{nombre_base}.xlsx"
    ruta = dir_reportes / nombre
    shutil.copy2(str(PLANTILLA_REPORTE), str(ruta))

    # Extraer elementos XML que openpyxl no preserva (x14:dataValidations, extensiones)
    _xml_extensions = _extraer_xml_extensions(PLANTILLA_REPORTE)

    wb = openpyxl.load_workbook(str(ruta))
    ws = wb['Masivo']

    # CECO en C4
    if ceco:
        try:
            ws['C4'] = int(ceco)
        except ValueError:
            ws['C4'] = ceco

    # Código de empleado en F6
    if codigo_empleado:
        try:
            ws['F6'] = int(codigo_empleado)
        except ValueError:
            ws['F6'] = codigo_empleado

    # Limpiar datos existentes en Masivo (filas 10 en adelante) — solo valores, no fórmulas de estructura
    for row in range(MASIVO_DATA_START, ws.max_row + 1):
        for col_letter in ['C', 'E', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws[f'{col_letter}{row}'].value = None

    # ── Calcular km recorridos para facturas de gasolina ──
    # Se ordenan TODAS las de gasolina del usuario (con odómetro) cronológicamente,
    # sin filtrar por rango, para poder usar la carga anterior aunque esté fuera del rango.
    # Luego asignamos km a las que están en `facturas`.
    perfil_usr = leer_perfil_usuario(user_id, username)
    odo_inicial = perfil_usr.get('odometro_inicial')
    km_por_clave = {}  # {"nit|serie|numero": km_recorridos}
    try:
        # Todas las facturas del registro (no solo el rango) para tener contexto
        reg_path = carpeta / "registro.json"
        todas_gas = []
        if reg_path.exists():
            with open(reg_path, 'r', encoding='utf-8') as _f:
                _regs = json.load(_f)
            for _r in _regs:
                _enriquecer_desde_pdf(_r)
                _clasificar_por_razon_social(_r)
                if (_r.get('tipo_comida') or '').lower() == 'gasolina' and _r.get('odometro') is not None:
                    todas_gas.append(_r)
        todas_gas.sort(key=lambda r: _fecha_para_orden(r.get('fecha_emision') or '') or datetime.min)
        # Si el usuario no configuró odometro_inicial, usamos el más antiguo (= más bajo en
        # secuencia monotónica) como punto de partida implícito. Esa primera factura queda
        # sin km calculables (no tenemos referencia previa), las demás se calculan normal.
        if odo_inicial is not None:
            odo_prev = int(odo_inicial)
        elif todas_gas:
            odo_prev = int(todas_gas[0]['odometro'])
            # La primera factura no tiene km (es el punto de arranque)
            todas_gas = todas_gas[1:]
        else:
            odo_prev = None
        for _g in todas_gas:
            _odo = int(_g['odometro'])
            _clave = _clave_de_registro(_g)
            if odo_prev is not None and _odo >= odo_prev:
                km_por_clave[_clave] = _odo - odo_prev
            odo_prev = _odo
    except Exception as e:
        logger.warning(f"No se pudieron calcular km recorridos: {e}")

    # Llenar con facturas del usuario
    TIPO_MAP = {
        'Gasolina':  ('GASOLINA',  'GASOLINA'),
        'Desayuno':  ('ALIMENTOS', 'DESAYUNO'),
        'Almuerzo':  ('ALIMENTOS', 'ALMUERZO'),
        'Cena':      ('ALIMENTOS', 'CENA'),
        'Hospedaje': ('HOSPEDAJE', 'HOSPEDAJE'),
        'Parqueo':   ('PARQUEO',   'PARQUEO'),
    }
    for idx, fac in enumerate(facturas):
        row = MASIVO_DATA_START + idx
        monto_num = _monto_efectivo(fac.get('monto', ''))
        # Reclasificar Hospedaje con monto bajo (facturas viejas guardadas
        # antes del fix del verificador). Se escribe al Excel el tipo corregido.
        tipo_efectivo = _reclasificar_tipo_si_necesario(fac)
        tipo_gasto, concepto = TIPO_MAP.get(tipo_efectivo, ('OTROS', 'OTROS'))

        fecha_raw = fac.get('fecha_emision', '') or ''
        fecha_str = fecha_raw.replace('.', '/')
        # Intentar convertir a objeto date para que Excel lo entienda como fecha real
        # Si la fecha es inválida o vacía, dejarla como string vacío (no "00/00/00")
        fecha_excel = None
        if fecha_str:
            for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                try:
                    fecha_excel = datetime.strptime(fecha_str, fmt).date()
                    break
                except ValueError:
                    continue

        # Departamento: misma fuente que usa el TXT (helper compartido).
        # Si la factura no lo tiene guardado, se extrae del PDF de la SAT.
        depto = _resolver_departamento(fac, proyecto_dir, pdf_sat_dir)

        # C = NIVEL (siempre 5)
        ws[f'C{row}'] = 5
        # E = CÓDIGO DE EMPLEADO (como número para evitar inlineStr en tabla)
        # Font color forzado a negro porque la plantilla usa theme="1" (blanco) en fontId="0"
        if codigo_empleado:
            try:
                ws[f'E{row}'] = int(codigo_empleado)
            except ValueError:
                ws[f'E{row}'] = codigo_empleado
            ws[f'E{row}'].font = XLFont(color="FF000000", size=11, name="Aptos Narrow")
        # G = FECHA — escribir como número Excel para evitar inlineStr y forzar formato DD/MM/YYYY
        if fecha_excel:
            ws[f'G{row}'] = _date_to_excel(fecha_excel)
            ws[f'G{row}'].number_format = 'DD/MM/YYYY'
        elif fecha_str:
            ws[f'G{row}'] = fecha_str
        # H = DEPARTAMENTO (valor exacto del dropdown)
        depto_norm = normalizar_departamento(depto) if depto else ''
        if depto_norm:
            ws[f'H{row}'] = depto_norm

        # Para Hospedaje: detectar número de noches según monto ÷ política T1/T2
        # y escribir "1 NOCHE" / "2 NOCHES" / "N NOCHES" en columna J.
        if tipo_efectivo == 'Hospedaje' and monto_num > 0:
            _is_t1 = depto_norm in DEPTOS_TABLA1_BD
            _rate = POLITICA_VIATICOS['Hospedaje']['T1' if _is_t1 else 'T2']
            _noches = round(monto_num / _rate) if _rate > 0 else 1
            _noches = max(1, min(_noches, 10))
            concepto = '1 NOCHE' if _noches == 1 else f'{_noches} NOCHES'

        # I = TIPO DE GASTO
        ws[f'I{row}'] = tipo_gasto
        # J = CONCEPTO
        ws[f'J{row}'] = concepto
        # K = KM'S RECORRIDOS (solo para gasolina)
        if fac.get('tipo_comida') == 'Gasolina':
            _clave_k = _clave_de_registro(fac)
            _km = km_por_clave.get(_clave_k)
            if _km is not None:
                ws[f'K{row}'] = _km
        # L = MONTO FACTURA
        ws[f'L{row}'] = monto_num

        # M = MONTO POLÍTICA  /  N = MÁXIMO REEMBOLSABLE
        # Fórmula IF directa (no VLOOKUP del BD) — funciona aunque la tabla
        # Tabla15 no esté perfectamente preservada, y se recalcula si el
        # usuario cambia manualmente el TIPO DE GASTO o CONCEPTO.
        _T1 = ','.join(f'H{row}="{d}"' for d in DEPTOS_TABLA1_BD)
        # Fórmula HOSPEDAJE: multiplica la tarifa (300 T1 / 250 T2) por las noches.
        # Las noches se leen de J ("1 NOCHE" → 1, "2 NOCHES" → 2, …).
        # Si J = "HOSPEDAJE" (sin número al inicio), IFERROR devuelve 1 → sin cambio.
        ws[f'M{row}'] = (
            f'=IF(I{row}="ALIMENTOS",'
            f'IF(OR({_T1}),'
            f'IF(J{row}="ALMUERZO",90,45),'
            f'IF(J{row}="ALMUERZO",80,40)),'
            f'IF(I{row}="HOSPEDAJE",'
            f'IF(OR({_T1}),300,250)*IFERROR(VALUE(LEFT(J{row},FIND(" ",J{row})-1)),1),'
            f'L{row}))'
        )
        ws[f'N{row}'] = f'=MIN(M{row},L{row})'

    # Quitar protección de hoja (solo el flag, sin romper el XML)
    for sheet in wb.sheetnames:
        wb[sheet].protection.sheet = False
        wb[sheet].protection.password = ''

    wb.save(str(ruta))

    # Eliminar calcChain.xml del ZIP generado por openpyxl.
    # openpyxl preserva el calcChain de la plantilla con valores cacheados obsoletos.
    # Sin calcChain, Excel recalcula TODAS las fórmulas al abrir (VLOOKUP incluido).
    try:
        _tmp_cc = ruta.with_suffix('.tmp.xlsx')
        with zipfile.ZipFile(str(ruta), 'r') as zin, \
             zipfile.ZipFile(str(_tmp_cc), 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'xl/calcChain.xml':
                    continue
                data = zin.read(item.filename)
                if item.filename == '[Content_Types].xml':
                    data = re.sub(rb'<Override[^>]*calcChain[^>]*/>\s*', b'', data)
                zout.writestr(item.filename, data)
        shutil.move(str(_tmp_cc), str(ruta))
    except Exception as e:
        logger.warning(f"No se pudo eliminar calcChain: {e}")

    # Reinyectar extensiones XML que openpyxl eliminó (dropdowns x14, etc.)
    if _xml_extensions:
        _reinyectar_xml_extensions(ruta, _xml_extensions)

    # Generar TXT con datos para copiar/pegar
    TIPO_CATEGORIA = {
        'Gasolina':  'GASOLINA',
        'Desayuno':  'ALIMENTACION',
        'Almuerzo':  'ALIMENTACION',
        'Cena':      'ALIMENTACION',
        'Hospedaje': 'HOSPEDAJE',
        'Parqueo':   'PARQUEO',
    }
    # Política de viáticos nivel 5 (según hoja BD de la plantilla)
    # Tabla 1: Guatemala + 7 departamentos con monto mayor
    # Tabla 2: resto de departamentos
    # Versión normalizada (sin tildes) de la tabla T1 para comparación robusta.
    # Se deriva de DEPTOS_TABLA1_BD — única fuente de verdad compartida con
    # la fórmula del Excel para que TXT y Excel siempre coincidan.
    DEPTOS_TABLA1_NORM = {_normalizar(d) for d in DEPTOS_TABLA1_BD}
    ruta_txt = ruta.with_suffix('.txt')
    with open(ruta_txt, 'w', encoding='utf-8') as txt:
        # Marca de versión — si no aparece, el bot está corriendo código viejo.
        txt.write("# Reporte v3 (reclasifica Hospedaje+<Q100 al vuelo por hora)\n")
        txt.write(f"{'#':<4} {'NIT EMISOR':<15} {'SERIE':<12} {'NUMERO':<15} {'FECHA':<12} {'MONTO':<12} {'A COBRAR':<12} {'TIPO':<15} {'CONCEPTO'}\n")
        txt.write("-" * 115 + "\n")
        for idx, fac in enumerate(facturas, 1):
            nit = fac.get('nit', '').replace('-', '')   # sin guión para fácil selección
            serie = fac.get('serie', '')
            numero = fac.get('numero', '')
            fecha = fac.get('fecha_emision', '').replace('.', '/')
            monto_raw = fac.get('monto', 'Q0').replace('Q', '').strip()
            # Mismo reclasificador que usa el Excel — así TXT y Excel siempre
            # muestran el mismo tipo de comida para cada factura.
            tipo_comida = _reclasificar_tipo_si_necesario(fac)
            categoria = TIPO_CATEGORIA.get(tipo_comida, 'OTROS')
            concepto = tipo_comida.upper() if tipo_comida else 'OTROS'
            # A COBRAR = mismo cálculo que la fórmula del Excel (celdas M y N).
            # Replicamos la lógica en Python porque openpyxl no evalúa fórmulas
            # y el TXT se genera sin que Excel haya abierto el archivo. Como
            # ambos caminos parten de DEPTOS_TABLA1_BD y POLITICA_VIATICOS,
            # quedan siempre sincronizados con lo que muestra el Excel.
            # Misma fuente que el Excel: si el registro no tiene departamento,
            # se extrae del PDF oficial de la SAT. Toda factura SAT trae su
            # lugar de emision, asi que el resultado siempre debe ser no vacio.
            depto_fac_raw = _resolver_departamento(fac, proyecto_dir, pdf_sat_dir)
            depto_fac_norm = _normalizar(depto_fac_raw) if depto_fac_raw else ''
            if not depto_fac_norm:
                logger.warning(
                    f"Factura sin departamento (NIT {fac.get('nit')}, "
                    f"serie {fac.get('serie')}, num {fac.get('numero')}): "
                    f"PDF SAT no disponible o sin direccion legible. "
                    f"Usando T2 (igual que la fórmula del Excel)."
                )
                tabla_key = 'T2'
            else:
                tabla_key = 'T1' if depto_fac_norm in DEPTOS_TABLA1_NORM else 'T2'
            pol = POLITICA_VIATICOS.get(tipo_comida, {})
            if pol:
                limite = pol[tabla_key]
                # Para Hospedaje multiplicar por noches (misma lógica que el Excel)
                if tipo_comida == 'Hospedaje':
                    try:
                        _monto_txt = float(monto_raw.replace(',', ''))
                        _noches_txt = round(_monto_txt / limite) if limite > 0 else 1
                        _noches_txt = max(1, min(_noches_txt, 10))
                        limite = limite * _noches_txt
                    except (ValueError, TypeError):
                        pass
                try:
                    gastado = float(monto_raw.replace(',', ''))
                    cobrar = min(gastado, limite)
                    # Mostrar sin decimales si es entero
                    cobrar_str = f"Q{int(cobrar)}" if cobrar == int(cobrar) else f"Q{cobrar:.2f}"
                except (ValueError, TypeError):
                    cobrar_str = f"Q{limite}"
            else:
                cobrar_str = monto_raw  # Gasolina/Otros: monto real sin límite
            txt.write(f"{idx:<4} {nit:<15} {serie:<12} {numero:<15} {fecha:<12} {monto_raw:<12} {cobrar_str:<12} {categoria:<15} {concepto}\n")
        txt.write("-" * 115 + "\n")
        txt.write(f"Total facturas: {len(facturas)}\n")

    return ruta, ruta_txt, facturas, carpeta




def _filtrar_facturas_para_reporte(user_id: int, username: str, fecha_inicio: str = "", fecha_fin: str = "") -> list:
    """Devuelve la lista filtrada y deduplicada de facturas para el reporte, sin generar archivos."""
    carpeta = carpeta_usuario(user_id, username)
    registro_path = carpeta / "registro.json"
    if not registro_path.exists():
        return []
    try:
        with open(registro_path, 'r', encoding='utf-8') as f:
            registros = json.load(f)
    except Exception:
        return []
    for r in registros:
        _enriquecer_desde_pdf(r)
        _clasificar_por_razon_social(r)
    facturas = [r for r in registros if _monto_efectivo(r.get('monto', '')) > 0]
    # Deduplicar
    vistos = {}
    for fac in facturas:
        clave = _clave_de_registro(fac)
        existente = vistos.get(clave)
        if not existente:
            vistos[clave] = fac
        else:
            campos_nuevos = sum(1 for k in ['tipo_comida', 'hora_emision', 'establecimiento', 'valida_sat'] if fac.get(k))
            campos_exist = sum(1 for k in ['tipo_comida', 'hora_emision', 'establecimiento', 'valida_sat'] if existente.get(k))
            if campos_nuevos > campos_exist:
                vistos[clave] = fac
    facturas = list(vistos.values())
    # Filtrar por fechas
    if fecha_inicio or fecha_fin:
        def _pf(f_str):
            try:
                return datetime.strptime(f_str.replace('.', '/').replace('-', '/'), "%d/%m/%Y")
            except ValueError:
                return None
        fi = _pf(fecha_inicio) if fecha_inicio else None
        ff = _pf(fecha_fin) if fecha_fin else None
        filtradas = []
        for fac in facturas:
            fecha_fac = _pf(fac.get('fecha_emision', ''))
            if not fecha_fac:
                filtradas.append(fac)
                continue
            if fi and fecha_fac < fi:
                continue
            if ff and fecha_fac > ff:
                continue
            filtradas.append(fac)
        facturas = filtradas

    # Ordenar por fecha y hora de emisión ascendente
    def _clave_fecha_hora(fac):
        try:
            fecha = fac.get('fecha_emision', '').replace('.', '/')
            hora = fac.get('hora_emision', '00:00') or '00:00'
            return datetime.strptime(f"{fecha} {hora}", "%d/%m/%Y %H:%M")
        except Exception:
            return datetime.min
    facturas.sort(key=_clave_fecha_hora)
    return facturas


