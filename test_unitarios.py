#!/usr/bin/env python3
"""
Tests unitarios de LÓGICA PURA del bot (sin red, sin API, sin Telegram).

Cubre las funciones que históricamente han tenido bugs sutiles:
  - NITs excluidos (receptor: Banco Azteca, Elektra) en todos los formatos
  - Corrección de dígito verificador que no debe recrear un NIT excluido
  - Detección de Consumidor Final en el OCR
  - Clave de identidad y deduplicación de facturas
  - Variantes OCR para corrección asistida
  - Validación de match de catálogo (regla de 2+ palabras distintivas)
  - Clasificación por razón social (Hotel "Rashell" no es gasolina)
  - Thread-safety del catálogo (lectura/escritura concurrente)

Uso:
    python test_unitarios.py          # runner propio, sin dependencias
    pytest test_unitarios.py          # también funciona con pytest
"""

import sys
import threading
import tempfile
import json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from sat_invoice_bot import (
    _es_nit_excluido, SistemaExtraccion, ParserFactura, CatalogoEmpresas,
    ExtraccionFactura, NITS_EXCLUIDOS,
)
import telegram_bot as tb


# ─────────────────────────── NITs excluidos ───────────────────────────

def test_nit_excluido_todos_los_formatos():
    # Banco Azteca 5074333-3 y Elektra 657697-4 son SIEMPRE receptor
    for nit in ("5074333-3", "5074333", "50743333", "657697-4", "657697", "6576974"):
        assert _es_nit_excluido(nit), f"{nit} debería estar excluido"
    # NITs normales NO excluidos
    for nit in ("1234567-K", "90494-5", "507433-3"):
        assert not _es_nit_excluido(nit), f"{nit} NO debería estar excluido"


def test_nit_excluido_vacio():
    assert not _es_nit_excluido("")
    assert not _es_nit_excluido(None)


def test_correccion_no_recrea_nit_excluido():
    # OCR lee mal el verificador de Banco Azteca → la corrección NO debe
    # aterrizar en 5074333-3 (debe devolver None y descartar)
    assert SistemaExtraccion._corregir_nit_ocr("5074333-1") is None
    assert SistemaExtraccion._corregir_nit_ocr("5074333-X") is None
    # Un NIT normal con verificador mal SÍ se corrige
    corregido = SistemaExtraccion._corregir_nit_ocr("1234567-9")
    assert corregido and corregido.startswith("1234567-")


# ─────────────────────── Consumidor Final (OCR) ───────────────────────

def _campos(texto):
    """extraer_campos devuelve (nit, serie, numero, razon_social, receptor_cf)."""
    return ParserFactura.extraer_campos(texto)[4]

def test_consumidor_final_detectado():
    casos_cf = [
        "Emisor X NIT 1234567-K Comprobante\nCONSUMIDOR FINAL",
        "Emisor X NIT 1234567-K\nDATOS DEL COMPRADOR\nConsumidor Final",
        "Emisor X NIT 1234567-K\nDATOS DEL RECEPTOR\nNIT: CF",
        "Emisor X NIT 1234567-K\nDATOS DEL COMPRADOR\nC/F",
        "Emisor X NIT 1234567-K\nDATOS DEL CLIENTE\nNIT: 0",
    ]
    for t in casos_cf:
        assert _campos(t) is True, f"debería detectar CF: {t!r}"

def test_consumidor_final_no_falso_positivo():
    casos_no_cf = [
        "Emisor X NIT 1234567-K\nNIT RECEPTOR: 9876543-2",
        "Emisor X NIT 1234567-K\nNombre: Juan Perez NIT: 9876543-2",
        "Emisor X NIT 1234567-K Mi Empresa\nNombre Cliente: ABC\nNIT: 8765432-1",
    ]
    for t in casos_no_cf:
        assert _campos(t) is False, f"NO debería detectar CF: {t!r}"


# ──────────────────── Clave de identidad / duplicados ────────────────────

def test_clave_factura_normaliza():
    assert tb._clave_factura("1234567-k", " a3d7a47f ", "1304644738") == "1234567-K|A3D7A47F|1304644738"
    assert tb._clave_factura(None, None, None) == "||"

def test_clave_resultado_y_registro_coinciden():
    reg = {"nit": "1234567-K", "serie": "A3D7A47F", "numero": "1304644738"}
    res = ExtraccionFactura(nit="1234567-k", serie="a3d7a47f", numero_factura="1304644738")
    assert tb._clave_de_registro(reg) == tb._clave_de_resultado(res)

def test_buscar_duplicado_en_registros():
    clave = tb._clave_factura("1234567-K", "A3D7A47F", "1304644738")
    regs_ok = [{"nit": "1234567-K", "serie": "A3D7A47F", "numero": "1304644738", "monto": "Q450"}]
    # verificar_pdf=False: con monto>0 es duplicado
    assert tb._buscar_duplicado_en_registros(regs_ok, Path("."), clave, verificar_pdf=False) is not None
    # monto 0 → NO es duplicado
    regs_0 = [{"nit": "1234567-K", "serie": "A3D7A47F", "numero": "1304644738", "monto": "Q0"}]
    assert tb._buscar_duplicado_en_registros(regs_0, Path("."), clave, verificar_pdf=False) is None
    # clave distinta → NO es duplicado
    otra = tb._clave_factura("9999999-9", "BBBBBBBB", "999")
    assert tb._buscar_duplicado_en_registros(regs_ok, Path("."), otra, verificar_pdf=False) is None


# ──────────────────────── Variantes OCR ────────────────────────

def test_variantes_ocr_serie_incluye_correcta():
    # Caso real: Haiku leyó 'A3D7447F', lo correcto era 'A3D7A47F'
    variantes = tb.BotFacturasGT._generar_variantes_ocr("A3D7447F", "serie")
    assert "A3D7A47F" in variantes, f"variante correcta ausente: {variantes}"

def test_variantes_ocr_respeta_tipo_campo():
    # numero: solo dígitos en las variantes
    for v in tb.BotFacturasGT._generar_variantes_ocr("1304644738", "numero"):
        assert v.isdigit(), f"variante de numero con no-dígitos: {v}"
    # serie: solo hex
    for v in tb.BotFacturasGT._generar_variantes_ocr("A3D7447F", "serie"):
        assert all(c in "0123456789ABCDEF" for c in v), f"variante de serie no-hex: {v}"

def test_variantes_ocr_vacio():
    assert tb.BotFacturasGT._generar_variantes_ocr("", "serie") == []


# ──────────────────── Reparación hex de serie / número ────────────────────

def test_reparar_serie_hex():
    # Letras imposibles en hex → su gemelo
    assert ParserFactura.reparar_serie_hex("A3D7O47F") == "A3D7047F"   # O->0
    assert ParserFactura.reparar_serie_hex("I3D7A47F") == "13D7A47F"   # I->1
    assert ParserFactura.reparar_serie_hex("S3D7A47G") == "53D7A476"   # S->5, G->6
    # Una serie ya válida (solo hex) NO se toca
    assert ParserFactura.reparar_serie_hex("A3D7A47F") == "A3D7A47F"
    assert ParserFactura.reparar_serie_hex("63CAC965") == "63CAC965"
    # Normaliza a mayúsculas
    assert ParserFactura.reparar_serie_hex("a3d7a47f") == "A3D7A47F"
    # Vacío
    assert ParserFactura.reparar_serie_hex("") == ""

def test_reparar_serie_no_toca_hex_validos():
    # B y D son hex válidos: NO deben mapearse aunque el OCR a veces los
    # confunda en otros contextos
    for s in ("BBBBBBBB", "DDDDDDDD", "ABCDEF01"):
        assert ParserFactura.reparar_serie_hex(s) == s

def test_limpiar_numero_repara_letra_ocr():
    # Un número con una letra que el OCR confundió se repara (no se rechaza)
    assert ParserFactura.limpiar_numero_factura("130464O738") == "1304640738"  # O->0
    # Número limpio pasa igual
    assert ParserFactura.limpiar_numero_factura("1304644738") == "1304644738"
    # Basura (demasiadas letras) NO se "inventa" → None
    assert ParserFactura.limpiar_numero_factura("ABCDEFGH") is None


# ──────────────────── Validación match de catálogo ────────────────────

def test_catalogo_match_requiere_2_palabras_distintivas():
    cat = {"razon_social": "POLLO CAMPERO SOCIEDAD ANONIMA", "nit": "90494-5"}
    # Solo 'POLLO' en común → NO debe matchear (evita pegar NIT de Campero)
    assert SistemaExtraccion._catalogo_match_valido("POLLO LANDIA", cat) is False
    assert SistemaExtraccion._catalogo_match_valido("POLLO", cat) is False
    # 'POLLO' + 'CAMPERO' → sí matchea
    assert SistemaExtraccion._catalogo_match_valido("POLLO CAMPERO", cat) is True

def test_catalogo_match_una_distintiva_en_catalogo():
    # Si el catálogo solo tiene 1 palabra distintiva, basta el match fuzzy
    cat = {"razon_social": "INTELAF, S.A.", "nit": "538207-6"}
    assert SistemaExtraccion._catalogo_match_valido("INTELAF", cat) is True


# ──────────────────── Clasificación por razón social ────────────────────

def _clasif(razon, est=""):
    r = {"razon_social": razon, "establecimiento": est}
    tb._clasificar_por_razon_social(r)
    return r.get("tipo_comida")

def test_hotel_rashell_no_es_gasolina():
    # 'shell' dentro de 'Rashell' NO debe clasificar como gasolina
    assert _clasif("HOTEL CASA RASHELL") == "Hospedaje"
    assert _clasif("HOTEL CASA RA SHELL") == "Hospedaje"  # OCR separó la palabra

def test_gasolinera_real_si_es_gasolina():
    assert _clasif("GASOLINERA SHELL") == "Gasolina"
    assert _clasif("PUMA ENERGY GUATEMALA") == "Gasolina"

def test_parqueo():
    assert _clasif("PARQUEO CENTRAL") == "Parqueo"


# ──────────────────── Thread-safety del catálogo ────────────────────

def test_catalogo_concurrencia_sin_errores():
    """3 lectores + 2 escritores en paralelo no deben lanzar
    'dictionary changed size during iteration'."""
    tmp = Path(tempfile.mktemp(suffix=".json"))
    data = {
        f"empresa numero {i} sociedad anonima": {
            "razon_social": f"EMPRESA {i}", "nit": f"{1000000 + i}-1", "fuente": "sat"
        } for i in range(200)
    }
    tmp.write_text(json.dumps(data), encoding="utf-8")
    try:
        cat = CatalogoEmpresas(file_path=tmp)
        errores = []

        def lector():
            try:
                for _ in range(300):
                    cat.buscar_por_razon_social("empresa numero 50 sociedad anonima")
                    cat.listar_empresas()
            except Exception as e:
                errores.append(("lector", repr(e)))

        def escritor():
            try:
                for i in range(300):
                    cat.reconciliar_con_sat(f"NUEVA {i} SOCIEDAD ANONIMA", f"{2000000 + i}-5")
            except Exception as e:
                errores.append(("escritor", repr(e)))

        hilos = [threading.Thread(target=lector) for _ in range(3)]
        hilos += [threading.Thread(target=escritor) for _ in range(2)]
        for h in hilos:
            h.start()
        for h in hilos:
            h.join()
        assert not errores, f"errores de concurrencia: {errores}"
    finally:
        tmp.unlink(missing_ok=True)


# ─────────────────────────── Runner propio ───────────────────────────

def _ejecutar():
    funcs = [v for k, v in sorted(globals().items())
             if k.startswith("test_") and callable(v)]
    passed, failed = 0, 0
    for f in funcs:
        try:
            f()
            print(f"  PASS  {f.__name__}")
            passed += 1
        except AssertionError as e:
            print(f"  FAIL  {f.__name__}: {e}")
            failed += 1
        except Exception as e:
            print(f"  ERROR {f.__name__}: {type(e).__name__}: {e}")
            failed += 1
    print(f"\n{passed} passed, {failed} failed (de {len(funcs)} tests)")
    return failed == 0


if __name__ == "__main__":
    print("Tests unitarios de lógica pura\n" + "=" * 40)
    ok = _ejecutar()
    sys.exit(0 if ok else 1)
