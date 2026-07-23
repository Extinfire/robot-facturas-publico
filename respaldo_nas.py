"""
Respaldo automático al NAS - Bot Facturas SAT
Copia la carpeta 'datos/' al NAS con fecha y hora.
Mantiene los últimos 7 respaldos, borra los más antiguos.
"""

import shutil
import os
from pathlib import Path
from datetime import datetime

# ── Configuración ──────────────────────────────────────────
ORIGEN      = Path(__file__).parent / "datos"
NAS_BASE    = Path(r"\\TU-NAS\Respaldos\RobotFacturas")  # <-- cambia por la ruta de tu NAS
MAX_RESPALDOS = 7   # cuántos respaldos conservar
# ───────────────────────────────────────────────────────────

def log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}")

def hacer_respaldo():
    if not ORIGEN.exists():
        log(f"ERROR: Carpeta origen no existe: {ORIGEN}")
        return False

    if not NAS_BASE.exists():
        log(f"ERROR: No se puede acceder al NAS: {NAS_BASE}")
        log("Verifica que la ruta de red esté montada y accesible.")
        return False

    # Nombre del respaldo con fecha y hora
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    destino = NAS_BASE / f"respaldo_{timestamp}"

    log(f"Iniciando respaldo...")
    log(f"  Origen : {ORIGEN}")
    log(f"  Destino: {destino}")

    try:
        shutil.copytree(str(ORIGEN), str(destino))
        log(f"Respaldo completado exitosamente.")
    except Exception as e:
        log(f"ERROR al copiar: {e}")
        return False

    # Limpiar respaldos antiguos
    limpiar_respaldos_antiguos()
    return True

def limpiar_respaldos_antiguos():
    try:
        respaldos = sorted(
            [d for d in NAS_BASE.iterdir() if d.is_dir() and d.name.startswith("respaldo_")],
            key=lambda d: d.name
        )
        if len(respaldos) > MAX_RESPALDOS:
            a_borrar = respaldos[:len(respaldos) - MAX_RESPALDOS]
            for viejo in a_borrar:
                log(f"Borrando respaldo antiguo: {viejo.name}")
                shutil.rmtree(str(viejo))
    except Exception as e:
        log(f"Advertencia al limpiar respaldos antiguos: {e}")

if __name__ == "__main__":
    log("=== Respaldo NAS - Bot Facturas SAT ===")
    exito = hacer_respaldo()
    if exito:
        log("=== Respaldo finalizado OK ===")
    else:
        log("=== Respaldo FALLIDO ===")
        exit(1)
