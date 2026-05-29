#!/usr/bin/env python3
"""
Crea los tres archivos de plantilla para SST-MTZ-005 según categoría Res. 0312.
Ejecución única: python crear_plantillas_mtz005.py

Resultado en la plantilla:
  SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL [CAT-I].xlsx   (7 estándares Art. 3)
  SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL [CAT-II].xlsx  (21 estándares Art. 9)
  SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL [CAT-III].xlsx (60 estándares Art. 16)

El archivo original sin tag se elimina al final.
"""
import sys
import shutil
import re
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

import openpyxl
from engine.matriz_documental import ESTANDARES_CAT_I, ESTANDARES_CAT_II

PLANTILLA = Path("Plantilla SG-SST 2026")
CARPETA_MTZ = PLANTILLA / "I. PLANEAR" / "2.5.1 ARCHIVO Y RETENCIÓN DOCUMENTAL"
NOMBRE_BASE = "SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL"
ORIGINAL = CARPETA_MTZ / f"{NOMBRE_BASE}.xlsx"

CATEGORIAS = {
    "CAT-I":   ESTANDARES_CAT_I,
    "CAT-II":  ESTANDARES_CAT_II,
    "CAT-III": None,   # None = sin filtro, todos visibles
}

_RE_ESTANDAR = re.compile(r"^(\d+\.\d+\.\d+)")


def extraer_codigo(valor) -> str | None:
    if not valor:
        return None
    m = _RE_ESTANDAR.match(str(valor).strip())
    return m.group(1) if m else None


def crear_plantilla(origen: Path, destino: Path, estandares_validos: set | None):
    shutil.copy2(str(origen), str(destino))
    if estandares_validos is None:
        print(f"  [CAT-III] copiado sin cambios → {destino.name}")
        return

    wb = openpyxl.load_workbook(str(destino))
    ws = wb["LISTADO MAESTRO PHVA"]

    estandar_actual = None
    filas_ocultas = 0

    for fila_num in range(7, ws.max_row + 1):
        val_c = ws.cell(row=fila_num, column=3).value
        codigo = extraer_codigo(val_c)
        if codigo:
            estandar_actual = codigo

        if estandar_actual and estandar_actual not in estandares_validos:
            ws.row_dimensions[fila_num].hidden = True
            filas_ocultas += 1

    wb.save(str(destino))
    print(f"  {destino.name}: {filas_ocultas} filas ocultas")


def main():
    if not ORIGINAL.exists():
        # Buscar por glob si el nombre tiene caracteres especiales
        encontrado = next(
            (f for f in CARPETA_MTZ.glob("*.xlsx")
             if "MTZ-005" in f.name and "[CAT" not in f.name and "~$" not in f.name),
            None,
        )
        if not encontrado:
            print(f"ERROR: No se encontró {NOMBRE_BASE}.xlsx en {CARPETA_MTZ}")
            return
        origen = encontrado
    else:
        origen = ORIGINAL

    print(f"Archivo origen: {origen.name}\n")

    for tag, estandares in CATEGORIAS.items():
        destino = CARPETA_MTZ / f"{NOMBRE_BASE} [{tag}].xlsx"
        crear_plantilla(origen, destino, estandares)

    # Eliminar el archivo original sin tag
    if origen.exists():
        origen.unlink()
        print(f"\nArchivo original eliminado: {origen.name}")

    print("\nPlantillas creadas correctamente.")
    print("Recuerda hacer commit de los 3 archivos nuevos.")


if __name__ == "__main__":
    main()
