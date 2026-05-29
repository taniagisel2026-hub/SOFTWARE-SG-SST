#!/usr/bin/env python3
"""
Crea los tres archivos de plantilla para SST-MTZ-005 según categoría Res. 0312.
Ejecución única: python crear_plantillas_mtz005.py

Resultado en la plantilla:
  SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL [CAT-I].xlsx   (7 estándares Art. 3)
  SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL [CAT-II].xlsx  (21 estándares Art. 9)
  SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL [CAT-III].xlsx (60 estándares Art. 16)

Las filas que no aplican se ELIMINAN (no se ocultan) para que los números
de fila queden continuos y el documento se vea hecho desde cero.
"""
import sys
import shutil
import re
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

import openpyxl
from engine.matriz_documental import ESTANDARES_CAT_I, ESTANDARES_CAT_II

PLANTILLA = Path("Plantilla SG-SST 2026")
NOMBRE_BASE = "SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL"
FILA_DATOS = 7          # primera fila con datos reales (las anteriores son encabezado)

_RE_ESTANDAR = re.compile(r"^(\d+\.\d+\.\d+)")

CATEGORIAS = {
    "CAT-I":   ESTANDARES_CAT_I,
    "CAT-II":  ESTANDARES_CAT_II,
    "CAT-III": None,
}


def extraer_codigo(valor) -> str | None:
    if not valor:
        return None
    m = _RE_ESTANDAR.match(str(valor).strip())
    return m.group(1) if m else None


def crear_plantilla(origen: Path, destino: Path, estandares_validos: set | None):
    if estandares_validos is None:
        if origen != destino:
            shutil.copy2(str(origen), str(destino))
        print(f"  [CAT-III] sin cambios → {destino.name}")
        return 0

    shutil.copy2(str(origen), str(destino))

    wb = openpyxl.load_workbook(str(destino))
    ws = wb["LISTADO MAESTRO PHVA"]

    # ── 1. Determinar qué filas eliminar ─────────────────────────────────
    est_actual = None
    filas_eliminar: set[int] = set()

    for fila in range(FILA_DATOS, ws.max_row + 1):
        val_c = str(ws.cell(row=fila, column=3).value or "").strip()
        codigo = extraer_codigo(val_c)
        if codigo:
            est_actual = codigo
        if est_actual and est_actual not in estandares_validos:
            filas_eliminar.add(fila)

    if not filas_eliminar:
        wb.save(str(destino))
        return 0

    # ── 2. Guardar y deshacer celdas combinadas del área de datos ─────────
    rangos_a_rehacer: list[tuple[int, int, int, int]] = []

    for rango in list(ws.merged_cells.ranges):
        min_r, min_c = rango.min_row, rango.min_col
        max_r, max_c = rango.max_row, rango.max_col

        # Solo trabajar rangos que tocan el área de datos
        if max_r < FILA_DATOS:
            continue

        sobreviven = [r for r in range(min_r, max_r + 1) if r not in filas_eliminar]
        ws.unmerge_cells(str(rango))

        if len(sobreviven) > 1:
            # Guardar rango original — se recalcula después de la eliminación
            rangos_a_rehacer.append((min(sobreviven), min_c, max(sobreviven), max_c))

    # ── 3. Eliminar filas de abajo hacia arriba ───────────────────────────
    for fila in sorted(filas_eliminar, reverse=True):
        ws.delete_rows(fila)

    # ── 4. Re-aplicar celdas combinadas con posiciones ajustadas ─────────
    for orig_min_r, min_c, orig_max_r, max_c in rangos_a_rehacer:
        elim_antes_min = sum(1 for f in filas_eliminar if f < orig_min_r)
        elim_antes_max = sum(1 for f in filas_eliminar if f < orig_max_r)
        nuevo_min_r = orig_min_r - elim_antes_min
        nuevo_max_r = orig_max_r - elim_antes_max
        if nuevo_min_r < nuevo_max_r:
            ws.merge_cells(
                start_row=nuevo_min_r, start_column=min_c,
                end_row=nuevo_max_r, end_column=max_c,
            )

    wb.save(str(destino))
    print(f"  [{destino.name.split('[')[1].split(']')[0]}] {len(filas_eliminar)} filas eliminadas → {destino.name}")
    return len(filas_eliminar)


def main():
    base = Path("Plantilla SG-SST 2026")
    origen = next(
        (f for f in base.rglob("*.xlsx")
         if "MTZ-005" in f.name and "[CAT-III]" in f.name and "~$" not in f.name),
        None,
    )
    if not origen:
        print("ERROR: No se encontró SST-MTZ-005 [CAT-III].xlsx en la plantilla")
        return

    carpeta = origen.parent
    print(f"Origen: {origen.name}\n")

    for tag, estandares in CATEGORIAS.items():
        destino = carpeta / f"{NOMBRE_BASE} [{tag}].xlsx"
        crear_plantilla(origen, destino, estandares)

    print("\nPlantillas creadas. Verifica en Excel que los números de fila son continuos.")


if __name__ == "__main__":
    main()
