"""Verifica que los links del Listado Maestro (SST-MTZ-005) apunten al
documento correcto por nombre.

Para cada fila con codigo SST-XXX-NNN, compara:
  - el NOMBRE del documento esperado (col F en PHVA, col E en PESV)
  - el NOMBRE del archivo dentro del HYPERLINK (=HYPERLINK("url","nombre"))
Calcula similitud Jaccard sobre palabras y reporta:
  - OK   (>= 0.70): el link coincide con el nombre esperado
  - WARN (0.30-0.70): coincidencia parcial; revisar manualmente
  - FAIL (< 0.30 o link a documento distinto)
  - VACIO: codigo sin link asignado
  - DUPL: el codigo aparece en multiples filas con LINKS DIFERENTES (correcto si los nombres difieren)

Uso:
    python verificar_links_listado.py "clientes/<EMPRESA>"

Ejemplo:
    python verificar_links_listado.py "clientes/FUNDACIÓN_CRISTOVISIÓN"
"""
import re
import sys
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

from engine.listado_maestro import (
    HOJAS_MATRIZ, RUTA_MATRIZ_GESTION,
    PATRON_CODIGO_GENERAL, _similitud,
)

UMBRAL_OK = 0.65
UMBRAL_WARN = 0.40


def _extraer_url_y_nombre(valor) -> tuple[str, str]:
    """Si el valor es =HYPERLINK("url","nombre") devuelve (url, nombre).
    Si no, devuelve ("", str(valor or ""))."""
    if not valor:
        return "", ""
    s = str(valor)
    m = re.match(r'\s*=HYPERLINK\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)', s, re.IGNORECASE)
    if m:
        return m.group(1), m.group(2)
    return "", s


def verificar(output_dir: Path) -> int:
    matriz_path = output_dir / RUTA_MATRIZ_GESTION
    if not matriz_path.exists():
        print(f"[ERROR] No existe la matriz: {matriz_path}")
        return 1

    wb = load_workbook(str(matriz_path))

    total_general = 0
    ok_general = 0
    fallos_general = 0

    for nombre_hoja, cfg in HOJAS_MATRIZ.items():
        if nombre_hoja not in wb.sheetnames:
            continue
        ws = wb[nombre_hoja]
        col_codigo = cfg["col_codigo"]
        col_nombre = cfg["col_nombre"]
        col_link = cfg["col_link"]

        print(f"\n{'='*80}")
        print(f"HOJA: {nombre_hoja}")
        print(f"{'='*80}")

        codigo_a_links = defaultdict(list)  # para detectar duplicados
        filas = []

        for row in range(cfg["fila_inicio"], ws.max_row + 1):
            valor_codigo = ws.cell(row=row, column=col_codigo).value
            if not valor_codigo:
                continue
            m = PATRON_CODIGO_GENERAL.search(str(valor_codigo))
            if not m:
                continue
            codigo = m.group(0)
            nombre_esperado = ws.cell(row=row, column=col_nombre).value or ""
            valor_link = ws.cell(row=row, column=col_link).value
            url, nombre_link = _extraer_url_y_nombre(valor_link)

            if url:
                codigo_a_links[codigo].append(url)

            filas.append({
                "row": row, "codigo": codigo,
                "nombre_esperado": str(nombre_esperado).strip(),
                "url": url, "nombre_link": nombre_link.strip(),
            })

        # Reportar
        ok = warn = fail = vacio = 0
        col_w = 75
        for f in filas:
            estado = ""
            sim = 0.0
            if not f["url"] and not f["nombre_link"]:
                estado = "VACIO"
                vacio += 1
            else:
                sim = _similitud(f["nombre_esperado"], f["nombre_link"])
                if sim >= UMBRAL_OK:
                    estado = "OK   "
                    ok += 1
                elif sim >= UMBRAL_WARN:
                    estado = "WARN "
                    warn += 1
                else:
                    estado = "FAIL "
                    fail += 1

            # Marcar duplicados con misma URL distintos nombres
            urls_codigo = codigo_a_links[f["codigo"]]
            urls_distintas = len(set(urls_codigo))
            sufijo_dup = ""
            if len(urls_codigo) > 1 and urls_distintas == 1:
                sufijo_dup = "  [DUPL: mismo link en varias filas con codigo repetido]"

            esperado = f["nombre_esperado"][:col_w]
            obtenido = f["nombre_link"][:col_w] if f["nombre_link"] else "(sin link)"
            print(f"  {estado} R{f['row']:3d} {f['codigo']} sim={sim:.2f}")
            print(f"        ESPERADO: {esperado}")
            print(f"        EN LINK : {obtenido}{sufijo_dup}")

        print(f"\n  RESUMEN {nombre_hoja}: OK={ok} WARN={warn} FAIL={fail} VACIO={vacio} (total {len(filas)})")

        total_general += len(filas)
        ok_general += ok
        fallos_general += fail

    print(f"\n{'='*80}")
    print(f"TOTAL GENERAL: {ok_general}/{total_general} OK | {fallos_general} FAIL")
    print(f"{'='*80}")
    return 0 if fallos_general == 0 else 2


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    sys.exit(verificar(Path(sys.argv[1])))
