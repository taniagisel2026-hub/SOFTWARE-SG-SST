"""Extractor de normas citadas en las plantillas del SG-SST.

Recorre "Plantilla SG-SST 2026" (todos los .docx y .xlsx, excepto la carpeta
NORMAS que contiene material de referencia y no plantillas generables) y
busca citas a Ley/Decreto/Resolución/Circular/Acuerdo. Arma un inventario
único con, para cada norma, en qué archivos aparece.

Este script SOLO extrae y reporta. No modifica nada ni decide si una norma
está vigente o derogada — eso requiere verificación contra fuente oficial
(ver auditar_normativa_web.py o revisión manual).

Uso:
    python auditar_normativa.py
    python auditar_normativa.py --json inventario_normas.json
"""
import json
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

RUTA_PLANTILLAS = Path("Plantilla SG-SST 2026")
CARPETAS_EXCLUIDAS = {"NORMAS"}

TIPOS_NORMA = {
    "decreto unico reglamentario": "Decreto",
    "decreto ley": "Decreto Ley",
    "decreto-ley": "Decreto Ley",
    "decreto": "Decreto",
    "ley": "Ley",
    "resolucion conjunta": "Resolución",
    "resolucion": "Resolución",
    "circular unificada": "Circular",
    "circular": "Circular",
    "acuerdo": "Acuerdo",
}

# Alternativas ordenadas de más específica a menos específica
PATRON_NORMA = re.compile(
    r"\b("
    r"Decreto\s+[UÚuú]nico\s+Reglamentario"
    r"|Decreto[\s-]+Ley"
    r"|Decreto"
    r"|Ley"
    r"|Resoluci[oó]n(?:\s+Conjunta)?"
    r"|Circular(?:\s+Unificada)?"
    r"|Acuerdo"
    r")\s+(?:No\.?\s*)?0*(\d{1,6})\s*(?:del|de)\s+(\d{4})\b",
    re.IGNORECASE,
)


def _quitar_tildes(texto: str) -> str:
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _normalizar_tipo(tipo_raw: str) -> str:
    clave = _quitar_tildes(tipo_raw.lower())
    clave = re.sub(r"\s+", " ", clave).strip()
    return TIPOS_NORMA.get(clave, tipo_raw.strip().title())


def _extraer_texto_docx(path: Path) -> str:
    try:
        doc = Document(str(path))
    except Exception as e:
        print(f"  [AVISO] No se pudo abrir {path}: {e}", file=sys.stderr)
        return ""

    partes = []
    partes.extend(p.text for p in doc.paragraphs)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                partes.extend(p.text for p in celda.paragraphs)
    for section in doc.sections:
        for elemento in (section.header, section.footer):
            partes.extend(p.text for p in elemento.paragraphs)
            for tabla in elemento.tables:
                for fila in tabla.rows:
                    for celda in fila.cells:
                        partes.extend(p.text for p in celda.paragraphs)
    return "\n".join(partes)


def _extraer_texto_xlsx(path: Path) -> str:
    try:
        wb = load_workbook(str(path), data_only=True, read_only=True)
    except Exception as e:
        print(f"  [AVISO] No se pudo abrir {path}: {e}", file=sys.stderr)
        return ""

    partes = []
    for ws in wb.worksheets:
        for fila in ws.iter_rows():
            for celda in fila:
                if isinstance(celda.value, str):
                    partes.append(celda.value)
    wb.close()
    return "\n".join(partes)


def _archivos_plantilla(base: Path):
    for path in base.rglob("*"):
        if path.name.startswith("~$"):
            continue
        if any(parte in CARPETAS_EXCLUIDAS for parte in path.relative_to(base).parts):
            continue
        if path.suffix.lower() in (".docx", ".xlsx"):
            yield path


def construir_inventario(base: Path = RUTA_PLANTILLAS) -> dict:
    """Retorna {clave_norma: {"tipo", "numero", "anio", "citas": [...], "archivos": [...]}}"""
    inventario = defaultdict(lambda: {"citas": set(), "archivos": set()})

    archivos = list(_archivos_plantilla(base))
    print(f"Escaneando {len(archivos)} archivo(s) en '{base}'...\n")

    for path in archivos:
        if path.suffix.lower() == ".docx":
            texto = _extraer_texto_docx(path)
        else:
            texto = _extraer_texto_xlsx(path)

        if not texto:
            continue

        for m in PATRON_NORMA.finditer(texto):
            tipo_raw, numero, anio = m.groups()
            tipo = _normalizar_tipo(tipo_raw)
            clave = f"{tipo} {int(numero)} de {anio}"
            entrada = inventario[clave]
            entrada["tipo"] = tipo
            entrada["numero"] = int(numero)
            entrada["anio"] = int(anio)
            entrada["citas"].add(m.group(0).strip())
            entrada["archivos"].add(str(path.relative_to(base)))

    return inventario


def imprimir_reporte(inventario: dict) -> None:
    if not inventario:
        print("No se encontraron citas de normas.")
        return

    ordenado = sorted(inventario.items(), key=lambda kv: (kv[1]["tipo"], kv[1]["anio"], kv[1]["numero"]))
    print(f"{len(ordenado)} norma(s) única(s) citadas en las plantillas:\n")
    for clave, datos in ordenado:
        n_archivos = len(datos["archivos"])
        variantes = ", ".join(sorted(datos["citas"]))
        print(f"- {clave}  ({n_archivos} archivo(s))")
        if len(datos["citas"]) > 1:
            print(f"    variantes de texto: {variantes}")


def main():
    args = sys.argv[1:]
    json_out = None
    if "--json" in args:
        idx = args.index("--json")
        json_out = args[idx + 1] if idx + 1 < len(args) else "inventario_normas.json"

    if not RUTA_PLANTILLAS.exists():
        print(f"[ERROR] No existe la carpeta de plantillas: {RUTA_PLANTILLAS}", file=sys.stderr)
        sys.exit(1)

    inventario = construir_inventario()
    imprimir_reporte(inventario)

    if json_out:
        serializable = {
            clave: {
                "tipo": datos["tipo"],
                "numero": datos["numero"],
                "anio": datos["anio"],
                "citas": sorted(datos["citas"]),
                "archivos": sorted(datos["archivos"]),
            }
            for clave, datos in inventario.items()
        }
        Path(json_out).write_text(
            json.dumps(serializable, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"\nInventario guardado en {json_out}")


if __name__ == "__main__":
    main()
