"""Tipo de letra (fuente) para todos los documentos Word generados.

Se configura en el YAML del cliente con:

    tipo_letra: "" | "Arial" | "Calibri" | "Times New Roman"

"" mantiene la fuente original de la plantilla (Calibri). Los documentos
de la plantilla casi no fijan la fuente por corrida (run) -- todo el texto
hereda el tipo de letra de `w:docDefaults` en styles.xml (confirmado
revisando varios documentos reales: título, subtítulos y contenido son
todos párrafo "Normal" con negrilla/tamaño puestos a mano, sin fuente
propia). Por eso reemplazar el nombre de fuente en cada `<w:rFonts>` de
document.xml/styles.xml/headers/footers cambia la letra de TODO el
documento de una sola vez, con la misma técnica de reemplazo directo en
el XML que usan engine/colores.py y engine/pie_pagina.py.
"""
import os
import re
import time
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from rich.console import Console

console = Console(highlight=False)

FUENTES_VALIDAS = ("Arial", "Calibri", "Times New Roman")

# Los 4 atributos de nombre de fuente que trae cualquier <w:rFonts> (latina,
# alta ansi, complex script, este asiatico) -- se reemplazan sin importar en
# que elemento aparezcan, igual de simple que el swap de w:jc en pie_pagina.py.
_RE_RFONTS_ATTR = re.compile(r'(w:(?:ascii|hAnsi|cs|eastAsia)=")[^"]*(")')

_PARTES_OBJETIVO_RE = re.compile(
    r"^word/(document\.xml|styles\.xml|header\d*\.xml|footer\d*\.xml)$"
)


def _reemplazar_fuente_en_xml(xml_text: str, fuente: str) -> tuple[str, int]:
    cambios = [0]

    def _procesar(m):
        cambios[0] += 1
        return f"{m.group(1)}{fuente}{m.group(2)}"

    nuevo_texto = _RE_RFONTS_ATTR.sub(_procesar, xml_text)
    return nuevo_texto, cambios[0]


def aplicar_tipo_letra_docx(docx_path: str, fuente: str) -> int:
    """Cambia el tipo de letra de un .docx. Retorna cuántos nombres de
    fuente se reemplazaron (0 si la fuente no es válida o no había nada
    que cambiar)."""
    if fuente not in FUENTES_VALIDAS:
        return 0

    temp_path = docx_path + ".tmp_tipo_letra"
    cambios_totales = [0]

    try:
        with ZipFile(docx_path, "r") as zin, ZipFile(temp_path, "w", compression=ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if _PARTES_OBJETIVO_RE.match(item.filename):
                    try:
                        text = data.decode("utf-8", errors="replace")
                        text, n = _reemplazar_fuente_en_xml(text, fuente)
                        if n:
                            cambios_totales[0] += n
                            data = text.encode("utf-8")
                    except Exception as e:
                        console.print(f"   [yellow]Aviso tipo de letra ({item.filename}): {e}[/yellow]")
                zout.writestr(item, data)

        if cambios_totales[0] == 0:
            os.remove(temp_path)
            return 0

        for intento in range(5):
            try:
                os.replace(temp_path, docx_path)
                break
            except PermissionError:
                if intento < 4:
                    time.sleep(0.5 * (intento + 1))
                else:
                    raise
    except Exception as e:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        console.print(f"   [yellow]Aviso procesando tipo de letra: {e}[/yellow]")
        return 0

    return cambios_totales[0]


def aplicar_tipo_letra(output_base: Path, fuente: str) -> dict:
    """Aplica el tipo de letra a TODOS los .docx generados.

    Args:
        output_base: Directorio raíz del SG-SST generado
        fuente: "Arial", "Calibri" o "Times New Roman"

    Returns:
        Dict {archivos, cambios}
    """
    if fuente not in FUENTES_VALIDAS:
        return {"archivos": 0, "cambios": 0}

    stats = {"archivos": 0, "cambios": 0}
    for archivo in output_base.rglob("*.docx"):
        if not archivo.is_file() or archivo.name.startswith("~$"):
            continue
        cambios = aplicar_tipo_letra_docx(str(archivo), fuente)
        if cambios > 0:
            stats["archivos"] += 1
            stats["cambios"] += cambios

    return stats
