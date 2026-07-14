"""Personalización de la posición del número de página en el pie de todos
los documentos Word generados.

Se configura en el YAML del cliente con:

    pie_pagina_numero: "izquierda" | "centro" | "derecha" | "ninguno"

Por defecto la plantilla trae el número centrado -- este módulo permite
cambiarlo (o quitarlo) para TODOS los documentos de una vez, con la misma
técnica de reemplazo directo en el XML que usa engine/colores.py.
"""
import os
import re
import time
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from rich.console import Console

console = Console(highlight=False)

ALINEACIONES_VALIDAS = ("izquierda", "centro", "derecha", "ninguno")
_ALINEACION_OOXML = {"izquierda": "left", "centro": "center", "derecha": "right"}

# Captura el <w:p>...</w:p> completo que contiene el campo de numero de
# pagina (instrText PAGE) -- sin importar cuantos <w:r> tenga alrededor.
_RE_PARRAFO_CON_PAGE = re.compile(
    r"<w:p\b[^>]*>(?:(?!</w:p>).)*?<w:instrText[^>]*>\s*PAGE\b(?:(?!</w:p>).)*?</w:p>",
    re.DOTALL,
)
_RE_JC = re.compile(r'<w:jc\s+w:val="[^"]*"\s*/>')


def _reemplazar_en_footer_xml(xml_text: str, posicion: str) -> tuple[str, int]:
    cambios = [0]

    def _procesar(m):
        parrafo = m.group(0)
        if posicion == "ninguno":
            cambios[0] += 1
            return "<w:p/>"
        nuevo_jc = f'<w:jc w:val="{_ALINEACION_OOXML[posicion]}"/>'
        if _RE_JC.search(parrafo):
            nuevo = _RE_JC.sub(nuevo_jc, parrafo, count=1)
        else:
            nuevo = parrafo.replace("<w:pPr>", f"<w:pPr>{nuevo_jc}", 1)
        if nuevo != parrafo:
            cambios[0] += 1
        return nuevo

    nuevo_texto = _RE_PARRAFO_CON_PAGE.sub(_procesar, xml_text)
    return nuevo_texto, cambios[0]


def aplicar_posicion_pagina_docx(docx_path: str, posicion: str) -> int:
    """Cambia la alineación del número de página en el pie de un .docx (o lo
    quita si posicion == "ninguno"). Retorna cuántos pies se modificaron."""
    if posicion not in ALINEACIONES_VALIDAS:
        return 0

    temp_path = docx_path + ".tmp_pie_pagina"
    cambios_totales = [0]

    try:
        with ZipFile(docx_path, "r") as zin, ZipFile(temp_path, "w", compression=ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("word/footer") and item.filename.endswith(".xml"):
                    try:
                        text = data.decode("utf-8", errors="replace")
                        text, n = _reemplazar_en_footer_xml(text, posicion)
                        if n:
                            cambios_totales[0] += n
                            data = text.encode("utf-8")
                    except Exception as e:
                        console.print(f"   [yellow]Aviso pie de pagina ({item.filename}): {e}[/yellow]")
                zout.writestr(item, data)

        if cambios_totales[0] == 0:
            os.remove(temp_path)
            return 0

        # Reintentos para tolerar OneDrive sincronizando en ese momento
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
        console.print(f"   [yellow]Aviso procesando pie de pagina: {e}[/yellow]")
        return 0

    return cambios_totales[0]


def aplicar_posicion_pagina(output_base: Path, posicion: str) -> dict:
    """Aplica la posición del número de página a TODOS los .docx generados.

    Args:
        output_base: Directorio raíz del SG-SST generado
        posicion: "izquierda", "centro", "derecha" o "ninguno"

    Returns:
        Dict {archivos, cambios}
    """
    if posicion not in ALINEACIONES_VALIDAS:
        return {"archivos": 0, "cambios": 0}

    stats = {"archivos": 0, "cambios": 0}
    for archivo in output_base.rglob("*.docx"):
        if not archivo.is_file() or archivo.name.startswith("~$"):
            continue
        cambios = aplicar_posicion_pagina_docx(str(archivo), posicion)
        if cambios > 0:
            stats["archivos"] += 1
            stats["cambios"] += cambios

    return stats
