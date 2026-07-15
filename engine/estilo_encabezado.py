"""Estilos de encabezado seleccionables para los documentos Word generados.

Se configura en el YAML del cliente con:

    estilo_encabezado: "" | "1" | "2" | "3" | "4"

"" mantiene el encabezado original de la plantilla. Los 4 estilos reacomodan
los mismos datos que ya trae el encabezado actual (logo, categoría/"FORMATO",
título del documento, código, versión, fecha) en distintas grillas con
bordes -- nada se agrega ni se quita, solo se reubica.

A diferencia de colores.py/pie_pagina.py (que hacen swaps triviales de un
atributo vía regex sobre el XML crudo), aquí hay que reconstruir la tabla
del encabezado por completo, así que se usa python-docx directamente: el
header real de ~150 documentos de la plantilla es 1 tabla, 3 filas, con
forma de celdas por fila (2, 2, 4):

    fila 0: {{LOGO}} (vMerge, 3 filas) | categoría (gridSpan 3)
    fila 1: (continuación)             | título (gridSpan 3)
    fila 2: (continuación)             | código | versión | fecha

Este módulo corre DESPUÉS de que docxtpl ya renderizó el documento (mismo
punto del pipeline que pie_pagina.py, sobre las copias del cliente en
output_base) -- para entonces {{LOGO}} ya es una imagen real embebida
(InlineImage), así que el contenido de esa celda se MUEVE tal cual a la
tabla nueva en vez de reconstruirse.
"""
import os
import time
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt
from rich.console import Console

console = Console(highlight=False)

ESTILOS_VALIDOS = ("1", "2", "3", "4")


# ══════════════════════════════════════════════════════════════════════
#  HELPERS DE CONSTRUCCIÓN (mismo patrón validado con Word real en el
#  prototipo scratchpad/encabezados/construir_bocetos.py)
# ══════════════════════════════════════════════════════════════════════

def _fijar_anchos(tabla, anchos):
    """tabla.columns[i].width solo cambia la grilla, no el ancho real de
    cada celda (que es lo que Word usa para pintar) -- hay que fijarlo
    también celda por celda de cada fila."""
    for i, ancho in enumerate(anchos):
        tabla.columns[i].width = ancho
    for row in tabla.rows:
        for i, ancho in enumerate(anchos):
            row.cells[i].width = ancho


def _borde_caja(table, color="000000", sz=4):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{lado}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def _escribir(cell, texto, *, tamano=10, alineacion=WD_ALIGN_PARAGRAPH.CENTER):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = alineacion
    p.text = ""
    run = p.add_run(texto)
    run.bold = True
    run.font.size = Pt(tamano)
    run.font.name = "Arial"
    return run


def _escribir_dos_lineas(cell, linea1, linea2, *, tamano1=10, tamano2=10):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.text = ""
    r1 = p.add_run(linea1 + "\n")
    r1.bold = True
    r1.font.size = Pt(tamano1)
    r1.font.name = "Arial"
    r2 = p.add_run(linea2)
    r2.bold = True
    r2.font.size = Pt(tamano2)
    r2.font.name = "Arial"


def _mover_logo(celda_destino, logo_parrafos):
    """Mueve los <w:p> capturados del logo original (con la imagen real ya
    embebida por docxtpl) a la celda nueva, reemplazando el párrafo vacío
    por defecto que trae toda celda nueva."""
    celda_destino.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    tc = celda_destino._tc
    for p in list(celda_destino.paragraphs):
        p._p.getparent().remove(p._p)
    if not logo_parrafos:
        tc.append(OxmlElement("w:p"))
        return
    for p_element in logo_parrafos:
        tc.append(p_element)


def _limpiar_header(header):
    header.is_linked_to_previous = False
    for t in list(header.tables):
        t._element.getparent().remove(t._element)
    for p in header.paragraphs[1:]:
        p._element.getparent().remove(p._element)
    header.paragraphs[0].text = ""
    return header


# ══════════════════════════════════════════════════════════════════════
#  EXTRACCIÓN DEL CONTENIDO DEL HEADER ACTUAL
# ══════════════════════════════════════════════════════════════════════

def _extraer_contenido_header(header):
    """Extrae logo/categoría/título/código/versión/fecha del header actual.

    Devuelve None si el header no tiene tabla o si la estructura no es la
    esperada (se omite ese archivo sin lanzar error, igual que hace
    pie_pagina.py cuando no encuentra el campo de página)."""
    if not header.tables:
        return None
    tabla = header.tables[0]
    try:
        logo_cell = tabla.cell(0, 0)
        logo_parrafos = [p._p for p in logo_cell.paragraphs]
        contenido = {
            "logo_parrafos": logo_parrafos,
            "categoria": tabla.cell(0, 1).text.strip(),
            "titulo": tabla.cell(1, 1).text.strip(),
            "codigo": tabla.cell(2, 1).text.strip(),
            "version": tabla.cell(2, 2).text.strip(),
            "fecha": tabla.cell(2, 3).text.strip(),
        }
    except IndexError:
        return None
    return contenido


# ══════════════════════════════════════════════════════════════════════
#  LOS 4 ESTILOS (anchos y estructura calcados de los .docx que la
#  usuaria ajustó a mano en propuestas_encabezado/boceto_1..4.docx)
# ══════════════════════════════════════════════════════════════════════

def _construir_estilo_1(header, c):
    tabla = header.add_table(rows=3, cols=3, width=Cm(17))
    tabla.autofit = False
    _fijar_anchos(tabla, [Cm(2.49), Cm(10.01), Cm(4.50)])
    _borde_caja(tabla)

    celda_logo = tabla.cell(0, 0)
    celda_logo.merge(tabla.cell(2, 0))
    _mover_logo(celda_logo, c["logo_parrafos"])

    celda_titulo = tabla.cell(0, 1)
    celda_titulo.merge(tabla.cell(1, 1))
    _escribir_dos_lineas(celda_titulo, c["categoria"], c["titulo"])

    _escribir(tabla.cell(2, 1), c["codigo"], tamano=9)

    celda_fecha = tabla.cell(0, 2)
    celda_fecha.merge(tabla.cell(1, 2))
    _escribir(celda_fecha, c["fecha"], tamano=9)

    _escribir(tabla.cell(2, 2), c["version"], tamano=9)


def _construir_estilo_2(header, c):
    tabla = header.add_table(rows=3, cols=3, width=Cm(17))
    tabla.autofit = False
    _fijar_anchos(tabla, [Cm(2.74), Cm(9.50), Cm(4.76)])
    _borde_caja(tabla)

    celda_logo = tabla.cell(0, 0)
    celda_logo.merge(tabla.cell(2, 0))
    _mover_logo(celda_logo, c["logo_parrafos"])

    celda_titulo = tabla.cell(0, 1)
    celda_titulo.merge(tabla.cell(2, 1))
    _escribir_dos_lineas(celda_titulo, c["categoria"], c["titulo"])

    _escribir(tabla.cell(0, 2), c["codigo"], tamano=9)
    _escribir(tabla.cell(1, 2), c["version"], tamano=9)
    _escribir(tabla.cell(2, 2), c["fecha"], tamano=9)


def _construir_estilo_3(header, c):
    tabla = header.add_table(rows=2, cols=4, width=Cm(17))
    tabla.autofit = False
    _fijar_anchos(tabla, [Cm(2.74), Cm(4.21), Cm(6.54), Cm(3.51)])
    _borde_caja(tabla)

    celda_logo = tabla.cell(0, 0)
    celda_logo.merge(tabla.cell(1, 0))
    _mover_logo(celda_logo, c["logo_parrafos"])

    celda_titulo = tabla.cell(0, 1)
    celda_titulo.merge(tabla.cell(0, 2))
    _escribir_dos_lineas(celda_titulo, c["categoria"], c["titulo"])

    _escribir(tabla.cell(1, 1), c["version"], tamano=9)
    _escribir(tabla.cell(1, 2), c["fecha"], tamano=9)

    celda_codigo = tabla.cell(0, 3)
    celda_codigo.merge(tabla.cell(1, 3))
    _escribir(celda_codigo, c["codigo"], tamano=9)


def _construir_estilo_4(header, c):
    tabla = header.add_table(rows=4, cols=4, width=Cm(17))
    tabla.autofit = False
    _fijar_anchos(tabla, [Cm(3.24), Cm(8.75), Cm(2.50), Cm(2.51)])
    _borde_caja(tabla)

    celda_logo = tabla.cell(0, 0)
    celda_logo.merge(tabla.cell(3, 0))
    _mover_logo(celda_logo, c["logo_parrafos"])

    celda_titulo = tabla.cell(0, 1)
    celda_titulo.merge(tabla.cell(3, 1))
    _escribir_dos_lineas(celda_titulo, c["categoria"], c["titulo"])

    _escribir(tabla.cell(0, 2), "CÓDIGO", tamano=7)
    _escribir(tabla.cell(0, 3), "VERSIÓN", tamano=7)
    _escribir(tabla.cell(1, 2), c["codigo"], tamano=7)
    _escribir(tabla.cell(1, 3), c["version"], tamano=7)

    celda_fecha_lbl = tabla.cell(2, 2)
    celda_fecha_lbl.merge(tabla.cell(2, 3))
    _escribir(celda_fecha_lbl, "FECHA", tamano=7)

    celda_fecha_val = tabla.cell(3, 2)
    celda_fecha_val.merge(tabla.cell(3, 3))
    _escribir(celda_fecha_val, c["fecha"], tamano=7)


_CONSTRUCTORES = {
    "1": _construir_estilo_1,
    "2": _construir_estilo_2,
    "3": _construir_estilo_3,
    "4": _construir_estilo_4,
}


# ══════════════════════════════════════════════════════════════════════
#  API PÚBLICA
# ══════════════════════════════════════════════════════════════════════

def aplicar_estilo_encabezado_docx(docx_path: str, estilo: str) -> int:
    """Reconstruye el encabezado de un .docx con el estilo elegido.

    Retorna 1 si se modificó, 0 si se omitió (estilo inválido, sin tabla
    en el header, o estructura inesperada)."""
    if estilo not in ESTILOS_VALIDOS:
        return 0

    try:
        doc = Document(docx_path)
    except Exception as e:
        console.print(f"   [yellow]Aviso estilo de encabezado ({docx_path}): {e}[/yellow]")
        return 0

    modificado = False
    for section in doc.sections:
        header = section.header
        contenido = _extraer_contenido_header(header)
        if contenido is None:
            continue
        _limpiar_header(header)
        _CONSTRUCTORES[estilo](header, contenido)
        modificado = True

    if not modificado:
        return 0

    temp_path = docx_path + ".tmp_estilo_encabezado"
    try:
        doc.save(temp_path)
    except Exception as e:
        console.print(f"   [yellow]Aviso guardando estilo de encabezado ({docx_path}): {e}[/yellow]")
        if os.path.exists(temp_path):
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
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                console.print(f"   [yellow]No se pudo actualizar el encabezado (archivo bloqueado): {docx_path}[/yellow]")
                return 0

    return 1


def aplicar_estilo_encabezado(output_base: Path, estilo: str) -> dict:
    """Aplica el estilo de encabezado a TODOS los .docx generados.

    Args:
        output_base: Directorio raíz del SG-SST generado
        estilo: "1", "2", "3" o "4"

    Returns:
        Dict {archivos, omitidos}
    """
    if estilo not in ESTILOS_VALIDOS:
        return {"archivos": 0, "omitidos": 0}

    stats = {"archivos": 0, "omitidos": 0}
    for archivo in output_base.rglob("*.docx"):
        if not archivo.is_file() or archivo.name.startswith("~$"):
            continue
        cambiado = aplicar_estilo_encabezado_docx(str(archivo), estilo)
        if cambiado:
            stats["archivos"] += 1
        else:
            stats["omitidos"] += 1

    return stats
