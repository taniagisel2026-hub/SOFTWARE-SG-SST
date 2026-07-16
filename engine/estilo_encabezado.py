"""Estilos de encabezado seleccionables para los documentos Word y las
matrices de Excel generadas.

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

En Excel el bloque de encabezado NO está en una posición fija (varía de
columna según la matriz) -- se ubica buscando la celda que empieza con
"CÓDIGO" en las primeras filas, y se reconstruye usando las MISMAS
proporciones de Word pero como fracciones de columna, sin salir nunca de
las filas que ya ocupaba el encabezado (para no invadir la fila de
encabezados de la tabla de datos de abajo). Ver la sección "EXCEL" más
abajo para el detalle.
"""
import os
import re
import time
import unicodedata
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
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


_RE_ETIQUETA = re.compile(r"^[^:]*:\s*")


def _quitar_etiqueta(texto: str) -> str:
    """Quita el prefijo "CÓDIGO: "/"VERSIÓN: "/"FECHA: " de un valor ya
    extraido, para estilos que dibujan la etiqueta por separado (estilo 4)
    -- sin esto el valor queda duplicado ("CÓDIGO" arriba y "CÓDIGO: X" abajo)."""
    return _RE_ETIQUETA.sub("", texto, count=1).strip()


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
    _escribir(tabla.cell(1, 2), _quitar_etiqueta(c["codigo"]), tamano=7)
    _escribir(tabla.cell(1, 3), _quitar_etiqueta(c["version"]), tamano=7)

    celda_fecha_lbl = tabla.cell(2, 2)
    celda_fecha_lbl.merge(tabla.cell(2, 3))
    _escribir(celda_fecha_lbl, "FECHA", tamano=7)

    celda_fecha_val = tabla.cell(3, 2)
    celda_fecha_val.merge(tabla.cell(3, 3))
    _escribir(celda_fecha_val, _quitar_etiqueta(c["fecha"]), tamano=7)


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


# ══════════════════════════════════════════════════════════════════════
#  EXCEL (.xlsx) -- mismos 5 estilos, adaptados a la realidad de Excel:
#  el bloque de encabezado usa las MISMAS columnas que la tabla de datos
#  de abajo (a diferencia de Word, que tiene su propia tabla con anchos en
#  cm), así que las proporciones se aplican como fracciones de columna, y
#  la reconstrucción NUNCA sale de las filas que ya ocupaba el encabezado
#  original (para no invadir la fila de encabezados de la tabla de datos).
#
#  La posición del bloque tampoco es fija (varía de columna B a J según
#  la matriz) -- se localiza buscando la celda que empieza con "CÓDIGO",
#  confirmado sobre 53 archivos reales: 46 (87%) se detectan así, los
#  demás (matrices tipo consolidado/indicador con estructura distinta)
#  se omiten sin error, igual que los .docx sin tabla de encabezado.
#
#  Solo se procesa la PRIMERA hoja de cada archivo (la que trae el
#  encabezado principal).
# ══════════════════════════════════════════════════════════════════════

# Fracciones (logo, título, bloque derecho) por estilo -- calcadas de los
# anchos en cm de cada _construir_estilo_N de Word (logo/título/resto).
_FRACCIONES_XLSX = {
    "1": (0.1465, 0.5888, 0.2647),
    "2": (0.1612, 0.5588, 0.2800),
    "3": (0.1612, 0.6323, 0.2065),
    "4": (0.1906, 0.5147, 0.2947),
}

_BORDE_FINO_XLSX = Side(style="thin", color="000000")


def _normalizar_xlsx(texto: str) -> str:
    t = unicodedata.normalize("NFD", texto)
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return t.strip().upper()


def _buscar_celda_por_prefijo(ws, filas, prefijo, col_desde=None, col_hasta=None):
    for fila in filas:
        for celda in ws[fila]:
            if col_desde is not None and celda.column < col_desde:
                continue
            if col_hasta is not None and celda.column > col_hasta:
                continue
            valor = celda.value
            if isinstance(valor, str) and _normalizar_xlsx(valor).startswith(prefijo):
                return celda
    return None


def _extraer_contenido_header_xlsx(ws):
    """Ubica el bloque de encabezado buscando la celda "CÓDIGO..." en las
    primeras 8 filas (sin asumir columna fija). Devuelve None si no la
    encuentra (matriz con estructura distinta -- se omite sin error)."""
    celda_codigo = _buscar_celda_por_prefijo(ws, range(1, 9), "CODIGO")
    if celda_codigo is None:
        return None

    fila_codigo, col = celda_codigo.row, celda_codigo.column
    celda_version = _buscar_celda_por_prefijo(ws, [fila_codigo], "VERSION")
    celda_fecha = _buscar_celda_por_prefijo(ws, [fila_codigo], "FECHA")

    def _valor_o_none(fila, columna):
        if fila < 1:
            return None
        v = ws.cell(row=fila, column=columna).value
        return v.strip() if isinstance(v, str) and v.strip() else None

    titulo = _valor_o_none(fila_codigo - 1, col)
    categoria = _valor_o_none(fila_codigo - 2, col) if titulo else None
    fila_top = fila_codigo - (2 if categoria else 1 if titulo else 0)

    # Ancho del bloque: columna mas a la derecha entre las combinaciones
    # que tocan las filas del encabezado (fila_top..fila_codigo), y las
    # propias celdas de version/fecha si existen (por si no vienen
    # combinadas, esas celdas ya prueban que el encabezado llega hasta ahi).
    col_derecha = col
    if celda_version:
        col_derecha = max(col_derecha, celda_version.column)
    if celda_fecha:
        col_derecha = max(col_derecha, celda_fecha.column)
    for m in ws.merged_cells.ranges:
        if m.min_row <= fila_codigo and m.max_row >= fila_top:
            col_derecha = max(col_derecha, m.max_col)

    imagen = ws._images[0] if getattr(ws, "_images", None) else None

    return {
        "fila_top": fila_top,
        "fila_codigo": fila_codigo,
        "col_izquierda": 1,
        "col_derecha": col_derecha,
        "categoria": categoria,
        "titulo": titulo,
        "codigo": celda_codigo.value.strip(),
        "version": celda_version.value.strip() if celda_version else None,
        "fecha": celda_fecha.value.strip() if celda_fecha else None,
        "imagen": imagen,
    }


_ANCHO_COLUMNA_XLSX_POR_DEFECTO = 8.43  # ancho por defecto de Excel cuando la columna no trae uno propio


def _ancho_columna_xlsx(ws, col_idx):
    letra = get_column_letter(col_idx)
    dim = ws.column_dimensions.get(letra)
    if dim and dim.width:
        return dim.width
    return _ANCHO_COLUMNA_XLSX_POR_DEFECTO


def _asignar_columnas(ws, fracciones, col_izq, col_der):
    """Reparte el rango [col_izq, col_der] entre las fracciones dadas por
    ANCHO REAL de columna (no por cantidad de columnas) -- las columnas de
    una matriz varian mucho de ancho (ej. 6 vs 13), asi que repartir por
    cantidad dejaba textos como "CÓDIGO: SST-MTZ-001" apretados en
    columnas angostas mientras versión/fecha caian en columnas anchas.
    Devuelve una lista de (col_ini, col_fin), una tupla por fraccion."""
    anchos_col = [_ancho_columna_xlsx(ws, c) for c in range(col_izq, col_der + 1)]
    n = len(anchos_col)
    total_ancho = sum(anchos_col) or 1.0

    cortes = []
    acumulado = 0.0
    objetivo = 0.0
    idx = 0
    for frac in fracciones:
        objetivo += frac * total_ancho
        inicio = idx
        while idx < n - 1 and acumulado < objetivo:
            acumulado += anchos_col[idx]
            idx += 1
        if idx == inicio:
            idx = min(inicio + 1, n)
        cortes.append(idx)
    cortes[-1] = n

    rangos = []
    prev = 0
    for corte in cortes:
        fin = max(corte, prev + 1)
        fin = min(fin, n)
        rangos.append((prev, fin - 1))
        prev = fin
    return [(col_izq + a, col_izq + b) for a, b in rangos]


def _limpiar_bloque_xlsx(ws, fila_top, fila_codigo, col_izq, col_der):
    for m in [str(m) for m in ws.merged_cells.ranges
              if m.min_row <= fila_codigo and m.max_row >= fila_top
              and m.min_col <= col_der and m.max_col >= col_izq]:
        ws.unmerge_cells(m)
    for fila in range(fila_top, fila_codigo + 1):
        for columna in range(col_izq, col_der + 1):
            ws.cell(row=fila, column=columna).value = None


def _escribir_celda_xlsx(ws, fila_ini, col_ini, fila_fin, col_fin, texto, *, tamano=10):
    if fila_fin > fila_ini or col_fin > col_ini:
        ws.merge_cells(start_row=fila_ini, start_column=col_ini, end_row=fila_fin, end_column=col_fin)
    celda = ws.cell(row=fila_ini, column=col_ini)
    celda.value = texto
    celda.font = Font(name="Arial", bold=True, size=tamano)
    celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for fila in range(fila_ini, fila_fin + 1):
        for columna in range(col_ini, col_fin + 1):
            c = ws.cell(row=fila, column=columna)
            c.border = Border(
                top=_BORDE_FINO_XLSX if fila == fila_ini else c.border.top,
                bottom=_BORDE_FINO_XLSX if fila == fila_fin else c.border.bottom,
                left=_BORDE_FINO_XLSX if columna == col_ini else c.border.left,
                right=_BORDE_FINO_XLSX if columna == col_fin else c.border.right,
            )


_EMU_POR_PIXEL = 9525
_EMU_POR_PUNTO = 12700
_ALTO_FILA_XLSX_POR_DEFECTO = 15.0  # puntos, alto de fila por defecto de Excel
_MARGEN_LOGO_XLSX = 0.85  # deja margen dentro del bloque para que el logo no toque los bordes


def _alto_fila_xlsx(ws, fila_idx):
    dim = ws.row_dimensions.get(fila_idx)
    if dim and dim.height:
        return dim.height
    return _ALTO_FILA_XLSX_POR_DEFECTO


def _ancho_columna_xlsx_en_pixeles(ancho_caracteres):
    # Aproximacion estandar de Excel (fuente Calibri 11, Maximum Digit Width=7)
    return round(ancho_caracteres * 7 + 5)


def _tamano_bloque_xlsx_en_emu(ws, fila_ini, fila_fin, col_ini, col_fin):
    ancho_emu = sum(
        _ancho_columna_xlsx_en_pixeles(_ancho_columna_xlsx(ws, c)) * _EMU_POR_PIXEL
        for c in range(col_ini, col_fin + 1)
    )
    alto_emu = sum(_alto_fila_xlsx(ws, f) * _EMU_POR_PUNTO for f in range(fila_ini, fila_fin + 1))
    return ancho_emu, alto_emu


def _mover_logo_xlsx(ws, imagen, fila_top, fila_codigo, col_izq, col_der):
    """Reposiciona y redimensiona la imagen ya embebida (no se toca el
    archivo de imagen en si, solo su ancla/tamaño en el XML) para que
    quede centrada y llene el bloque del logo. A diferencia de Word (donde
    la celda del logo ya viene del tamaño correcto porque la trae la
    plantilla), en Excel el logo original esta pensado para el ancho de
    columna que traiga la matriz -- que varia por estilo y por matriz --
    asi que hay que escalarlo al bloque real en vez de solo reubicarlo."""
    if imagen is None:
        return
    ancla = imagen.anchor
    origen = ancla._from if hasattr(ancla, "_from") else None
    if origen is None:
        return

    ancho_bloque, alto_bloque = _tamano_bloque_xlsx_en_emu(ws, fila_top, fila_codigo, col_izq, col_der)
    ext = getattr(ancla, "ext", None)
    if ext is not None and ext.cx and ext.cy and ancho_bloque > 0 and alto_bloque > 0:
        escala = min(
            (ancho_bloque * _MARGEN_LOGO_XLSX) / ext.cx,
            (alto_bloque * _MARGEN_LOGO_XLSX) / ext.cy,
        )
        nuevo_cx = round(ext.cx * escala)
        nuevo_cy = round(ext.cy * escala)
        ext.cx, ext.cy = nuevo_cx, nuevo_cy
        origen.colOff = round(max(0, (ancho_bloque - nuevo_cx) / 2))
        origen.rowOff = round(max(0, (alto_bloque - nuevo_cy) / 2))
    else:
        origen.colOff = 0
        origen.rowOff = 0
    origen.col = col_izq - 1
    origen.row = fila_top - 1


def _construir_header_xlsx(ws, c, estilo):
    fila_top, fila_codigo = c["fila_top"], c["fila_codigo"]
    col_izq, col_der = c["col_izquierda"], c["col_derecha"]

    frac_logo, frac_titulo, frac_resto = _FRACCIONES_XLSX[estilo]
    (col_logo_ini, col_logo_fin), (col_titulo_ini, col_titulo_fin), (col_resto_ini, col_resto_fin) = (
        _asignar_columnas(ws, (frac_logo, frac_titulo, frac_resto), col_izq, col_der)
    )

    _limpiar_bloque_xlsx(ws, fila_top, fila_codigo, col_izq, col_der)

    # Logo: toda la altura del bloque
    _escribir_celda_xlsx(ws, fila_top, col_logo_ini, fila_codigo, col_logo_fin, None)
    _mover_logo_xlsx(ws, c["imagen"], fila_top, fila_codigo, col_logo_ini, col_logo_fin)

    # Titulo (categoria + titulo): todas las filas menos la de codigo
    fila_titulo_fin = max(fila_top, fila_codigo - 1)
    texto_titulo = "\n".join(t for t in (c["categoria"], c["titulo"]) if t) or ""
    _escribir_celda_xlsx(ws, fila_top, col_titulo_ini, fila_titulo_fin, col_titulo_fin, texto_titulo)

    # Fila de codigo/version/fecha: se reparte el bloque derecho (por ancho
    # real de columna) entre los que existan -- siempre hay codigo;
    # version/fecha son opcionales.
    etiquetas = [c["codigo"]]
    if c["version"]:
        etiquetas.append(c["version"])
    if c["fecha"]:
        etiquetas.append(c["fecha"])
    rangos_resto = _asignar_columnas(ws, [1 / len(etiquetas)] * len(etiquetas), col_resto_ini, col_resto_fin)
    for etiqueta, (col_ini, col_fin) in zip(etiquetas, rangos_resto):
        _escribir_celda_xlsx(ws, fila_codigo, col_ini, fila_codigo, col_fin, etiqueta, tamano=9)


def aplicar_estilo_encabezado_xlsx(xlsx_path: str, estilo: str) -> int:
    """Reconstruye el encabezado de la primera hoja de un .xlsx con el
    estilo elegido. Retorna 1 si se modificó, 0 si se omitió (estilo
    inválido, o no se encontró un bloque de encabezado reconocible)."""
    if estilo not in ESTILOS_VALIDOS:
        return 0

    try:
        wb = load_workbook(xlsx_path)
    except Exception as e:
        console.print(f"   [yellow]Aviso estilo de encabezado ({xlsx_path}): {e}[/yellow]")
        return 0

    ws = wb.worksheets[0]
    contenido = _extraer_contenido_header_xlsx(ws)
    if contenido is None:
        return 0

    try:
        _construir_header_xlsx(ws, contenido, estilo)
    except Exception as e:
        console.print(f"   [yellow]Aviso reconstruyendo encabezado ({xlsx_path}): {e}[/yellow]")
        return 0

    temp_path = xlsx_path + ".tmp_estilo_encabezado"
    try:
        wb.save(temp_path)
    except Exception as e:
        console.print(f"   [yellow]Aviso guardando estilo de encabezado ({xlsx_path}): {e}[/yellow]")
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return 0

    for intento in range(5):
        try:
            os.replace(temp_path, xlsx_path)
            break
        except PermissionError:
            if intento < 4:
                time.sleep(0.5 * (intento + 1))
            else:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                console.print(f"   [yellow]No se pudo actualizar el encabezado (archivo bloqueado): {xlsx_path}[/yellow]")
                return 0

    return 1


def aplicar_estilo_encabezado(output_base: Path, estilo: str) -> dict:
    """Aplica el estilo de encabezado a TODOS los .docx y .xlsx generados.

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

    for archivo in output_base.rglob("*.xlsx"):
        if not archivo.is_file() or archivo.name.startswith("~$"):
            continue
        cambiado = aplicar_estilo_encabezado_xlsx(str(archivo), estilo)
        if cambiado:
            stats["archivos"] += 1
        else:
            stats["omitidos"] += 1

    return stats
