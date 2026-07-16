"""Tests de engine/estilo_encabezado.py -- reconstruye la tabla del
encabezado de los .docx generados en uno de 4 estilos, preservando el
logo, categoria/"FORMATO", titulo, codigo, version y fecha ya presentes."""
import io
from pathlib import Path

import openpyxl
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Cm
from openpyxl.drawing.image import Image as XlsxImage
from PIL import Image as PILImage

from engine.estilo_encabezado import (
    ESTILOS_VALIDOS,
    aplicar_estilo_encabezado,
    aplicar_estilo_encabezado_docx,
    aplicar_estilo_encabezado_xlsx,
)

# Marcador distintivo que hace las veces de "imagen real" del logo (docxtpl
# ya la habria reemplazado por un InlineImage para cuando este modulo corre;
# aca solo nos interesa verificar que el contenido de esa celda sobrevive
# tal cual a la reconstruccion, sea imagen o texto).
_LOGO_STANDIN = "__LOGO_IMG__"


def _crear_docx_con_header_real(directorio: Path, nombre: str = "doc.docx", *,
                                 categoria: str = "FORMATO",
                                 titulo: str = "ASIGNACIÓN DEL RESPONSABLE DEL SG-SST",
                                 codigo: str = "CÓDIGO: SST-FMT-001",
                                 version: str = "VERSIÓN: {{VERSION}}",
                                 fecha: str = "FECHA: {{FECHA_EMISION}}") -> Path:
    """Construye un .docx cuyo header replica la estructura real confirmada
    en la plantilla: 1 tabla, 3 filas, forma de celdas por fila (2, 2, 4)."""
    doc = Document()
    doc.add_paragraph("Contenido de prueba")

    header = doc.sections[0].header
    header.is_linked_to_previous = False
    tabla = header.add_table(rows=3, cols=4, width=Cm(17))

    celda_logo = tabla.cell(0, 0)
    celda_logo.merge(tabla.cell(2, 0))
    celda_logo.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    celda_logo.paragraphs[0].text = _LOGO_STANDIN

    celda_categoria = tabla.cell(0, 1)
    celda_categoria.merge(tabla.cell(0, 3))
    celda_categoria.paragraphs[0].text = categoria

    celda_titulo = tabla.cell(1, 1)
    celda_titulo.merge(tabla.cell(1, 3))
    celda_titulo.paragraphs[0].text = titulo

    tabla.cell(2, 1).paragraphs[0].text = codigo
    tabla.cell(2, 2).paragraphs[0].text = version
    tabla.cell(2, 3).paragraphs[0].text = fecha

    ruta = directorio / nombre
    doc.save(str(ruta))
    return ruta


def _texto_completo_header(ruta: Path) -> str:
    doc = Document(str(ruta))
    header = doc.sections[0].header
    assert header.tables, "se esperaba una tabla en el header reconstruido"
    return "\n".join(
        p.text for t in header.tables for row in t.rows for c in row.cells for p in c.paragraphs
    )


def _header_tabla(ruta: Path):
    doc = Document(str(ruta))
    header = doc.sections[0].header
    assert header.tables, "se esperaba una tabla en el header reconstruido"
    return header.tables[0]


def test_estilos_validos_son_1_a_4():
    assert ESTILOS_VALIDOS == ("1", "2", "3", "4")


def test_aplica_estilo_1_reordena_contenido_y_preserva_todo(tmp_path):
    ruta = _crear_docx_con_header_real(tmp_path)
    cambiado = aplicar_estilo_encabezado_docx(str(ruta), "1")
    assert cambiado == 1

    texto = _texto_completo_header(ruta)
    assert _LOGO_STANDIN in texto
    assert "FORMATO" in texto
    assert "ASIGNACIÓN DEL RESPONSABLE DEL SG-SST" in texto
    assert "CÓDIGO: SST-FMT-001" in texto
    assert "VERSIÓN: {{VERSION}}" in texto
    assert "FECHA: {{FECHA_EMISION}}" in texto


def test_aplica_estilo_2_reordena_contenido_y_preserva_todo(tmp_path):
    ruta = _crear_docx_con_header_real(tmp_path)
    cambiado = aplicar_estilo_encabezado_docx(str(ruta), "2")
    assert cambiado == 1
    texto = _texto_completo_header(ruta)
    for esperado in (_LOGO_STANDIN, "FORMATO", "ASIGNACIÓN DEL RESPONSABLE DEL SG-SST",
                     "CÓDIGO: SST-FMT-001", "VERSIÓN: {{VERSION}}", "FECHA: {{FECHA_EMISION}}"):
        assert esperado in texto


def test_aplica_estilo_3_reordena_contenido_y_preserva_todo(tmp_path):
    ruta = _crear_docx_con_header_real(tmp_path)
    cambiado = aplicar_estilo_encabezado_docx(str(ruta), "3")
    assert cambiado == 1
    texto = _texto_completo_header(ruta)
    for esperado in (_LOGO_STANDIN, "FORMATO", "ASIGNACIÓN DEL RESPONSABLE DEL SG-SST",
                     "CÓDIGO: SST-FMT-001", "VERSIÓN: {{VERSION}}", "FECHA: {{FECHA_EMISION}}"):
        assert esperado in texto


def test_aplica_estilo_4_reordena_contenido_y_preserva_todo(tmp_path):
    ruta = _crear_docx_con_header_real(tmp_path)
    cambiado = aplicar_estilo_encabezado_docx(str(ruta), "4")
    assert cambiado == 1
    texto = _texto_completo_header(ruta)
    assert _LOGO_STANDIN in texto
    assert "FORMATO" in texto
    assert "ASIGNACIÓN DEL RESPONSABLE DEL SG-SST" in texto
    assert "SST-FMT-001" in texto  # estilo 4 separa la etiqueta "CÓDIGO" del valor
    assert "{{VERSION}}" in texto
    assert "{{FECHA_EMISION}}" in texto


def test_estilo_4_no_duplica_la_etiqueta_con_el_valor(tmp_path):
    """Bug real: estilo 4 dibuja "CÓDIGO"/"VERSIÓN"/"FECHA" como etiqueta
    aparte, pero el valor extraido YA trae el prefijo ("CÓDIGO: SST-FMT-001")
    -- sin quitarlo, el resultado quedaba "CÓDIGO" seguido de
    "CÓDIGO: SST-FMT-001" duplicado y con la letra corrida."""
    ruta = _crear_docx_con_header_real(tmp_path)
    aplicar_estilo_encabezado_docx(str(ruta), "4")
    tabla = _header_tabla(ruta)
    textos_celdas = [c.text for row in tabla.rows for c in row.cells]
    assert "CÓDIGO: SST-FMT-001" not in textos_celdas
    assert "VERSIÓN: {{VERSION}}" not in textos_celdas
    assert "FECHA: {{FECHA_EMISION}}" not in textos_celdas
    assert "SST-FMT-001" in textos_celdas
    assert "{{VERSION}}" in textos_celdas
    assert "{{FECHA_EMISION}}" in textos_celdas


def test_categoria_no_literal_formato_se_preserva(tmp_path):
    """La categoria real varia por tipo de documento (POLÍTICA, PROCEDIMIENTO,
    MANUAL, PLAN, REGLAMENTO) -- no debe hardcodearse como "FORMATO"."""
    ruta = _crear_docx_con_header_real(tmp_path, categoria="POLÍTICA",
                                        titulo="POLÍTICA DE SEGURIDAD Y SALUD EN EL TRABAJO")
    aplicar_estilo_encabezado_docx(str(ruta), "2")
    texto = _texto_completo_header(ruta)
    assert "POLÍTICA" in texto
    assert "FORMATO" not in texto


def test_estilo_invalido_no_hace_nada(tmp_path):
    ruta = _crear_docx_con_header_real(tmp_path)
    cambiado = aplicar_estilo_encabezado_docx(str(ruta), "9")
    assert cambiado == 0


def test_docx_sin_tabla_en_header_no_truena(tmp_path):
    doc = Document()
    doc.add_paragraph("Sin tabla en el header")
    ruta = tmp_path / "sin_tabla.docx"
    doc.save(str(ruta))
    cambiado = aplicar_estilo_encabezado_docx(str(ruta), "1")
    assert cambiado == 0


def test_aplicar_estilo_encabezado_recorre_la_carpeta(tmp_path):
    output_base = tmp_path / "salida"
    output_base.mkdir()
    _crear_docx_con_header_real(output_base, "uno.docx")
    _crear_docx_con_header_real(output_base, "dos.docx")
    (output_base / "sin_header.docx").write_bytes(b"")  # archivo invalido, no debe tronar

    stats = aplicar_estilo_encabezado(output_base, "2")

    assert stats["archivos"] == 2
    assert stats["omitidos"] == 1


def test_aplicar_estilo_encabezado_estilo_invalido_no_recorre(tmp_path):
    output_base = tmp_path / "salida"
    output_base.mkdir()
    _crear_docx_con_header_real(output_base)
    stats = aplicar_estilo_encabezado(output_base, "no-existe")
    assert stats == {"archivos": 0, "omitidos": 0}


# ══════════════════════════════════════════════════════════════════════
#  EXCEL -- misma logica pero el bloque de encabezado se ubica buscando
#  la celda "CODIGO..." (columna variable), no por coordenada fija.
# ══════════════════════════════════════════════════════════════════════

def _imagen_xlsx_de_prueba():
    buf = io.BytesIO()
    PILImage.new("RGB", (20, 20), color="green").save(buf, format="PNG")
    buf.seek(0)
    return XlsxImage(buf)


def _crear_xlsx_con_header_real(directorio: Path, nombre: str = "matriz.xlsx", *,
                                 col_codigo: str = "F",
                                 categoria: str = "MATRIZ",
                                 titulo: str = "ROLES Y RESPONSABILIDADES",
                                 codigo: str = "CÓDIGO: SST-MTZ-001",
                                 version: str = "VERSIÓN: 002",
                                 fecha: str = "FECHA: 20/01/2025") -> Path:
    """Replica la estructura real confirmada sobre 53 archivos: categoria/
    titulo/codigo apilados en una columna (variable), version y fecha en
    la misma fila del codigo, mas una imagen (el logo)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws[f"{col_codigo}1"] = categoria
    ws[f"{col_codigo}2"] = titulo
    ws[f"{col_codigo}3"] = codigo
    col_version = chr(ord(col_codigo) + 3)
    col_fecha = chr(ord(col_codigo) + 6)
    ws[f"{col_version}3"] = version
    ws[f"{col_fecha}3"] = fecha
    ws.add_image(_imagen_xlsx_de_prueba(), "A1")
    # Filas de "datos" despues del encabezado, para verificar que no se tocan
    ws["A5"] = "CARGO"
    ws["B5"] = "PROCESO"
    ruta = directorio / nombre
    wb.save(str(ruta))
    return ruta


def _leer_valores_xlsx(ruta: Path):
    wb = openpyxl.load_workbook(str(ruta))
    ws = wb.worksheets[0]
    return [c.value for row in ws.iter_rows(min_row=1, max_row=6) for c in row if c.value not in (None, "")]


def test_aplica_estilo_1_xlsx_columna_f(tmp_path):
    ruta = _crear_xlsx_con_header_real(tmp_path, col_codigo="F")
    cambiado = aplicar_estilo_encabezado_xlsx(str(ruta), "1")
    assert cambiado == 1
    valores = _leer_valores_xlsx(ruta)
    texto = "\n".join(valores)
    assert "MATRIZ" in texto
    assert "ROLES Y RESPONSABILIDADES" in texto
    assert "CÓDIGO: SST-MTZ-001" in texto
    assert "VERSIÓN: 002" in texto
    assert "FECHA: 20/01/2025" in texto


def test_aplica_estilo_2_xlsx_columna_distinta_b(tmp_path):
    """Misma prueba pero con el bloque empezando en otra columna (B en vez
    de F) -- confirma que la ubicacion es por busqueda de texto, no por
    coordenada fija."""
    ruta = _crear_xlsx_con_header_real(tmp_path, col_codigo="B")
    cambiado = aplicar_estilo_encabezado_xlsx(str(ruta), "2")
    assert cambiado == 1
    valores = _leer_valores_xlsx(ruta)
    texto = "\n".join(valores)
    assert "MATRIZ" in texto
    assert "ROLES Y RESPONSABILIDADES" in texto
    assert "CÓDIGO: SST-MTZ-001" in texto


def test_xlsx_no_toca_las_filas_de_datos(tmp_path):
    ruta = _crear_xlsx_con_header_real(tmp_path)
    aplicar_estilo_encabezado_xlsx(str(ruta), "4")
    wb = openpyxl.load_workbook(str(ruta))
    ws = wb.worksheets[0]
    assert ws["A5"].value == "CARGO"
    assert ws["B5"].value == "PROCESO"


def test_xlsx_estilo_invalido_no_hace_nada(tmp_path):
    ruta = _crear_xlsx_con_header_real(tmp_path)
    cambiado = aplicar_estilo_encabezado_xlsx(str(ruta), "9")
    assert cambiado == 0


def test_xlsx_sin_codigo_detectable_no_truena(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["C1"] = "MATRIZ DE REQUISITOS LEGALES"  # categoria+titulo juntos, sin "codigo" separado
    ruta = tmp_path / "sin_codigo.xlsx"
    wb.save(str(ruta))
    cambiado = aplicar_estilo_encabezado_xlsx(str(ruta), "1")
    assert cambiado == 0


def test_xlsx_columna_de_datos_no_se_confunde_con_encabezado(tmp_path):
    """Regresion: una columna de datos que empieza con "codigo" (ej. una
    columna "CÓDIGO CIE 10" en una tabla de diagnosticos) no trae los dos
    puntos que si trae el bloque de encabezado real ("CÓDIGO: SST-..."),
    asi que no debe tratarse como un bloque de encabezado."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["F1"] = "MATRIZ"
    ws["F2"] = "AUSENTISMO"
    ws["F3"] = "CÓDIGO: SST-MTZ-011"
    ws["O4"] = "CÓDIGO CIE 10"  # encabezado de columna, no un bloque nuevo
    ruta = tmp_path / "ausentismo.xlsx"
    wb.save(str(ruta))
    cambiado = aplicar_estilo_encabezado_xlsx(str(ruta), "1")
    assert cambiado == 1
    wb2 = openpyxl.load_workbook(str(ruta))
    ws2 = wb2.worksheets[0]
    # La columna de datos debe seguir intacta, no reescrita como encabezado
    assert ws2["O4"].value == "CÓDIGO CIE 10"


def test_xlsx_reestila_todos_los_bloques_repetidos(tmp_path):
    """Regresion: SST-MTZ-002 (y otras matrices reales) repiten el bloque
    de encabezado completo varias veces en la misma hoja -- una por
    seccion. Antes solo se ubicaba el primero (busqueda limitada a las
    primeras 8 filas); ahora debe encontrar y restylar todos, cada uno con
    su propio logo."""
    wb = openpyxl.Workbook()
    ws = wb.active
    # Bloque 1
    ws["D1"] = "MATRIZ"
    ws["D2"] = "RECURSOS Y PRESUPUESTO PARA EL SG-SST"
    ws["H3"] = "CÓDIGO: SST-MTZ-002"
    ws["I3"] = "VERSIÓN: 001"
    ws["J3"] = "FECHA: 16/07/2026"
    ws.add_image(_imagen_xlsx_de_prueba(), "A1")
    ws["A5"] = "FINANCIEROS"
    ws["C5"] = "Asesoría Medina SST"
    # Bloque 2, mas abajo -- fuera del rango de busqueda anterior (filas 1-8)
    ws["C13"] = "MATRIZ"
    ws["C14"] = "RECURSOS Y PRESUPUESTO PARA EL SG-SST"
    ws["C15"] = "CÓDIGO: SST-MTZ-002"
    ws["E15"] = "VERSIÓN: 001"
    ws["G15"] = "FECHA: 16/07/2026"
    ws.add_image(_imagen_xlsx_de_prueba(), "A13")
    ws["A16"] = "HUMANOS"
    ws["C16"] = "Responsable SST"
    ruta = tmp_path / "mtz002.xlsx"
    wb.save(str(ruta))

    cambiado = aplicar_estilo_encabezado_xlsx(str(ruta), "2")
    assert cambiado == 1

    wb2 = openpyxl.load_workbook(str(ruta))
    ws2 = wb2.worksheets[0]
    valores = [c.value for row in ws2.iter_rows() for c in row if c.value not in (None, "")]
    texto = "\n".join(str(v) for v in valores)
    # Ambos bloques deben seguir teniendo su contenido, reordenado
    assert texto.count("MATRIZ") == 2
    assert texto.count("CÓDIGO: SST-MTZ-002") == 2
    assert "FINANCIEROS" in texto and "Asesoría Medina SST" in texto
    assert "HUMANOS" in texto and "Responsable SST" in texto
    # Cada bloque debe conservar su propia imagen, ninguna se pierde ni se
    # reutiliza en el otro bloque
    assert len(ws2._images) == 2
    filas_logo = sorted(img.anchor._from.row for img in ws2._images)
    assert filas_logo[0] != filas_logo[1]


def test_xlsx_logo_se_centra_y_se_escala_al_bloque(tmp_path):
    """Regresion: antes, el logo solo se reubicaba (ancla en la esquina
    superior izquierda del bloque, colOff/rowOff=0) sin tocar su tamano
    original -- si ese tamano no correspondia al bloque real (que varia por
    matriz/estilo), el logo quedaba desproporcionado y pegado a una esquina
    en vez de centrado. Ahora debe escalarse al bloque y centrarse."""
    ruta = _crear_xlsx_con_header_real(tmp_path, col_codigo="F")
    aplicar_estilo_encabezado_xlsx(str(ruta), "1")
    wb = openpyxl.load_workbook(str(ruta))
    ws = wb.worksheets[0]
    img = ws._images[0]
    ancla = img.anchor
    origen = ancla._from

    # La imagen de prueba es cuadrada de 20x20px (190500 EMU) -- si el
    # bloque del logo no es cuadrado, el tamano final no puede seguir
    # siendo el original sin escalar.
    assert (ancla.ext.cx, ancla.ext.cy) != (190500, 190500)
    assert ancla.ext.cx > 0 and ancla.ext.cy > 0
    # Debe quedar centrada dentro del bloque (offset > 0 en al menos una
    # dimension), no pegada a la esquina superior izquierda como antes.
    assert origen.colOff > 0 or origen.rowOff > 0


def test_aplicar_estilo_encabezado_recorre_xlsx_tambien(tmp_path):
    output_base = tmp_path / "salida"
    output_base.mkdir()
    _crear_xlsx_con_header_real(output_base, "uno.xlsx", col_codigo="F")
    _crear_xlsx_con_header_real(output_base, "dos.xlsx", col_codigo="C")
    _crear_docx_con_header_real(output_base, "tres.docx")

    stats = aplicar_estilo_encabezado(output_base, "3")

    assert stats["archivos"] == 3
