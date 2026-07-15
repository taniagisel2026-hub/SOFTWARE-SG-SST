"""Tests de engine/estilo_encabezado.py -- reconstruye la tabla del
encabezado de los .docx generados en uno de 4 estilos, preservando el
logo, categoria/"FORMATO", titulo, codigo, version y fecha ya presentes."""
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Cm

from engine.estilo_encabezado import (
    ESTILOS_VALIDOS,
    aplicar_estilo_encabezado,
    aplicar_estilo_encabezado_docx,
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
