"""Tests de engine/tipo_letra.py -- cambia el tipo de letra de todos los
.docx y .xlsx generados (titulos, subtitulos y contenido a la vez, ya que
la plantilla no usa Estilos de Word separados por nivel), con la misma
tecnica de reemplazo directo en XML que engine/colores.py."""
from pathlib import Path
from zipfile import ZipFile

import openpyxl
from docx import Document
from openpyxl.styles import Font

from engine.tipo_letra import (
    aplicar_tipo_letra,
    aplicar_tipo_letra_docx,
    aplicar_tipo_letra_xlsx,
)


def _crear_docx_con_fuente(directorio: Path, nombre: str = "doc.docx") -> Path:
    doc = Document()
    doc.add_paragraph("TÍTULO DE EJEMPLO")
    doc.add_paragraph("Contenido de prueba en Calibri.")
    ruta = directorio / nombre
    doc.save(str(ruta))
    return ruta


def _crear_xlsx_con_fuente(directorio: Path, nombre: str = "matriz.xlsx") -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "MATRIZ"
    ws["A1"].font = Font(name="Arial", bold=True, size=10)
    ws["A2"] = "Contenido de prueba"
    ws["A2"].font = Font(name="Calibri", size=11)
    ruta = directorio / nombre
    wb.save(str(ruta))
    return ruta


def _leer_styles_xml(docx_path: Path) -> str:
    with ZipFile(docx_path) as z:
        return z.read("word/styles.xml").decode("utf-8")


def _leer_xlsx_styles_xml(xlsx_path: Path) -> str:
    with ZipFile(xlsx_path) as z:
        return z.read("xl/styles.xml").decode("utf-8")


def test_cambia_el_tipo_de_letra(tmp_path):
    ruta = _crear_docx_con_fuente(tmp_path)
    cambios = aplicar_tipo_letra_docx(str(ruta), "Arial")
    assert cambios > 0
    xml = _leer_styles_xml(ruta)
    assert "Calibri" not in xml
    assert 'w:ascii="Arial"' in xml


def test_fuente_invalida_no_hace_nada(tmp_path):
    ruta = _crear_docx_con_fuente(tmp_path)
    cambios = aplicar_tipo_letra_docx(str(ruta), "Comic Sans")
    assert cambios == 0


def test_aplicar_tipo_letra_recorre_la_carpeta(tmp_path):
    output_base = tmp_path / "salida"
    output_base.mkdir()
    _crear_docx_con_fuente(output_base, "uno.docx")
    _crear_docx_con_fuente(output_base, "dos.docx")

    stats = aplicar_tipo_letra(output_base, "Times New Roman")

    assert stats["archivos"] == 2
    assert stats["cambios"] > 0


def test_aplicar_tipo_letra_fuente_invalida_no_recorre(tmp_path):
    output_base = tmp_path / "salida"
    output_base.mkdir()
    _crear_docx_con_fuente(output_base)
    stats = aplicar_tipo_letra(output_base, "no-existe")
    assert stats == {"archivos": 0, "cambios": 0}


def test_cambia_el_tipo_de_letra_en_xlsx(tmp_path):
    ruta = _crear_xlsx_con_fuente(tmp_path)
    cambios = aplicar_tipo_letra_xlsx(str(ruta), "Times New Roman")
    assert cambios > 0
    xml = _leer_xlsx_styles_xml(ruta)
    assert "Arial" not in xml
    assert "Calibri" not in xml
    assert 'val="Times New Roman"' in xml


def test_xlsx_fuente_invalida_no_hace_nada(tmp_path):
    ruta = _crear_xlsx_con_fuente(tmp_path)
    cambios = aplicar_tipo_letra_xlsx(str(ruta), "Comic Sans")
    assert cambios == 0


def test_aplicar_tipo_letra_recorre_docx_y_xlsx(tmp_path):
    output_base = tmp_path / "salida"
    output_base.mkdir()
    _crear_docx_con_fuente(output_base, "uno.docx")
    _crear_xlsx_con_fuente(output_base, "matriz.xlsx")

    stats = aplicar_tipo_letra(output_base, "Arial")

    assert stats["archivos"] == 2
    assert stats["cambios"] > 0
