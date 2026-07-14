"""Tests de engine/pie_pagina.py -- cambia la alineacion del numero de
pagina en el pie de los .docx generados (o lo quita), con la misma tecnica
de reemplazo directo en XML que engine/colores.py."""
import shutil
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from engine.pie_pagina import aplicar_posicion_pagina, aplicar_posicion_pagina_docx


def _inyectar_campo_pagina(docx_path: Path):
    """Agrega un campo PAGE al primer parrafo de word/footer1.xml, tal como
    lo produce Word real (visto en la plantilla real del proyecto)."""
    tmp = str(docx_path) + ".tmp"
    with ZipFile(docx_path, "r") as zin, ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/footer1.xml":
                xml = data.decode("utf-8")
                campo = (
                    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                    '<w:r><w:instrText>PAGE</w:instrText></w:r>'
                    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                    '<w:r><w:t>1</w:t></w:r>'
                    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
                )
                xml = xml.replace("</w:p>", campo + "</w:p>", 1)
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, docx_path)


def _crear_docx_con_pagina(directorio: Path, nombre: str = "doc.docx") -> Path:
    doc = Document()
    doc.add_paragraph("Contenido de prueba")
    doc.sections[0].footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    ruta = directorio / nombre
    doc.save(str(ruta))
    _inyectar_campo_pagina(ruta)
    return ruta


def _leer_footer_xml(docx_path: Path) -> str:
    with ZipFile(docx_path) as z:
        return z.read("word/footer1.xml").decode("utf-8")


def test_cambia_alineacion_del_numero_de_pagina(tmp_path):
    ruta = _crear_docx_con_pagina(tmp_path)
    cambios = aplicar_posicion_pagina_docx(str(ruta), "izquierda")
    assert cambios == 1
    xml = _leer_footer_xml(ruta)
    assert 'w:jc w:val="left"' in xml
    assert "PAGE" in xml  # el campo de numero de pagina se conserva


def test_quita_el_numero_de_pagina(tmp_path):
    ruta = _crear_docx_con_pagina(tmp_path)
    cambios = aplicar_posicion_pagina_docx(str(ruta), "ninguno")
    assert cambios == 1
    xml = _leer_footer_xml(ruta)
    assert "PAGE" not in xml


def test_posicion_invalida_no_hace_nada(tmp_path):
    ruta = _crear_docx_con_pagina(tmp_path)
    cambios = aplicar_posicion_pagina_docx(str(ruta), "arriba")
    assert cambios == 0


def test_docx_sin_footer_con_pagina_no_truena(tmp_path):
    doc = Document()
    doc.add_paragraph("Sin campo de pagina en el pie")
    ruta = tmp_path / "sin_pagina.docx"
    doc.save(str(ruta))
    cambios = aplicar_posicion_pagina_docx(str(ruta), "derecha")
    assert cambios == 0


def test_aplicar_posicion_pagina_recorre_la_carpeta(tmp_path):
    output_base = tmp_path / "salida"
    output_base.mkdir()
    _crear_docx_con_pagina(output_base, "uno.docx")
    _crear_docx_con_pagina(output_base, "dos.docx")

    stats = aplicar_posicion_pagina(output_base, "derecha")

    assert stats["archivos"] == 2
    assert stats["cambios"] == 2
