"""Tests de engine/colores.py: reemplazo de colores corporativos."""
from pathlib import Path

import docx
import openpyxl
import pytest
from docx.oxml.ns import qn
from openpyxl.styles import PatternFill

from engine.colores import (
    _construir_mapa_colores,
    aplicar_colores_docx,
    aplicar_colores_xlsx,
    resolver_paleta,
)


def _shade_docx_cell(cell, hex_color: str) -> None:
    """Aplica w:shd (fondo de celda) directo en XML, como las plantillas reales."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.makeelement(qn("w:shd"), {
        qn("w:val"): "clear", qn("w:color"): "auto", qn("w:fill"): hex_color,
    })
    tcPr.append(shd)


def _crear_docx_con_celdas(path: Path, colores: list[str]) -> None:
    doc = docx.Document()
    tabla = doc.add_table(rows=1, cols=len(colores))
    for cell, color in zip(tabla.rows[0].cells, colores):
        _shade_docx_cell(cell, color)
    doc.save(str(path))


def _leer_fills_docx(path: Path) -> list[str]:
    doc = docx.Document(str(path))
    tabla = doc.tables[0]
    fills = []
    for cell in tabla.rows[0].cells:
        shd = cell._tc.find(".//" + qn("w:shd"))
        fills.append(shd.get(qn("w:fill")).upper() if shd is not None else None)
    return fills


def _crear_xlsx_con_celdas(path: Path, colores: list[str]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, color in enumerate(colores, start=1):
        ws.cell(row=1, column=i).fill = PatternFill(
            start_color=color, end_color=color, fill_type="solid"
        )
    wb.save(str(path))


def _leer_fills_xlsx(path: Path) -> list[str]:
    wb = openpyxl.load_workbook(str(path))
    ws = wb.active
    fills = []
    for i in range(1, ws.max_column + 1):
        rgb = ws.cell(row=1, column=i).fill.fgColor.rgb
        fills.append(rgb[-6:].upper() if isinstance(rgb, str) else None)
    return fills


class TestColorAzulPesvSeAplica:
    """Bug real: las tablas del PESV usan un azul propio (1C4587/1F3864) que
    no formaba parte del mapa de reemplazo, asi que se quedaban azules sin
    importar el color_corporativo elegido por el cliente."""

    def test_docx_1c4587_cambia_con_color_corporativo(self, tmp_path):
        f = tmp_path / "pesv.docx"
        _crear_docx_con_celdas(f, ["1C4587"])
        mapa = _construir_mapa_colores(resolver_paleta("#33CC33"))
        cambios = aplicar_colores_docx(str(f), mapa)
        assert cambios >= 1
        assert _leer_fills_docx(f) == ["33CC33"]

    def test_xlsx_1c4587_cambia_con_color_corporativo(self, tmp_path):
        f = tmp_path / "pesv.xlsx"
        _crear_xlsx_con_celdas(f, ["1C4587"])
        mapa = _construir_mapa_colores(resolver_paleta("#33CC33"))
        cambios = aplicar_colores_xlsx(str(f), mapa)
        assert cambios >= 1
        assert _leer_fills_xlsx(f) == ["33CC33"]

    def test_docx_sin_color_corporativo_no_toca_nada(self, tmp_path):
        """Si el mapa esta vacio (sin color_corporativo en el YAML), el 1C4587
        original de la plantilla debe quedar intacto."""
        f = tmp_path / "pesv.docx"
        _crear_docx_con_celdas(f, ["1C4587"])
        cambios = aplicar_colores_docx(str(f), {})
        assert cambios == 0
        assert _leer_fills_docx(f) == ["1C4587"]


class TestSinColisionEncadenada:
    """Bug potencial: la paleta 'azul' predefinida tiene oscuro=1F3864, que
    es EXACTAMENTE uno de los colores originales de plantilla PESV. Si los
    reemplazos se aplican en cadena (uno tras otro sobre el mismo texto), la
    celda que empezo en 1C4587 (-> 1F3864) podria ser atrapada de nuevo por
    la regla de 1F3864 (-> 2E4A7A), perdiendo el efecto de dos tonos."""

    def test_docx_1c4587_y_1f3864_no_se_mezclan_con_paleta_azul(self, tmp_path):
        f = tmp_path / "pesv.docx"
        _crear_docx_con_celdas(f, ["1C4587", "1F3864"])
        paleta = resolver_paleta("azul")
        assert paleta["oscuro"] == "1F3864"  # precondicion del escenario de colision
        mapa = _construir_mapa_colores(paleta)
        aplicar_colores_docx(str(f), mapa)
        fills = _leer_fills_docx(f)
        assert fills == ["1F3864", "2E4A7A"], (
            f"Colision en cadena: {fills} (la celda original 1C4587 no debe "
            "terminar igual a la celda original 1F3864)"
        )

    def test_xlsx_1c4587_y_1f3864_no_se_mezclan_con_paleta_azul(self, tmp_path):
        f = tmp_path / "pesv.xlsx"
        _crear_xlsx_con_celdas(f, ["1C4587", "1F3864"])
        paleta = resolver_paleta("azul")
        mapa = _construir_mapa_colores(paleta)
        aplicar_colores_xlsx(str(f), mapa)
        fills = _leer_fills_xlsx(f)
        assert fills == ["1F3864", "2E4A7A"]


class TestReemplazoVinotintoExistenteSigueFuncionando:
    """Regresion: el refactor a un solo patron generico no debe romper el
    reemplazo original de la paleta vino tinto."""

    def test_docx_660000_cambia(self, tmp_path):
        f = tmp_path / "doc.docx"
        _crear_docx_con_celdas(f, ["660000", "F4CCCC"])
        mapa = _construir_mapa_colores(resolver_paleta("verde"))
        aplicar_colores_docx(str(f), mapa)
        paleta = resolver_paleta("verde")
        assert _leer_fills_docx(f) == [paleta["oscuro"], paleta["claro1"]]

    def test_xlsx_c00000_cambia(self, tmp_path):
        f = tmp_path / "doc.xlsx"
        _crear_xlsx_con_celdas(f, ["C00000"])
        paleta = resolver_paleta("morado")
        mapa = _construir_mapa_colores(paleta)
        aplicar_colores_xlsx(str(f), mapa)
        assert _leer_fills_xlsx(f) == [paleta["medio"]]
