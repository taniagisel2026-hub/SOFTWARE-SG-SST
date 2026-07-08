"""Tests de las correcciones legales de normativa.py.

Cubre la funcionalidad agregada en la auditoria normativa 2026-07-08:
  - Cada regla de CORRECCIONES puede tener un "vigente_desde" (año)
  - Las reglas sin fecha (errores de cita) se aplican siempre
  - Las reglas con fecha solo aplican si el año del cliente ya la alcanzo
  - aplicar_correcciones_docx() respeta anio_cliente de punta a punta
  - Sin anio_cliente, se aplican todas las reglas (comportamiento previo,
    para no romper generaciones sin el campo "año")
"""
from docx import Document

from engine.normativa import (
    CORRECCIONES,
    _patrones_activos,
    aplicar_correcciones_docx,
    listar_correcciones,
)


def _crear_docx(tmp_path, texto) -> str:
    doc = Document()
    doc.add_paragraph(texto)
    path = tmp_path / "test.docx"
    doc.save(str(path))
    return str(path)


# ---------------------------------------------------------------------------
# Filtrado de reglas por año (_patrones_activos)
# ---------------------------------------------------------------------------

class TestPatronesActivos:

    def test_sin_anio_incluye_todas_las_reglas(self):
        assert len(_patrones_activos(None)) == len(CORRECCIONES)

    def test_reglas_sin_fecha_siempre_incluidas(self):
        """Los errores de cita (vigente_desde=None) no dependen del año."""
        activas_2000 = {nueva for _, nueva in _patrones_activos(2000)}
        for _vieja, nueva, vigente_desde in CORRECCIONES:
            if vigente_desde is None:
                assert nueva in activas_2000

    def test_regla_no_aplica_antes_de_su_vigencia(self):
        """Resolución 3461 de 2025 (vigente desde 2025) no debe verse en 2020."""
        textos_nuevos = [nueva for _, nueva in _patrones_activos(2020)]
        assert "Resolución 3461 de 2025" not in textos_nuevos

    def test_regla_aplica_en_su_propio_año(self):
        textos_nuevos = [nueva for _, nueva in _patrones_activos(2015)]
        assert "Decreto 1072 de 2015" in textos_nuevos

    def test_mas_años_nunca_reduce_reglas_activas(self):
        """Monotonía: a mayor año del cliente, nunca hay menos reglas activas."""
        conteos = [len(_patrones_activos(anio)) for anio in range(1990, 2027)]
        assert conteos == sorted(conteos)


# ---------------------------------------------------------------------------
# listar_correcciones() (usado por `python cli.py normas`)
# ---------------------------------------------------------------------------

class TestListarCorrecciones:

    def test_incluye_vigente_desde(self):
        correcciones = listar_correcciones()
        assert len(correcciones) == len(CORRECCIONES)
        assert all("vigente_desde" in c for c in correcciones)


# ---------------------------------------------------------------------------
# aplicar_correcciones_docx() end-to-end con anio_cliente
# ---------------------------------------------------------------------------

class TestAplicarCorreccionesDocxPorAnio:

    def test_typo_se_corrige_sin_importar_el_año(self, tmp_path):
        """Decreto 1010 de 2006 nunca existió: se corrige incluso para un
        cliente muy anterior a la auditoría."""
        path = _crear_docx(tmp_path, "Cumplir el DECRETO 1010 DE 2006.")
        reporte = aplicar_correcciones_docx(path, detallado=True, anio_cliente=1995)
        assert reporte.total_cambios == 1
        assert Document(path).paragraphs[0].text == "Cumplir el Ley 1010 de 2006."

    def test_correccion_con_fecha_no_aplica_a_cliente_anterior(self, tmp_path):
        """Un SG-SST de 2010 debe conservar el Decreto 614 de 1984 tal cual,
        porque el Decreto 1072 de 2015 todavía no existía."""
        path = _crear_docx(tmp_path, "Cumplir el DECRETO 614 DE 1984.")
        reporte = aplicar_correcciones_docx(path, detallado=True, anio_cliente=2010)
        assert reporte.total_cambios == 0
        assert Document(path).paragraphs[0].text == "Cumplir el DECRETO 614 DE 1984."

    def test_correccion_con_fecha_aplica_desde_su_año(self, tmp_path):
        path = _crear_docx(tmp_path, "Cumplir el DECRETO 614 DE 1984.")
        reporte = aplicar_correcciones_docx(path, detallado=True, anio_cliente=2015)
        assert reporte.total_cambios == 1
        assert "Decreto 1072 de 2015" in Document(path).paragraphs[0].text

    def test_correccion_reciente_no_aplica_a_cliente_de_hace_unos_años(self, tmp_path):
        """Resolución 652 de 2012 fue derogada por la 3461 de 2025: un
        cliente de 2020 debe conservar la cita original de 2012."""
        path = _crear_docx(tmp_path, "Cumplir la RESOLUCIÓN 652 DE 2012.")
        reporte = aplicar_correcciones_docx(path, detallado=True, anio_cliente=2020)
        assert reporte.total_cambios == 0
        assert "RESOLUCIÓN 652 DE 2012" in Document(path).paragraphs[0].text

    def test_correccion_reciente_aplica_a_cliente_actual(self, tmp_path):
        path = _crear_docx(tmp_path, "Cumplir la RESOLUCIÓN 652 DE 2012.")
        reporte = aplicar_correcciones_docx(path, detallado=True, anio_cliente=2026)
        assert reporte.total_cambios == 1
        assert "Resolución 3461 de 2025" in Document(path).paragraphs[0].text

    def test_sin_anio_cliente_aplica_todas_las_reglas(self, tmp_path):
        """Compatibilidad: si no se pasa anio_cliente, se corrige todo
        (comportamiento previo a la auditoría 2026-07-08)."""
        path = _crear_docx(tmp_path, "Cumplir la RESOLUCIÓN 652 DE 2012.")
        reporte = aplicar_correcciones_docx(path, detallado=True)
        assert reporte.total_cambios == 1
        assert "Resolución 3461 de 2025" in Document(path).paragraphs[0].text

    def test_multiples_reglas_de_distintos_años_en_un_solo_documento(self, tmp_path):
        """Reproduce el caso real verificado sobre SST-PRG-001: un cliente de
        2020 recibe la corrección de 2015 pero no la de 2025."""
        texto = (
            "Cumplir con el DECRETO 614 DE 1984 y la RESOLUCIÓN 652 DE 2012 "
            "y el DECRETO 1010 DE 2006."
        )
        path = _crear_docx(tmp_path, texto)
        reporte = aplicar_correcciones_docx(path, detallado=True, anio_cliente=2020)
        assert reporte.total_cambios == 2  # typo + Decreto 1072/2015, no la de 2025

        resultado = Document(path).paragraphs[0].text
        assert "Decreto 1072 de 2015" in resultado
        assert "Ley 1010 de 2006" in resultado
        assert "RESOLUCIÓN 652 DE 2012" in resultado  # sin corregir, año insuficiente
