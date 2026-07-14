"""Tests de engine/matriz_documental.py.

BUG: filtrar_matriz_documental() (filtra SST-MTZ-005 segun la categoria del
cliente, Res. 0312/2019) existia pero nunca se llamaba desde generador.py --
solo se llamaba _adaptar_vigia_xlsx() por separado, que ni siquiera es una
funcion publica. El resultado: SST-MTZ-005 siempre mostraba los 60 estandares
completos sin importar la categoria real de la empresa.

Detectado en conversacion real (2026-07-14), al preguntar en que parte del
flujo se determinaban los estandares minimos aplicables por empresa."""
import openpyxl
import pytest

from engine.matriz_documental import determinar_categoria, filtrar_matriz_documental


@pytest.mark.parametrize("nivel_riesgo,num_trabajadores,esperado", [
    ("I", 5, "I"),
    ("III", 10, "I"),
    ("I", 11, "II"),
    ("III", 50, "II"),
    ("IV", 3, "III"),
    ("V", 3, "III"),
    ("I", 51, "III"),
])
def test_determinar_categoria(nivel_riesgo, num_trabajadores, esperado):
    assert determinar_categoria(nivel_riesgo, num_trabajadores) == esperado


def _crear_mtz005(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "LISTADO MAESTRO PHVA"
    for _ in range(6):
        ws.append([None, None, None])
    ws.append([None, None, "1.1.1 Responsable del SG-SST"])       # Cat I
    ws.append([None, None, None])                                  # continuacion 1.1.1
    ws.append([None, None, "1.2.2 Estandar que NO aplica a Cat I"])
    ws.append([None, None, "2.5.1 Archivo y retención documental"])  # Cat I
    ruta = tmp_path / "SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL.xlsx"
    wb.save(str(ruta))
    return ruta


def test_filtrar_matriz_documental_elimina_filas_no_aplicables_categoria_i(tmp_path):
    _crear_mtz005(tmp_path)
    resultado = filtrar_matriz_documental(tmp_path, nivel_riesgo="I", num_trabajadores=5)
    assert resultado["categoria"] == "I"
    assert resultado["filas_eliminadas"] == 1

    wb = openpyxl.load_workbook(str(next(tmp_path.glob("*MTZ-005*"))))
    ws = wb["LISTADO MAESTRO PHVA"]
    textos = [ws.cell(row=r, column=3).value for r in range(7, ws.max_row + 1)]
    assert not any(t and "1.2.2" in t for t in textos)


def test_filtrar_matriz_documental_categoria_iii_no_elimina_filas(tmp_path):
    _crear_mtz005(tmp_path)
    resultado = filtrar_matriz_documental(tmp_path, nivel_riesgo="V", num_trabajadores=3)
    assert resultado["categoria"] == "III"
    assert resultado["filas_eliminadas"] == 0


def test_filtrar_matriz_documental_sin_archivo_reporta_error(tmp_path):
    resultado = filtrar_matriz_documental(tmp_path, nivel_riesgo="I", num_trabajadores=5)
    assert resultado["categoria"] is None
    assert resultado["error"]
