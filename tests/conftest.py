"""Fixtures compartidas entre tests."""
import shutil
from pathlib import Path

import pytest

ROOT = Path(__file__).parent.parent
PLANTILLA = ROOT / "Plantilla SG-SST 2026"


@pytest.fixture
def cliente_minimo() -> dict:
    """Dict con los campos minimos de un cliente para tests."""
    return {
        "razon_social": "EMPRESA TEST S.A.S",
        "nit": "900000000-1",
        "actividad_economica": "Test",
        "nivel_riesgo": "I",
        "arl": "SURA",
        "numero_trabajadores_empresa": "25",
        "ciudad_empresa": "Bogota",
        "descripcion_empresa": "Empresa de pruebas",
        "generar_pesv": False,
    }


@pytest.fixture
def cliente_con_pesv(cliente_minimo) -> dict:
    """Cliente con PESV activado y cargos viales."""
    cliente_minimo["generar_pesv"] = True
    cliente_minimo["cargos"] = [
        {
            "nombre": "Conductor de Camion",
            "proceso": "LOGISTICA",
            "tipo": "conductor",
            "num_expuestos": 2,
            "actividades": "Conduccion de camiones para entrega de productos.",
        },
        {
            "nombre": "Mensajero Motorizado",
            "proceso": "LOGISTICA",
            "tipo": "operativo",
            "num_expuestos": 4,
            "actividades": "Entregas en motocicleta.",
        },
        {
            "nombre": "Auxiliar de Tramites",
            "proceso": "ADMINISTRATIVO",
            "tipo": "administrativo",
            "num_expuestos": 1,
            "actividades": "Diligencias externas en bancos y notarias.",
        },
    ]
    return cliente_minimo


@pytest.fixture
def cliente_con_cargos(cliente_minimo) -> dict:
    """Cliente con cargos para matriz IPVR."""
    cliente_minimo["cargos"] = [
        {
            "nombre": "Director General",
            "proceso": "ADMINISTRATIVO",
            "tipo": "administrativo",
            "num_expuestos": 1,
            "actividades": "Dirige la empresa.",
            "tareas": "- Definir estrategia",
        },
        {
            "nombre": "Operario de Produccion",
            "proceso": "PRODUCCION",
            "tipo": "operativo",
            "num_expuestos": 5,
            "actividades": "Opera maquinaria.",
            "tareas": "- Manejar maquinaria",
        },
    ]
    return cliente_minimo


@pytest.fixture
def output_ipvr(tmp_path) -> Path:
    """Carpeta temporal con la matriz IPVR copiada desde la plantilla."""
    src = PLANTILLA / "II. HACER" / "4.1.2 MATRIZ DE IPVR" / "SST-MTZ-012 MATRIZ DE IPVR.xlsx"
    if not src.exists():
        pytest.skip(f"No existe la plantilla IPVR: {src}")
    dst = tmp_path / "II. HACER" / "4.1.2 MATRIZ DE IPVR" / "SST-MTZ-012 MATRIZ DE IPVR.xlsx"
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(str(src), str(dst))
    return tmp_path


@pytest.fixture
def output_pesv(tmp_path) -> Path:
    """Carpeta temporal con la matriz vial PESV copiada desde la plantilla."""
    src = (PLANTILLA / "PLAN ESTRATÉGICO DE SEGURIDAD VIAL" / "FASE 1"
           / "6. CARACTERIZACIÓN, EVALUACIÓN Y CONTROL DE RIESGOS"
           / "PESV-MTZ-002 MATRIZ DE PELIGROS Y VALORACIÓN DE RIESGOS VIALES.xlsx")
    if not src.exists():
        pytest.skip(f"No existe la plantilla PESV: {src}")
    dst = (tmp_path / "PLAN ESTRATÉGICO DE SEGURIDAD VIAL" / "FASE 1"
           / "6. CARACTERIZACIÓN, EVALUACIÓN Y CONTROL DE RIESGOS"
           / "PESV-MTZ-002 MATRIZ DE PELIGROS Y VALORACIÓN DE RIESGOS VIALES.xlsx")
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(str(src), str(dst))
    return tmp_path


@pytest.fixture
def output_listado_maestro(tmp_path) -> Path:
    """Carpeta temporal con SST-MTZ-005 copiado desde la plantilla.

    La plantilla ya no tiene un SST-MTZ-005 sin etiqueta -- desde que se
    agrego el filtro por categoria Res. 0312 (ver generador.py, tag [CAT-I]/
    [CAT-II]/[CAT-III], mismo mecanismo que [COPASST]/[VIGIA]) solo existen
    las 3 variantes etiquetadas; generador.py elige la que aplica y le quita
    la etiqueta al copiarla. Estos tests no dependen de la categoria (prueban
    el matcher de links, no el filtrado por estandares), asi que cualquiera
    de las 3 sirve -- se copia a un nombre sin etiqueta, igual que quedaria
    en una generacion real."""
    carpeta = PLANTILLA / "I. PLANEAR" / "2.5.1 ARCHIVO Y RETENCIÓN DOCUMENTAL"
    candidatos = sorted(carpeta.glob("SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL*.xlsx"))
    candidatos = [c for c in candidatos if not c.name.startswith("~$")]
    if not candidatos:
        pytest.skip(f"No existe ninguna plantilla de SST-MTZ-005 en: {carpeta}")
    src = candidatos[0]
    dst = (tmp_path / "I. PLANEAR" / "2.5.1 ARCHIVO Y RETENCIÓN DOCUMENTAL"
           / "SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL.xlsx")
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(str(src), str(dst))
    return tmp_path


@pytest.fixture(autouse=True)
def _stub_claude(monkeypatch):
    """Por defecto, evita llamadas reales a Claude API en los tests.
    Las funciones que llamen a Claude devolveran respuestas vacias o por defecto."""
    monkeypatch.setenv("ANTHROPIC_API_KEY", "")
    # Si engine.matriz_peligros._generar_datos_cargo existe, hacerlo devolver fallback
    try:
        from engine import matriz_peligros
        def _stub(cliente, cargo, tipo_cargo):
            return {
                "modo": "legacy",
                "actividades": cargo.get("actividades") or f"Funciones del cargo {cargo.get('nombre')}",
                "tareas": cargo.get("tareas") or "- Tarea generica",
                "peligros_adicionales": [],
            }
        monkeypatch.setattr(matriz_peligros, "_generar_datos_cargo", _stub)
    except ImportError:
        pass
