"""Tests del almacen SQLite del asistente de onboarding (sin HTTP)."""
import pytest

from engine import wizard_store


@pytest.fixture(autouse=True)
def _wizard_db_temporal(tmp_path, monkeypatch):
    """Cada test usa su propia base de datos SQLite temporal."""
    monkeypatch.setattr(wizard_store, "WIZARD_DB_PATH", tmp_path / "wizard_test.db")
    wizard_store.inicializar_db()
    yield


def test_crear_borrador_arranca_en_modulo_1():
    bid = wizard_store.crear_borrador()
    borrador = wizard_store.obtener_borrador(bid)
    assert borrador["modulo_actual"] == 1
    assert borrador["modulos_confirmados"] == []
    assert borrador["estado"] == "en_progreso"


def test_obtener_borrador_inexistente_retorna_none():
    assert wizard_store.obtener_borrador("no-existe") is None


def test_confirmar_modulo_avanza_el_puntero():
    bid = wizard_store.crear_borrador()
    borrador = wizard_store.guardar_modulo(bid, "modulo_1", {"razon_social": "ACME"})
    assert borrador["modulo_actual"] == 2
    assert borrador["modulos_confirmados"] == ["modulo_1"]

    borrador = wizard_store.guardar_modulo(bid, "modulo_2", {"nivel_riesgo": "III"})
    assert borrador["modulo_actual"] == 3
    assert borrador["modulos_confirmados"] == ["modulo_1", "modulo_2"]


def test_guardar_extraccion_no_combina_documentos():
    """El conflicto RUT-vs-Camara de Comercio depende de que cada documento
    se guarde por separado -- ver test_extraccion.py::test_detectar_conflictos."""
    bid = wizard_store.crear_borrador()
    wizard_store.guardar_extraccion(bid, "rut", {"direccion_empresa": "Calle 1"})
    wizard_store.guardar_extraccion(bid, "camara_comercio", {"direccion_empresa": "Carrera 2"})

    borrador = wizard_store.obtener_borrador(bid)
    assert borrador["extraccion"]["rut"]["direccion_empresa"] == "Calle 1"
    assert borrador["extraccion"]["camara_comercio"]["direccion_empresa"] == "Carrera 2"


def test_saltar_modulo_lo_marca_confirmado():
    bid = wizard_store.crear_borrador()
    borrador = wizard_store.saltar_modulo(bid, "modulo_7")
    assert "modulo_7" in borrador["modulos_confirmados"]
    assert borrador["datos"]["modulo_7"] == {"omitido": True}


def test_marcar_finalizado_guarda_estado_y_ruta():
    bid = wizard_store.crear_borrador()
    wizard_store.marcar_finalizado(bid, "clientes_config/ACME.yaml")
    borrador = wizard_store.obtener_borrador(bid)
    assert borrador["estado"] == "finalizado"
    assert borrador["config_path_final"] == "clientes_config/ACME.yaml"


def test_listar_borradores_filtra_por_estado():
    bid1 = wizard_store.crear_borrador()
    bid2 = wizard_store.crear_borrador()
    wizard_store.marcar_finalizado(bid1, "clientes_config/X.yaml")

    finalizados = wizard_store.listar_borradores(estado="finalizado")
    en_progreso = wizard_store.listar_borradores(estado="en_progreso")

    assert {b["borrador_id"] for b in finalizados} == {bid1}
    assert {b["borrador_id"] for b in en_progreso} == {bid2}


def test_eliminar_borrador():
    bid = wizard_store.crear_borrador()
    wizard_store.eliminar_borrador(bid)
    assert wizard_store.obtener_borrador(bid) is None
