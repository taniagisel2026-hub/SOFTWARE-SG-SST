"""Tests de engine/cliente.py::generar_contexto()."""
from engine.cliente import generar_contexto


def test_descripcion_empresa_llega_al_contexto(cliente_minimo):
    """BUG: {{DESCRIPCION_EMPRESA}} quedaba sin reemplazar (SST-PRG-009) porque
    generar_contexto() nunca mapeaba descripcion_empresa a un marcador.

    Detectado en auditoria real (2026-07-10, cliente EMPRESA_DE_PRUEBA_SAS)."""
    contexto = generar_contexto(cliente_minimo)
    assert contexto["DESCRIPCION_EMPRESA"] == cliente_minimo["descripcion_empresa"]
    assert contexto["DESCRIPCIÓN_EMPRESA"] == cliente_minimo["descripcion_empresa"]
    assert contexto["descripcion_empresa"] == cliente_minimo["descripcion_empresa"]


def test_descripcion_empresa_vacia_no_truena(cliente_minimo):
    del cliente_minimo["descripcion_empresa"]
    contexto = generar_contexto(cliente_minimo)
    assert contexto["DESCRIPCION_EMPRESA"] == ""
