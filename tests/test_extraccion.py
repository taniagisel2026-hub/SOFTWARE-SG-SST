"""Tests de extraccion estructurada (RUT/Camara de Comercio/Cedula).

Los tests rapidos mockean la respuesta de Claude (mismo espiritu que
conftest.py::_stub_claude, que ya evita llamadas reales por defecto).
El test marcado `integration` si llama a la API real -- requiere
ANTHROPIC_API_KEY con saldo.
"""
import json
from unittest.mock import MagicMock

import pytest

from engine import extraccion


def _mock_response(json_text: str) -> MagicMock:
    """Simula una respuesta de client.messages.create(...)."""
    bloque = MagicMock()
    bloque.text = json_text
    respuesta = MagicMock()
    respuesta.content = [bloque]
    respuesta.usage = MagicMock(input_tokens=10, output_tokens=10)
    return respuesta


def test_texto_vacio_retorna_diccionario_vacio():
    """PDF escaneado sin capa de texto -- no debe llamar a Claude ni tronar."""
    assert extraccion.extraer_datos_documento("rut", "") == {}
    assert extraccion.extraer_datos_documento("rut", "   ") == {}


def test_sin_api_key_retorna_diccionario_vacio(monkeypatch):
    monkeypatch.setattr(extraccion, "ANTHROPIC_API_KEY", "")
    assert extraccion.extraer_datos_documento("rut", "texto real de un RUT") == {}


def test_tipo_desconocido_lanza_value_error():
    with pytest.raises(ValueError):
        extraccion.extraer_datos_documento("pasaporte", "algun texto")


def test_extraer_datos_documento_parsea_json_de_claude(monkeypatch):
    monkeypatch.setattr(extraccion, "ANTHROPIC_API_KEY", "sk-ant-fake-para-test")

    esperado = {
        "razon_social": "ACME S.A.S", "nit": "900123456-7",
        "direccion_empresa": "Calle 1 # 2-3", "ciudad_empresa": "Bogota",
        "departamento_empresa": "Cundinamarca", "actividad_economica": "",
        "telefono_empresa": "", "email_empresa": "",
    }

    fake_client = MagicMock()
    fake_client.messages.create.return_value = _mock_response(json.dumps(esperado))
    monkeypatch.setattr(extraccion.claude_service, "_crear_client", lambda: fake_client)
    monkeypatch.setattr(extraccion.claude_service, "_registrar_uso", lambda r: None)

    resultado = extraccion.extraer_datos_documento("rut", "texto de un RUT de prueba")
    assert resultado == esperado


def test_extraer_datos_documento_quita_fences_de_markdown(monkeypatch):
    """Claude a veces envuelve el JSON en ```json ... ``` a pesar de que se le pide que no lo haga."""
    monkeypatch.setattr(extraccion, "ANTHROPIC_API_KEY", "sk-ant-fake-para-test")

    con_fences = '```json\n{"razon_social": "ACME S.A.S", "nit": "900123456-7"}\n```'
    fake_client = MagicMock()
    fake_client.messages.create.return_value = _mock_response(con_fences)
    monkeypatch.setattr(extraccion.claude_service, "_crear_client", lambda: fake_client)
    monkeypatch.setattr(extraccion.claude_service, "_registrar_uso", lambda r: None)

    resultado = extraccion.extraer_datos_documento("rut", "texto de un RUT")
    assert resultado["razon_social"] == "ACME S.A.S"
    assert resultado["nit"] == "900123456-7"


def test_extraer_datos_documento_json_invalido_retorna_vacio(monkeypatch):
    """Si Claude responde con basura (o la API falla), nunca debe tronar."""
    monkeypatch.setattr(extraccion, "ANTHROPIC_API_KEY", "sk-ant-fake-para-test")

    fake_client = MagicMock()
    fake_client.messages.create.return_value = _mock_response("esto no es JSON")
    monkeypatch.setattr(extraccion.claude_service, "_crear_client", lambda: fake_client)
    monkeypatch.setattr(extraccion.claude_service, "_registrar_uso", lambda r: None)

    assert extraccion.extraer_datos_documento("rut", "texto de un RUT") == {}


def test_detectar_conflictos_sin_diferencias():
    extracciones = {
        "rut": {"razon_social": "ACME SAS", "nit": "900123456-7", "direccion_empresa": "Calle 1 # 2-3"},
        "camara_comercio": {"razon_social": "ACME SAS", "nit": "900123456-7", "direccion_empresa": "calle 1  #  2-3"},
    }
    # Misma direccion, solo difieren en espacios/mayusculas -- no debe marcar conflicto
    assert extraccion.detectar_conflictos(extracciones) == []


def test_detectar_conflictos_direccion_distinta_rut_vs_camara():
    """El caso concreto que motivo este requisito: RUT y Camara de Comercio
    con direcciones distintas para la misma empresa."""
    extracciones = {
        "rut": {"razon_social": "ACME SAS", "nit": "900123456-7", "direccion_empresa": "Calle 45 # 20-15"},
        "camara_comercio": {"razon_social": "ACME SAS", "nit": "900123456-7", "direccion_empresa": "Carrera 68 # 12-30"},
    }
    conflictos = extraccion.detectar_conflictos(extracciones)
    campos_en_conflicto = {c["campo"] for c in conflictos}
    assert "direccion_empresa" in campos_en_conflicto
    assert "razon_social" not in campos_en_conflicto

    conflicto_direccion = next(c for c in conflictos if c["campo"] == "direccion_empresa")
    assert conflicto_direccion["valores"]["rut"] == "Calle 45 # 20-15"
    assert conflicto_direccion["valores"]["camara_comercio"] == "Carrera 68 # 12-30"


def test_detectar_conflictos_ignora_campos_con_un_solo_documento():
    extracciones = {"rut": {"direccion_empresa": "Calle 1"}}
    assert extraccion.detectar_conflictos(extracciones) == []


@pytest.mark.integration
def test_extraccion_real_contra_claude():
    """Requiere ANTHROPIC_API_KEY con saldo -- llama a la API de verdad."""
    from engine.config import ANTHROPIC_API_KEY
    if not ANTHROPIC_API_KEY:
        pytest.skip("ANTHROPIC_API_KEY no configurada")

    texto_rut = (
        "REGISTRO UNICO TRIBUTARIO - RUT\n"
        "Razon Social: TRANSPORTE Y LOGISTICA YATS S.A.S\n"
        "NIT: 900.460.873-0\n"
        "Direccion: Calle 45 # 20-15\n"
        "Ciudad: Bogota D.C.\n"
    )
    resultado = extraccion.extraer_datos_documento("rut", texto_rut)
    assert resultado.get("nit", "").replace(".", "").replace("-", "").startswith("9004608730")
