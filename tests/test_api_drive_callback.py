"""Tests de api.py::auth_drive_callback -- la unica ruta de la API que es
publica a proposito (Google la llama sin API key). Bug real: el `state` que
trae se usaba tal cual para nombrar el archivo de token (path traversal) y
se interpolaba sin escapar en la respuesta HTML (XSS reflejado). El fix real
(firma del state + validacion de cliente_id) esta cubierto en
tests/test_drive_service.py -- aqui se prueba especificamente que la capa
HTTP escapa lo que muestra, sin importar de donde venga el texto.
"""
from fastapi.testclient import TestClient

import api


client = TestClient(api.app)


def test_error_con_caracteres_html_se_escapa_en_la_respuesta(monkeypatch):
    """Si completar_auth_web() falla con un mensaje que contenga HTML (por
    ejemplo, un error de Google reflejando algo del request), la pagina de
    error no debe devolverlo tal cual -- debe quedar escapado."""
    def _falla(*a, **kw):
        raise ValueError("state con firma invalida: <script>alert(1)</script>")

    monkeypatch.setattr("engine.drive_service.completar_auth_web", _falla)

    r = client.get("/auth/drive/callback", params={"code": "x", "state": "cualquiera"})
    assert r.status_code == 200
    assert "<script>alert(1)</script>" not in r.text
    assert "&lt;script&gt;alert(1)&lt;/script&gt;" in r.text


def test_exito_muestra_el_cliente_id_verificado(monkeypatch):
    monkeypatch.setattr(
        "engine.drive_service.completar_auth_web",
        lambda state, code, redirect_uri: "EMPRESA_DE_PRUEBA_SAS",
    )

    r = client.get("/auth/drive/callback", params={"code": "x", "state": "cualquiera"})
    assert r.status_code == 200
    assert "EMPRESA_DE_PRUEBA_SAS" in r.text
    assert "Autorizacion completada" in r.text or "Autorización completada" in r.text
