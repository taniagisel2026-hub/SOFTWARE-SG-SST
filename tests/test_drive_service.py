"""Tests de engine/drive_service.py: reintentos ante errores transitorios.

Bug real: ninguna llamada a la API de Drive tenia logica de reintento, asi
que una subida con muchos documentos podia fallar completa por un solo
error de red pasajero. El fix usa el mecanismo propio de googleapiclient
(HttpRequest.execute(num_retries=...)) en vez de reimplementar backoff a
mano -- estos tests verifican que quedo conectado en cada punto de llamada.
"""
from unittest.mock import MagicMock

import pytest

from engine import drive_service


def _service_con_ejecutar(retorno: dict) -> MagicMock:
    """Un mock de `service` donde cualquier .execute(...) devuelve `retorno`,
    y permite inspeccionar despues con que argumentos se llamo a execute."""
    service = MagicMock()
    service.files.return_value.list.return_value.execute.return_value = retorno
    service.files.return_value.create.return_value.execute.return_value = retorno
    service.files.return_value.update.return_value.execute.return_value = retorno
    return service


class TestCrearCarpetaReintenta:
    def test_carpeta_existente_usa_num_retries_en_list(self):
        service = _service_con_ejecutar({"files": [{"id": "carpeta-existente"}]})
        resultado = drive_service.crear_carpeta(service, "I. PLANEAR")
        assert resultado == "carpeta-existente"
        _, kwargs = service.files.return_value.list.return_value.execute.call_args
        assert kwargs["num_retries"] == drive_service._NUM_REINTENTOS

    def test_carpeta_nueva_usa_num_retries_en_list_y_create(self):
        service = _service_con_ejecutar({"files": []})
        service.files.return_value.create.return_value.execute.return_value = {"id": "carpeta-nueva"}
        resultado = drive_service.crear_carpeta(service, "I. PLANEAR")
        assert resultado == "carpeta-nueva"

        _, kwargs_list = service.files.return_value.list.return_value.execute.call_args
        assert kwargs_list["num_retries"] == drive_service._NUM_REINTENTOS
        _, kwargs_create = service.files.return_value.create.return_value.execute.call_args
        assert kwargs_create["num_retries"] == drive_service._NUM_REINTENTOS


class TestSubirArchivoReintenta:
    def test_archivo_nuevo_usa_num_retries_en_list_y_create(self, tmp_path):
        archivo_local = tmp_path / "politica.docx"
        archivo_local.write_bytes(b"contenido de prueba")

        service = _service_con_ejecutar({"files": []})
        service.files.return_value.create.return_value.execute.return_value = {
            "id": "archivo-nuevo", "name": "politica.docx", "webViewLink": "https://drive/x",
        }
        resultado = drive_service.subir_archivo(service, archivo_local, parent_id="carpeta-1")
        assert resultado["id"] == "archivo-nuevo"

        _, kwargs_list = service.files.return_value.list.return_value.execute.call_args
        assert kwargs_list["num_retries"] == drive_service._NUM_REINTENTOS
        _, kwargs_create = service.files.return_value.create.return_value.execute.call_args
        assert kwargs_create["num_retries"] == drive_service._NUM_REINTENTOS

    def test_archivo_existente_usa_num_retries_en_update(self, tmp_path):
        archivo_local = tmp_path / "politica.docx"
        archivo_local.write_bytes(b"contenido de prueba")

        service = _service_con_ejecutar({"files": [{"id": "archivo-existente"}]})
        service.files.return_value.update.return_value.execute.return_value = {
            "id": "archivo-existente", "name": "politica.docx", "webViewLink": "https://drive/x",
        }
        resultado = drive_service.subir_archivo(service, archivo_local, parent_id="carpeta-1")
        assert resultado["id"] == "archivo-existente"

        _, kwargs_update = service.files.return_value.update.return_value.execute.call_args
        assert kwargs_update["num_retries"] == drive_service._NUM_REINTENTOS


class TestTokenPathRechazaClienteIdInvalido:
    """Bug real: _token_path() armaba tokens/token_<cliente_id>.json sin
    validar cliente_id -- un cliente_id con '/' o '..' permitia escribir el
    archivo de token fuera de TOKENS_DIR (path traversal)."""

    @pytest.mark.parametrize("cliente_id", [
        "../../etc/passwd",
        "sub/dir",
        "sub\\dir",
        "",
        "cliente con espacios",
        "cliente;rm -rf",
    ])
    def test_cliente_id_con_separadores_o_caracteres_invalidos_lanza(self, cliente_id):
        with pytest.raises(ValueError):
            drive_service._token_path(cliente_id)

    def test_cliente_id_normal_funciona(self, tmp_path, monkeypatch):
        monkeypatch.setattr(drive_service, "TOKENS_DIR", tmp_path)
        ruta = drive_service._token_path("EMPRESA_DE_PRUEBA_SAS")
        assert ruta == tmp_path / "token_EMPRESA_DE_PRUEBA_SAS.json"

    def test_cliente_id_con_signo_mas_funciona(self, tmp_path, monkeypatch):
        """Nombres de cliente reales usan '+' (ej. CONSTRUCTORA_P+V) -- la
        validacion no debe rechazar casos legitimos."""
        monkeypatch.setattr(drive_service, "TOKENS_DIR", tmp_path)
        ruta = drive_service._token_path("CONSTRUCTORA_P+V")
        assert ruta == tmp_path / "token_CONSTRUCTORA_P+V.json"


class TestFirmaDeState:
    """Bug real: el callback publico /auth/drive/callback (Google lo llama
    sin API key) confiaba en el `state` tal cual, sin verificar que lo
    hubiera generado iniciar_auth_web(). Firmar el state con HMAC (usando
    el mismo API_AUTH_TOKEN que ya protege el resto de la API) cierra esa
    puerta: solo pasa la verificacion un state generado por alguien que
    ya tenia la API key."""

    def test_state_firmado_se_verifica_correctamente(self, monkeypatch):
        monkeypatch.setattr(drive_service, "API_AUTH_TOKEN", "secreto-de-prueba")
        state = drive_service._firmar_state("EMPRESA_DE_PRUEBA_SAS")
        assert drive_service._verificar_state(state) == "EMPRESA_DE_PRUEBA_SAS"

    def test_state_falsificado_es_rechazado(self, monkeypatch):
        monkeypatch.setattr(drive_service, "API_AUTH_TOKEN", "secreto-de-prueba")
        with pytest.raises(ValueError):
            drive_service._verificar_state("OTRO_CLIENTE.firma-inventada")

    def test_state_de_otro_cliente_firmado_con_otra_clave_es_rechazado(self, monkeypatch):
        monkeypatch.setattr(drive_service, "API_AUTH_TOKEN", "secreto-de-prueba")
        state = drive_service._firmar_state("EMPRESA_DE_PRUEBA_SAS")
        monkeypatch.setattr(drive_service, "API_AUTH_TOKEN", "otra-clave-distinta")
        with pytest.raises(ValueError):
            drive_service._verificar_state(state)

    def test_state_sin_punto_es_rechazado(self, monkeypatch):
        monkeypatch.setattr(drive_service, "API_AUTH_TOKEN", "secreto-de-prueba")
        with pytest.raises(ValueError):
            drive_service._verificar_state("sin-firma-adjunta")

    def test_state_con_cliente_id_invalido_se_rechaza_en_completar_auth_web(self, monkeypatch):
        """Aunque la firma sea valida, completar_auth_web() debe propagar el
        ValueError de _token_path() si el cliente_id firmado no es valido
        (defensa en profundidad -- no deberia poder ocurrir en la practica
        porque _firmar_state() solo firma lo que ya se valido al crear el
        borrador, pero si el secreto se filtrara, esto sigue protegiendo)."""
        monkeypatch.setattr(drive_service, "API_AUTH_TOKEN", "secreto-de-prueba")
        state = drive_service._firmar_state("../../fuera")
        with pytest.raises(ValueError):
            drive_service.completar_auth_web(state, "codigo-falso", "https://x/callback")
