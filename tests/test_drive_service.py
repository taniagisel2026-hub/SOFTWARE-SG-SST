"""Tests de engine/drive_service.py: reintentos ante errores transitorios.

Bug real: ninguna llamada a la API de Drive tenia logica de reintento, asi
que una subida con muchos documentos podia fallar completa por un solo
error de red pasajero. El fix usa el mecanismo propio de googleapiclient
(HttpRequest.execute(num_retries=...)) en vez de reimplementar backoff a
mano -- estos tests verifican que quedo conectado en cada punto de llamada.
"""
from unittest.mock import MagicMock

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
