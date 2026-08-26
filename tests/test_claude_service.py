"""Tests de engine/claude_service.py: reintentos ante errores transitorios,
construccion del contexto para el prompt, y humanizar_secciones (batching,
cache en disco, fallback ante error de API).

Bug real (reintentos): antes de este fix, _llamar_claude solo reintentaba
ante anthropic.RateLimitError -- un corte de red, un timeout o un 5xx del
lado de Anthropic tumbaban la llamada (y con ella, la generacion completa
del SG-SST) en el primer intento en vez de reintentarse solos.
"""
from unittest.mock import MagicMock

import anthropic
import httpx
import pytest

from engine import claude_service


class _CacheFalso:
    """Reemplaza el diskcache.Cache real por un dict en memoria -- evita
    tocar disco y deja que cada test empiece con la cache vacia."""

    def __init__(self):
        self._datos = {}

    def get(self, key):
        return self._datos.get(key)

    def set(self, key, value):
        self._datos[key] = value


def _seccion(id_, texto):
    return {"id": id_, "parrafos": [{"idx": id_, "texto": texto}], "texto_completo": texto}


def _respuesta_con_texto(texto: str):
    bloque = MagicMock()
    bloque.text = texto
    respuesta = MagicMock()
    respuesta.content = [bloque]
    respuesta.usage = MagicMock(input_tokens=5, output_tokens=5)
    return respuesta


def _respuesta_ok():
    bloque = MagicMock()
    bloque.text = "ok"
    respuesta = MagicMock()
    respuesta.content = [bloque]
    respuesta.usage = MagicMock(input_tokens=1, output_tokens=1)
    return respuesta


def _error_conexion() -> anthropic.APIConnectionError:
    return anthropic.APIConnectionError(
        request=httpx.Request("POST", "https://api.anthropic.com/v1/messages")
    )


def _error_servidor() -> anthropic.InternalServerError:
    respuesta = httpx.Response(
        500, request=httpx.Request("POST", "https://api.anthropic.com/v1/messages")
    )
    return anthropic.InternalServerError("error interno", response=respuesta, body=None)


@pytest.fixture(autouse=True)
def _sin_espera_real(monkeypatch):
    """Los reintentos hacen time.sleep(30/60/120) -- no hay que esperar eso
    de verdad para que el test sea rapido."""
    monkeypatch.setattr(claude_service.time, "sleep", lambda segundos: None)


class TestReintentosAnteErroresTransitorios:
    def test_reintenta_tras_error_de_conexion_y_termina_bien(self, monkeypatch):
        client = MagicMock()
        client.messages.create.side_effect = [_error_conexion(), _respuesta_ok()]
        monkeypatch.setattr(claude_service, "_crear_client", lambda: client)

        resultado = claude_service._llamar_claude("hola", max_reintentos=2)
        assert resultado is not None
        assert client.messages.create.call_count == 2

    def test_reintenta_tras_error_5xx_y_termina_bien(self, monkeypatch):
        client = MagicMock()
        client.messages.create.side_effect = [_error_servidor(), _respuesta_ok()]
        monkeypatch.setattr(claude_service, "_crear_client", lambda: client)

        resultado = claude_service._llamar_claude("hola", max_reintentos=2)
        assert resultado is not None
        assert client.messages.create.call_count == 2

    def test_reintenta_tras_rate_limit_como_antes(self, monkeypatch):
        """Regresion: el comportamiento original (rate limit) sigue igual
        despues de generalizar la tupla de excepciones reintentables."""
        respuesta_429 = httpx.Response(
            429, request=httpx.Request("POST", "https://api.anthropic.com/v1/messages")
        )
        error = anthropic.RateLimitError("rate limited", response=respuesta_429, body=None)
        client = MagicMock()
        client.messages.create.side_effect = [error, _respuesta_ok()]
        monkeypatch.setattr(claude_service, "_crear_client", lambda: client)

        resultado = claude_service._llamar_claude("hola", max_reintentos=2)
        assert resultado is not None
        assert client.messages.create.call_count == 2

    def test_relanza_si_se_agotan_los_reintentos(self, monkeypatch):
        client = MagicMock()
        client.messages.create.side_effect = _error_conexion()
        monkeypatch.setattr(claude_service, "_crear_client", lambda: client)

        with pytest.raises(anthropic.APIConnectionError):
            claude_service._llamar_claude("hola", max_reintentos=2)
        assert client.messages.create.call_count == 2

    def test_no_reintenta_errores_no_transitorios(self, monkeypatch):
        """Un error que no es de red/rate-limit/5xx (ej. request invalida)
        debe fallar de inmediato, sin agotar reintentos ni esperar."""
        respuesta_400 = httpx.Response(
            400, request=httpx.Request("POST", "https://api.anthropic.com/v1/messages")
        )
        error = anthropic.BadRequestError("prompt invalido", response=respuesta_400, body=None)
        client = MagicMock()
        client.messages.create.side_effect = error
        monkeypatch.setattr(claude_service, "_crear_client", lambda: client)

        with pytest.raises(anthropic.BadRequestError):
            claude_service._llamar_claude("hola", max_reintentos=3)
        assert client.messages.create.call_count == 1


class TestConstruirContexto:
    def test_diez_o_mas_trabajadores_usa_copasst(self, cliente_minimo):
        cliente_minimo["numero_trabajadores_empresa"] = "10"
        contexto = claude_service._construir_contexto(cliente_minimo)
        assert "COPASST" in contexto
        assert "USAR SIEMPRE 'COPASST'" in contexto

    def test_menos_de_diez_trabajadores_usa_vigia(self, cliente_minimo):
        cliente_minimo["numero_trabajadores_empresa"] = "9"
        contexto = claude_service._construir_contexto(cliente_minimo)
        assert "Vigía de SST" in contexto or "Vigía de Seguridad" in contexto
        assert "REEMPLAZARLO por 'Vigía de SST'" in contexto

    def test_incluye_sedes_en_el_contexto(self, cliente_minimo):
        cliente_minimo["sedes"] = [{"nombre": "Sede Norte", "direccion": "Calle 1", "ciudad": "Bogota"}]
        contexto = claude_service._construir_contexto(cliente_minimo)
        assert "Sede Norte" in contexto

    def test_numero_de_trabajadores_no_numerico_no_truena(self, cliente_minimo):
        """El calculo de COPASST/Vigia castea a int -- si el YAML trae un
        valor no numerico, debe caer al default (Vigia) en vez de lanzar."""
        cliente_minimo["numero_trabajadores_empresa"] = "no-definido"
        contexto = claude_service._construir_contexto(cliente_minimo)
        assert "Vigía" in contexto


class TestHumanizarSecciones:
    def test_sin_api_key_devuelve_secciones_sin_tocar(self, monkeypatch, cliente_minimo):
        monkeypatch.setattr(claude_service, "ANTHROPIC_API_KEY", "")
        secciones = [_seccion(0, "Texto de prueba numero uno de la seccion.")]
        resultado = claude_service.humanizar_secciones(secciones, cliente_minimo)
        assert resultado == secciones
        assert "texto_humanizado" not in resultado[0]

    def test_lista_vacia_no_llama_a_claude(self, monkeypatch, cliente_minimo):
        monkeypatch.setattr(claude_service, "ANTHROPIC_API_KEY", "sk-ant-fake")
        client = MagicMock()
        monkeypatch.setattr(claude_service, "_crear_client", lambda: client)
        assert claude_service.humanizar_secciones([], cliente_minimo) == []
        client.messages.create.assert_not_called()

    def test_seccion_ya_en_cache_no_llama_a_claude(self, monkeypatch, cliente_minimo):
        monkeypatch.setattr(claude_service, "ANTHROPIC_API_KEY", "sk-ant-fake")
        cache = _CacheFalso()
        monkeypatch.setattr(claude_service, "_cache", cache)
        contexto = claude_service._construir_contexto(cliente_minimo)
        texto = "Texto de prueba numero uno de la seccion."
        cache.set(claude_service._cache_key(texto, contexto, "completa"), "Ya humanizado desde cache.")

        client = MagicMock()
        monkeypatch.setattr(claude_service, "_crear_client", lambda: client)

        resultado = claude_service.humanizar_secciones([_seccion(0, texto)], cliente_minimo)
        assert resultado[0]["texto_humanizado"] == "Ya humanizado desde cache."
        client.messages.create.assert_not_called()

    def test_llama_a_claude_parsea_batch_y_guarda_en_cache(self, monkeypatch, cliente_minimo):
        monkeypatch.setattr(claude_service, "ANTHROPIC_API_KEY", "sk-ant-fake")
        cache = _CacheFalso()
        monkeypatch.setattr(claude_service, "_cache", cache)
        client = MagicMock()
        client.messages.create.return_value = _respuesta_con_texto(
            "Contenido humanizado uno.\n---SECCION---\nContenido humanizado dos."
        )
        monkeypatch.setattr(claude_service, "_crear_client", lambda: client)

        secciones = [
            _seccion(0, "Texto de prueba numero uno de la seccion."),
            _seccion(1, "Texto de prueba numero dos de la seccion."),
        ]
        resultado = claude_service.humanizar_secciones(secciones, cliente_minimo)

        assert resultado[0]["texto_humanizado"] == "Contenido humanizado uno."
        assert resultado[1]["texto_humanizado"] == "Contenido humanizado dos."
        assert client.messages.create.call_count == 1  # las 2 secciones fueron en un solo batch

        # Quedo guardado en cache para la proxima vez
        contexto = claude_service._construir_contexto(cliente_minimo)
        key = claude_service._cache_key(secciones[0]["texto_completo"], contexto, "completa")
        assert cache.get(key) == "Contenido humanizado uno."

    def test_error_de_claude_usa_texto_original_como_fallback(self, monkeypatch, cliente_minimo):
        """Si Claude falla incluso tras agotar los reintentos de
        _llamar_claude, humanizar_secciones no debe tronar: las secciones
        afectadas conservan su texto_completo original."""
        monkeypatch.setattr(claude_service, "ANTHROPIC_API_KEY", "sk-ant-fake")
        monkeypatch.setattr(claude_service, "_cache", _CacheFalso())
        client = MagicMock()
        client.messages.create.side_effect = _error_conexion()
        monkeypatch.setattr(claude_service, "_crear_client", lambda: client)

        texto = "Texto de prueba numero uno de la seccion."
        resultado = claude_service.humanizar_secciones([_seccion(0, texto)], cliente_minimo)
        assert resultado[0]["texto_humanizado"] == texto

    def test_reordena_resultado_por_id_original(self, monkeypatch, cliente_minimo):
        monkeypatch.setattr(claude_service, "ANTHROPIC_API_KEY", "sk-ant-fake")
        cache = _CacheFalso()
        contexto = claude_service._construir_contexto(cliente_minimo)
        # La seccion 0 ya esta en cache; la 1 y la 2 se piden a Claude.
        cache.set(claude_service._cache_key("Texto A de prueba con longitud suficiente.", contexto, "completa"),
                   "A humanizado.")
        monkeypatch.setattr(claude_service, "_cache", cache)

        client = MagicMock()
        client.messages.create.return_value = _respuesta_con_texto(
            "B humanizado.\n---SECCION---\nC humanizado."
        )
        monkeypatch.setattr(claude_service, "_crear_client", lambda: client)

        secciones = [
            _seccion(0, "Texto A de prueba con longitud suficiente."),
            _seccion(1, "Texto B de prueba con longitud suficiente."),
            _seccion(2, "Texto C de prueba con longitud suficiente."),
        ]
        resultado = claude_service.humanizar_secciones(secciones, cliente_minimo)
        assert [s["id"] for s in resultado] == [0, 1, 2]
        assert resultado[0]["texto_humanizado"] == "A humanizado."
        assert resultado[1]["texto_humanizado"] == "B humanizado."
        assert resultado[2]["texto_humanizado"] == "C humanizado."
