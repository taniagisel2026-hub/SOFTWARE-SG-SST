"""Tests de los endpoints del asistente de onboarding, via TestClient (sin
levantar un servidor real). Cubre: subida+extraccion del modulo 1, el
bloqueo secuencial entre modulos (409), y finalizar produciendo un YAML
que carga sin errores en el pipeline existente.
"""
import io

import pytest
from fastapi.testclient import TestClient

import api
import wizard_api
from engine import wizard_store


@pytest.fixture(autouse=True)
def _wizard_aislado(tmp_path, monkeypatch):
    """Aisla cada test: su propia base de datos, carpeta de subidas, y
    clientes_config -- nunca toca el jobs.db/wizard_drafts.db reales ni
    escribe YAML en la carpeta real de clientes_config."""
    monkeypatch.setattr(wizard_store, "WIZARD_DB_PATH", tmp_path / "wizard_test.db")
    wizard_store.inicializar_db()
    monkeypatch.setattr(wizard_api, "WIZARD_UPLOADS_DIR", tmp_path / "wizard_uploads")
    monkeypatch.setattr(wizard_api, "CLIENTES_CONFIG_PATH", tmp_path / "clientes_config")
    (tmp_path / "clientes_config").mkdir(parents=True, exist_ok=True)

    import api_auth
    monkeypatch.setattr(api_auth, "API_AUTH_TOKEN", "token-de-prueba")
    yield


@pytest.fixture
def client():
    return TestClient(api.app)


@pytest.fixture
def headers():
    return {"Authorization": "Bearer token-de-prueba"}


def _pdf_falso(texto: str) -> bytes:
    """PDF minimo pero valido -- no necesita texto real extraible para
    estos tests (la extraccion de Claude ya se prueba aparte, mockeada,
    en test_extraccion.py). Aqui solo importa que el archivo se guarde."""
    return b"%PDF-1.4\n%mock\n" + texto.encode()


def test_crear_borrador(client, headers):
    r = client.post("/wizard/borradores", headers=headers)
    assert r.status_code == 200
    assert r.json()["modulo_actual"] == 1


def test_modulo_2_sin_confirmar_modulo_1_da_409(client, headers):
    bid = client.post("/wizard/borradores", headers=headers).json()["borrador_id"]
    r = client.post(
        f"/wizard/borradores/{bid}/modulo-2/confirmar",
        headers=headers,
        json={"nivel_riesgo": "III", "arl": "SURA", "numero_trabajadores_empresa": "9", "color_corporativo": "azul"},
    )
    assert r.status_code == 409


def test_subir_archivos_modulo_1_guarda_y_extrae(client, headers, monkeypatch):
    from engine import extraccion
    monkeypatch.setattr(extraccion, "extraer_datos_documento", lambda tipo, texto: {"nit": "900123456-7"})

    bid = client.post("/wizard/borradores", headers=headers).json()["borrador_id"]
    r = client.post(
        f"/wizard/borradores/{bid}/modulo-1/archivos",
        headers=headers,
        files={"rut": ("rut.pdf", io.BytesIO(_pdf_falso("RUT de prueba")), "application/pdf")},
    )
    assert r.status_code == 200
    body = r.json()
    assert body["extraccion"]["rut"]["nit"] == "900123456-7"
    assert body["archivos"]["rut"]


def test_subir_archivos_modulo_4_guarda_documentos_soporte(client, headers):
    bid = client.post("/wizard/borradores", headers=headers).json()["borrador_id"]
    r = client.post(
        f"/wizard/borradores/{bid}/modulo-4/archivos",
        headers=headers,
        files={
            "firma_sst": ("firma.png", io.BytesIO(b"fake-png"), "image/png"),
            "hoja_vida_sst": ("hv.pdf", io.BytesIO(_pdf_falso("hoja de vida")), "application/pdf"),
            "licencia_sst_pdf": ("lic.pdf", io.BytesIO(_pdf_falso("licencia")), "application/pdf"),
            "cedula_sst_pdf": ("ced.pdf", io.BytesIO(_pdf_falso("cedula")), "application/pdf"),
            "diploma_cursos_sst": ("dip.pdf", io.BytesIO(_pdf_falso("diploma")), "application/pdf"),
        },
    )
    assert r.status_code == 200
    body = r.json()
    assert set(body["archivos"]) == {
        "firma_sst", "hoja_vida_sst", "licencia_sst_pdf", "cedula_sst_pdf", "diploma_cursos_sst",
    }


def test_flujo_completo_hasta_finalizar_produce_yaml_valido(client, headers):
    bid = client.post("/wizard/borradores", headers=headers).json()["borrador_id"]

    r = client.post(
        f"/wizard/borradores/{bid}/modulo-1/confirmar",
        headers=headers,
        json={
            "razon_social": "ACME PRUEBAS S.A.S", "nit": "900123456-7",
            "direccion_empresa": "Calle 1 # 2-3", "ciudad_empresa": "Bogota",
        },
    )
    assert r.status_code == 200
    assert r.json()["modulo_actual"] == 2

    r = client.post(
        f"/wizard/borradores/{bid}/modulo-2/confirmar", headers=headers,
        json={"nivel_riesgo": "III", "arl": "SURA", "numero_trabajadores_empresa": "9", "color_corporativo": "azul"},
    )
    assert r.status_code == 200

    r = client.post(
        f"/wizard/borradores/{bid}/modulo-3/confirmar", headers=headers,
        json={"sedes": [{"nombre": "Principal", "direccion": "Calle 1", "ciudad": "Bogota", "telefono": "123"}]},
    )
    assert r.status_code == 200

    r = client.post(
        f"/wizard/borradores/{bid}/modulo-4/archivos", headers=headers,
        files={"hoja_vida_sst": ("hv.pdf", io.BytesIO(_pdf_falso("hoja de vida")), "application/pdf")},
    )
    assert r.status_code == 200

    r = client.post(
        f"/wizard/borradores/{bid}/modulo-4/confirmar", headers=headers,
        json={
            "nombre_responsable_sst": "Juan Perez", "cedula_responsable_sst": "1234",
            "cargo_sst": "Profesional SST", "ciudad_cedula_sst": "Bogota",
            "licencia_sst": "1234 de 2020", "empresa_consultora_sst": "MEDINA SST EMPRESARIAL S.A.S",
            "nit_consultora_sst": "900.988.068-5", "licencia_empresa_consultora": "25-996 de 2017",
        },
    )
    assert r.status_code == 200

    r = client.post(
        f"/wizard/borradores/{bid}/modulo-5/confirmar", headers=headers,
        json={"version": "1", "año": "2026", "ciudad": "Bogota", "generar_pesv": False},
    )
    assert r.status_code == 200

    r = client.post(
        f"/wizard/borradores/{bid}/modulo-6/confirmar", headers=headers,
        json={"descripcion_empresa": "Empresa de pruebas."},
    )
    assert r.status_code == 200

    # modulo_7 es opcional -- lo saltamos, pero igual queda "confirmado" y
    # habilita el gate del modulo 8 (_verificar_gate exige el anterior).
    r = client.post(f"/wizard/borradores/{bid}/modulo-7/saltar", headers=headers)
    assert r.status_code == 200

    r = client.post(
        f"/wizard/borradores/{bid}/modulo-8/confirmar", headers=headers,
        json={"pie_pagina_numero": "izquierda", "estilo_encabezado": "2"},
    )
    assert r.status_code == 200

    r = client.post(f"/wizard/borradores/{bid}/finalizar", headers=headers)
    assert r.status_code == 200
    resultado = r.json()
    assert resultado["cliente_id"] == "ACME_PRUEBAS_SAS"
    assert "generar" in resultado["siguiente_paso"]

    # El YAML resultante debe cargar sin errores en el pipeline existente
    from pathlib import Path
    from engine.cliente import (
        cargar_cliente, generar_contexto, obtener_imagenes,
        obtener_documentos_soporte_responsable,
    )
    from engine.validacion import validar_cliente_dict

    config_path = Path(resultado["config_path"])
    assert config_path.exists()

    cliente = cargar_cliente(config_path)
    assert validar_cliente_dict(cliente, config_path.name) == []
    contexto = generar_contexto(cliente)
    assert contexto["RAZON_SOCIAL"] == "ACME PRUEBAS S.A.S"
    assert obtener_imagenes(cliente, config_path.parent) == {}  # no se subio logo/firma en este test
    assert cliente["pie_pagina_numero"] == "izquierda"
    assert cliente["estilo_encabezado"] == "2"

    # La hoja de vida subida en el modulo 4 debe quedar copiada y referenciada
    documentos = obtener_documentos_soporte_responsable(cliente, config_path.parent)
    assert "Hoja de Vida" in documentos
    assert documentos["Hoja de Vida"].exists()


def test_finalizar_sin_todos_los_modulos_da_409(client, headers):
    bid = client.post("/wizard/borradores", headers=headers).json()["borrador_id"]
    client.post(
        f"/wizard/borradores/{bid}/modulo-1/confirmar", headers=headers,
        json={"razon_social": "ACME S.A.S", "nit": "900123456-7"},
    )
    r = client.post(f"/wizard/borradores/{bid}/finalizar", headers=headers)
    assert r.status_code == 409


def _avanzar_hasta_modulo_7(client, headers) -> str:
    """Crea un borrador y confirma los modulos 1-6 con datos minimos validos.
    Retorna el borrador_id, listo para probar el modulo 7."""
    bid = client.post("/wizard/borradores", headers=headers).json()["borrador_id"]
    client.post(
        f"/wizard/borradores/{bid}/modulo-1/confirmar", headers=headers,
        json={"razon_social": "ACME CARGOS S.A.S", "nit": "900987654-3",
              "direccion_empresa": "Calle 1", "ciudad_empresa": "Bogota"},
    )
    client.post(
        f"/wizard/borradores/{bid}/modulo-2/confirmar", headers=headers,
        json={"nivel_riesgo": "III", "arl": "SURA", "numero_trabajadores_empresa": "9", "color_corporativo": "azul"},
    )
    client.post(
        f"/wizard/borradores/{bid}/modulo-3/confirmar", headers=headers,
        json={"sedes": [{"nombre": "Principal", "direccion": "Calle 1", "ciudad": "Bogota", "telefono": "123"}]},
    )
    client.post(
        f"/wizard/borradores/{bid}/modulo-4/confirmar", headers=headers,
        json={"nombre_responsable_sst": "Juan Perez", "cedula_responsable_sst": "1234",
              "cargo_sst": "Profesional SST", "ciudad_cedula_sst": "Bogota",
              "licencia_sst": "1234 de 2020", "empresa_consultora_sst": "MEDINA SST EMPRESARIAL S.A.S",
              "nit_consultora_sst": "900.988.068-5", "licencia_empresa_consultora": "25-996 de 2017"},
    )
    client.post(
        f"/wizard/borradores/{bid}/modulo-5/confirmar", headers=headers,
        json={"version": "1", "año": "2026", "ciudad": "Bogota", "generar_pesv": False},
    )
    client.post(
        f"/wizard/borradores/{bid}/modulo-6/confirmar", headers=headers,
        json={"descripcion_empresa": "Empresa de pruebas."},
    )
    return bid


def test_modulo_7_lista_sin_cargos_da_400(client, headers):
    bid = _avanzar_hasta_modulo_7(client, headers)
    r = client.post(f"/wizard/borradores/{bid}/modulo-7/lista", headers=headers, json={"cargos": []})
    assert r.status_code == 400


def test_modulo_7_lista_de_cargos_llega_al_yaml_final(client, headers):
    from pathlib import Path
    from engine.cliente import cargar_cliente

    bid = _avanzar_hasta_modulo_7(client, headers)
    r = client.post(
        f"/wizard/borradores/{bid}/modulo-7/lista", headers=headers,
        json={"cargos": [
            {"nombre": "Gerente General", "proceso": "ADMINISTRATIVO", "tipo": "administrativo", "num_expuestos": 1},
            {"nombre": "Conductor", "proceso": "LOGISTICA", "tipo": "conductor", "num_expuestos": 2},
        ]},
    )
    assert r.status_code == 200
    assert r.json()["modulo_actual"] == 8  # avanza al ultimo modulo (estilo del documento)

    r = client.post(
        f"/wizard/borradores/{bid}/modulo-8/confirmar", headers=headers,
        json={"pie_pagina_numero": "centro"},
    )
    assert r.status_code == 200

    r = client.post(f"/wizard/borradores/{bid}/finalizar", headers=headers)
    assert r.status_code == 200
    cliente = cargar_cliente(Path(r.json()["config_path"]))
    assert len(cliente["cargos"]) == 2
    assert cliente["cargos"][0]["nombre"] == "Gerente General"
    assert cliente["cargos"][1]["tipo"] == "conductor"


def test_modulo_7_subir_manuales_llega_al_yaml_final(client, headers):
    from pathlib import Path
    from engine.cliente import cargar_cliente

    bid = _avanzar_hasta_modulo_7(client, headers)
    r = client.post(
        f"/wizard/borradores/{bid}/modulo-7/manuales", headers=headers,
        files=[
            ("archivos", ("gerente.docx", io.BytesIO(b"contenido docx de prueba"),
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.document")),
        ],
    )
    assert r.status_code == 200
    assert r.json()["datos"]["modulo_7"]["manuales_subidos"] == ["gerente.docx"]

    r = client.post(
        f"/wizard/borradores/{bid}/modulo-8/confirmar", headers=headers,
        json={"pie_pagina_numero": "centro"},
    )
    assert r.status_code == 200

    r = client.post(f"/wizard/borradores/{bid}/finalizar", headers=headers)
    assert r.status_code == 200
    resultado = r.json()
    cliente = cargar_cliente(Path(resultado["config_path"]))
    assert cliente["manuales_cargos_dir"] == "ACME_CARGOS_SAS_assets/manuales_cargos"

    carpeta_manuales = Path(resultado["config_path"]).parent / cliente["manuales_cargos_dir"]
    assert (carpeta_manuales / "gerente.docx").exists()


def test_modulo_7_archivo_no_soportado_da_400(client, headers):
    bid = _avanzar_hasta_modulo_7(client, headers)
    r = client.post(
        f"/wizard/borradores/{bid}/modulo-7/manuales", headers=headers,
        files=[("archivos", ("foto.png", io.BytesIO(b"no es un manual"), "image/png"))],
    )
    assert r.status_code == 400


def test_modulo_7_saltar_no_agrega_cargos_al_yaml(client, headers):
    from pathlib import Path
    from engine.cliente import cargar_cliente

    bid = _avanzar_hasta_modulo_7(client, headers)
    r = client.post(f"/wizard/borradores/{bid}/modulo-7/saltar", headers=headers)
    assert r.status_code == 200

    r = client.post(
        f"/wizard/borradores/{bid}/modulo-8/confirmar", headers=headers,
        json={"pie_pagina_numero": "centro"},
    )
    assert r.status_code == 200

    r = client.post(f"/wizard/borradores/{bid}/finalizar", headers=headers)
    assert r.status_code == 200
    cliente = cargar_cliente(Path(r.json()["config_path"]))
    assert "cargos" not in cliente
    assert "manuales_cargos_dir" not in cliente


def test_endpoints_protegidos_requieren_api_key(client):
    r = client.post("/wizard/borradores")
    assert r.status_code == 401
