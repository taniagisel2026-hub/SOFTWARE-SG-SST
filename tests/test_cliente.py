"""Tests de engine/cliente.py::generar_contexto()."""
import os
from pathlib import Path

from engine.cliente import generar_contexto, obtener_imagenes


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


def test_obtener_imagenes_devuelve_ruta_absoluta_con_config_dir_relativo(tmp_path, monkeypatch, cliente_minimo):
    """BUG: si el CLI se invoca con una ruta relativa (ej. "clientes_config\\X.yaml",
    tal como lo documenta CLAUDE.md), config_dir tambien queda relativo y la ruta de
    imagen guardada en el contexto ("clientes_config\\X_assets\\logo.jpeg") no es
    absoluta. docx_engine.resolve_image_path() solo usa la ruta tal cual cuando
    es absoluta -- si no, la busca junto a la plantilla en vez de junto al YAML del
    cliente y nunca la encuentra, dejando el logo/firma como texto plano en el
    documento en vez de insertar la imagen.

    Detectado en generacion real (2026-07-14, cliente EMPRESA_DE_PRUEBA_SAS)."""
    monkeypatch.chdir(tmp_path)
    config_dir = tmp_path / "clientes_config"
    assets_dir = config_dir / "X_assets"
    assets_dir.mkdir(parents=True)
    (assets_dir / "logo.jpeg").write_bytes(b"fake-image-data")

    cliente_minimo["logo"] = os.path.join("X_assets", "logo.jpeg")
    imagenes = obtener_imagenes(cliente_minimo, config_dir=Path("clientes_config"))

    assert os.path.isabs(imagenes["LOGO"])
    assert Path(imagenes["LOGO"]) == (assets_dir / "logo.jpeg").resolve()
