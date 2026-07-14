"""Tests de engine/generador.py::_replicar_carpetas_vacias().

Las carpetas de estandares que aun no tienen ningun documento cargado en la
plantilla (ej. "1.1.4 AFILIACION AL SGRL") no aparecian en la salida del
cliente -- el encargado no tenia forma de saber que ese estandar existe y
falta subir el archivo. Detectado en conversacion real (2026-07-14), al
comparar la tabla oficial de estandares minimos (Res. 0312/2019) contra la
salida generada."""
from pathlib import Path

import engine.generador as generador


def _crear_plantilla(tmp_path: Path) -> Path:
    plantilla = tmp_path / "Plantilla"
    (plantilla / "I. PLANEAR" / "1.1.1 CON ARCHIVO").mkdir(parents=True)
    (plantilla / "I. PLANEAR" / "1.1.1 CON ARCHIVO" / "doc.docx").write_text("x")
    (plantilla / "I. PLANEAR" / "1.1.4 SIN ARCHIVO").mkdir(parents=True)
    (plantilla / "I. PLANEAR" / "9.9.9 NO APLICA A LA CATEGORIA").mkdir(parents=True)
    return plantilla


def test_crea_carpeta_vacia_solo_si_el_estandar_aplica_a_la_categoria(tmp_path, monkeypatch):
    plantilla = _crear_plantilla(tmp_path)
    monkeypatch.setattr(generador, "PLANTILLA_PATH", plantilla)
    output_base = tmp_path / "salida"

    creadas = generador._replicar_carpetas_vacias(
        output_base, usar_vigia=False, categoria="I",
        estandares_validos=frozenset({"1.1.1", "1.1.4"}),
    )

    assert creadas == 1
    assert (output_base / "I. PLANEAR" / "1.1.4 SIN ARCHIVO").is_dir()
    assert not (output_base / "I. PLANEAR" / "9.9.9 NO APLICA A LA CATEGORIA").exists()
    # La carpeta con archivo la crea el loop principal, no esta funcion.
    assert not (output_base / "I. PLANEAR" / "1.1.1 CON ARCHIVO").exists()


def test_categoria_iii_sin_filtro_crea_todas_las_carpetas_vacias(tmp_path, monkeypatch):
    plantilla = _crear_plantilla(tmp_path)
    monkeypatch.setattr(generador, "PLANTILLA_PATH", plantilla)
    output_base = tmp_path / "salida"

    creadas = generador._replicar_carpetas_vacias(
        output_base, usar_vigia=False, categoria="III", estandares_validos=None,
    )

    assert creadas == 2
    assert (output_base / "I. PLANEAR" / "1.1.4 SIN ARCHIVO").is_dir()
    assert (output_base / "I. PLANEAR" / "9.9.9 NO APLICA A LA CATEGORIA").is_dir()
