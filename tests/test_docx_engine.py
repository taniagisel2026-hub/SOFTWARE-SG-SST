"""Tests de engine/docx_engine.py: reemplazo de marcadores {{}}, deteccion
de texto humanizable, y preservacion de negrillas al aplicar humanizacion.

Este es el motor central del pipeline (cli.py -> generador.py ->
docx_engine.py, ver CLAUDE.md) y hasta ahora no tenia ninguna prueba
propia.
"""
from pathlib import Path
from zipfile import ZipFile

import docx
import pytest

from engine.docx_engine import (
    _aplicar_humanizacion_post,
    _autonegrilla_termino_definicion,
    _dividir_por_negrillas,
    _es_parrafo_humanizable,
    _parsear_markdown_bold,
    extraer_secciones_contenido,
    extraer_texto_docx,
    listar_marcadores_en_docx,
    procesar_docx,
    verificar_residuos_docx,
)


# ---------------------------------------------------------------------------
# Helpers para construir documentos .docx de prueba
# ---------------------------------------------------------------------------

def _doc_con_parrafos(path: Path, textos: list, estilos: list = None) -> None:
    doc = docx.Document()
    estilos = estilos or [None] * len(textos)
    for texto, estilo in zip(textos, estilos):
        doc.add_paragraph(texto, style=estilo)
    doc.save(str(path))


def _doc_con_tabla(path: Path, celdas: list) -> None:
    doc = docx.Document()
    tabla = doc.add_table(rows=1, cols=len(celdas))
    for cell, texto in zip(tabla.rows[0].cells, celdas):
        cell.text = texto
    doc.save(str(path))


def _png_1x1() -> bytes:
    """PNG 1x1 minimo pero valido -- suficiente para que docxtpl/Pillow lo
    abran como InlineImage sin necesitar un logo real."""
    import io
    from PIL import Image as PILImage
    buf = io.BytesIO()
    PILImage.new("RGB", (1, 1), color="white").save(buf, format="PNG")
    return buf.getvalue()


class TestProcesarDocxReemplazaMarcadores:
    def test_reemplaza_marcador_en_parrafo(self, tmp_path):
        origen = tmp_path / "plantilla.docx"
        _doc_con_parrafos(origen, ["Empresa: {{RAZON_SOCIAL}}, NIT: {{NIT}}"])
        destino = tmp_path / "salida.docx"

        resultado = procesar_docx(
            origen, destino, {"RAZON_SOCIAL": "ACME S.A.S", "NIT": "900123456-7"}, str(tmp_path)
        )

        doc_salida = docx.Document(str(destino))
        assert doc_salida.paragraphs[0].text == "Empresa: ACME S.A.S, NIT: 900123456-7"
        assert resultado["faltantes"] == []
        assert resultado["residuos"] == []

    def test_reemplaza_marcador_en_tabla(self, tmp_path):
        origen = tmp_path / "plantilla.docx"
        _doc_con_tabla(origen, ["{{RAZON_SOCIAL}}", "{{NIVEL_RIESGO}}"])
        destino = tmp_path / "salida.docx"

        procesar_docx(origen, destino, {"RAZON_SOCIAL": "ACME S.A.S", "NIVEL_RIESGO": "III"}, str(tmp_path))

        celdas = docx.Document(str(destino)).tables[0].rows[0].cells
        assert celdas[0].text == "ACME S.A.S"
        assert celdas[1].text == "III"

    def test_marcador_sin_valor_en_contexto_queda_vacio_no_lanza_error(self, tmp_path):
        """docxtpl usa Jinja2: una variable no declarada en el contexto se
        renderiza como cadena vacia (Undefined -> ''), no queda el marcador
        literal ni truena el render. Se reporta en 'faltantes'."""
        origen = tmp_path / "plantilla.docx"
        _doc_con_parrafos(origen, ["Empresa: {{RAZON_SOCIAL}}, Sede: {{SEDE_NO_DEFINIDA}}"])
        destino = tmp_path / "salida.docx"

        resultado = procesar_docx(origen, destino, {"RAZON_SOCIAL": "ACME S.A.S"}, str(tmp_path))

        assert docx.Document(str(destino)).paragraphs[0].text == "Empresa: ACME S.A.S, Sede: "
        assert "SEDE_NO_DEFINIDA" in resultado["faltantes"]

    def test_marcador_con_tilde_es_distinto_del_marcador_ascii(self, tmp_path):
        """Jinja2 SI soporta identificadores unicode: {{RAZÓN_SOCIAL}} y
        {{RAZON_SOCIAL}} son variables distintas para docxtpl -- por eso
        engine/cliente.py debe expandir el contexto con ambas variantes
        (ver CLAUDE.md, seccion 'Key flow')."""
        origen = tmp_path / "plantilla.docx"
        _doc_con_parrafos(origen, ["{{RAZÓN_SOCIAL}} vs {{RAZON_SOCIAL}}"])
        destino = tmp_path / "salida.docx"

        procesar_docx(
            origen, destino,
            {"RAZÓN_SOCIAL": "Con tilde", "RAZON_SOCIAL": "Sin tilde"},
            str(tmp_path),
        )
        assert docx.Document(str(destino)).paragraphs[0].text == "Con tilde vs Sin tilde"

    def test_llaves_mal_cerradas_copia_el_original_sin_tronar(self, tmp_path):
        origen = tmp_path / "plantilla.docx"
        _doc_con_parrafos(origen, ["Bloque roto: {% if X %} sin cierre {{RAZON_SOCIAL}}"])
        destino = tmp_path / "salida.docx"

        resultado = procesar_docx(origen, destino, {"RAZON_SOCIAL": "ACME"}, str(tmp_path))

        assert destino.exists()
        assert "error_render" in resultado
        assert resultado["residuos"] == ["Documento copiado intacto por error de sintaxis en llaves"]


class TestProcesarDocxResuelveImagenes:
    def test_logo_se_resuelve_desde_carpeta_de_la_plantilla(self, tmp_path):
        origen = tmp_path / "plantilla.docx"
        _doc_con_parrafos(origen, ["Logo: {{LOGO}}"])
        destino = tmp_path / "salida.docx"
        (tmp_path / "logo.png").write_bytes(_png_1x1())

        resultado = procesar_docx(origen, destino, {"LOGO": "logo.png"}, str(tmp_path))

        assert destino.exists()
        assert resultado["faltantes"] == []
        with ZipFile(destino) as z:
            assert any(n.startswith("word/media/") for n in z.namelist())

    def test_logo_sin_archivo_existente_no_lanza_error(self, tmp_path):
        """Si la ruta de imagen no se puede resolver (archivo no existe),
        _resolver_imagenes_en_contexto no reemplaza el valor por InlineImage
        y lo deja como string sin tocar -- Jinja2 lo imprime tal cual (texto
        plano con la ruta), no queda en 'faltantes' porque la clave SI esta
        en el contexto (solo que no se pudo convertir a imagen). El render
        no debe tronar en ningun caso."""
        origen = tmp_path / "plantilla.docx"
        _doc_con_parrafos(origen, ["Logo: {{LOGO}}"])
        destino = tmp_path / "salida.docx"

        resultado = procesar_docx(origen, destino, {"LOGO": "no_existe.png"}, str(tmp_path))

        assert destino.exists()
        assert "error_render" not in resultado
        assert docx.Document(str(destino)).paragraphs[0].text == "Logo: no_existe.png"


class TestVerificarResiduosDocx:
    def test_detecta_marcador_sin_reemplazar(self, tmp_path):
        # Documento crudo (sin pasar por docxtpl) con un {{...}} literal --
        # simula el caso real de un marcador que quedo sin fusionar/reemplazar.
        p = tmp_path / "con_residuo.docx"
        _doc_con_parrafos(p, ["Texto con {{MARCADOR_HUERFANO}} sin reemplazar"])
        residuos = verificar_residuos_docx(str(p))
        assert any("MARCADOR_HUERFANO" in r for r in residuos)

    def test_sin_residuos_si_no_hay_llaves(self, tmp_path):
        p = tmp_path / "limpio.docx"
        _doc_con_parrafos(p, ["Texto limpio sin marcadores"])
        assert verificar_residuos_docx(str(p)) == []


class TestListarMarcadoresEnDocx:
    def test_detecta_marcadores_en_parrafos_y_tablas(self, tmp_path):
        p = tmp_path / "doc.docx"
        doc = docx.Document()
        doc.add_paragraph("{{RAZON_SOCIAL}} - {{NIT}}")
        tabla = doc.add_table(rows=1, cols=1)
        tabla.rows[0].cells[0].text = "{{NIVEL_RIESGO}}"
        doc.save(str(p))

        assert listar_marcadores_en_docx(p) == ["{{NIT}}", "{{NIVEL_RIESGO}}", "{{RAZON_SOCIAL}}"]

    def test_no_detecta_marcadores_con_tilde(self, tmp_path):
        """El regex de esta funcion (_PATRON_MARCADOR = [A-ZÑ_]+) no incluye
        vocales acentuadas, a diferencia de docxtpl/Jinja2 que si reconoce
        {{RAZÓN_SOCIAL}} como variable (ver
        TestProcesarDocxReemplazaMarcadores::test_marcador_con_tilde...).
        Esta funcion solo la usa `cli.py probar-doc` para depuracion manual,
        no el pipeline de generacion real -- se documenta el comportamiento
        actual tal como esta, no se considera un bug del pipeline principal."""
        p = tmp_path / "doc.docx"
        _doc_con_parrafos(p, ["{{RAZÓN_SOCIAL}}"])
        assert listar_marcadores_en_docx(p) == []


class TestExtraerTextoDocx:
    def test_extrae_texto_de_parrafos_y_tablas(self, tmp_path):
        p = tmp_path / "doc.docx"
        doc = docx.Document()
        doc.add_paragraph("Primer parrafo")
        tabla = doc.add_table(rows=1, cols=1)
        tabla.rows[0].cells[0].text = "Celda de la tabla"
        doc.save(str(p))

        texto = extraer_texto_docx(p)
        assert "Primer parrafo" in texto
        assert "Celda de la tabla" in texto


class TestEsParrafoHumanizable:
    def test_texto_muy_corto_no_es_humanizable(self):
        assert _es_parrafo_humanizable("Texto corto") is False

    def test_texto_largo_normal_es_humanizable(self):
        texto = "Esta es una politica de seguridad y salud en el trabajo con longitud suficiente."
        assert _es_parrafo_humanizable(texto) is True

    def test_titulo_numerado_no_es_humanizable(self):
        assert _es_parrafo_humanizable("7.6.1 RESPONSABILIDADES DEL EMPLEADOR EN EL SISTEMA") is False

    def test_titulo_todo_mayusculas_no_es_humanizable(self):
        assert _es_parrafo_humanizable("POLITICA DE SEGURIDAD Y SALUD EN EL TRABAJO") is False

    def test_parrafo_solo_con_marcadores_no_es_humanizable(self):
        assert _es_parrafo_humanizable("{{RAZON_SOCIAL}} {{NIT}} {{FECHA}}") is False

    def test_excluido_por_estilo_heading(self):
        texto = "Este parrafo tiene texto largo pero su estilo es un heading asi que se excluye igual"
        assert _es_parrafo_humanizable(texto, style_name="Heading 1") is False


class TestExtraerSeccionesContenido:
    def test_agrupa_parrafos_humanizables_en_una_seccion(self, tmp_path):
        p = tmp_path / "doc.docx"
        _doc_con_parrafos(p, [
            "TITULO DE LA POLITICA",
            "Primer parrafo de contenido con longitud suficiente para humanizar en la politica.",
            "Segundo parrafo de contenido, tambien con longitud suficiente para ser humanizado.",
        ])
        secciones = extraer_secciones_contenido(p)
        assert len(secciones) == 1
        assert len(secciones[0]["parrafos"]) == 2

    def test_bibliografia_no_se_incluye_en_secciones(self, tmp_path):
        p = tmp_path / "doc.docx"
        _doc_con_parrafos(p, [
            "Este es un parrafo normal con longitud suficiente para humanizarlo sin problema.",
            "BIBLIOGRAFIA",
            "Decreto 1072 de 2015. Ministerio del Trabajo. Republica de Colombia, texto legal.",
            "Resolucion 0312 de 2019. Ministerio del Trabajo. Republica de Colombia, texto legal.",
        ])
        textos = [s["texto_completo"] for s in extraer_secciones_contenido(p)]
        assert not any("Decreto 1072" in t for t in textos)
        assert any("parrafo normal" in t for t in textos)

    def test_seccion_despues_de_bibliografia_si_se_incluye(self, tmp_path):
        p = tmp_path / "doc.docx"
        _doc_con_parrafos(p, [
            "BIBLIOGRAFIA",
            "Decreto 1072 de 2015, texto de referencia legal que no debe humanizarse nunca.",
            "11. CONTROL DE CAMBIOS",
            "Este parrafo si debe humanizarse porque ya salimos de la zona de bibliografia completa.",
        ])
        textos = [s["texto_completo"] for s in extraer_secciones_contenido(p)]
        assert any("si debe humanizarse" in t for t in textos)
        assert not any("Decreto 1072" in t for t in textos)

    def test_marcador_de_ruta_de_imagen_no_se_expone_como_texto(self, tmp_path):
        """extraer_secciones_contenido no debe filtrar rutas de imagen
        (logo/firma) hacia el texto que se manda a Claude, porque Claude las
        reproduciria como texto visible en el documento humanizado."""
        p = tmp_path / "doc.docx"
        _doc_con_parrafos(p, [
            "Firma del responsable: {{FIRMA_RESPONSABLE}} y descripcion larga de la politica interna.",
        ])
        secciones = extraer_secciones_contenido(p, marcadores={"FIRMA_RESPONSABLE": "firma.png"})
        assert not any("firma.png" in s["texto_completo"] for s in secciones)


class TestParsearMarkdownBold:
    def test_segmenta_negrita_simple(self):
        segmentos = _parsear_markdown_bold("La empresa **ACME S.A.S** cumple con la norma vigente.")
        assert ("ACME S.A.S", True) in segmentos
        assert any(not b and "cumple con la norma vigente" in t for t, b in segmentos)

    def test_sin_negrita_devuelve_un_solo_segmento(self):
        assert _parsear_markdown_bold("Texto sin negrillas") == [("Texto sin negrillas", False)]

    def test_asteriscos_sueltos_no_generan_basura(self):
        segmentos = _parsear_markdown_bold("Texto normal ** con asterisco suelto")
        assert all("**" not in t for t, _ in segmentos)


class TestAutonegrillaTerminoDefinicion:
    def test_envuelve_termino_de_definicion_en_negrita(self):
        texto = "Riesgo locativo: caidas al mismo nivel por pisos humedos o desorden en el area"
        resultado = _autonegrilla_termino_definicion(texto)
        assert resultado.startswith("**Riesgo locativo:**")

    def test_no_envuelve_lista_numerada(self):
        texto = "1. Este es un item de una lista numerada con longitud suficiente para el chequeo"
        assert _autonegrilla_termino_definicion(texto) == texto

    def test_no_envuelve_si_ya_tiene_negrita(self):
        texto = "**Riesgo locativo:** ya viene con negrita desde antes de llamar la funcion"
        assert _autonegrilla_termino_definicion(texto) == texto


class TestDividirPorNegrillas:
    def test_marca_solo_las_frases_que_estaban_en_negrita_originalmente(self):
        segmentos = _dividir_por_negrillas(
            "La empresa ACME S.A.S cumple con el Decreto 1072 de 2015 sin excepcion.",
            frases_bold=["ACME S.A.S"],
        )
        assert ("ACME S.A.S", True) in segmentos
        assert not any(b for t, b in segmentos if "Decreto 1072" in t)


class TestAplicarHumanizacionPostPreservaNegrillas:
    def test_reemplaza_texto_y_preserva_negrita_de_la_razon_social(self, tmp_path):
        p = tmp_path / "doc.docx"
        doc = docx.Document()
        parrafo = doc.add_paragraph()
        parrafo.add_run("ACME S.A.S").bold = True
        parrafo.add_run(" es una empresa comprometida con la seguridad de sus trabajadores.")
        doc.save(str(p))

        secciones = [{
            "parrafos": [{"idx": 0, "texto": "..."}],
            "texto_humanizado": "En **ACME S.A.S** creemos firmemente en el bienestar de nuestro equipo.",
        }]
        _aplicar_humanizacion_post(str(p), secciones)

        parrafo2 = docx.Document(str(p)).paragraphs[0]
        assert parrafo2.text == "En ACME S.A.S creemos firmemente en el bienestar de nuestro equipo."
        assert any(r.bold and r.text == "ACME S.A.S" for r in parrafo2.runs)

    def test_parrafo_sin_lineas_humanizadas_conserva_texto_original(self, tmp_path):
        p = tmp_path / "doc.docx"
        _doc_con_parrafos(p, ["Texto original que no debe cambiar si no hay texto humanizado."])

        _aplicar_humanizacion_post(str(p), [{
            "parrafos": [{"idx": 0, "texto": "..."}],
            "texto_humanizado": "",
        }])

        assert docx.Document(str(p)).paragraphs[0].text == (
            "Texto original que no debe cambiar si no hay texto humanizado."
        )
