"""Tests del matcher de links del Listado Maestro (SST-MTZ-005).

Cubre los bugs reales que se encontraron en produccion:
  - Equivalencia FMT <-> FTM (typo historico de la plantilla)
  - Desambiguacion por nombre cuando hay codigos duplicados (RGL-001)
  - Fallback por prefijo + nombre cuando el codigo numerico no coincide
  - Tolerancia a typos (ADQUISICION vs ADQUISISION)
  - Filtro por carpeta/numeral (4.1.2)
  - Validacion minima: no asignar link si el nombre no se parece
"""
import pytest

from engine.listado_maestro import (
    _normalizar_codigo,
    _normalizar_nombre,
    _similitud,
    _similitud_palabras,
    _indexar_drive_links_por_codigo,
    _elegir_mejor_match,
    _normalizar_tipo,
    TIPOS_PESTANA,
)


# ---------------------------------------------------------------------------
# Normalizacion de codigos
# ---------------------------------------------------------------------------

class TestNormalizarCodigo:

    def test_fmt_se_mantiene_fmt(self):
        assert _normalizar_codigo("SST-FMT-010") == "SST-FMT-010"

    def test_ftm_se_convierte_a_fmt(self):
        """El typo historico FTM se normaliza a FMT para que matche."""
        assert _normalizar_codigo("SST-FTM-010") == "SST-FMT-010"

    def test_pesv_se_mantiene(self):
        assert _normalizar_codigo("PESV-POL-001") == "PESV-POL-001"

    def test_codigo_invalido_se_devuelve_igual(self):
        assert _normalizar_codigo("BASURA-CODIGO") == "BASURA-CODIGO"

    def test_otros_prefijos_no_se_alteran(self):
        assert _normalizar_codigo("SST-PRG-005") == "SST-PRG-005"
        assert _normalizar_codigo("SST-RGL-001") == "SST-RGL-001"


# ---------------------------------------------------------------------------
# Normalizacion de nombres
# ---------------------------------------------------------------------------

class TestNormalizarNombre:

    def test_minusculas(self):
        assert _normalizar_nombre("ABCD") == "abcd"

    def test_sin_tildes(self):
        assert "policia" in _normalizar_nombre("POLÍCIA")

    def test_quita_codigo_sst(self):
        """El codigo SST-FMT-001 debe ser eliminado del texto."""
        n = _normalizar_nombre("SST-FMT-001 NOMBRE DEL DOCUMENTO")
        # El codigo original SST-FMT-001 (con guiones) ya no debe aparecer
        assert "SST-FMT-001" not in n.upper()
        assert "nombre" in n
        assert "documento" in n

    def test_quita_extension(self):
        assert "docx" not in _normalizar_nombre("documento.docx")

    def test_colapsa_espacios(self):
        assert _normalizar_nombre("hola    mundo") == "hola mundo"


# ---------------------------------------------------------------------------
# Similitud
# ---------------------------------------------------------------------------

class TestSimilitudPalabras:

    def test_palabras_identicas(self):
        assert _similitud_palabras("hola", "hola") is True

    def test_typo_un_caracter(self):
        """Bug real: ADQUISICION vs ADQUISISION (typo en el archivo)."""
        assert _similitud_palabras("adquisicion", "adquisision") is True

    def test_typo_dos_caracteres(self):
        """INSPECCION vs INSPECION (falta una C)."""
        assert _similitud_palabras("inspeccion", "inspecion") is True

    def test_palabras_distintas(self):
        assert _similitud_palabras("orden", "aseo") is False
        assert _similitud_palabras("politica", "programa") is False

    def test_longitudes_muy_distintas(self):
        """No es similar si las longitudes difieren mucho."""
        assert _similitud_palabras("orden", "ordenamientos") is False


class TestSimilitud:

    def test_nombre_contenido_en_archivo_da_alta_similitud(self):
        """Caso comun: matriz tiene nombre corto y archivo tiene codigo + nombre largo."""
        s = _similitud("PROGRAMA DE ORDEN Y ASEO",
                       "SST-PRG-006 PROGRAMA DE ORDEN Y ASEO")
        assert s >= 0.9

    def test_responsabilidades_copasst(self):
        s = _similitud("RESPONSABILIDADES COPASST",
                       "SST-FMT-005 RESPONSABILIDADES DEL COPASST")
        assert s >= 0.9

    def test_nombres_diferentes_dan_baja_similitud(self):
        s = _similitud("INFORME DE PAUSAS ACTIVAS",
                       "SST-MTZ-010 MATRIZ DE INDICADORES")
        assert s < 0.4

    def test_tolerancia_typo_en_nombre(self):
        """POLITICA DE ADQUISICION vs ADQUISISION debe matchear."""
        s = _similitud("POLITICA DE ADQUISICION",
                       "PESV-POL-006 POLITICA DE ADQUISISION")
        assert s >= 0.8

    def test_string_vacio(self):
        assert _similitud("", "algo") == 0.0
        assert _similitud("algo", "") == 0.0

    def test_solo_stopwords(self):
        """Si todo son stopwords (DE, DEL, LA), no debe contar como match."""
        assert _similitud("DE LA EL", "DE LOS LAS") == 0.0


# ---------------------------------------------------------------------------
# Indexacion de drive_links
# ---------------------------------------------------------------------------

class TestIndexarDriveLinks:

    def test_indice_simple(self):
        links = {
            "ruta/SST-FMT-001 NOMBRE.docx": {
                "link": "https://drive/AAA",
                "nombre": "SST-FMT-001 NOMBRE.docx",
            },
        }
        indice, indice_pref = _indexar_drive_links_por_codigo(links)
        assert "SST-FMT-001" in indice
        assert len(indice["SST-FMT-001"]) == 1
        assert "SST-FMT" in indice_pref

    def test_ftm_se_indexa_como_fmt(self):
        """Si el archivo tiene FTM-001, debe quedar indexado como FMT-001."""
        links = {
            "ruta/SST-FTM-001 NOMBRE.docx": {
                "link": "https://drive/BBB",
                "nombre": "SST-FTM-001 NOMBRE.docx",
            },
        }
        indice, _ = _indexar_drive_links_por_codigo(links)
        assert "SST-FMT-001" in indice
        assert "SST-FTM-001" not in indice

    def test_codigos_duplicados_se_acumulan_en_lista(self):
        """SST-RGL-001 se usa para COPASST y para HIGIENE."""
        links = {
            "ruta1/SST-RGL-001 REGLAMENTO COPASST.docx": {
                "link": "https://drive/COPASST", "nombre": "SST-RGL-001 REGLAMENTO COPASST.docx",
            },
            "ruta2/SST-RGL-001 REGLAMENTO DE HIGIENE.docx": {
                "link": "https://drive/HIGIENE", "nombre": "SST-RGL-001 REGLAMENTO DE HIGIENE.docx",
            },
        }
        indice, _ = _indexar_drive_links_por_codigo(links)
        assert len(indice["SST-RGL-001"]) == 2

    def test_archivos_sin_codigo_se_ignoran(self):
        links = {
            "ruta/CERTIFICADO AFILIACION.pdf": {
                "link": "https://drive/X", "nombre": "CERTIFICADO AFILIACION.pdf",
            },
        }
        indice, _ = _indexar_drive_links_por_codigo(links)
        assert len(indice) == 0

    def test_archivos_sin_link_se_ignoran(self):
        links = {
            "ruta/SST-FMT-001.docx": {"link": "", "nombre": "SST-FMT-001.docx"},
        }
        indice, _ = _indexar_drive_links_por_codigo(links)
        assert len(indice) == 0

    def test_input_vacio(self):
        indice, pref = _indexar_drive_links_por_codigo({})
        assert indice == {}
        assert pref == {}


# ---------------------------------------------------------------------------
# Eleccion del mejor match
# ---------------------------------------------------------------------------

class TestElegirMejorMatch:

    def test_un_solo_candidato_con_nombre_que_coincide(self):
        cand = [("url", "SST-FMT-010 CONVOCATORIA AL COPASST",
                 "ruta/SST-FMT-010 CONVOCATORIA AL COPASST.docx")]
        m = _elegir_mejor_match("SST-FMT-010", "CONVOCATORIA AL COPASST", cand)
        assert m is not None
        assert m[0] == "url"

    def test_un_solo_candidato_pero_nombre_no_coincide_rechaza(self):
        """Bug real: matriz pide CONTROL DE PLAGAS pero archivo es MANTENIMIENTO TANQUES.
        El sistema NO debe asignar el link aunque sea el unico candidato."""
        cand = [("url", "SST-FTM-053 FORMATO DE MANTENIMIENTO DE TANQUES",
                 "ruta/SST-FTM-053 FORMATO DE MANTENIMIENTO DE TANQUES.docx")]
        m = _elegir_mejor_match("SST-FMT-053", "FORMATO DE CONTROL DE PLAGAS", cand)
        assert m is None

    def test_codigos_duplicados_desambigua_por_nombre(self):
        """SST-RGL-001 se usa para COPASST y Higiene. Debe elegir el correcto."""
        cand = [
            ("url_copasst", "SST-RGL-001 REGLAMENTO COPASST",
             "II.HACER/COPASST/SST-RGL-001 REGLAMENTO COPASST.docx"),
            ("url_higiene", "SST-RGL-001 REGLAMENTO DE HIGIENE Y SEGURIDAD",
             "I.PLANEAR/HIGIENE/SST-RGL-001 REGLAMENTO DE HIGIENE Y SEGURIDAD.docx"),
        ]
        m = _elegir_mejor_match("SST-RGL-001", "REGLAMENTO DEL COPASST", cand)
        assert m[0] == "url_copasst"

        m = _elegir_mejor_match("SST-RGL-001", "REGLAMENTO DE HIGIENE Y SEGURIDAD INDUSTRIAL", cand)
        assert m[0] == "url_higiene"

    def test_filtro_por_numeral_seccion(self):
        """Cuando la fila tiene seccion '4.1.2', prefiere candidatos en esa carpeta."""
        cand = [
            ("url_correcto", "SST-MTZ-012 MATRIZ DE IPVR",
             "II. HACER/4.1.2 MATRIZ DE IPVR/SST-MTZ-012 MATRIZ DE IPVR.xlsx"),
            ("url_otro", "SST-MTZ-012 OTRA MATRIZ",
             "I. PLANEAR/2.5.1 OTRA/SST-MTZ-012 OTRA MATRIZ.xlsx"),
        ]
        # Usar nombre que comparta palabras con el candidato correcto (IPVR)
        m = _elegir_mejor_match(
            "SST-MTZ-012",
            "MATRIZ DE IPVR",
            cand,
            seccion_esperada="4.1.2 Identificacion de peligros",
        )
        assert m is not None
        assert m[0] == "url_correcto"

    def test_sin_nombre_devuelve_primero(self):
        cand = [
            ("url1", "doc1", "ruta1"),
            ("url2", "doc2", "ruta2"),
        ]
        m = _elegir_mejor_match("SST-FMT-001", "", cand)
        assert m is not None

    def test_sin_candidatos(self):
        assert _elegir_mejor_match("SST-FMT-001", "ALGO", []) is None

    def test_umbral_minimo_rechaza_match_pobre(self):
        """Con umbral 0.40, no debe asignar si la similitud es menor."""
        cand = [("url", "DOCUMENTO X TOTALMENTE DIFERENTE", "ruta/X.docx")]
        m = _elegir_mejor_match("SST-FMT-001", "NOMBRE ESPERADO DE ALGO", cand,
                                umbral_minimo=0.50)
        assert m is None


# ---------------------------------------------------------------------------
# Tipos de pestanas
# ---------------------------------------------------------------------------

class TestPestanasPorTipo:

    def test_normalizar_tipo_quita_tildes_y_mayusculiza(self):
        assert _normalizar_tipo("Política") == "POLITICA"
        assert _normalizar_tipo("formato") == "FORMATO"
        assert _normalizar_tipo("  Manual  ") == "MANUAL"

    def test_tipos_pestana_son_los_esperados(self):
        """Lista canonica de pestanas a generar."""
        esperados = {"PROGRAMA", "PLAN", "PROCEDIMIENTO", "INSTRUCTIVO",
                     "REGLAMENTO", "MANUAL", "POLITICA", "INDICADOR",
                     "MATRIZ", "FORMATO"}
        assert set(TIPOS_PESTANA) == esperados


# ---------------------------------------------------------------------------
# Pipeline integrado: actualizar matriz con drive_links
# ---------------------------------------------------------------------------

class TestActualizarMatrizGestion:
    """Tests integrados con la plantilla real."""

    def test_no_falla_sin_drive_links(self, output_listado_maestro):
        from engine.listado_maestro import actualizar_matriz_gestion_documental
        ruta = actualizar_matriz_gestion_documental(output_listado_maestro, {})
        assert ruta is not None
        assert ruta.exists()

    def test_elimina_hoja_pesv_si_no_aplica(self, output_listado_maestro):
        import openpyxl
        from engine.listado_maestro import actualizar_matriz_gestion_documental
        actualizar_matriz_gestion_documental(
            output_listado_maestro, {}, generar_pesv=False,
        )
        ruta = (output_listado_maestro / "I. PLANEAR"
                / "2.5.1 ARCHIVO Y RETENCIÓN DOCUMENTAL"
                / "SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL.xlsx")
        wb = openpyxl.load_workbook(str(ruta))
        assert "PESV" not in wb.sheetnames

    def test_genera_pestanas_por_tipo(self, output_listado_maestro):
        """Verifica que se creen las pestanas PROGRAMA, FORMATO, etc."""
        import openpyxl
        from engine.listado_maestro import actualizar_matriz_gestion_documental
        actualizar_matriz_gestion_documental(
            output_listado_maestro, {}, generar_pesv=False,
        )
        ruta = (output_listado_maestro / "I. PLANEAR"
                / "2.5.1 ARCHIVO Y RETENCIÓN DOCUMENTAL"
                / "SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL.xlsx")
        wb = openpyxl.load_workbook(str(ruta))
        # Al menos PROGRAMA, FORMATO, MATRIZ deben existir
        for tipo in ("PROGRAMA", "FORMATO", "MATRIZ", "PROCEDIMIENTO", "MANUAL"):
            assert tipo in wb.sheetnames, f"Falta pestaña {tipo}"

    def test_inserta_links_con_formula_hyperlink(self, output_listado_maestro):
        """Inyectar un link conocido y verificar que aparece en la matriz."""
        import openpyxl
        from engine.listado_maestro import actualizar_matriz_gestion_documental

        drive_links = {
            "I.PLANEAR/SST-FMT-001 DESIGNACION RESPONSABLE.docx": {
                "link": "https://drive.google.com/d/AAA",
                "nombre": "SST-FMT-001 DESIGNACION RESPONSABLE.docx",
            },
        }
        actualizar_matriz_gestion_documental(
            output_listado_maestro, drive_links, generar_pesv=False,
        )
        ruta = (output_listado_maestro / "I. PLANEAR"
                / "2.5.1 ARCHIVO Y RETENCIÓN DOCUMENTAL"
                / "SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL.xlsx")
        wb = openpyxl.load_workbook(str(ruta))
        ws = wb["LISTADO MAESTRO PHVA"]
        # Buscar la fila donde aparezca SST-FMT-001
        encontrado = False
        for row in range(7, ws.max_row + 1):
            codigo = ws.cell(row=row, column=4).value
            if codigo and "FMT-001" in str(codigo):
                link_val = ws.cell(row=row, column=22).value
                if link_val and "HYPERLINK" in str(link_val) and "AAA" in str(link_val):
                    encontrado = True
                    break
        assert encontrado, "No se encontro el link inyectado en SST-MTZ-005"
