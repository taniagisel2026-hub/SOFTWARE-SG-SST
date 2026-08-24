"""Tests de la Matriz Vial PESV-MTZ-002."""
import pytest

from engine.matriz_peligros_viales import (
    _inferir_actores_viales,
    _calcular_nivel_riesgo,
    CATALOGO,
    ORDEN_FACTORES,
    RUTA_MATRIZ_VIAL,
    _FILL_RIESGO_CRITICO,
    _FILL_RIESGO_MODERADO,
    _FILL_RIESGO_BAJO,
)


# ---------------------------------------------------------------------------
# Inferencia de actor vial por cargo
# ---------------------------------------------------------------------------

class TestInferirActoresViales:

    def test_conductor_de_vehiculo(self):
        cargo = {"nombre": "Conductor de Camion",
                 "actividades": "Conduccion de camion para entregas"}
        assert "conductor" in _inferir_actores_viales(cargo)

    def test_motociclista(self):
        cargo = {"nombre": "Mensajero Motorizado",
                 "actividades": "Entregas en motocicleta"}
        assert "motociclista" in _inferir_actores_viales(cargo)

    def test_peaton_por_diligencias(self):
        cargo = {"nombre": "Auxiliar de Tramites",
                 "actividades": "Realiza diligencias externas en bancos"}
        assert "peaton" in _inferir_actores_viales(cargo)

    def test_pasajero_por_visitas_comerciales(self):
        cargo = {"nombre": "Asesor Comercial Externo",
                 "actividades": "Visitas comerciales a clientes mayoristas"}
        assert "pasajero" in _inferir_actores_viales(cargo)

    def test_admin_sin_desplazamiento_no_aplica(self):
        """Bug real: Coordinador TH no debe clasificarse como ningun actor vial."""
        cargo = {"nombre": "Coordinador de Talento Humano",
                 "actividades": "Procesos de seleccion, contratacion, nomina y bienestar"}
        actores = _inferir_actores_viales(cargo)
        assert actores == []

    def test_motociclista_no_se_clasifica_tambien_como_peaton(self):
        """Bug real: 'Auxiliar de Reparto Motorizado' tenia AUXILIAR y se
        clasificaba como peaton ademas de motociclista. Solo debe ser
        motociclista."""
        cargo = {"nombre": "Auxiliar de Reparto Motorizado",
                 "actividades": "Entregas de pedidos en motocicleta"}
        actores = _inferir_actores_viales(cargo)
        assert "motociclista" in actores
        assert "peaton" not in actores

    def test_conductor_no_es_motociclista(self):
        """Conductor de camion no debe ser motociclista."""
        cargo = {"nombre": "Conductor de Camion",
                 "actividades": "Conduccion de vehiculo de carga"}
        actores = _inferir_actores_viales(cargo)
        assert "conductor" in actores
        assert "motociclista" not in actores


# ---------------------------------------------------------------------------
# Calculo del nivel de riesgo (semaforo)
# ---------------------------------------------------------------------------

class TestCalcularNivelRiesgo:

    def test_critico_alta_probabilidad_y_exposicion(self):
        nivel, texto, fill = _calcular_nivel_riesgo("Muy Probable", "Frecuente")
        assert nivel == 9
        assert texto == "CRITICO"
        assert fill is _FILL_RIESGO_CRITICO

    def test_moderado_probabilidad_media(self):
        nivel, texto, fill = _calcular_nivel_riesgo("Poco Probable", "Ocasional")
        assert nivel == 4
        assert texto == "II MODERADO"
        assert fill is _FILL_RIESGO_MODERADO

    def test_bajo_minima_exposicion(self):
        nivel, texto, fill = _calcular_nivel_riesgo("No es Probable", "Esporádica")
        assert nivel == 1
        assert texto == "III BAJO"
        assert fill is _FILL_RIESGO_BAJO

    def test_tolera_sin_tilde(self):
        """'Esporadica' sin tilde tambien debe funcionar."""
        nivel, _, _ = _calcular_nivel_riesgo("No es Probable", "Esporadica")
        assert nivel >= 1


# ---------------------------------------------------------------------------
# Catalogo de peligros por factor
# ---------------------------------------------------------------------------

class TestCatalogoFactores:

    def test_los_5_factores_estan(self):
        esperados = {"FACTOR HUMANO", "FACTOR MECÁNICO",
                     "FACTOR INFRAESTRUCTURA VIAL",
                     "FACTOR AMBIENTAL", "FACTOR SOCIAL"}
        assert set(ORDEN_FACTORES) == esperados
        assert set(CATALOGO.keys()) == esperados

    def test_factor_humano_tiene_4_actores(self):
        """Humano debe tener conductor, motociclista, peaton, pasajero."""
        actores = CATALOGO["FACTOR HUMANO"]
        for a in ("conductor", "motociclista", "peaton", "pasajero"):
            assert a in actores, f"Falta actor {a} en FACTOR HUMANO"

    def test_factor_mecanico_solo_conductor_y_motociclista(self):
        """No tiene sentido factor mecanico para peatones."""
        actores = CATALOGO["FACTOR MECÁNICO"]
        assert "conductor" in actores
        assert "motociclista" in actores
        # peaton y pasajero no aplican
        assert "peaton" not in actores

    def test_peligros_tienen_estructura(self):
        """Cada peligro debe tener todos los campos GTC 45."""
        claves = {"peligro", "efectos", "fuente", "medio", "individuo",
                  "nivel_probabilidad", "nivel_exposicion", "no_expuestos",
                  "peor", "eliminacion", "sustitucion", "ingenieria",
                  "administrativos", "epp", "plan_accion"}
        for factor, actores in CATALOGO.items():
            for actor, peligros in actores.items():
                for p in peligros:
                    assert claves.issubset(p.keys()), \
                        f"Falta clave en {factor}/{actor}: {claves - p.keys()}"


# ---------------------------------------------------------------------------
# Generacion completa
# ---------------------------------------------------------------------------

class TestGenerarMatrizVial:

    def test_genera_con_pesv_activado(self, cliente_con_pesv, output_pesv):
        from engine.matriz_peligros_viales import generar_matriz_peligros_viales
        ruta = generar_matriz_peligros_viales(cliente_con_pesv, output_pesv)
        assert ruta is not None
        assert ruta.exists()

    def test_no_genera_sin_pesv(self, cliente_minimo, output_pesv):
        """generar_pesv: false debe abortar."""
        from engine.matriz_peligros_viales import generar_matriz_peligros_viales
        # cliente_minimo tiene generar_pesv: False
        ruta = generar_matriz_peligros_viales(cliente_minimo, output_pesv)
        assert ruta is None

    def test_no_genera_sin_cargos_viales(self, cliente_minimo, output_pesv):
        """Si todos los cargos son admin de oficina, no se genera matriz."""
        from engine.matriz_peligros_viales import generar_matriz_peligros_viales
        cliente_minimo["generar_pesv"] = True
        cliente_minimo["cargos"] = [
            {"nombre": "Gerente", "actividades": "Direccion estrategica desde oficina"},
        ]
        ruta = generar_matriz_peligros_viales(cliente_minimo, output_pesv)
        assert ruta is None

    def test_matriz_tiene_los_5_factores(self, cliente_con_pesv, output_pesv):
        import openpyxl
        from engine.matriz_peligros_viales import generar_matriz_peligros_viales
        generar_matriz_peligros_viales(cliente_con_pesv, output_pesv)
        ruta = output_pesv / RUTA_MATRIZ_VIAL
        wb = openpyxl.load_workbook(str(ruta))
        ws = wb["Matriz de peligros"]
        # Recolectar las bandas de factor (en col A)
        factores_encontrados = set()
        for r in range(7, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if v and "FACTOR" in str(v).upper():
                factores_encontrados.add(str(v).upper())
        assert len(factores_encontrados) >= 5, f"Solo {len(factores_encontrados)} factores generados"

    def test_columna_R_tiene_color_semaforo(self, cliente_con_pesv, output_pesv):
        """La columna R (interpretacion riesgo) debe tener color rojo/naranja/verde."""
        import openpyxl
        from engine.matriz_peligros_viales import generar_matriz_peligros_viales
        generar_matriz_peligros_viales(cliente_con_pesv, output_pesv)
        ruta = output_pesv / RUTA_MATRIZ_VIAL
        wb = openpyxl.load_workbook(str(ruta))
        ws = wb["Matriz de peligros"]
        colores_validos = {"00C00000", "00FFC000", "0000B050"}  # rojo, naranja, verde
        encontrados = set()
        for r in range(7, ws.max_row + 1):
            cel = ws.cell(row=r, column=18)
            if cel.fill and cel.fill.fgColor:
                rgb = cel.fill.fgColor.rgb
                if rgb in colores_validos:
                    encontrados.add(rgb)
        # Debe haber al menos un color de cada o varios validos
        assert len(encontrados) >= 1

    def test_banda_factor_usa_color_corporativo(self, cliente_con_pesv, output_pesv):
        """Bug real: la banda roja de FACTOR quedaba fija en C00000 aunque el
        cliente tuviera color_corporativo configurado, porque esta matriz se
        genera DESPUES del paso aplicar_colores_corporativos (nunca la toca).
        La banda debe seguir la paleta del cliente."""
        import openpyxl
        from engine.matriz_peligros_viales import generar_matriz_peligros_viales

        cliente_con_pesv["color_corporativo"] = "azul"
        generar_matriz_peligros_viales(cliente_con_pesv, output_pesv)
        ruta = output_pesv / RUTA_MATRIZ_VIAL
        wb = openpyxl.load_workbook(str(ruta))
        ws = wb["Matriz de peligros"]

        fila_factor = next(
            r for r in range(7, ws.max_row + 1)
            if ws.cell(row=r, column=1).value
            and "FACTOR" in str(ws.cell(row=r, column=1).value).upper()
        )
        rgb = ws.cell(row=fila_factor, column=1).fill.fgColor.rgb
        assert rgb != "00C00000", "La banda de factor sigue en rojo vinotinto pese a color_corporativo=azul"
        assert rgb == "002E75B6"  # paleta azul -> "medio"

    def test_banda_factor_sin_color_corporativo_es_roja(self, cliente_con_pesv, output_pesv):
        """Sin color_corporativo configurado, la banda debe seguir siendo
        el rojo original de la plantilla (comportamiento por defecto)."""
        import openpyxl
        from engine.matriz_peligros_viales import generar_matriz_peligros_viales

        cliente_con_pesv.pop("color_corporativo", None)
        generar_matriz_peligros_viales(cliente_con_pesv, output_pesv)
        ruta = output_pesv / RUTA_MATRIZ_VIAL
        wb = openpyxl.load_workbook(str(ruta))
        ws = wb["Matriz de peligros"]

        fila_factor = next(
            r for r in range(7, ws.max_row + 1)
            if ws.cell(row=r, column=1).value
            and "FACTOR" in str(ws.cell(row=r, column=1).value).upper()
        )
        rgb = ws.cell(row=fila_factor, column=1).fill.fgColor.rgb
        assert rgb == "00C00000"


# ---------------------------------------------------------------------------
# Limpieza de hojas PESV
# ---------------------------------------------------------------------------

class TestPesvCleanup:

    def test_elimina_hoja_pesv_pero_conserva_seguridad_vial(self, tmp_path):
        """Bug real: 'SEGURIDAD VIAL' aplica por Resolucion 1565/2014 incluso
        cuando PESV no aplica. Solo se borra la hoja PESV propiamente."""
        from engine.pesv_cleanup import limpiar_hojas_pesv_en_matrices
        import openpyxl

        # Crear un xlsx con varias hojas
        xlsx = tmp_path / "test.xlsx"
        wb = openpyxl.Workbook()
        wb.active.title = "MATRIZ DE SST"
        wb.create_sheet("SEGURIDAD VIAL ")  # con espacio al final como en plantilla
        wb.create_sheet("PESV")
        wb.create_sheet("CONTROL DE VERSIONES")
        wb.save(str(xlsx))

        resumen = limpiar_hojas_pesv_en_matrices(tmp_path)
        assert "test.xlsx" in resumen
        assert "PESV" in resumen["test.xlsx"]

        wb2 = openpyxl.load_workbook(str(xlsx))
        assert "PESV" not in wb2.sheetnames
        assert "SEGURIDAD VIAL " in wb2.sheetnames  # se conserva
        assert "MATRIZ DE SST" in wb2.sheetnames
        assert "CONTROL DE VERSIONES" in wb2.sheetnames

    def test_no_falla_sin_archivos(self, tmp_path):
        from engine.pesv_cleanup import limpiar_hojas_pesv_en_matrices
        resumen = limpiar_hojas_pesv_en_matrices(tmp_path)
        assert resumen == {}

    def test_no_elimina_si_dejaria_libro_vacio(self, tmp_path):
        """Proteccion: si todas las hojas son PESV, no borrar nada."""
        from engine.pesv_cleanup import limpiar_hojas_pesv_en_matrices
        import openpyxl
        xlsx = tmp_path / "solo_pesv.xlsx"
        wb = openpyxl.Workbook()
        wb.active.title = "PESV"
        wb.save(str(xlsx))
        resumen = limpiar_hojas_pesv_en_matrices(tmp_path)
        assert "solo_pesv.xlsx" not in resumen
