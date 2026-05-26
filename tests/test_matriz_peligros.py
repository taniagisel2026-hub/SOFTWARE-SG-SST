"""Tests de la Matriz IPVR SST-MTZ-012."""
import pytest

from engine.matriz_peligros import (
    _detectar_tipo_cargo,
    _extraer_denominacion,
    PELIGROS_BASE_ADMINISTRATIVO,
    EXTRAS_POR_TIPO,
    RUTA_MATRIZ_IPVR,
)


# ---------------------------------------------------------------------------
# Deteccion automatica del tipo de cargo
# ---------------------------------------------------------------------------

class TestDetectarTipoCargo:

    def test_conductor(self):
        assert _detectar_tipo_cargo("Conductor de Camion") == "conductor"
        assert _detectar_tipo_cargo("Chofer Ejecutivo") == "conductor"
        assert _detectar_tipo_cargo("Mensajero Motorizado") == "conductor"

    def test_servicios_generales(self):
        assert _detectar_tipo_cargo("Auxiliar de Servicios Generales") == "servicios_generales"
        assert _detectar_tipo_cargo("Personal de Aseo") == "servicios_generales"
        assert _detectar_tipo_cargo("Limpieza") == "servicios_generales"

    def test_operativo(self):
        assert _detectar_tipo_cargo("Operario de Torno") == "operativo"
        assert _detectar_tipo_cargo("Soldador") == "operativo"
        assert _detectar_tipo_cargo("Mecanico Industrial") == "operativo"
        assert _detectar_tipo_cargo("Tecnico de Mantenimiento") == "operativo"

    def test_administrativo_default(self):
        """Cargos sin patron claro caen en administrativo."""
        assert _detectar_tipo_cargo("Gerente General") == "administrativo"
        assert _detectar_tipo_cargo("Contador") == "administrativo"
        assert _detectar_tipo_cargo("Director Comercial") == "administrativo"


# ---------------------------------------------------------------------------
# Catalogo base de peligros
# ---------------------------------------------------------------------------

class TestCatalogoBase:

    def test_peligros_base_tiene_estructura_correcta(self):
        """Cada peligro debe tener las claves obligatorias."""
        claves_obligatorias = {"peligro", "clasificacion", "efectos",
                               "fuente", "medio", "individuo",
                               "nd", "ne", "nc", "peor", "req_legal"}
        for pel in PELIGROS_BASE_ADMINISTRATIVO:
            assert claves_obligatorias.issubset(pel.keys()), \
                f"Falta clave en {pel.get('peligro')}: {claves_obligatorias - pel.keys()}"

    def test_valores_nd_ne_nc_validos(self):
        """ND debe ser 2/6/10/14, NE 1-4, NC 10/25/60/100 segun GTC 45."""
        for pel in PELIGROS_BASE_ADMINISTRATIVO:
            assert pel["nd"] in (2, 6, 10, 14), f"ND invalido en {pel['peligro']}"
            assert pel["ne"] in (1, 2, 3, 4), f"NE invalido en {pel['peligro']}"
            assert pel["nc"] in (10, 25, 60, 100), f"NC invalido en {pel['peligro']}"

    def test_clasificaciones_validas(self):
        """Las clasificaciones deben ser las 7 categorias GTC 45."""
        validas = {"BIOLÓGICO", "FÍSICO", "QUÍMICO", "PSICOSOCIAL",
                   "BIOMECÁNICO", "CONDICIONES DE SEGURIDAD",
                   "FENÓMENOS NATURALES"}
        for pel in PELIGROS_BASE_ADMINISTRATIVO:
            assert pel["clasificacion"] in validas, \
                f"Clasificacion invalida: {pel['clasificacion']}"

    def test_extras_por_tipo_existen(self):
        for tipo in ("administrativo", "operativo", "conductor", "servicios_generales"):
            assert tipo in EXTRAS_POR_TIPO

    def test_extras_operativos_tienen_mecanico(self):
        """Operativos deben tener peligro mecanico (atrapamiento)."""
        peligros_op = EXTRAS_POR_TIPO["operativo"]
        clasifs = [p["clasificacion"] for p in peligros_op]
        textos = " ".join(p["peligro"] for p in peligros_op).upper()
        assert "MECÁNICO" in textos or "CONDICIONES DE SEGURIDAD" in clasifs


# ---------------------------------------------------------------------------
# Extraccion de DENOMINACION desde texto de manual
# ---------------------------------------------------------------------------

class TestExtraerDenominacion:

    def test_extrae_denominacion_de_tabla(self):
        texto = """MANUAL DE FUNCIONES
DENOMINACIÓN | DIRECTOR GENERAL
DEPENDENCIA | GERENCIA"""
        assert _extraer_denominacion(texto, "fallback.docx") == "DIRECTOR GENERAL"

    def test_fallback_a_nombre_archivo(self):
        nombre = _extraer_denominacion("texto sin denominacion", "5. CONDUCTOR.docx")
        assert "CONDUCTOR" in nombre.upper()
        assert nombre.startswith("CONDUCTOR")  # quita la numeracion "5. "

    def test_fallback_quita_numeracion(self):
        nombre = _extraer_denominacion("", "10. EDITOR AUDIOVISUAL.docx")
        assert nombre == "EDITOR AUDIOVISUAL"


# ---------------------------------------------------------------------------
# Generacion completa (integrado con plantilla)
# ---------------------------------------------------------------------------

class TestGenerarMatrizIPVR:

    def test_genera_matriz_con_cargos(self, cliente_con_cargos, output_ipvr):
        from engine.matriz_peligros import generar_matriz_peligros
        ruta = generar_matriz_peligros(cliente_con_cargos, output_ipvr)
        assert ruta is not None
        assert ruta.exists()

    def test_sin_cargos_no_genera(self, cliente_minimo, output_ipvr):
        from engine.matriz_peligros import generar_matriz_peligros
        # cliente_minimo no tiene 'cargos' ni 'manuales_cargos_dir'
        ruta = generar_matriz_peligros(cliente_minimo, output_ipvr)
        assert ruta is None

    def test_matriz_tiene_proceso_administrativo(self, cliente_con_cargos, output_ipvr):
        """La matriz debe contener los procesos del cliente."""
        import openpyxl
        from engine.matriz_peligros import generar_matriz_peligros
        generar_matriz_peligros(cliente_con_cargos, output_ipvr)
        ruta = output_ipvr / RUTA_MATRIZ_IPVR
        wb = openpyxl.load_workbook(str(ruta))
        ws = wb["MATRIZ DE PELIGROS"]
        procesos = set()
        for r in range(6, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if v:
                procesos.add(str(v).upper())
        assert "ADMINISTRATIVO" in procesos

    def test_cargos_se_agrupan_por_proceso(self, cliente_minimo, output_ipvr):
        """Si paso 2 cargos administrativos mezclados con 1 operativo,
        los administrativos deben quedar JUNTOS."""
        import openpyxl
        from engine.matriz_peligros import generar_matriz_peligros

        cliente_minimo["cargos"] = [
            {"nombre": "Director", "proceso": "ADMINISTRATIVO", "tipo": "administrativo",
             "actividades": "x", "tareas": "y"},
            {"nombre": "Operario", "proceso": "PRODUCCION", "tipo": "operativo",
             "actividades": "x", "tareas": "y"},
            {"nombre": "Asistente", "proceso": "ADMINISTRATIVO", "tipo": "administrativo",
             "actividades": "x", "tareas": "y"},
        ]
        generar_matriz_peligros(cliente_minimo, output_ipvr)
        ruta = output_ipvr / RUTA_MATRIZ_IPVR
        wb = openpyxl.load_workbook(str(ruta))
        ws = wb["MATRIZ DE PELIGROS"]
        # Recolectar orden de cargos
        cargos_en_orden = []
        for r in range(6, ws.max_row + 1):
            c = ws.cell(row=r, column=3).value
            if c and c not in cargos_en_orden:
                cargos_en_orden.append(c)
        # Director y Asistente deben estar consecutivos (ambos ADMIN)
        idx_dir = cargos_en_orden.index("Director") if "Director" in cargos_en_orden else -1
        idx_asis = cargos_en_orden.index("Asistente") if "Asistente" in cargos_en_orden else -1
        idx_op = cargos_en_orden.index("Operario") if "Operario" in cargos_en_orden else -1
        assert idx_dir >= 0 and idx_asis >= 0 and idx_op >= 0
        # Director y Asistente consecutivos (sin operativo en medio)
        assert abs(idx_dir - idx_asis) == 1, f"Cargos ADMIN no consecutivos: {cargos_en_orden}"

    def test_negrillas_en_columnas_correctas(self, cliente_con_cargos, output_ipvr):
        """PROCESO (A) y CARGO (C) deben ir en negrilla en la primera fila del bloque."""
        import openpyxl
        from engine.matriz_peligros import generar_matriz_peligros
        generar_matriz_peligros(cliente_con_cargos, output_ipvr)
        ruta = output_ipvr / RUTA_MATRIZ_IPVR
        wb = openpyxl.load_workbook(str(ruta))
        ws = wb["MATRIZ DE PELIGROS"]
        # La fila 6 deberia ser la primera fila del primer cargo
        cel_proceso = ws.cell(row=6, column=1)
        cel_cargo = ws.cell(row=6, column=3)
        assert cel_proceso.font.bold is True
        assert cel_cargo.font.bold is True
