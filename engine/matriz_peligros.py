"""Generador de la Matriz de Identificacion de Peligros y Valoracion de
Riesgos (IPVR) SST-MTZ-012, segun GTC 45.

Para cada cargo definido en el YAML del cliente, se inserta un bloque de
filas con los peligros estandar (catalogo base) + peligros adicionales que
Claude genera segun la naturaleza del cargo (administrativo / operativo /
conduccion / aseo / etc.).

YAML esperado:

    cargos:
      - nombre: "Gerente General"
        proceso: "ADMINISTRATIVO"          # opcional, default ADMINISTRATIVO
        zona: "Oficinas Administrativas"   # opcional
        num_expuestos: 1                   # opcional, default 1
        tipo: "administrativo"             # opcional: administrativo|operativo|conductor|servicios_generales
        actividades: ""                    # opcional, IA lo genera si vacio
        tareas: ""                         # opcional, IA lo genera si vacio
"""
from __future__ import annotations

import json
import logging
import re
from copy import copy
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from openpyxl.utils import get_column_letter

def _set_bold(cell, bold: bool) -> None:
    """Cambia solo el bold de una celda preservando nombre, tamano y color de la fuente."""
    f = cell.font
    cell.font = Font(
        name=f.name, size=f.size, color=f.color,
        italic=f.italic, underline=f.underline, strike=f.strike,
        vertAlign=f.vertAlign, bold=bold,
    )

# Color FIJO para columnas A (PROCESO) y B (ZONA/LUGAR): siempre durazno
_FILL_PROCESO_ZONA = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

# Colores ALTERNADOS por cargo aplicados a columnas C..L
_FILL_AZUL_CLARO = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")   # azul cielo
_FILL_VERDE_CLARO = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # verde claro
_FILLS_CARGOS = [_FILL_AZUL_CLARO, _FILL_VERDE_CLARO]

# Columnas con color del cargo: C (3) hasta L (12)
_COL_FILL_INI = 3
_COL_FILL_FIN = 12

# Columnas con color fijo durazno: A (1) y B (2)
_COL_FIJO_INI = 1
_COL_FIJO_FIN = 2

logger = logging.getLogger(__name__)

# Ruta de la matriz dentro del SG-SST generado
RUTA_MATRIZ_IPVR = Path("II. HACER") / "4.1.2 MATRIZ DE IPVR" / "SST-MTZ-012 MATRIZ DE IPVR.xlsx"

NOMBRE_HOJA = "MATRIZ DE PELIGROS"
FILA_INICIO = 6  # Primera fila de datos (debajo de los headers)

# ---------------------------------------------------------------------------
# Catalogo base de peligros segun GTC 45 (aplican a todos los cargos)
# Cada entrada produce UNA fila en la matriz.
# ND, NE, NC son los valores numericos de la GTC 45:
#   ND: 2 (Bajo), 6 (Medio), 10 (Alto), 14 (Muy Alto)
#   NE: 1 (Esporadica), 2 (Ocasional), 3 (Frecuente), 4 (Continua)
#   NC: 10 (Leve), 25 (Grave), 60 (Muy Grave), 100 (Mortal)
# ---------------------------------------------------------------------------

PELIGROS_BASE_ADMINISTRATIVO: list[dict[str, Any]] = [
    {
        "peligro": "EXPOSICIÓN A MICROORGANISMOS TALES COMO VIRUS, BACTERIAS, HONGOS, RICKETSIAS, PARÁSITOS, PICADURAS, MORDEDURAS, FLUIDOS Y EXCREMENTOS",
        "clasificacion": "BIOLÓGICO",
        "efectos": "VIROSIS\nINFECCIÓN RESPIRATORIA AGUDA\nALERGIAS",
        "fuente": "NO OBSERVADO",
        "medio": "NO OBSERVADO",
        "individuo": "LAVADO DE MANOS\nDESINFECCIÓN DE ÁREA DE TRABAJO\nINDUCCIÓN SG-SST\nEXÁMENES MÉDICO OCUPACIONALES DE INGRESO Y PERIÓDICOS",
        "nd": 2, "ne": 3, "nc": 25,
        "peor": "VIROSIS\nINFECCIÓN RESPIRATORIA AGUDA\nALERGIAS",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A", "ingenieria": "N/A",
        "administrativos": "INFORMACIÓN PUBLICADA SOBRE EL ADECUADO LAVADO DE MANOS\nCAPACITACIÓN EN AUTOCUIDADO\nHIGIENE PERSONAL\nCAPACITACIÓN EN RIESGO BIOLÓGICO",
        "epp": "N/A",
    },
    {
        "peligro": "RUIDO INTERMITENTE GENERADO POR EQUIPOS DE OFICINA Y TELÉFONOS",
        "clasificacion": "FÍSICO",
        "efectos": "PROBLEMAS DE AUDICIÓN\nPÉRDIDA DE CAPACIDAD AUDITIVA",
        "fuente": "NO OBSERVADO",
        "medio": "NO OBSERVADO",
        "individuo": "PAUSAS DURANTE LA JORNADA\nINDUCCIÓN SG-SST\nEXÁMENES OCUPACIONALES DE INGRESO, PERIÓDICOS Y DE RETIRO",
        "nd": 2, "ne": 3, "nc": 10,
        "peor": "PROBLEMAS DE AUDICIÓN\nPÉRDIDA DE CAPACIDAD AUDITIVA",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A", "ingenieria": "N/A",
        "administrativos": "CAPACITACIÓN EN PROTECCIÓN AUDITIVA\nAPLICACIÓN DE PROCEDIMIENTOS SEGUROS\nEXÁMENES MÉDICOS OCUPACIONALES PERIÓDICOS\nMEDICIONES AMBIENTALES",
        "epp": "N/A",
    },
    {
        "peligro": "RADIACIONES NO IONIZANTES POR EXPOSICIÓN A PANTALLAS DE EQUIPOS DE CÓMPUTO",
        "clasificacion": "FÍSICO",
        "efectos": "FATIGA OCULAR, CEFALEA, ARDOR OCULAR, NÁUSEAS, VÉRTIGOS, PROBLEMAS DE VISIÓN Y LAGRIMEO",
        "fuente": "NO OBSERVADO",
        "medio": "NO OBSERVADO",
        "individuo": "PAUSAS ACTIVAS VISUALES\nEXÁMENES MÉDICOS OCUPACIONALES DE INGRESO, PERIÓDICOS Y DE EGRESO",
        "nd": 2, "ne": 1, "nc": 10,
        "peor": "PROBLEMAS VISUALES Y DISMINUCIÓN DE LA VISIÓN",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A", "ingenieria": "N/A",
        "administrativos": "CAPACITACIÓN EN PAUSAS ACTIVAS Y YOGA VISUAL\nEXÁMENES MÉDICOS DE INGRESO Y PERIÓDICOS\nRECOMENDACIONES MÉDICAS\nSENSIBILIZACIÓN CONTROL POR PARTE DEL TRABAJADOR",
        "epp": "N/A",
    },
    {
        "peligro": "(GESTIÓN ORGANIZACIONAL) ESTILO DE MANDO, PAGO, CONTRATACIÓN, PARTICIPACIÓN, INDUCCIÓN Y CAPACITACIÓN, BIENESTAR SOCIAL, EVALUACIÓN DEL DESEMPEÑO, MANEJO DE CAMBIOS",
        "clasificacion": "PSICOSOCIAL",
        "efectos": "CEFALEAS, ESTRÉS LABORAL, GASTRITIS, MALESTARES ESTOMACALES, DOLORES CERVICALES",
        "fuente": "NO OBSERVADO",
        "medio": "NO OBSERVADO",
        "individuo": "PAUSAS ACTIVAS\nAPLICACIÓN BATERÍA PSICOSOCIAL\nENTREGA DE MANUALES DE CARGOS\nDIVULGACIÓN DE RESPONSABILIDADES",
        "nd": 6, "ne": 3, "nc": 25,
        "peor": "ESTRÉS LABORAL",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A", "ingenieria": "N/A",
        "administrativos": "CAPACITACIÓN EN MANEJO DE ESTRÉS\nCAPACITACIÓN EN LIDERAZGO\nCAPACITACIÓN EN COMUNICACIÓN ASERTIVA",
        "epp": "N/A",
    },
    {
        "peligro": "POSTURA SEDENTE PROLONGADA",
        "clasificacion": "BIOMECÁNICO",
        "efectos": "DESÓRDENES MUSCULOESQUELÉTICOS\nDOLOR LUMBAR\nDOLOR CERVICAL\nVÁRICES",
        "fuente": "NO OBSERVADO",
        "medio": "PROGRAMA DE PAUSAS ACTIVAS Y GIMNASIA LABORAL\nPVE BIOMECÁNICO",
        "individuo": "PAUSAS DURANTE LA JORNADA\nINDUCCIÓN SG-SST\nEXÁMENES OCUPACIONALES DE INGRESO, PERIÓDICOS Y DE RETIRO\nENTREGA DE MANUALES DE CARGO",
        "nd": 2, "ne": 3, "nc": 10,
        "peor": "FATIGA MUSCULAR\nDESÓRDENES MUSCULOESQUELÉTICOS",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A", "ingenieria": "N/A",
        "administrativos": "PAUSAS ACTIVAS\nCAPACITACIÓN EN RIESGO BIOMECÁNICO\nCHARLAS DE 5 MINUTOS",
        "epp": "N/A",
    },
    {
        "peligro": "MOVIMIENTO REPETITIVO DE MIEMBROS SUPERIORES POR DIGITACIÓN",
        "clasificacion": "BIOMECÁNICO",
        "efectos": "SÍNDROME DE TÚNEL CARPIANO\nDOLOR CERVICAL",
        "fuente": "NO OBSERVADO",
        "medio": "NO OBSERVADO",
        "individuo": "PAUSAS DURANTE LA JORNADA\nINDUCCIÓN SG-SST\nEXÁMENES OCUPACIONALES DE INGRESO\nENTREGA DE MANUALES DE CARGOS",
        "nd": 2, "ne": 1, "nc": 10,
        "peor": "SÍNDROME DE TÚNEL CARPIANO",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A", "ingenieria": "N/A",
        "administrativos": "PAUSAS ACTIVAS\nCAPACITACIÓN EN RIESGO BIOMECÁNICO\nPROGRAMA DE VIGILANCIA EPIDEMIOLÓGICA POR DESÓRDENES OSTEOMUSCULARES",
        "epp": "N/A",
    },
    {
        "peligro": "(LOCATIVO) SUPERFICIES DE TRABAJO (IRREGULARES, DESLIZANTES, CON DIFERENCIA DE NIVEL), CONDICIONES DE ORDEN Y ASEO",
        "clasificacion": "CONDICIONES DE SEGURIDAD",
        "efectos": "CAÍDA DE PERSONAS, GOLPES, CONTUSIONES, TRAUMAS, HERIDAS, LESIONES OSTEOMUSCULARES",
        "fuente": "DELIMITACIÓN ÁREAS DE TRABAJO",
        "medio": "BOTIQUÍN DE PRIMEROS AUXILIOS\nINSPECCIONES LOCATIVAS",
        "individuo": "CAPACITACIÓN EN AUTOCUIDADO\nINDUCCIÓN SG-SST\nELEMENTOS DE PROTECCIÓN PERSONAL",
        "nd": 2, "ne": 3, "nc": 10,
        "peor": "LESIONES FÍSICAS",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A", "ingenieria": "N/A",
        "administrativos": "CAPACITACIÓN EN ORDEN Y ASEO\nCAPACITACIONES EN AUTOCUIDADO\nCAPACITACIÓN EN ACTOS Y CONDICIONES INSEGURAS\nINSPECCIONES DE SEGURIDAD",
        "epp": "CALZADO ANTIDESLIZANTE\nDOTACIÓN LEGAL",
    },
    {
        "peligro": "(ELÉCTRICO) CONTACTO CON CONEXIONES, TOMAS Y EQUIPOS ENERGIZADOS",
        "clasificacion": "CONDICIONES DE SEGURIDAD",
        "efectos": "QUEMADURAS, CHOQUE ELÉCTRICO, FIBRILACIÓN VENTRICULAR",
        "fuente": "INSTALACIONES ELÉCTRICAS REGLAMENTADAS",
        "medio": "MANTENIMIENTO ELÉCTRICO PREVENTIVO",
        "individuo": "CAPACITACIÓN EN RIESGO ELÉCTRICO\nINDUCCIÓN SG-SST",
        "nd": 2, "ne": 1, "nc": 25,
        "peor": "QUEMADURAS\nMUERTE",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A",
        "ingenieria": "PROTECCIONES ELÉCTRICAS Y POLO A TIERRA",
        "administrativos": "INSPECCIONES DE INSTALACIONES ELÉCTRICAS\nCAPACITACIÓN EN RIESGO ELÉCTRICO",
        "epp": "N/A",
    },
    {
        "peligro": "(ACCIDENTES DE TRÁNSITO) PEATÓN, CICLISTA, MOTOCICLISTA, CONDUCTOR, PASAJERO EN DESPLAZAMIENTOS LABORALES",
        "clasificacion": "CONDICIONES DE SEGURIDAD",
        "efectos": "LESIONES, TRAUMAS, FRACTURAS, INVALIDEZ, MUERTE",
        "fuente": "NO OBSERVADO",
        "medio": "PLAN ESTRATÉGICO DE SEGURIDAD VIAL",
        "individuo": "AUTOCUIDADO",
        "nd": 2, "ne": 1, "nc": 25,
        "peor": "LESIONES\nMUERTE",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A", "ingenieria": "N/A",
        "administrativos": "IMPLEMENTAR Y EJECUTAR PLAN ESTRATÉGICO DE SEGURIDAD VIAL\nCAPACITACIÓN EN SEGURIDAD VIAL\nCAPACITACIÓN EN COMO ACTUAR EN CASO DE EMERGENCIA",
        "epp": "N/A",
    },
    {
        "peligro": "(PÚBLICOS) ROBOS, ATRACOS, ASALTOS, ATENTADOS, ALTERACIONES DE ORDEN PÚBLICO",
        "clasificacion": "CONDICIONES DE SEGURIDAD",
        "efectos": "LESIONES, TRAUMAS, FRACTURAS, GOLPES, HERIDAS, INVALIDEZ, MUERTE",
        "fuente": "NO OBSERVADO",
        "medio": "CIRCUITO CERRADO DE TELEVISIÓN",
        "individuo": "AUTOCUIDADO",
        "nd": 2, "ne": 3, "nc": 60,
        "peor": "LESIONES\nMUERTE",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A", "ingenieria": "N/A",
        "administrativos": "CAPACITACIÓN DE AUTOCUIDADO Y DETECCIÓN DE PERSONAS SOSPECHOSAS\nDIVULGACIÓN DE NÚMEROS DE EMERGENCIAS\nDIVULGACIÓN DE TIPS DE SEGURIDAD",
        "epp": "N/A",
    },
    {
        "peligro": "SISMO, TERREMOTO, VENDAVAL, INUNDACIÓN, DERRUMBES, PRECIPITACIONES (LLUVIAS, GRANIZADAS, HELADAS)",
        "clasificacion": "FENÓMENOS NATURALES",
        "efectos": "LESIONES, TRAUMAS, FRACTURAS, INVALIDEZ, MUERTE",
        "fuente": "NO OBSERVADO",
        "medio": "BOTIQUÍN DE PRIMEROS AUXILIOS\nEQUIPOS DE EMERGENCIAS",
        "individuo": "CAPACITACIÓN DE COMO ACTUAR ANTE EMERGENCIAS\nSIMULACROS DE EVACUACIÓN",
        "nd": 6, "ne": 1, "nc": 25,
        "peor": "LESIONES\nMUERTE",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A", "ingenieria": "N/A",
        "administrativos": "SIMULACROS DE EMERGENCIA AVISADOS Y NO AVISADOS\nCAPACITACIÓN EN PRIMEROS AUXILIOS, EVACUACIÓN Y RESCATE Y CONTRA INCENDIOS\nSEÑALIZACIÓN DE RUTAS DE EVACUACIÓN",
        "epp": "N/A",
    },
]

# Peligros adicionales para cargos OPERATIVOS (taller, planta, produccion)
PELIGROS_EXTRA_OPERATIVO: list[dict[str, Any]] = [
    {
        "peligro": "RUIDO CONTINUO POR OPERACIÓN DE MAQUINARIA Y EQUIPOS",
        "clasificacion": "FÍSICO",
        "efectos": "HIPOACUSIA NEUROSENSORIAL, CEFALEA, ESTRÉS, IRRITABILIDAD",
        "fuente": "NO OBSERVADO",
        "medio": "MEDICIONES AMBIENTALES",
        "individuo": "PAUSAS DURANTE LA JORNADA\nINDUCCIÓN SG-SST\nEXÁMENES OCUPACIONALES (AUDIOMETRÍAS)",
        "nd": 6, "ne": 4, "nc": 25,
        "peor": "PÉRDIDA DE CAPACIDAD AUDITIVA",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A",
        "ingenieria": "MANTENIMIENTO PREVENTIVO DE MAQUINAS\nAISLAMIENTO ACÚSTICO",
        "administrativos": "PVE AUDITIVO\nCAPACITACIÓN EN PROTECCIÓN AUDITIVA\nMEDICIONES AMBIENTALES PERIÓDICAS",
        "epp": "PROTECTOR AUDITIVO TIPO COPA O INSERCIÓN",
    },
    {
        "peligro": "GASES Y VAPORES POR USO DE PRODUCTOS QUÍMICOS",
        "clasificacion": "QUÍMICO",
        "efectos": "CEFALEAS, IRRITACIÓN VÍAS RESPIRATORIAS, POSIBLE INTOXICACIÓN",
        "fuente": "VENTILACIÓN NATURAL",
        "medio": "HOJAS DE SEGURIDAD\nVENTILACIÓN MECÁNICA",
        "individuo": "USO DE TAPABOCAS\nCAPACITACIÓN RIESGO QUÍMICO",
        "nd": 6, "ne": 3, "nc": 25,
        "peor": "INTOXICACIÓN",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "SUSTITUCIÓN POR PRODUCTOS DE MENOR TOXICIDAD",
        "ingenieria": "EXTRACCIÓN LOCALIZADA",
        "administrativos": "INFORMACIÓN DE MANEJO DE QUÍMICOS\nHOJAS DE SEGURIDAD\nCAPACITACIÓN EN RIESGO QUÍMICO",
        "epp": "PROTECCIÓN RESPIRATORIA\nGUANTES NITRILO\nMONOGAFAS",
    },
    {
        "peligro": "POSTURA BÍPEDA PROLONGADA",
        "clasificacion": "BIOMECÁNICO",
        "efectos": "DESÓRDENES MUSCULOESQUELÉTICOS\nALTERACIÓN SISTEMA CIRCULATORIO\nDOLOR DE PIERNAS\nDOLOR LUMBAR\nVÁRICES",
        "fuente": "NO OBSERVADO",
        "medio": "PROGRAMA DE PAUSAS ACTIVAS Y GIMNASIA LABORAL\nPVE BIOMECÁNICO",
        "individuo": "PAUSAS DURANTE LA JORNADA\nINDUCCIÓN SG-SST\nEXÁMENES OCUPACIONALES",
        "nd": 6, "ne": 4, "nc": 10,
        "peor": "DESÓRDENES MUSCULOESQUELÉTICOS",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A",
        "ingenieria": "TAPETES ANTIFATIGA",
        "administrativos": "PAUSAS ACTIVAS\nCAPACITACIÓN EN HIGIENE POSTURAL\nROTACIÓN DE TAREAS",
        "epp": "CALZADO DE SEGURIDAD CON SUELA AMORTIGUADORA",
    },
    {
        "peligro": "MANIPULACIÓN MANUAL DE CARGAS",
        "clasificacion": "BIOMECÁNICO",
        "efectos": "LUMBALGIAS, HERNIAS DISCALES, DESÓRDENES MUSCULOESQUELÉTICOS",
        "fuente": "AYUDAS MECÁNICAS DISPONIBLES",
        "medio": "PVE BIOMECÁNICO",
        "individuo": "INDUCCIÓN SG-SST\nCAPACITACIÓN HIGIENE POSTURAL",
        "nd": 6, "ne": 3, "nc": 25,
        "peor": "HERNIAS DISCALES\nLUMBALGIA CRÓNICA",
        "req_legal": "SI",
        "eliminacion": "N/A",
        "sustitucion": "USO DE AYUDAS MECÁNICAS",
        "ingenieria": "CARRETILLAS, MONTACARGAS, GATOS HIDRÁULICOS",
        "administrativos": "CAPACITACIÓN EN MANIPULACIÓN MANUAL DE CARGAS Resolución 4272/2018\nROTACIÓN DE TAREAS",
        "epp": "FAJA LUMBAR (BAJO PRESCRIPCIÓN)\nGUANTES DE CARNAZA",
    },
    {
        "peligro": "(MECÁNICO) CONTACTO CON MAQUINARIA, HERRAMIENTAS, PARTES MÓVILES, PROYECCIÓN DE PARTÍCULAS",
        "clasificacion": "CONDICIONES DE SEGURIDAD",
        "efectos": "ATRAPAMIENTOS, APLASTAMIENTOS, CORTADURAS, CONTUSIÓN, FRACTURAS, AMPUTACIONES",
        "fuente": "GUARDAS DE PROTECCIÓN EN MAQUINAS",
        "medio": "PROGRAMA DE MANTENIMIENTO PREVENTIVO",
        "individuo": "CAPACITACIÓN EN AUTOCUIDADO\nINDUCCIÓN SG-SST\nUSO ADECUADO DE MAQUINAS\nPROCEDIMIENTOS SEGUROS",
        "nd": 6, "ne": 4, "nc": 60,
        "peor": "AMPUTACIÓN\nMUERTE",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A",
        "ingenieria": "MANTENIMIENTO CORRECTIVO Y PREVENTIVO\nGUARDAS Y BLOQUEOS DE SEGURIDAD",
        "administrativos": "PROCEDIMIENTOS SEGUROS DE OPERACIÓN\nINSPECCIÓN DE MAQUINAS Y HERRAMIENTAS\nCAPACITACIÓN EN BLOQUEO Y ETIQUETADO LOTO",
        "epp": "CALZADO DE SEGURIDAD\nGUANTES DE PROTECCIÓN MECÁNICA\nMONOGAFAS\nDOTACIÓN LEGAL",
    },
    {
        "peligro": "(TECNOLÓGICO) DERRAMES, INCENDIOS Y EXPLOSIONES POR MANEJO DE COMBUSTIBLES, ACEITES Y QUÍMICOS",
        "clasificacion": "CONDICIONES DE SEGURIDAD",
        "efectos": "QUEMADURAS, INTOXICACIÓN, TRAUMAS, MUERTE",
        "fuente": "NO OBSERVADO",
        "medio": "EXTINTORES\nBOTIQUÍN DE PRIMEROS AUXILIOS\nHOJAS DE DATOS DE SEGURIDAD",
        "individuo": "CAPACITACIÓN EN AUTOCUIDADO\nINDUCCIÓN SG-SST\nELEMENTOS DE PROTECCIÓN PERSONAL\nPLAN DE EMERGENCIAS",
        "nd": 6, "ne": 2, "nc": 60,
        "peor": "QUEMADURAS\nMUERTE",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A",
        "ingenieria": "MANTENIMIENTO PREVENTIVO Y CORRECTIVO A MAQUINAS",
        "administrativos": "CAPACITACIÓN DE COMO ACTUAR ANTE EMERGENCIAS\nAUTOREPORTE DE CONDICIONES INSEGURAS\nINSPECCIONES DE SEGURIDAD",
        "epp": "EQUIPO DE EMERGENCIAS",
    },
]

PELIGROS_EXTRA_CONDUCTOR: list[dict[str, Any]] = [
    {
        "peligro": "(ACCIDENTES DE TRÁNSITO) CONDUCCIÓN DE VEHÍCULO COMO ACTIVIDAD PRINCIPAL",
        "clasificacion": "CONDICIONES DE SEGURIDAD",
        "efectos": "POLITRAUMATISMOS, FRACTURAS, INVALIDEZ, MUERTE",
        "fuente": "MANTENIMIENTO PREVENTIVO DEL VEHÍCULO",
        "medio": "PLAN ESTRATÉGICO DE SEGURIDAD VIAL\nKIT DE CARRETERA Y BOTIQUÍN",
        "individuo": "LICENCIA VIGENTE\nCAPACITACIÓN SEGURIDAD VIAL\nEXÁMENES MÉDICOS",
        "nd": 6, "ne": 4, "nc": 60,
        "peor": "MUERTE",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A",
        "ingenieria": "MANTENIMIENTO PREVENTIVO DEL VEHÍCULO",
        "administrativos": "PESV\nCAPACITACIÓN EN CONDUCCIÓN DEFENSIVA\nCONTROL DE DOCUMENTACIÓN DEL CONDUCTOR Y VEHÍCULO\nCONTROL DE FATIGA Y JORNADAS",
        "epp": "CINTURÓN DE SEGURIDAD",
    },
    {
        "peligro": "POSTURA SEDENTE PROLONGADA EN VEHÍCULO",
        "clasificacion": "BIOMECÁNICO",
        "efectos": "LUMBALGIA, DOLOR CERVICAL, DESÓRDENES MUSCULOESQUELÉTICOS",
        "fuente": "NO OBSERVADO",
        "medio": "PVE BIOMECÁNICO",
        "individuo": "PAUSAS ACTIVAS EN PARADAS\nAJUSTE ERGONÓMICO DE LA SILLA",
        "nd": 6, "ne": 4, "nc": 10,
        "peor": "DESÓRDENES MUSCULOESQUELÉTICOS",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A", "ingenieria": "N/A",
        "administrativos": "PAUSAS ACTIVAS\nCAPACITACIÓN EN HIGIENE POSTURAL",
        "epp": "N/A",
    },
]

PELIGROS_EXTRA_SERVICIOS_GENERALES: list[dict[str, Any]] = [
    {
        "peligro": "EXPOSICIÓN A PRODUCTOS QUÍMICOS DE ASEO (DETERGENTES, DESINFECTANTES, BLANQUEADORES)",
        "clasificacion": "QUÍMICO",
        "efectos": "DERMATITIS, IRRITACIÓN OCULAR Y DE VÍAS RESPIRATORIAS",
        "fuente": "ETIQUETADO DE PRODUCTOS",
        "medio": "HOJAS DE SEGURIDAD\nVENTILACIÓN",
        "individuo": "CAPACITACIÓN RIESGO QUÍMICO\nUSO DE EPP",
        "nd": 6, "ne": 4, "nc": 10,
        "peor": "DERMATITIS CRÓNICA",
        "req_legal": "SI",
        "eliminacion": "N/A",
        "sustitucion": "PRODUCTOS BIODEGRADABLES Y DE MENOR TOXICIDAD",
        "ingenieria": "VENTILACIÓN",
        "administrativos": "CAPACITACIÓN EN MANEJO DE QUÍMICOS\nHOJAS DE SEGURIDAD ACCESIBLES",
        "epp": "GUANTES DE NITRILO\nTAPABOCAS\nMONOGAFAS\nDELANTAL",
    },
    {
        "peligro": "POSTURA BÍPEDA PROLONGADA Y MOVIMIENTOS REPETITIVOS POR ACTIVIDADES DE LIMPIEZA",
        "clasificacion": "BIOMECÁNICO",
        "efectos": "DESÓRDENES MUSCULOESQUELÉTICOS, LUMBALGIA, EPICONDILITIS",
        "fuente": "NO OBSERVADO",
        "medio": "PVE BIOMECÁNICO",
        "individuo": "PAUSAS ACTIVAS\nINDUCCIÓN SG-SST",
        "nd": 6, "ne": 4, "nc": 10,
        "peor": "DESÓRDENES MUSCULOESQUELÉTICOS",
        "req_legal": "SI",
        "eliminacion": "N/A", "sustitucion": "N/A",
        "ingenieria": "AYUDAS ERGONÓMICAS (CARROS, ESCOBAS DE MANGO LARGO)",
        "administrativos": "ROTACIÓN DE TAREAS\nCAPACITACIÓN EN HIGIENE POSTURAL",
        "epp": "CALZADO ANTIDESLIZANTE\nGUANTES",
    },
]

# Mapeo nombre tipo -> peligros extra a anexar
EXTRAS_POR_TIPO = {
    "administrativo": [],
    "operativo": PELIGROS_EXTRA_OPERATIVO,
    "conductor": PELIGROS_EXTRA_CONDUCTOR,
    "servicios_generales": PELIGROS_EXTRA_SERVICIOS_GENERALES,
}


def _extraer_texto_docx(ruta: Path) -> str:
    """Extrae todo el texto util de un .docx (parrafos + tablas)."""
    try:
        from docx import Document
    except ImportError:
        logger.error("python-docx no instalado")
        return ""
    try:
        doc = Document(str(ruta))
    except Exception as e:
        logger.warning("No se pudo leer %s: %s", ruta, e)
        return ""

    lineas = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            lineas.append(t)
    # Tablas: deduplicar celdas merged (mismo texto repetido)
    for tbl in doc.tables:
        for row in tbl.rows:
            vistos = set()
            partes = []
            for cell in row.cells:
                t = cell.text.strip()
                if t and t not in vistos:
                    vistos.add(t)
                    partes.append(t)
            if partes:
                lineas.append(" | ".join(partes))
    return "\n".join(lineas)


def _extraer_texto_pdf(ruta: Path) -> str:
    """Extrae texto util de un .pdf (parrafos + tablas).
    Usa pdfplumber por buena extraccion de tablas; cae a pypdf si falla.
    """
    # Primer intento: pdfplumber (mejor para tablas)
    try:
        import pdfplumber
        lineas = []
        with pdfplumber.open(str(ruta)) as pdf:
            for pagina in pdf.pages:
                # Texto plano
                texto = pagina.extract_text() or ""
                for ln in texto.split("\n"):
                    ln = ln.strip()
                    if ln:
                        lineas.append(ln)
                # Tablas (formateadas como "col1 | col2 | col3")
                try:
                    tablas = pagina.extract_tables() or []
                except Exception:
                    tablas = []
                for tbl in tablas:
                    for row in tbl:
                        vistos = set()
                        partes = []
                        for cell in row:
                            if cell is None:
                                continue
                            t = str(cell).strip().replace("\n", " ")
                            if t and t not in vistos:
                                vistos.add(t)
                                partes.append(t)
                        if partes:
                            lineas.append(" | ".join(partes))
        return "\n".join(lineas)
    except ImportError:
        logger.warning("pdfplumber no instalado, usando pypdf")
    except Exception as e:
        logger.warning("pdfplumber fallo en %s: %s — intentando pypdf", ruta, e)

    # Fallback: pypdf (mas basico pero suele funcionar)
    try:
        from pypdf import PdfReader
    except ImportError:
        logger.error("Ni pdfplumber ni pypdf instalados, no se puede leer %s", ruta)
        return ""
    try:
        reader = PdfReader(str(ruta))
        lineas = []
        for pagina in reader.pages:
            texto = pagina.extract_text() or ""
            for ln in texto.split("\n"):
                ln = ln.strip()
                if ln:
                    lineas.append(ln)
        return "\n".join(lineas)
    except Exception as e:
        logger.warning("No se pudo leer PDF %s: %s", ruta, e)
        return ""


def _extraer_texto_xlsx(ruta: Path) -> str:
    """Extrae todo el texto util de un .xlsx (todas las hojas y celdas)."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        logger.error("openpyxl no instalado")
        return ""
    try:
        wb = load_workbook(str(ruta), data_only=True, read_only=True)
    except Exception as e:
        logger.warning("No se pudo leer %s: %s", ruta, e)
        return ""

    lineas = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            partes = []
            vistos = set()
            for cell in row:
                if cell is None:
                    continue
                t = str(cell).strip()
                if t and t not in vistos:
                    vistos.add(t)
                    partes.append(t)
            if partes:
                lineas.append(" | ".join(partes))
    wb.close()
    return "\n".join(lineas)


def _extraer_denominacion(texto: str, fallback_nombre_archivo: str) -> str:
    """Intenta extraer la DENOMINACION del cargo desde el texto del manual."""
    # Patron tabla: "DENOMINACIÓN | <nombre>"
    m = re.search(r"DENOMINACI[ÓO]N\s*\|\s*([^\n|]+)", texto, re.IGNORECASE)
    if m:
        nombre = m.group(1).strip()
        if nombre and nombre.upper() != "DENOMINACIÓN":
            return nombre
    # Patron "Nombre del cargo: ..."
    m = re.search(r"NOMBRE DEL CARGO\s*[:\|]\s*([^\n|]+)", texto, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    # Patron "CARGO | <nombre>" (comun en plantillas Excel)
    m = re.search(r"(?:^|\n)CARGO\s*\|\s*([^\n|]+)", texto, re.IGNORECASE)
    if m:
        nombre = m.group(1).strip()
        if nombre and nombre.upper() != "CARGO":
            return nombre
    # Fallback: nombre del archivo, limpiando numeracion y extension
    nombre = re.sub(r"^\d+[\.\-\s]+", "", fallback_nombre_archivo)
    nombre = re.sub(r"\.(docx?|xlsx?|pdf)$", "", nombre, flags=re.IGNORECASE)
    return nombre.strip()


def cargar_cargos_desde_manuales(directorio: Path | str) -> list[dict]:
    """Recorre una carpeta con manuales de funciones y devuelve una lista de
    dicts {nombre, texto_manual, archivo}.

    Soporta:
      - .docx : Word de Microsoft (recomendado)
      - .xlsx : Excel de Microsoft (todas las hojas y celdas)
      - .pdf  : PDF (texto + tablas via pdfplumber)
    """
    directorio = Path(directorio)
    if not directorio.exists() or not directorio.is_dir():
        logger.warning("Carpeta de manuales no existe: %s", directorio)
        return []

    extractores = {
        ".docx": _extraer_texto_docx,
        ".xlsx": _extraer_texto_xlsx,
        ".pdf": _extraer_texto_pdf,
    }
    cargos = []
    for archivo in sorted(directorio.iterdir()):
        if not archivo.is_file():
            continue
        suffix = archivo.suffix.lower()
        extractor = extractores.get(suffix)
        if extractor is None:
            continue
        if archivo.name.startswith("~$"):
            continue
        texto = extractor(archivo)
        if not texto:
            continue
        nombre = _extraer_denominacion(texto, archivo.name)
        cargos.append({
            "nombre": nombre,
            "texto_manual": texto,
            "archivo": archivo.name,
        })
        logger.info(
            "Manual cargado: %s -> cargo '%s' (%d chars, formato %s)",
            archivo.name, nombre, len(texto), suffix,
        )
    return cargos


def _detectar_tipo_cargo(nombre: str) -> str:
    """Heuristica para inferir el tipo de cargo desde el nombre."""
    n = (nombre or "").lower()
    if any(k in n for k in ["conductor", "chofer", "motorizado", "mensajero"]):
        return "conductor"
    if any(k in n for k in ["servicios generales", "aseo", "limpieza", "auxiliar de servicios"]):
        return "servicios_generales"
    if any(k in n for k in [
        "operari", "operador", "tecnic", "mecanic", "soldador", "ayudant",
        "rectificador", "lavador", "tornero", "produccion", "planta",
        "almacenista", "bodeguero", "obrero", "instalador",
    ]):
        return "operativo"
    return "administrativo"


# ---------------------------------------------------------------------------
# Generacion con Claude (actividades + tareas + peligros adicionales)
# ---------------------------------------------------------------------------

def _prompt_cargo(cliente: dict, cargo: dict, tipo_cargo: str) -> str:
    """Prompt para Claude. Si hay 'texto_manual' (manual de funciones del cargo),
    se incluye y se le pide a Claude que GENERE TODA la matriz IPVR para ese cargo
    a partir de las funciones reales descritas. Si no hay manual, se usan los
    catalogos base + extras por tipo (modo legacy)."""
    contexto_empresa = f"""EMPRESA: {cliente.get('razon_social', 'N/A')}
ACTIVIDAD ECONÓMICA: {cliente.get('actividad_economica', 'N/A')}
NIVEL DE RIESGO ARL: {cliente.get('nivel_riesgo', 'N/A')}
DESCRIPCIÓN DE LA EMPRESA: {cliente.get('descripcion_empresa', 'N/A')}
CIUDAD: {cliente.get('ciudad_empresa', 'N/A')}"""

    texto_manual = cargo.get("texto_manual", "").strip()

    if texto_manual:
        # Modo MANUAL: Claude lee el manual y genera TODOS los peligros aplicables
        return f"""Eres un experto en SG-SST en Colombia, GTC 45 y elaboración de Matrices IPVR.

{contexto_empresa}

Te entrego el MANUAL DE FUNCIONES de un cargo. Analízalo y genera la matriz de peligros IPVR completa SOLO para los peligros que REALMENTE apliquen a las funciones descritas.

=== MANUAL DE FUNCIONES ===
{texto_manual[:8000]}
=== FIN DEL MANUAL ===

Genera EXCLUSIVAMENTE un objeto JSON válido (sin texto adicional, sin ```json) con esta estructura:

{{
  "nombre": "Nombre del cargo (extraído de DENOMINACIÓN del manual)",
  "proceso": "Nombre del proceso/área (ej: ADMINISTRATIVO, PRODUCCIÓN, OPERATIVO, etc.)",
  "zona": "Zona física donde labora (ej: Oficinas Administrativas, Estudios de Producción, Taller, Vehículo, etc.)",
  "actividades": "Resumen del propósito del cargo en 2-4 líneas (basado en OBJETIVOS DEL CARGO).",
  "tareas": "• Tarea 1\\n• Tarea 2\\n... (entre 6 y 10 bullets con las funciones más importantes del manual, redactadas concisamente)",
  "peligros": [
    {{
      "peligro": "Descripción detallada del peligro según GTC 45 (ej: 'EXPOSICIÓN A RADIACIONES NO IONIZANTES POR USO DE PANTALLAS DE COMPUTADOR')",
      "clasificacion": "FÍSICO|QUÍMICO|BIOLÓGICO|BIOMECÁNICO|PSICOSOCIAL|CONDICIONES DE SEGURIDAD|FENÓMENOS NATURALES",
      "efectos": "Efectos posibles en la salud (separar con \\n)",
      "fuente": "Controles existentes en la fuente (NO OBSERVADO si no hay)",
      "medio": "Controles existentes en el medio (NO OBSERVADO si no hay)",
      "individuo": "Controles existentes en el individuo",
      "nd": 2,
      "ne": 3,
      "nc": 25,
      "peor": "Peor consecuencia posible",
      "req_legal": "SI",
      "eliminacion": "N/A",
      "sustitucion": "N/A",
      "ingenieria": "Controles de ingeniería propuestos o N/A",
      "administrativos": "Controles administrativos, capacitaciones, señalización propuestos",
      "epp": "EPP requerido o N/A"
    }}
  ]
}}

REGLAS DE VALORACIÓN GTC 45:
- nd (Nivel de Deficiencia): 2=Bajo, 6=Medio, 10=Alto, 14=Muy Alto
- ne (Nivel de Exposición): 1=Esporádica, 2=Ocasional, 3=Frecuente, 4=Continua
- nc (Nivel de Consecuencia): 10=Leve, 25=Grave, 60=Muy Grave, 100=Mortal

PELIGROS QUE DEBES SIEMPRE EVALUAR (incluir si aplica al cargo):
1. BIOLÓGICO: exposición a microorganismos
2. FÍSICO ruido (intermitente o continuo según el caso)
3. FÍSICO radiaciones no ionizantes (pantallas) — siempre que use computador
4. PSICOSOCIAL — gestión organizacional (siempre aplica)
5. BIOMECÁNICO postura (sedente para oficina / bípeda para operativos)
6. BIOMECÁNICO movimiento repetitivo (digitación o manipulación)
7. CONDICIONES DE SEGURIDAD locativo (siempre aplica)
8. CONDICIONES DE SEGURIDAD eléctrico (siempre aplica)
9. CONDICIONES DE SEGURIDAD accidente de tránsito (siempre que se desplace)
10. CONDICIONES DE SEGURIDAD públicos (siempre aplica)
11. FENÓMENOS NATURALES (siempre aplica)

PELIGROS ESPECIALES SI EL MANUAL LOS INDICA:
- Trabajo en alturas (>1.5m)
- Espacios confinados
- Soldadura, oxicorte
- Manejo de químicos, gases, vapores
- Conducción de vehículos como tarea principal
- Manejo de cargas
- Trabajo con maquinaria (atrapamiento mecánico)
- Trabajo en exteriores (radiación solar UV, temperatura)
- Trabajo nocturno

REGLAS GENERALES:
- Devolver entre 8 y 16 peligros (lo que aplique realmente al cargo).
- USAR TILDES correctamente (Español Colombia: organización, función, capacitación, exámenes, periódicos, eléctrico, mecánico, biológico, químico, psicosocial, biomecánico, fenómenos, eléctrica, médico).
- Los textos de descripción y controles en MAYÚSCULAS (sin perder tildes).
- Devuelve SOLO el JSON, sin explicaciones.
"""

    # Modo LEGACY: solo actividades/tareas + peligros adicionales (catalogo base se anade en codigo)
    return f"""Eres un experto en SG-SST en Colombia y la GTC 45.

{contexto_empresa}

CARGO A ANALIZAR: {cargo.get('nombre')}
TIPO DE CARGO INFERIDO: {tipo_cargo}

Genera EXCLUSIVAMENTE un objeto JSON válido con esta estructura:

{{
  "actividades": "Propósito y responsabilidad general del cargo (2-4 líneas).",
  "tareas": "• Tarea 1\\n• Tarea 2\\n... (5 a 8 bullets reales del cargo).",
  "peligros_adicionales": []
}}

REGLAS:
- ESPECÍFICO para este cargo y empresa.
- Si hay peligros muy específicos no cubiertos por el catálogo base, agrégalos en "peligros_adicionales" como objetos con: peligro, clasificacion, efectos, fuente, medio, individuo, nd (2|6|10), ne (1|2|3|4), nc (10|25|60|100), peor, req_legal, eliminacion, sustitucion, ingenieria, administrativos, epp.
- USAR TILDES correctamente (Español Colombia).
- Devuelve SOLO el JSON.
"""


# Flag global de sesion: si Claude falla por saldo bajo, no volver a intentar
_API_SIN_SALDO = False


def _generar_datos_cargo(cliente: dict, cargo: dict, tipo_cargo: str) -> dict:
    """Llama a Claude.

    Modo MANUAL (cargo trae 'texto_manual'): devuelve la matriz IPVR completa
      generada por Claude para ese cargo.
        -> {actividades, tareas, peligros: [...full list...], proceso, zona,
            "modo": "manual"}
    Modo LEGACY (sin manual): devuelve actividades, tareas y peligros adicionales
      (luego se concatenan con los catalogos base).
        -> {actividades, tareas, peligros_adicionales, "modo": "legacy"}
    """
    actividades_default = (
        f"Desempeñar las funciones propias del cargo de {cargo.get('nombre')} "
        "en cumplimiento de las políticas, procedimientos y obligaciones legales de la empresa."
    )
    tareas_default = (
        "• Cumplir con las funciones asignadas al cargo\n"
        "• Asistir a capacitaciones y reuniones programadas\n"
        "• Reportar actos y condiciones inseguras\n"
        "• Usar adecuadamente los elementos de protección personal\n"
        "• Cumplir con el reglamento interno de SST"
    )
    tiene_manual = bool(cargo.get("texto_manual", "").strip())
    fallback = {
        "modo": "manual" if tiene_manual else "legacy",
        "actividades": cargo.get("actividades") or actividades_default,
        "tareas": cargo.get("tareas") or tareas_default,
        "peligros_adicionales": [],
        "peligros": [],
        "proceso": cargo.get("proceso"),
        "zona": cargo.get("zona"),
    }

    # Si una llamada anterior fallo por saldo bajo, no volver a intentar
    global _API_SIN_SALDO
    if _API_SIN_SALDO:
        return fallback

    try:
        from . import claude_service
        from .config import ANTHROPIC_API_KEY, CLAUDE_MODEL
        if not ANTHROPIC_API_KEY:
            return fallback
        client = claude_service._crear_client()
        prompt = _prompt_cargo(cliente, cargo, tipo_cargo)
        max_tokens = 8192 if tiene_manual else 2048
        response = client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}],
        )
        claude_service._registrar_uso(response)
        raw = response.content[0].text.strip()
        # Limpiar fences si los hay
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        data = json.loads(raw)
        if tiene_manual:
            return {
                "modo": "manual",
                "nombre": data.get("nombre") or cargo.get("nombre"),
                "proceso": cargo.get("proceso") or data.get("proceso"),
                "zona": cargo.get("zona") or data.get("zona"),
                "actividades": data.get("actividades") or actividades_default,
                "tareas": data.get("tareas") or tareas_default,
                "peligros": data.get("peligros") or [],
            }
        else:
            return {
                "modo": "legacy",
                "actividades": cargo.get("actividades") or data.get("actividades") or actividades_default,
                "tareas": cargo.get("tareas") or data.get("tareas") or tareas_default,
                "peligros_adicionales": data.get("peligros_adicionales") or [],
            }
    except Exception as e:
        msg = str(e)
        # Si es saldo insuficiente, activar flag para no volver a intentar
        if "credit balance" in msg.lower() or "insufficient_quota" in msg.lower():
            if not _API_SIN_SALDO:
                logger.info(
                    "Claude API sin saldo. Continuando sin humanizacion (usando catalogo base + manuales).",
                )
                _API_SIN_SALDO = True
        else:
            logger.warning(
                "No se pudo generar datos IA para cargo '%s': %s", cargo.get("nombre"), e,
            )
        return fallback


# ---------------------------------------------------------------------------
# Escritura del Excel
# ---------------------------------------------------------------------------

# Mapeo columna numero -> letra
def _col(letra: str) -> int:
    return {chr(ord('A') + i): i + 1 for i in range(26)}.get(letra) or {
        "AA": 27, "AB": 28
    }[letra]


def _formula_np(row: int) -> str:
    return f"=$M{row}*$N{row}"


def _formula_interpret_np(row: int) -> str:
    return (
        f'=IF(O{row}>=24,"MUY ALTO",'
        f'IF((O{row}>=10)*OR(O{row}<=20),"ALTO",'
        f'IF((O{row}>=6)*OR(O{row}<=8),"MEDIO",'
        f'IF((O{row}<=4)*OR(O{row}>=2),"BAJO"))))'
    )


def _formula_nr(row: int) -> str:
    return f"=O{row}*Q{row}"


def _formula_interpret_nr(row: int) -> str:
    return (
        f'=IF(R{row}>=600,"I",'
        f'IF((R{row}>=150)*OR(R{row}<=500),"II",'
        f'IF((R{row}>=40)*OR(R{row}<=120),"III",'
        f'IF((R{row}<=20)*OR(R{row}>=1),"IV"))))'
    )


def _formula_aceptabilidad(row: int) -> str:
    return (
        f'=IF(S{row}="I","NO ACEPTABLE",'
        f'IF(S{row}="II","NO ACEPTABLE O ACEPTABLE CON CONTROL ESPECIFICO",'
        f'IF(S{row}="III","MEJORABLE",'
        f'IF(S{row}="IV","ACEPTABLE"))))'
    )


# Estilos basicos de celda
_BORDE = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
_ALIGN_WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
_ALIGN_LEFT = Alignment(wrap_text=True, vertical="center", horizontal="left")


def _escribir_fila_peligro(ws, row: int, peligro: dict) -> None:
    """Escribe los campos del peligro (G..AB) en una fila."""
    ws.cell(row=row, column=7, value=peligro.get("peligro", "")).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=8, value=peligro.get("clasificacion", "")).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=9, value=peligro.get("efectos", "")).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=10, value=peligro.get("fuente", "NO OBSERVADO")).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=11, value=peligro.get("medio", "NO OBSERVADO")).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=12, value=peligro.get("individuo", "")).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=13, value=int(peligro.get("nd", 2))).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=14, value=int(peligro.get("ne", 1))).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=15, value=_formula_np(row)).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=16, value=_formula_interpret_np(row)).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=17, value=int(peligro.get("nc", 10))).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=18, value=_formula_nr(row)).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=19, value=_formula_interpret_nr(row)).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=20, value=_formula_aceptabilidad(row)).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=22, value=peligro.get("peor", "")).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=23, value=peligro.get("req_legal", "SI")).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=24, value=peligro.get("eliminacion", "N/A")).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=25, value=peligro.get("sustitucion", "N/A")).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=26, value=peligro.get("ingenieria", "N/A")).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=27, value=peligro.get("administrativos", "")).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=28, value=peligro.get("epp", "N/A")).alignment = _ALIGN_LEFT
    # Aplicar borde y normalizar negrilla en celdas G..AB
    # Bold: H (CLASIFICACION, col 8) y T (ACEPTABILIDAD, col 20)
    # Resto: bold=False
    for c in range(7, 29):
        cel = ws.cell(row=row, column=c)
        cel.border = _BORDE
        _set_bold(cel, c in (8, 20))


def _aplicar_color_bloque(ws, row_inicio: int, row_fin: int, fill: PatternFill) -> None:
    """Pinta el bloque del cargo:
    - Cols A..B (PROCESO, ZONA): siempre DURAZNO
    - Cols C..L (CARGO..INDIVIDUO): color alternado del cargo
    """
    for r in range(row_inicio, row_fin + 1):
        # Columnas fijas A y B siempre durazno
        for c in range(_COL_FIJO_INI, _COL_FIJO_FIN + 1):
            ws.cell(row=r, column=c).fill = _FILL_PROCESO_ZONA
        # Columnas C..L con el color del cargo (azul/verde alternado)
        for c in range(_COL_FILL_INI, _COL_FILL_FIN + 1):
            ws.cell(row=r, column=c).fill = fill


def _escribir_cabecera_cargo(
    ws, row_inicio: int, row_fin: int, cargo: dict, datos: dict, tipo_cargo: str,
) -> None:
    """Escribe y hace merge de los campos por cargo (A..F)."""
    proceso = (cargo.get("proceso")
               or datos.get("proceso")
               or ("ADMINISTRATIVO" if tipo_cargo == "administrativo" else "OPERATIVO"))
    zona = (cargo.get("zona")
            or datos.get("zona")
            or ("Oficinas Administrativas" if tipo_cargo == "administrativo" else "Área Operativa"))
    rutinaria = cargo.get("rutinaria") or "SI"

    # NO escribir A (proceso) ni B (zona) aqui: se hace despues a nivel de
    # GRUPO DE PROCESO (todos los cargos del mismo proceso comparten merge en A y B).
    celda_cargo = ws.cell(row=row_inicio, column=3, value=cargo.get("nombre", ""))
    celda_cargo.alignment = _ALIGN_WRAP
    _set_bold(celda_cargo, True)  # CARGO en negrilla
    ws.cell(row=row_inicio, column=4, value=datos.get("actividades", "")).alignment = _ALIGN_LEFT
    ws.cell(row=row_inicio, column=5, value=datos.get("tareas", "")).alignment = _ALIGN_LEFT
    ws.cell(row=row_inicio, column=6, value=rutinaria).alignment = _ALIGN_WRAP

    # Bordes y normalizar negrilla en celdas del bloque (A..F)
    # Bold: A (PROCESO) y C (CARGO) en row_inicio, F (TAREA RUTINARIA) en row_inicio
    # Resto: bold=False
    for r in range(row_inicio, row_fin + 1):
        for c in range(1, 7):
            cel = ws.cell(row=r, column=c)
            cel.border = _BORDE
            if r == row_inicio and c in (3, 6):  # CARGO y TAREA RUTINARIA en negrilla
                _set_bold(cel, True)
            elif r != row_inicio or c not in (1,):  # A se setea aparte en grupo
                _set_bold(cel, False)

    # Merge interno del cargo solo en C, D, E, F (A y B se fusionan despues
    # a nivel de PROCESO, abarcando todos los cargos del mismo proceso).
    if row_fin > row_inicio:
        for col_letter in ["C", "D", "E", "F"]:
            ws.merge_cells(f"{col_letter}{row_inicio}:{col_letter}{row_fin}")

    # Numero de expuestos en columna U (21) por cada fila del cargo
    num_exp = cargo.get("num_expuestos", 1)
    try:
        num_exp = int(num_exp)
    except (TypeError, ValueError):
        num_exp = 1
    for r in range(row_inicio, row_fin + 1):
        ws.cell(row=r, column=21, value=num_exp).alignment = _ALIGN_WRAP


def generar_matriz_peligros(
    cliente: dict,
    output_dir: Path,
    config_dir: Path | None = None,
) -> Path | None:
    """Llena la matriz SST-MTZ-012 con un bloque por cada cargo.

    Hay dos modos:

    1) **Modo MANUAL (recomendado)**: el YAML trae `manuales_cargos_dir: "ruta"`.
       Se escanea la carpeta y por cada .docx (un manual de funciones por cargo)
       Claude lee el manual completo y genera la matriz IPVR adaptada.

    2) **Modo LEGACY**: el YAML trae `cargos:` con lista de cargos. Se usa el
       catalogo base + extras por tipo + peligros adicionales generados por IA.

    Si trae ambos, se usan los manuales (modo 1).

    Args:
        cliente: dict cargado desde el YAML.
        output_dir: carpeta raiz del SG-SST generado.
        config_dir: carpeta del YAML del cliente (para resolver rutas relativas).

    Returns:
        Ruta del archivo Excel actualizado, o None si no aplica.
    """
    cargos: list[dict] = []

    # Modo 1: carpeta de manuales de funciones
    dir_manuales = cliente.get("manuales_cargos_dir")
    if dir_manuales:
        ruta_manuales = Path(dir_manuales)
        if not ruta_manuales.is_absolute() and config_dir:
            ruta_manuales = (Path(config_dir) / ruta_manuales).resolve()
        cargos = cargar_cargos_desde_manuales(ruta_manuales)
        # Mezclar configuracion explicita del YAML (num_expuestos, etc.)
        # match por nombre (case-insensitive)
        overrides = {
            (c.get("nombre") or "").strip().lower(): c
            for c in (cliente.get("cargos") or [])
        }
        for cg in cargos:
            ov = overrides.get((cg.get("nombre") or "").strip().lower())
            if ov:
                for k, v in ov.items():
                    if k != "nombre" and v not in (None, ""):
                        cg[k] = v

    # Modo 2: lista en YAML (sin manuales)
    if not cargos:
        cargos = cliente.get("cargos") or []

    if not cargos:
        logger.info("Cliente sin cargos definidos; matriz IPVR no se llena.")
        return None

    # Guardar el total de cargos procesados en el cliente para que el
    # mensaje en consola refleje el numero real (los cargos pueden venir de
    # manuales y no estar todavia en cliente['cargos']).
    cliente["_cargos_procesados"] = len(cargos)

    matriz_path = output_dir / RUTA_MATRIZ_IPVR
    if not matriz_path.exists():
        logger.warning("No se encontro la matriz IPVR: %s", matriz_path)
        return None

    wb = load_workbook(str(matriz_path))
    if NOMBRE_HOJA not in wb.sheetnames:
        logger.warning("Hoja '%s' no encontrada en %s", NOMBRE_HOJA, matriz_path)
        return None
    ws = wb[NOMBRE_HOJA]

    # Liberar merges previos en zona de datos para evitar conflictos
    rangos_a_quitar = []
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= FILA_INICIO:
            rangos_a_quitar.append(str(mr))
    for r in rangos_a_quitar:
        ws.unmerge_cells(r)

    # PASO 1: generar datos (Claude) para cada cargo y resolver su proceso
    cargos_con_datos = []
    for cargo in cargos:
        nombre = (cargo.get("nombre") or "").strip()
        if not nombre:
            continue
        tipo_cargo = (cargo.get("tipo") or _detectar_tipo_cargo(nombre)).lower()
        datos = _generar_datos_cargo(cliente, cargo, tipo_cargo)
        proceso_resuelto = (
            cargo.get("proceso")
            or datos.get("proceso")
            or ("ADMINISTRATIVO" if tipo_cargo == "administrativo" else "OPERATIVO")
        )
        cargos_con_datos.append({
            "cargo": cargo,
            "nombre": nombre,
            "tipo_cargo": tipo_cargo,
            "datos": datos,
            "proceso": str(proceso_resuelto).upper().strip(),
        })

    # PASO 2: ordenar por proceso (agrupar todos los del mismo proceso juntos)
    # Orden estable: respeta el orden de aparicion del primer cargo de cada proceso
    orden_procesos: list[str] = []
    for item in cargos_con_datos:
        if item["proceso"] not in orden_procesos:
            orden_procesos.append(item["proceso"])
    cargos_con_datos.sort(key=lambda x: orden_procesos.index(x["proceso"]))

    fila_actual = FILA_INICIO
    indice_cargo = 0
    # Rastreo de grupos por proceso para merge de columnas A y B al final.
    # Cada grupo: {"proceso": str, "zona": str, "row_ini": int, "row_fin": int}
    grupos_proceso: list[dict] = []

    for item in cargos_con_datos:
        cargo = item["cargo"]
        # Forzar el proceso resuelto en el cargo para que la celda lo refleje
        cargo["proceso"] = item["proceso"]
        nombre = item["nombre"]
        tipo_cargo = item["tipo_cargo"]
        datos = item["datos"]

        if datos.get("modo") == "manual" and datos.get("peligros"):
            # Claude genero la lista completa a partir del manual de funciones
            peligros = list(datos["peligros"])
            # Si Claude regreso menos de 5, complementar con catalogo base por seguridad
            if len(peligros) < 5:
                logger.info(
                    "Cargo '%s': solo %d peligros desde manual; complementando con catalogo base.",
                    nombre, len(peligros),
                )
                peligros = peligros + list(PELIGROS_BASE_ADMINISTRATIVO)
        else:
            # Modo legacy: catalogo base + extras por tipo + extras IA
            peligros_extras_tipo = EXTRAS_POR_TIPO.get(tipo_cargo, [])
            peligros_extra_ia = datos.get("peligros_adicionales", []) or []
            peligros = (
                list(PELIGROS_BASE_ADMINISTRATIVO)
                + list(peligros_extras_tipo)
                + list(peligros_extra_ia)
            )

        if not peligros:
            continue

        row_ini = fila_actual
        for i, pel in enumerate(peligros):
            _escribir_fila_peligro(ws, fila_actual, pel)
            fila_actual += 1
        row_fin = fila_actual - 1

        _escribir_cabecera_cargo(ws, row_ini, row_fin, cargo, datos, tipo_cargo)

        # Color alternado por cargo (azul claro / verde claro) en columnas C..L
        fill = _FILLS_CARGOS[indice_cargo % len(_FILLS_CARGOS)]
        _aplicar_color_bloque(ws, row_ini, row_fin, fill)
        indice_cargo += 1

        # Acumular en grupo de proceso (extender el grupo actual o iniciar uno nuevo)
        zona_cargo = (
            cargo.get("zona") or datos.get("zona")
            or ("Oficinas Administrativas" if tipo_cargo == "administrativo" else "Área Operativa")
        )
        if grupos_proceso and grupos_proceso[-1]["proceso"] == item["proceso"]:
            grupos_proceso[-1]["row_fin"] = row_fin
            # mantener primera zona; si no habia, asumir la actual
            if not grupos_proceso[-1].get("zona"):
                grupos_proceso[-1]["zona"] = zona_cargo
        else:
            grupos_proceso.append({
                "proceso": item["proceso"],
                "zona": zona_cargo,
                "row_ini": row_ini,
                "row_fin": row_fin,
            })

        logger.info(
            "IPVR: cargo '%s' (%s) -> %d filas (R%d-R%d)",
            nombre, tipo_cargo, row_fin - row_ini + 1, row_ini, row_fin,
        )

    # PASO 3: para cada grupo de proceso, escribir A y B en la primera fila
    # y hacer merge vertical sobre TODO el rango del grupo.
    for grupo in grupos_proceso:
        r_ini = grupo["row_ini"]
        r_fin = grupo["row_fin"]
        celda_proceso = ws.cell(row=r_ini, column=1, value=grupo["proceso"])
        celda_proceso.alignment = _ALIGN_WRAP
        _set_bold(celda_proceso, True)  # PROCESO en negrilla
        celda_zona = ws.cell(row=r_ini, column=2, value=grupo["zona"])
        celda_zona.alignment = _ALIGN_WRAP
        _set_bold(celda_zona, False)  # ZONA NO en negrilla
        # Bordes en todo el rango A..B del grupo
        for r in range(r_ini, r_fin + 1):
            ws.cell(row=r, column=1).border = _BORDE
            ws.cell(row=r, column=2).border = _BORDE
        if r_fin > r_ini:
            ws.merge_cells(f"A{r_ini}:A{r_fin}")
            ws.merge_cells(f"B{r_ini}:B{r_fin}")

    wb.save(str(matriz_path))
    logger.info("Matriz IPVR generada: %s (%d cargos)", matriz_path, len(cargos))
    return matriz_path
