"""Generador de la Matriz de Peligros y Valoracion de Riesgos Viales (PESV-MTZ-002).

A partir de los cargos del YAML/manuales de la empresa, identifica los actores
viales aplicables (conductor, motociclista, pasajero, peaton) y genera la
matriz agrupada por 5 FACTORES de la Resolucion 40595/2022 - PESV:

    1. FACTOR HUMANO
    2. FACTOR MECANICO
    3. FACTOR INFRAESTRUCTURA VIAL
    4. FACTOR AMBIENTAL
    5. FACTOR SOCIAL

Cada factor tiene un catalogo de peligros distintos por TIPO DE ACTOR VIAL.

Solo se genera si el cliente tiene `generar_pesv: true`.
"""
from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font

logger = logging.getLogger(__name__)

# Ruta dentro del SG-SST generado
RUTA_MATRIZ_VIAL = (
    Path("PLAN ESTRATÉGICO DE SEGURIDAD VIAL")
    / "FASE 1"
    / "6. CARACTERIZACIÓN, EVALUACIÓN Y CONTROL DE RIESGOS"
    / "PESV-MTZ-002 MATRIZ DE PELIGROS Y VALORACIÓN DE RIESGOS VIALES.xlsx"
)
NOMBRE_HOJA = "Matriz de peligros"
FILA_INICIO = 7  # primera fila de datos (despues de los headers)

# ---------------------------------------------------------------------------
# Catalogo de peligros por FACTOR x ACTOR VIAL
# ---------------------------------------------------------------------------

def _peligro(g, i, j, k, l, prob, expo, nc, peor, ingen, admin, epp, plan):
    return {
        "peligro": g, "efectos": i,
        "fuente": j, "medio": k, "individuo": l,
        "nivel_probabilidad": prob,  # "No es Probable" | "Poco Probable" | "Muy Probable"
        "nivel_exposicion": expo,    # "Esporadica" | "Ocasional" | "Frecuente"
        "no_expuestos": nc, "peor": peor,
        "eliminacion": "N/A", "sustitucion": "N/A",
        "ingenieria": ingen, "administrativos": admin,
        "epp": epp, "plan_accion": plan,
    }

# FACTOR HUMANO -------------------------------------------------------------
PELIGROS_HUMANO_CONDUCTOR = [
    _peligro("Exceso de velocidad.",
             "Siniestro vial, heridas, contusiones, fracturas, laceraciones, traumatismo craneoencefálico, muerte.",
             "SIN CONTROL", "Monitoreo GPS instalado en vehículo",
             "Capacitación en manejo defensivo y seguridad vial",
             "Poco Probable", "Ocasional", 1, "Muerte",
             "N/A", "1. Capacitación en manejo defensivo.\n2. Aplicar política de velocidades máximas.",
             "N/A", "1. Incluir en plan de capacitación PESV.\n2. Control de velocidad con GPS."),
    _peligro("Consumo de alcohol o sustancias psicoactivas antes o durante la conducción.",
             "Siniestro vial, heridas, fracturas, muerte.",
             "Política sobre consumo de alcohol, tabaco y sustancias psicoactivas.",
             "SIN CONTROL", "Política de consumo cero",
             "No es Probable", "Ocasional", 1, "Muerte",
             "N/A", "1. Pruebas aleatorias de alcoholimetría.\n2. Capacitación en autocuidado.",
             "N/A", "1. Implementar pruebas aleatorias.\n2. Socialización de política."),
    _peligro("Maniobras peligrosas (adelantamientos indebidos, cambios bruscos de carril).",
             "Siniestro vial, heridas, contusiones, fracturas, muerte.",
             "SIN CONTROL", "SIN CONTROL", "Capacitación en manejo defensivo",
             "Muy Probable", "Ocasional", 1, "Muerte",
             "N/A", "Capacitación sobre responsabilidad vial y manejo defensivo.",
             "N/A", "1. Plan de capacitación PESV.\n2. Reforzar normas internas."),
    _peligro("Sueño o fatiga durante la conducción.",
             "Siniestro vial, lesiones, muerte.",
             "SIN CONTROL", "SIN CONTROL", "Control de jornadas y descansos",
             "Poco Probable", "Ocasional", 1, "Muerte",
             "N/A", "Política de pausas activas durante conducción y control de jornadas.",
             "N/A", "1. Implementar control de horas de conducción.\n2. Capacitación en fatiga."),
    _peligro("Distracción por uso de celular o dispositivos electrónicos.",
             "Siniestro vial, heridas, muerte.",
             "Política de uso de celular durante conducción", "SIN CONTROL",
             "Capacitación sobre distractores",
             "Poco Probable", "Frecuente", 1, "Muerte",
             "Manos libres en vehículo institucional.",
             "1. Política de no uso de celular.\n2. Capacitación en distractores.",
             "N/A", "1. Refuerzo de política.\n2. Sanciones disciplinarias."),
    _peligro("Estado psicológico adverso (estrés, conflictos personales) durante la conducción.",
             "Siniestro vial, lesiones, muerte.",
             "SIN CONTROL", "Programa de bienestar psicosocial",
             "Acompañamiento psicológico",
             "Poco Probable", "Ocasional", 1, "Muerte",
             "N/A", "Aplicación batería psicosocial y seguimiento.",
             "N/A", "1. Incluir en programa de bienestar.\n2. Atención psicosocial."),
]

PELIGROS_HUMANO_MOTOCICLISTA = [
    _peligro("Exceso de velocidad en motocicleta.",
             "Siniestro vial, fracturas, traumatismo, muerte.",
             "SIN CONTROL", "Monitoreo de tiempos de entrega",
             "Capacitación en seguridad vial para motociclistas",
             "Muy Probable", "Frecuente", 1, "Muerte",
             "N/A", "1. Política de velocidades.\n2. Capacitación específica motociclistas.",
             "Casco certificado, chaleco reflectivo, rodilleras, guantes",
             "1. Implementar curso de motociclismo seguro.\n2. Auditoría de tiempos."),
    _peligro("Consumo de alcohol o sustancias psicoactivas antes o durante la conducción.",
             "Siniestro vial, lesiones, muerte.",
             "Política de consumo cero", "SIN CONTROL", "Política de consumo cero",
             "No es Probable", "Ocasional", 1, "Muerte",
             "N/A", "Pruebas aleatorias de alcoholimetría.",
             "N/A", "1. Pruebas periódicas.\n2. Socialización de política."),
    _peligro("Maniobras peligrosas en moto (zigzag entre vehículos, adelantamientos por la derecha).",
             "Siniestro vial, fracturas, muerte.",
             "SIN CONTROL", "SIN CONTROL", "Capacitación en manejo defensivo",
             "Muy Probable", "Frecuente", 1, "Muerte",
             "N/A", "Capacitación específica para motociclistas.",
             "EPP completo de motociclista", "1. Curso de motociclismo defensivo.\n2. Reforzar normas."),
    _peligro("Sueño o fatiga durante la conducción de motocicleta.",
             "Siniestro vial, lesiones, muerte.",
             "SIN CONTROL", "Control de turnos", "Pausas activas",
             "Poco Probable", "Frecuente", 1, "Muerte",
             "N/A", "Control de jornadas y rutas.",
             "N/A", "1. Política de descansos.\n2. Limitación de horas de operación."),
    _peligro("Uso de auriculares o celular durante la conducción de motocicleta.",
             "Siniestro vial, distracción, muerte.",
             "Política prohibición de auriculares", "SIN CONTROL",
             "Capacitación sobre distractores",
             "Poco Probable", "Frecuente", 1, "Muerte",
             "N/A", "1. Política de prohibición.\n2. Capacitación.",
             "N/A", "1. Refuerzo de política.\n2. Inspecciones aleatorias."),
]

PELIGROS_HUMANO_PEATON = [
    _peligro("Actos inseguros al cruzar la vía (cruzar fuera del paso peatonal, contra el semáforo).",
             "Atropellamiento, lesiones, fracturas, muerte.",
             "SIN CONTROL", "Capacitación peatonal", "Autocuidado",
             "Poco Probable", "Frecuente", 1, "Muerte",
             "N/A", "Capacitación en seguridad peatonal.",
             "N/A", "1. Charlas de autocuidado peatonal.\n2. Señalización."),
    _peligro("Distracción por uso de celular al caminar en vía pública.",
             "Atropellamiento, caídas, lesiones.",
             "SIN CONTROL", "SIN CONTROL", "Sensibilización autocuidado",
             "Poco Probable", "Frecuente", 1, "Lesiones graves",
             "N/A", "Capacitación en uso responsable de celular en vía.",
             "N/A", "1. Sensibilización.\n2. Refuerzo de normas."),
    _peligro("Sueño o fatiga durante desplazamiento peatonal.",
             "Caídas, atropellamiento, lesiones.",
             "SIN CONTROL", "Control de jornadas", "Pausas durante el día",
             "No es Probable", "Ocasional", 1, "Lesiones",
             "N/A", "Política de pausas activas.",
             "N/A", "1. Programa de pausas.\n2. Control de jornadas."),
]

PELIGROS_HUMANO_PASAJERO = [
    _peligro("Conducción imprudente del conductor del vehículo de transporte (taxi, app, transporte público).",
             "Siniestro vial, lesiones, muerte.",
             "SIN CONTROL", "Verificación de placas y conductor",
             "Capacitación en autocuidado como pasajero",
             "Poco Probable", "Ocasional", 1, "Muerte",
             "N/A", "Política de transporte seguro (apps verificadas).",
             "Cinturón de seguridad", "1. Política de transporte autorizado.\n2. Verificación previa."),
    _peligro("Distracción del conductor por instrucciones del pasajero o equipo durante el trayecto.",
             "Siniestro vial, lesiones.",
             "SIN CONTROL", "SIN CONTROL", "Capacitación en seguridad como pasajero",
             "Poco Probable", "Ocasional", 1, "Lesiones graves",
             "N/A", "Capacitación: no distraer al conductor.",
             "Cinturón de seguridad", "1. Sensibilización.\n2. Plan de capacitación."),
]

# FACTOR MECANICO -----------------------------------------------------------
PELIGROS_MECANICO_CONDUCTOR = [
    _peligro("Fallas mecánicas del vehículo (frenos, dirección, suspensión).",
             "Siniestro vial, lesiones, muerte.",
             "Mantenimiento preventivo del vehículo",
             "Inspección preoperacional",
             "Capacitación en inspección de vehículo",
             "Poco Probable", "Ocasional", 1, "Muerte",
             "Programa de mantenimiento preventivo y correctivo.",
             "Inspección preoperacional documentada.",
             "N/A", "1. Auditoría de mantenimiento.\n2. Hoja de vida del vehículo."),
    _peligro("Falla en el sistema de iluminación del vehículo (luces, direccionales).",
             "Siniestro vial nocturno, atropellamientos.",
             "Mantenimiento preventivo", "Inspección preoperacional",
             "Capacitación inspección",
             "Poco Probable", "Ocasional", 1, "Lesiones graves",
             "Mantenimiento preventivo de luces.",
             "Lista de chequeo preoperacional.",
             "N/A", "1. Inspección documentada.\n2. Reposición inmediata."),
]

PELIGROS_MECANICO_MOTOCICLISTA = [
    _peligro("Fallas mecánicas en la moto (frenos, llantas, dirección, cadena).",
             "Siniestro vial, lesiones, muerte.",
             "Mantenimiento preventivo de la moto",
             "Inspección preoperacional documentada",
             "Capacitación en inspección de motocicleta",
             "Poco Probable", "Frecuente", 1, "Muerte",
             "Programa de mantenimiento preventivo.",
             "Hoja de vida del vehículo y bitácora de revisión.",
             "EPP del motociclista",
             "1. Auditoría mensual.\n2. Plan de mantenimiento."),
    _peligro("Llanta pinchada o desinflada durante la conducción.",
             "Pérdida de control, caída, lesiones.",
             "Inspección de llantas pre-viaje",
             "Kit de reparación",
             "Capacitación en respuesta",
             "Poco Probable", "Ocasional", 1, "Lesiones graves",
             "Llantas certificadas y kit de emergencia.",
             "Capacitación en mantenimiento básico.",
             "N/A", "1. Inspección preoperacional diaria.\n2. Repuestos disponibles."),
]

# FACTOR INFRAESTRUCTURA VIAL -----------------------------------------------
PELIGROS_INFRA_CONDUCTOR = [
    _peligro("Condiciones físicas de la vía: baches, rugosidades, pavimento deteriorado.",
             "Siniestro vial, lesiones.",
             "SIN CONTROL", "SIN CONTROL", "Capacitación en manejo en malas vías",
             "Poco Probable", "Frecuente", 1, "Lesiones graves",
             "N/A", "Selección de rutas seguras y reporte a entes viales.",
             "N/A", "1. Estudio de rutas.\n2. Reporte a autoridades."),
    _peligro("Geometría y topología de la vía: curvas pronunciadas, pendientes peligrosas.",
             "Siniestro vial, volcamientos.",
             "SIN CONTROL", "Señalización vial",
             "Capacitación en manejo en zonas críticas",
             "Poco Probable", "Ocasional", 1, "Muerte",
             "N/A", "Estudio de rutas críticas y planeación.",
             "N/A", "1. Mapa de rutas críticas.\n2. Capacitación específica."),
    _peligro("Señalización vial deficiente o ausente.",
             "Siniestro vial por desorientación.",
             "SIN CONTROL", "GPS y mapas",
             "Capacitación en navegación segura",
             "Poco Probable", "Ocasional", 1, "Lesiones graves",
             "N/A", "Reporte de fallas de señalización.",
             "N/A", "1. Reporte a autoridades.\n2. Apoyo en GPS."),
]

PELIGROS_INFRA_MOTOCICLISTA = [
    _peligro("Baches, huecos y pavimento deteriorado (mayor riesgo en moto).",
             "Caída, lesiones, muerte.",
             "SIN CONTROL", "SIN CONTROL", "Capacitación en manejo de moto en malas vías",
             "Muy Probable", "Frecuente", 1, "Muerte",
             "N/A", "Selección de rutas y reporte de fallas viales.",
             "EPP completo de motociclista",
             "1. Mapa de zonas críticas.\n2. Capacitación en manejo defensivo."),
    _peligro("Señalización vial deficiente o ausente.",
             "Pérdida de orientación, siniestros.",
             "SIN CONTROL", "GPS", "Capacitación",
             "Poco Probable", "Frecuente", 1, "Lesiones graves",
             "N/A", "Reporte y registro de fallas.",
             "N/A", "1. Reporte autoridades.\n2. Uso obligatorio GPS."),
]

PELIGROS_INFRA_PEATON = [
    _peligro("Andenes en mal estado: baches, huecos, obstáculos.",
             "Caídas, esguinces, fracturas.",
             "SIN CONTROL", "SIN CONTROL", "Capacitación autocuidado peatonal",
             "Poco Probable", "Frecuente", 1, "Lesiones graves",
             "N/A", "Reporte de fallas de andenes.",
             "Calzado adecuado", "1. Reporte autoridades.\n2. Sensibilización."),
    _peligro("Señalización peatonal deficiente (semáforos, pasos peatonales).",
             "Atropellamientos por desorientación.",
             "SIN CONTROL", "SIN CONTROL", "Sensibilización en autocuidado",
             "Poco Probable", "Ocasional", 1, "Muerte",
             "N/A", "Capacitación en seguridad peatonal.",
             "N/A", "1. Charlas autocuidado.\n2. Reporte a autoridades."),
]

# FACTOR AMBIENTAL ----------------------------------------------------------
PELIGROS_AMBIENTAL_CONDUCTOR = [
    _peligro("Condiciones climáticas adversas: lluvia intensa, neblina, granizada.",
             "Siniestro vial por baja visibilidad o piso resbaladizo.",
             "SIN CONTROL", "Inspección de limpiaparabrisas y luces",
             "Capacitación en manejo en malas condiciones",
             "Poco Probable", "Ocasional", 1, "Muerte",
             "Mantenimiento de limpiaparabrisas y luces antiniebla.",
             "Capacitación específica en manejo bajo lluvia.",
             "N/A", "1. Capacitación en manejo defensivo.\n2. Plan de contingencia."),
]

PELIGROS_AMBIENTAL_MOTOCICLISTA = [
    _peligro("Condiciones climáticas adversas: lluvia, neblina, granizada (alto riesgo en moto).",
             "Caídas, siniestros, muerte.",
             "Política de no operación en condiciones extremas",
             "Equipo impermeable",
             "Capacitación en manejo en lluvia",
             "Muy Probable", "Frecuente", 1, "Muerte",
             "N/A", "Política: suspensión de operaciones en lluvia intensa.",
             "Impermeable, casco con visor antiniebla",
             "1. Política de suspensión.\n2. Dotación impermeable."),
    _peligro("Exposición a radiación solar intensa durante conducción.",
             "Quemaduras solares, deshidratación, fatiga.",
             "SIN CONTROL", "Hidratación durante el día",
             "Bloqueador solar",
             "Poco Probable", "Frecuente", 1, "Lesiones",
             "N/A", "Suministro de bloqueador y agua.",
             "Bloqueador solar, gafas UV",
             "1. Política de hidratación.\n2. EPP solar."),
]

PELIGROS_AMBIENTAL_PEATON = [
    _peligro("Condiciones climáticas adversas: lluvia intensa, piso mojado y resbaladizo.",
             "Caídas, fracturas, atropellamientos.",
             "SIN CONTROL", "Equipo impermeable",
             "Calzado antideslizante",
             "Poco Probable", "Ocasional", 1, "Lesiones graves",
             "N/A", "Suspensión de diligencias externas en lluvia.",
             "Calzado antideslizante, paraguas",
             "1. Política de protección.\n2. Dotación adecuada."),
]

# FACTOR SOCIAL -------------------------------------------------------------
PELIGROS_SOCIAL_CONDUCTOR = [
    _peligro("Manifestaciones sociales, bloqueos y protestas en vía.",
             "Atrapamiento del vehículo, atentados, lesiones.",
             "SIN CONTROL", "Monitoreo de noticias y rutas",
             "Capacitación en autocuidado",
             "Poco Probable", "Ocasional", 1, "Lesiones graves",
             "N/A", "Plan de contingencia y rutas alternas.",
             "N/A", "1. Monitoreo de medios.\n2. Plan de rutas alternas."),
    _peligro("Terrorismo, ataques o vandalismo en vía.",
             "Lesiones, muerte.",
             "SIN CONTROL", "Coordinación con autoridades",
             "Capacitación en respuesta",
             "No es Probable", "Esporádica", 1, "Muerte",
             "N/A", "Procedimiento de emergencia.",
             "N/A", "1. Coordinar con policía.\n2. Plan de emergencias."),
]

PELIGROS_SOCIAL_MOTOCICLISTA = [
    _peligro("Manifestaciones sociales y bloqueos de vías.",
             "Imposibilidad de operar, atrapamiento.",
             "SIN CONTROL", "Monitoreo de noticias",
             "Capacitación rutas alternas",
             "Poco Probable", "Ocasional", 1, "Lesiones",
             "N/A", "Plan de rutas alternas.",
             "N/A", "1. Monitoreo medios.\n2. Plan de contingencia."),
    _peligro("Atracos, hurtos y ataques a motociclistas en vía.",
             "Lesiones, muerte, robo.",
             "SIN CONTROL", "Sistema de comunicación con base",
             "Capacitación en autocuidado",
             "Poco Probable", "Frecuente", 1, "Muerte",
             "Sistema GPS de pánico.",
             "Procedimiento ante atracos.",
             "N/A", "1. Capacitación específica.\n2. Equipo de comunicación."),
]

PELIGROS_SOCIAL_PEATON = [
    _peligro("Manifestaciones sociales y disturbios en zonas de diligencia.",
             "Atrapamiento, lesiones.",
             "SIN CONTROL", "Monitoreo previo de zonas",
             "Capacitación autocuidado",
             "Poco Probable", "Ocasional", 1, "Lesiones",
             "N/A", "Plan de contingencia.",
             "N/A", "1. Monitoreo previo.\n2. Comunicación con base."),
    _peligro("Atracos peatonales en diligencias externas.",
             "Lesiones, muerte, robo.",
             "SIN CONTROL", "Sistema de acompañamiento o llamada",
             "Capacitación en autocuidado",
             "Poco Probable", "Ocasional", 1, "Muerte",
             "N/A", "Procedimiento ante atracos.",
             "N/A", "1. Capacitación.\n2. No portar elementos llamativos."),
]

# Catalogo: factor -> tipo_actor -> lista de peligros
CATALOGO: dict[str, dict[str, list[dict]]] = {
    "FACTOR HUMANO": {
        "conductor": PELIGROS_HUMANO_CONDUCTOR,
        "motociclista": PELIGROS_HUMANO_MOTOCICLISTA,
        "peaton": PELIGROS_HUMANO_PEATON,
        "pasajero": PELIGROS_HUMANO_PASAJERO,
    },
    "FACTOR MECÁNICO": {
        "conductor": PELIGROS_MECANICO_CONDUCTOR,
        "motociclista": PELIGROS_MECANICO_MOTOCICLISTA,
    },
    "FACTOR INFRAESTRUCTURA VIAL": {
        "conductor": PELIGROS_INFRA_CONDUCTOR,
        "motociclista": PELIGROS_INFRA_MOTOCICLISTA,
        "peaton": PELIGROS_INFRA_PEATON,
    },
    "FACTOR AMBIENTAL": {
        "conductor": PELIGROS_AMBIENTAL_CONDUCTOR,
        "motociclista": PELIGROS_AMBIENTAL_MOTOCICLISTA,
        "peaton": PELIGROS_AMBIENTAL_PEATON,
    },
    "FACTOR SOCIAL": {
        "conductor": PELIGROS_SOCIAL_CONDUCTOR,
        "motociclista": PELIGROS_SOCIAL_MOTOCICLISTA,
        "peaton": PELIGROS_SOCIAL_PEATON,
    },
}

ORDEN_FACTORES = [
    "FACTOR HUMANO", "FACTOR MECÁNICO", "FACTOR INFRAESTRUCTURA VIAL",
    "FACTOR AMBIENTAL", "FACTOR SOCIAL",
]

# Etiquetas para columna ACTOR VIAL (col C) por tipo
ETIQUETA_ACTOR = {
    "conductor": "CONDUCTOR - VEHÍCULO INSTITUCIONAL",
    "motociclista": "MOTOCICLISTA - MENSAJERO\n(Cargo: {cargos})",
    "peaton": "PEATÓN OPERATIVO\n(Cargo: {cargos})",
    "pasajero": "PASAJERO - DESPLAZAMIENTOS LABORALES",
}


# ---------------------------------------------------------------------------
# Inferencia de actores viales por cargo
# ---------------------------------------------------------------------------

def _inferir_actores_viales(cargo: dict) -> list[str]:
    """Devuelve la lista de tipos de actor vial que aplican a un cargo.
    Un cargo puede tener varios actores (ej. un conductor que tambien camina
    al cliente -> conductor + peaton).
    """
    nombre = (cargo.get("nombre") or "").lower()
    tipo = (cargo.get("tipo") or "").lower()
    actividades = (cargo.get("actividades", "") or "").lower()
    tareas = (cargo.get("tareas", "") or "").lower()
    texto_manual = (cargo.get("texto_manual", "") or "").lower()
    texto = " ".join([nombre, tipo, actividades, tareas, texto_manual])

    actores = set()

    # Conductor de vehiculo institucional (excluye motos)
    es_motorizado = any(k in texto for k in [
        "motociclista", "moto ", "motocicleta", "mensajero motorizado",
        "domiciliario", "reparto motorizado", "mensajero",
    ])
    if not es_motorizado and any(k in texto for k in [
        "conductor", "chofer", "vehiculo", "vehículo", "camion", "camión",
        "conducir vehiculo", "conducir vehículo", "conducir camion", "conducir camión",
        "conduccion de", "conducción de", "manejo de vehiculo", "manejo de vehículo",
    ]):
        actores.add("conductor")

    if es_motorizado:
        actores.add("motociclista")

    # Peaton: cargos que hacen DILIGENCIAS EXTERNAS A PIE (no auxiliar generico)
    senales_peaton = [
        "diligencia", "diligencias", "tramite externo", "tramites externos",
        "trámite externo", "trámites externos", "diligencias externas",
        "trámites bancarios", "tramites bancarios", "mensajeria fisica",
        "mensajería física", "diligenciamiento en", "visita a banc",
    ]
    if not es_motorizado and any(s in texto for s in senales_peaton):
        actores.add("peaton")

    # Pasajero: cargos que se desplazan EN TRANSPORTE como pasajero
    senales_pasajero = [
        "visita comercial", "visitas comerciales", "asesor comercial",
        "asesor externo", "ventas externas", "trabajo de campo",
        "viajes de trabajo", "viajes regionales", "produccion en exteriores",
        "producción en exteriores", "rodaje en exteriores",
    ]
    if any(s in texto for s in senales_pasajero):
        actores.add("pasajero")

    # Si todavia no detecto nada PERO el cargo claramente se desplaza, asumir peaton
    if not actores:
        if any(k in texto for k in ["desplazamiento", "ruta vial", "via publica", "vía pública"]):
            actores.add("peaton")

    return sorted(actores)


# ---------------------------------------------------------------------------
# Estilos
# ---------------------------------------------------------------------------
_BORDE = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
_ALIGN_WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
_ALIGN_LEFT = Alignment(wrap_text=True, vertical="center", horizontal="left")
_FONT_NEGRILLA = Font(bold=True)

_FILL_FACTOR = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")  # rojo banda factor
_FILL_AZUL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
_FILL_VERDE = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
_FILL_DURAZNO = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
_FILL_CLASIF = _FILL_DURAZNO  # color para columna H CLASIFICACION
_FILLS_ACTORES = [_FILL_AZUL, _FILL_VERDE]

# Colores para interpretacion de nivel de riesgo (semaforo)
_FILL_RIESGO_CRITICO = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
_FILL_RIESGO_MODERADO = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
_FILL_RIESGO_BAJO = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")

# Fuente estandar de la matriz: Arial 9
FUENTE_NOMBRE = "Arial"
FUENTE_SIZE = 9


def _set_bold(cel, bold: bool):
    """Aplica bold preservando color, italic, etc., pero FORZANDO Arial 9."""
    f = cel.font
    cel.font = Font(
        name=FUENTE_NOMBRE, size=FUENTE_SIZE, color=f.color,
        italic=f.italic, underline=f.underline, strike=f.strike,
        vertAlign=f.vertAlign, bold=bold,
    )


def _calcular_nivel_riesgo(prob: str, expo: str) -> tuple[int, str, PatternFill]:
    """Pre-calcula el nivel de riesgo, interpretacion y color de fondo
    basado en la probabilidad y exposicion (igual que la formula de Excel)."""
    valor_prob = {"No es Probable": 1, "Poco Probable": 2, "Muy Probable": 3}.get(prob, 1)
    valor_expo = {"Esporádica": 1, "Esporadica": 1, "Ocasional": 2, "Frecuente": 3}.get(expo, 1)
    nivel = valor_prob * valor_expo
    if nivel >= 6:
        return nivel, "CRITICO", _FILL_RIESGO_CRITICO
    elif nivel >= 3:
        return nivel, "II MODERADO", _FILL_RIESGO_MODERADO
    elif nivel >= 1:
        return nivel, "III BAJO", _FILL_RIESGO_BAJO
    return nivel, "N/A", PatternFill(fill_type=None)


# ---------------------------------------------------------------------------
# Escritura
# ---------------------------------------------------------------------------

def _escribir_fila_peligro(ws, row: int, peligro: dict):
    """Cols G..Z + S=expuestos, T=peor."""
    ws.cell(row=row, column=7, value=peligro["peligro"]).alignment = _ALIGN_LEFT
    # H (clasificacion) - el factor; se escribe afuera
    ws.cell(row=row, column=9, value=peligro["efectos"]).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=10, value=peligro["fuente"]).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=11, value=peligro["medio"]).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=12, value=peligro["individuo"]).alignment = _ALIGN_LEFT
    # M (probabilidad texto) y O (exposicion texto) + formulas en N, P
    ws.cell(row=row, column=13, value=peligro["nivel_probabilidad"]).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=14,
            value=f'=IF(M{row}="No es Probable","1",IF(M{row}="Poco Probable","2","3"))').alignment = _ALIGN_WRAP
    ws.cell(row=row, column=15, value=peligro["nivel_exposicion"]).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=16,
            value=f'=IF(O{row}="Esporádica","1",IF(O{row}="Ocasional","2","3"))').alignment = _ALIGN_WRAP
    ws.cell(row=row, column=17, value=f"=N{row}*P{row}").alignment = _ALIGN_WRAP
    # Mantener la FORMULA en columna R (asi si el usuario cambia M u O, se recalcula)
    # y aplicar el color como fill explicito basado en los valores actuales de
    # probabilidad y exposicion (asi siempre se ve el color correcto al abrir).
    cel_r = ws.cell(row=row, column=18,
        value=f'=IF(Q{row}>=6,"CRITICO",IF(Q{row}>=3,"II MODERADO",IF(Q{row}>=1,"III BAJO","N/A")))')
    cel_r.alignment = _ALIGN_WRAP
    _, _, fill_riesgo = _calcular_nivel_riesgo(
        peligro["nivel_probabilidad"], peligro["nivel_exposicion"]
    )
    cel_r.fill = fill_riesgo
    ws.cell(row=row, column=19, value=peligro["no_expuestos"]).alignment = _ALIGN_WRAP
    ws.cell(row=row, column=20, value=peligro["peor"]).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=21, value=peligro["eliminacion"]).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=22, value=peligro["sustitucion"]).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=23, value=peligro["ingenieria"]).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=24, value=peligro["administrativos"]).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=25, value=peligro["epp"]).alignment = _ALIGN_LEFT
    ws.cell(row=row, column=26, value=peligro["plan_accion"]).alignment = _ALIGN_LEFT
    for c in range(7, 27):
        cel = ws.cell(row=row, column=c)
        cel.border = _BORDE
        # Forzar Arial 9, bold solo en col 8 (CLASIFICACION) y 18 (INTERPRETACION)
        es_negrilla = c in (8, 18)
        # Texto siempre negro (incluso sobre rojo)
        cel.font = Font(name=FUENTE_NOMBRE, size=FUENTE_SIZE, bold=es_negrilla,
                        color="000000")
    # Altura uniforme para filas de peligro
    if (ws.row_dimensions[row].height or 0) < 75:
        ws.row_dimensions[row].height = 75


def _escribir_fila_factor(ws, row: int, factor: str, max_col: int = 26):
    """Banda roja con el nombre del factor."""
    cel = ws.cell(row=row, column=1, value=factor)
    cel.alignment = Alignment(horizontal="center", vertical="center")
    cel.font = Font(name=FUENTE_NOMBRE, size=11, bold=True, color="FFFFFF")
    for c in range(1, max_col + 1):
        ws.cell(row=row, column=c).fill = _FILL_FACTOR
        ws.cell(row=row, column=c).border = _BORDE
    if max_col > 1:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
    # Forzar altura de la banda de factor
    ws.row_dimensions[row].height = 28


def _escribir_bloque_actor(ws, row_ini: int, row_fin: int, proceso: str,
                           tipo_desplaz: str, actor_label: str,
                           actividades: str, fill: PatternFill):
    """Escribe A..F y aplica merges del bloque de actor."""
    ws.cell(row=row_ini, column=1, value=proceso).alignment = _ALIGN_WRAP
    ws.cell(row=row_ini, column=2, value=tipo_desplaz).alignment = _ALIGN_WRAP
    ws.cell(row=row_ini, column=3, value=actor_label).alignment = _ALIGN_WRAP
    ws.cell(row=row_ini, column=4, value=actividades).alignment = _ALIGN_LEFT
    ws.cell(row=row_ini, column=5, value="X").alignment = _ALIGN_WRAP
    # Bordes, color de fondo y fuente Arial 9 negro en TODAS las celdas A..F
    # Negrilla solo en PROCESO (col 1) y ACTORES (col 3) en row_ini
    for r in range(row_ini, row_fin + 1):
        for c in range(1, 7):
            cel = ws.cell(row=r, column=c)
            cel.border = _BORDE
            cel.fill = fill
            es_negrilla = (r == row_ini and c in (1, 3))
            cel.font = Font(name=FUENTE_NOMBRE, size=FUENTE_SIZE,
                            bold=es_negrilla, color="000000")
    if row_fin > row_ini:
        for col_letter in ["A", "B", "C", "D", "E", "F"]:
            ws.merge_cells(f"{col_letter}{row_ini}:{col_letter}{row_fin}")


def generar_matriz_peligros_viales(
    cliente: dict,
    output_dir: Path,
    config_dir: Path | None = None,
) -> Path | None:
    """Genera la PESV-MTZ-002 a partir de los cargos del cliente.

    Solo se ejecuta si el cliente tiene generar_pesv: true.
    """
    if not cliente.get("generar_pesv", False):
        logger.info("PESV no aplica para este cliente; matriz vial no se llena.")
        return None

    matriz_path = output_dir / RUTA_MATRIZ_VIAL
    if not matriz_path.exists():
        logger.warning("No se encontro la matriz PESV-MTZ-002: %s", matriz_path)
        return None

    # Recolectar cargos (mismo enfoque que IPVR)
    cargos: list[dict] = []
    dir_manuales = cliente.get("manuales_cargos_dir")
    if dir_manuales:
        from .matriz_peligros import cargar_cargos_desde_manuales
        ruta = Path(dir_manuales)
        if not ruta.is_absolute() and config_dir:
            ruta = (Path(config_dir) / ruta).resolve()
        cargos = cargar_cargos_desde_manuales(ruta)
    if not cargos:
        cargos = cliente.get("cargos") or []

    if not cargos:
        logger.info("Sin cargos definidos; matriz vial no se llena.")
        return None

    # Agrupar cargos por tipo de actor vial
    # actor -> list of {cargo, proceso, actividades}
    actores_a_cargos: dict[str, list[dict]] = {}
    for cargo in cargos:
        actores = _inferir_actores_viales(cargo)
        if not actores:
            continue
        proceso = (cargo.get("proceso") or "ADMINISTRATIVO").upper()
        for actor in actores:
            actores_a_cargos.setdefault(actor, []).append({
                "cargo": cargo,
                "proceso": proceso,
                "actividades": cargo.get("actividades") or f"Desplazamientos del cargo {cargo.get('nombre', '')}.",
            })

    if not actores_a_cargos:
        logger.info("Ningun cargo aplica como actor vial; matriz vial no se llena.")
        return None

    wb = load_workbook(str(matriz_path))
    if NOMBRE_HOJA not in wb.sheetnames:
        logger.warning("Hoja '%s' no encontrada", NOMBRE_HOJA)
        return None
    ws = wb[NOMBRE_HOJA]

    # Limpiar conditional formatting de la plantilla (la zona R8:R41) para que
    # nuestros fills explicitos del semaforo no se vean sobrescritos.
    try:
        ws.conditional_formatting._cf_rules.clear()
    except Exception:
        pass

    # Limpiar merges previos en zona de datos
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= FILA_INICIO:
            ws.unmerge_cells(str(mr))

    # Limpiar contenido en zona de datos
    for row in range(FILA_INICIO, ws.max_row + 1):
        for col in range(1, 27):
            ws.cell(row=row, column=col).value = None
            ws.cell(row=row, column=col).fill = PatternFill(fill_type=None)

    fila_actual = FILA_INICIO
    indice_actor = 0

    for factor in ORDEN_FACTORES:
        catalogo_factor = CATALOGO.get(factor, {})
        if not catalogo_factor:
            continue
        # Verificar si hay actores aplicables
        actores_aplicables = [a for a in catalogo_factor if a in actores_a_cargos]
        if not actores_aplicables:
            continue

        # Banda del factor
        _escribir_fila_factor(ws, fila_actual, factor)
        fila_actual += 1

        for actor in actores_aplicables:
            peligros = catalogo_factor[actor]
            cargos_actor = actores_a_cargos[actor]
            nombres_cargos = ", ".join(c["cargo"].get("nombre", "") for c in cargos_actor[:3])
            proceso = cargos_actor[0]["proceso"]
            actividades = cargos_actor[0]["actividades"]
            actor_label = ETIQUETA_ACTOR.get(actor, actor.upper())
            actor_label = actor_label.replace("{cargos}", nombres_cargos)
            if "{cargos}" not in ETIQUETA_ACTOR.get(actor, ""):
                actor_label = f"{actor_label}\n(Cargo: {nombres_cargos})"

            row_ini = fila_actual
            for pel in peligros:
                _escribir_fila_peligro(ws, fila_actual, pel)
                # Clasificacion en H = factor (con fondo durazno claro)
                cel_h = ws.cell(row=fila_actual, column=8, value=factor)
                cel_h.alignment = _ALIGN_WRAP
                cel_h.font = Font(name=FUENTE_NOMBRE, size=FUENTE_SIZE, bold=True, color="000000")
                cel_h.border = _BORDE
                cel_h.fill = _FILL_CLASIF
                fila_actual += 1
            row_fin = fila_actual - 1

            fill = _FILLS_ACTORES[indice_actor % len(_FILLS_ACTORES)]
            indice_actor += 1
            _escribir_bloque_actor(ws, row_ini, row_fin, proceso,
                                   "EN MISIÓN / DESPLAZAMIENTO LABORAL",
                                   actor_label, actividades, fill)
            logger.info("PESV vial: %s / actor=%s -> %d filas (R%d-R%d)",
                        factor, actor, row_fin - row_ini + 1, row_ini, row_fin)

    wb.save(str(matriz_path))
    logger.info("Matriz PESV-MTZ-002 generada: %s", matriz_path)
    return matriz_path
