#!/usr/bin/env python3
"""
Adapta el contenido de los documentos en la carpeta de plantilla
'1.1.6 DESIGNACIÓN DEL VIGÍA SST' para que sean correctos para Vigía.

Los archivos fueron creados copiando la carpeta COPASST y renombrando
los archivos, pero el contenido interno sigue siendo de COPASST.

Ejecución única sobre la carpeta de plantillas:
    python adaptar_vigia_plantilla.py
"""
import re
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn

VIGIA_DIR = Path(
    r"Plantilla SG-SST 2026\I. PLANEAR\1.1.6 DESIGNACIÓN DEL VIGÍA SST"
)

# ─── Reemplazos de texto aplicados a TODOS los documentos ───────────────────

# Orden importante: los más específicos van primero para evitar dobles reemplazos
REEMPLAZOS = [
    # Número de trabajadores
    ("más de 10 trabajadores", "menos de 10 trabajadores"),
    ("Más de 10 trabajadores", "Menos de 10 trabajadores"),
    # Nombre completo + sigla
    ("Comité Paritario de la Seguridad y Salud en el Trabajo - COPASST", "Vigía de Seguridad y Salud en el Trabajo"),
    ("Comité Paritario de Seguridad y Salud en el Trabajo - COPASST",    "Vigía de Seguridad y Salud en el Trabajo"),
    ("Comité Paritario de la Seguridad y Salud en el Trabajo (COPASST)", "Vigía de Seguridad y Salud en el Trabajo"),
    ("Comité Paritario de Seguridad y Salud en el Trabajo (COPASST)",    "Vigía de Seguridad y Salud en el Trabajo"),
    ("Comité Paritario de la Seguridad y Salud en el Trabajo",           "Vigía de Seguridad y Salud en el Trabajo"),
    ("Comité Paritario de Seguridad y Salud en el Trabajo",              "Vigía de Seguridad y Salud en el Trabajo"),
    ("COMITÉ PARITARIO DE LA SEGURIDAD Y SALUD EN EL TRABAJO",          "VIGÍA DE SEGURIDAD Y SALUD EN EL TRABAJO"),
    ("COMITÉ PARITARIO DE SEGURIDAD Y SALUD EN EL TRABAJO",             "VIGÍA DE SEGURIDAD Y SALUD EN EL TRABAJO"),
    ("Comité Paritario", "Vigía"),
    ("COMITÉ PARITARIO", "VIGÍA"),
    # Sigla sola
    ("COPASST", "VIGÍA DE SST"),
    # Conformación → Designación
    ("conformación y funcionamiento del Comité", "designación y funcionamiento del Vigía de SST"),
    ("conformación y funcionamiento del Vigía",  "designación y funcionamiento del Vigía de SST"),
    ("conformación del Comité", "designación del Vigía de SST"),
    ("conformación del Vigía",  "designación del Vigía de SST"),
    ("CONFORMACIÓN DEL COMITÉ", "DESIGNACIÓN DEL VIGÍA DE SST"),
    # Integrantes / miembros
    ("integrantes del Comité",    "Vigía de SST"),
    ("miembros del Comité",       "Vigía de SST"),
    ("integrantes del VIGÍA DE SST", "Vigía de SST"),
    ("miembros del VIGÍA DE SST",    "Vigía de SST"),
    # Período (2 años → 1 año)
    ("período de dos (2) años", "período de un (1) año"),
    ("periodo de dos (2) años",  "periodo de un (1) año"),
    ("dos (2) años a partir",    "un (1) año a partir"),
    # Resolución 2013 de 1986 (COPASST) → norma del Vigía
    (
        "Resolución 2013 (06 de 006 de 1986), Ministerio de Trabajo y Protección social,"
        " «Por la cual se reglamenta la organización y funcionamiento de los Comités Paritarios"
        " de Seguridad y Salud en el Trabajo».",
        "Decreto 1072 de 2015, artículo 2.2.4.6.8. Ministerio de Trabajo."
        " Designación del Vigía de SST en empresas con menos de diez (10) trabajadores.",
    ),
    ("Resolución 2013 de 1986", "Decreto 1072 de 2015, artículo 2.2.4.6.8"),
    # Roles propios del COPASST
    ("Presidente del VIGÍA DE SST", "Vigía de SST"),
    ("Secretaria del VIGÍA DE SST", "Representante Legal"),
    ("Presidente del Comité",       "Vigía de SST"),
    ("Secretaria del Comité",       "Representante Legal"),
    ("presidente del comité",       "Vigía de SST"),
    # Informe trimestral
    ("INFORME TRIMESTRAL PARA LA GESTIÓN DEL VIGÍA DE SST",   "INFORME TRIMESTRAL DEL VIGÍA DE SST"),
    ("INFORME TRIMESTRAL PARA LA GESTIÓN DEL COPASST",        "INFORME TRIMESTRAL DEL VIGÍA DE SST"),
    ("gestión adelantada por el Vigía de Seguridad y Salud en el Trabajo (VIGÍA DE SST)",
     "gestión adelantada por el Vigía de Seguridad y Salud en el Trabajo"),
    ("gestión adelantada por el Vigía de SST",
     "gestión adelantada por el Vigía de Seguridad y Salud en el Trabajo"),
    ("acciones ejecutadas por el VIGÍA DE SST",    "acciones ejecutadas por el Vigía de SST"),
    ("acciones ejecutadas por el Comité",          "acciones ejecutadas por el Vigía de SST"),
    ("Reuniones del VIGÍA DE SST",                 "Reuniones del Vigía de SST"),
    ("Reuniones del Comité",                       "Reuniones del Vigía de SST"),
    ("reuniones mensuales del mismo",              "reuniones mensuales con el empleador"),
    # Acta de reunión – título
    (
        "ACTA N° ________ VIGÍA DE SEGURIDAD Y SALUD EN EL TRABAJO",
        "ACTA N° ________ VIGÍA DE SEGURIDAD Y SALUD EN EL TRABAJO",
    ),
    (
        "ACTA N° ________ COMITÉ PARITARIO DE SEGURIDAD Y SALUD EN EL TRABAJO",
        "ACTA N° ________ VIGÍA DE SEGURIDAD Y SALUD EN EL TRABAJO",
    ),
    # Asistentes acta reunión
    ("Se verifica asistencia de todos los miembros del Comité",
     "Se verifica la asistencia del Vigía de SST y del Representante Legal"),
    ("Se verifica asistencia de todos los miembros del VIGÍA DE SST",
     "Se verifica la asistencia del Vigía de SST y del Representante Legal"),
    # Firmas acta reunión – suplentes (se eliminan dejando cadena vacía)
    ("Suplente del presidente",   ""),
    ("Suplente de la secretaria", ""),
    # Documentos relacionados en PRO-001
    ("ANEXO SST-FTM-010 CONVOCATORIA AL VIGÍA DE SST.",  "ANEXO SST-FTM-013 ACTA DE DESIGNACIÓN DEL VIGÍA DE SST."),
    ("ANEXO SST-FTM-011 INSCRIPCIÓN AL VIGÍA DE SST.",   "ANEXO SST-FMT-015 ACTA DE REUNIÓN VIGÍA DE SST."),
    ("ANEXO SST-FTM-012 VOTACIÓN DEL VIGÍA DE SST.",     "ANEXO SST-FMT-016 INFORME TRIMESTRAL VIGÍA DE SST."),
    ("ANEXO SST-FTM-013 ACTA DE RESULTADOS.",            "ANEXO SST-FTM-014 ACTA DE CONFORMACIÓN Y ACEPTACIÓN DEL VIGÍA DE SST."),
    ("ANEXO SST-FTM-014 ACTA DE CONFORMACIÓN.",          ""),
    # Indicadores en informe trimestral
    ("Reuniones del VIGÍA DE SST | Nivel de cumplimiento de las reuniones del VIGÍA DE SST",
     "Reuniones del Vigía de SST | Nivel de cumplimiento de las reuniones del Vigía de SST"),
]


# ─── Reemplazos específicos por documento ───────────────────────────────────

# PRO-001: Objetivo, alcance, generalidades y actividades adaptadas para Vigía
PRO001_REEMPLAZOS = [
    (
        "Dar a conocer los lineamientos para la conformación y funcionamiento del "
        "Vigía de Seguridad y Salud en el Trabajo, como organismo de promoción y "
        "vigilancia de las normas de Seguridad y Salud en el Trabajo.",
        "Establecer el procedimiento para la designación del Vigía de Seguridad y "
        "Salud en el Trabajo en {{RAZON_SOCIAL}}, de conformidad con el artículo "
        "2.2.4.6.8 del Decreto 1072 de 2015, aplicable a empresas con menos de "
        "diez (10) trabajadores.",
    ),
    (
        "Aplica a todos los trabajadores de {{RAZON_SOCIAL}}., desde el momento "
        "en que inicia la convocatoria para pertenecer al comité, hasta las "
        "reuniones mensuales del mismo.",
        "Aplica al proceso de designación y funcionamiento del Vigía de Seguridad "
        "y Salud en el Trabajo de {{RAZON_SOCIAL}}, empresa con menos de diez "
        "(10) trabajadores permanentes.",
    ),
    (
        "Vigía de Seguridad y Salud en el Trabajo: Organismo de promoción y "
        "vigilancia del funcionamiento de la Seguridad y Salud en el Trabajo.",
        "Vigía de Seguridad y Salud en el Trabajo: Persona designada directamente "
        "por el empleador en empresas con menos de diez (10) trabajadores, para "
        "promover y vigilar el funcionamiento del Sistema de Gestión de la "
        "Seguridad y Salud en el Trabajo (SG-SST). Ejerce funciones similares a "
        "las del COPASST pero de manera individual.",
    ),
    (
        "Velar por el funcionamiento del Vigía de Seguridad y Salud en el Trabajo"
        " - VIGÍA DE SST y por la renovación de sus integrantes cada vez que se "
        "requiera de acuerdo a la normatividad legal vigente.",
        "Apoyar al Vigía de SST en el cumplimiento de sus funciones, garantizar "
        "los recursos necesarios y velar por la renovación del cargo cada año.",
    ),
    (
        "Elegir a sus representantes mediante el voto secreto.",
        "Participar activamente en las actividades de SST promovidas por el Vigía.",
    ),
    (
        "{{RAZON_SOCIAL}}., por ser una empresa con menos de 10 trabajadores está "
        "en la obligación de conformar el Vigía de Seguridad y Salud en el Trabajo"
        " - VIGÍA DE SST según lo establecido en la Resolución 2013 de 1986.",
        "{{RAZON_SOCIAL}}, por tener menos de diez (10) trabajadores permanentes, "
        "está en la obligación de designar un Vigía de Seguridad y Salud en el "
        "Trabajo según lo establecido en el artículo 2.2.4.6.8 del Decreto 1072 "
        "de 2015 y la Resolución 0312 de 2019. La designación es realizada "
        "directamente por el empleador, sin proceso de elección. El Vigía ejerce "
        "sus funciones durante un (1) año y puede ser redesignado.",
    ),
]

# Actividades en la tabla del PRO-001 (por número de fila / descripción clave)
PRO001_ACTIVIDADES = {
    # clave: texto que identifica la fila → (responsable nuevo, unidad nueva, descripción nueva)
    "Identificación": (
        "Encargado SG-SST",
        "Identificación",
        "{{RAZON_SOCIAL}}, por tener menos de diez (10) trabajadores, debe "
        "designar un Vigía de SST. El Encargado del SG-SST identifica al "
        "trabajador más idóneo para el cargo, considerando su disponibilidad, "
        "conocimiento del proceso productivo y compromiso con la SST.",
    ),
    "Selección de los representantes por parte del empl": (
        "Gerente general",
        "Designación formal",
        "El Representante Legal designa formalmente al Vigía de SST mediante "
        "acto administrativo (Acta de Designación SST-FTM-013). La designación "
        "es directa; no requiere convocatoria ni elección por parte de los "
        "trabajadores.",
    ),
    "Convocatoria e inscripción de los candidatos por p": (
        "Encargado SG-SST",
        "Comunicación",
        "Se comunica formalmente al trabajador designado su nuevo rol como "
        "Vigía de SST, indicando el período de gestión (un año), sus funciones "
        "y responsabilidades según el Decreto 1072 de 2015.",
    ),
    "La documentación y papelería correspondiente a las": (
        "Vigía de SST designado",
        "Aceptación",
        "El trabajador designado firma el Acta de Conformación y Aceptación "
        "(SST-FTM-014), manifestando su aceptación del cargo y su compromiso "
        "de cumplir las funciones establecidas en la normativa.",
    ),
    "Conforme a los resultados registrados en el acta d": (
        "Encargado SG-SST / Vigía de SST",
        "Registro y capacitación",
        "Se registra la designación en el SG-SST y se programa la capacitación "
        "inicial del Vigía en SST. El registro debe quedar disponible para "
        "inspección de la ARL y el Ministerio de Trabajo.",
    ),
    "Según el artículo sexto. de la Resolución 2013 de ": (
        "Vigía de SST / Encargado SG-SST",
        "Reuniones y seguimiento",
        "El Vigía de SST se reúne mensualmente con el empleador para revisar "
        "el avance del SG-SST, las condiciones de trabajo, los accidentes "
        "e incidentes y las recomendaciones de mejora. Al finalizar el año "
        "se evalúa el desempeño del Vigía y se decide la redesignación.",
    ),
}

# FTM-013: Acta de designación (reemplaza el formulario de resultados de votación)
FTM013_PARRAFOS = {
    "se dio por finalizado el proceso de votación": (
        "En cumplimiento de lo establecido en el artículo 2.2.4.6.8 del Decreto "
        "1072 de 2015, el señor(a) {{REPRESENTANTE_LEGAL}}, en calidad de "
        "{{CARGO_RL}} de {{RAZON_SOCIAL}}, con NIT {{NIT}}, empresa con menos de "
        "diez (10) trabajadores, procede a DESIGNAR como Vigía de Seguridad y "
        "Salud en el Trabajo a la siguiente persona:"
    ),
    "Como jurado de mesa estuvo": (
        "Nombre del Vigía designado: ___________________________________________\n"
        "Número de identificación:   ___________________________________________\n"
        "Cargo que desempeña:        ___________________________________________"
    ),
    "RESULTADOS OBTENIDOS DEL CONTEO DE VOTOS": (
        "PERÍODO DE GESTIÓN"
    ),
    "Resultaron elegidas en esta votación las siguientes personas:": (
        "El Vigía de SST ejercerá sus funciones durante el período comprendido "
        "entre el _____ de _____________ de 202_ y el _____ de _____________ de "
        "202_, equivalente a un (1) año de gestión, renovable según evaluación "
        "de desempeño."
    ),
    "Categoría PRINCIPAL:": (
        "ACEPTACIÓN DEL CARGO\n\n"
        "Yo, _______________________________________, identificado(a) con cédula "
        "de ciudadanía N°. _________________, manifiesto expresamente mi "
        "aceptación para ejercer el cargo de Vigía de Seguridad y Salud en el "
        "Trabajo de {{RAZON_SOCIAL}}, comprometiéndome a cumplir las funciones y "
        "responsabilidades establecidas en el artículo 2.2.4.6.20 del Decreto "
        "1072 de 2015."
    ),
    "Categoría SUPLENTE:": "",
}

# FTM-014: Acta de conformación → acta de designación con funciones del Vigía
FTM014_PARRAFOS = {
    "se conformó el Vigía de Seguridad y Salud en el Trabajo": (
        "En cumplimiento del artículo 2.2.4.6.8 del Decreto 1072 de 2015 y la "
        "Resolución 0312 de 2019, en las instalaciones de {{RAZON_SOCIAL}}, el "
        "señor(a) {{REPRESENTANTE_LEGAL}}, en calidad de {{CARGO_RL}}, procede a "
        "formalizar la designación del Vigía de Seguridad y Salud en el Trabajo:"
    ),
    "El período de los miembros es de un (1) año": (
        "El período de gestión del Vigía de SST es de un (1) año, contado a "
        "partir del ___ de _____________ de 202_. Al término del período, el "
        "empleador podrá redesignar al mismo trabajador o designar a otro."
    ),
    "El representante legal de la empresa nombró a:": (
        "DATOS DEL VIGÍA DESIGNADO\n\n"
        "Nombre completo:       _______________________________________________\n"
        "Número de cédula:      _______________________________________________\n"
        "Cargo en la empresa:   _______________________________________________\n"
        "Fecha de designación:  _______________________________________________"
    ),
    "Los elegidos de los trabajadores son:": (
        "FUNCIONES DEL VIGÍA DE SST (Art. 2.2.4.6.20 Decreto 1072/2015)\n\n"
        "1. Proponer a la administración la adopción de medidas de SST y su "
        "verificación.\n"
        "2. Proponer y participar en actividades de capacitación en SST.\n"
        "3. Vigilar el cumplimiento de las actividades del SG-SST.\n"
        "4. Colaborar en el análisis de las causas de accidentes y enfermedades "
        "laborales.\n"
        "5. Visitar periódicamente los puestos de trabajo e informar al empleador "
        "sobre las condiciones de riesgo.\n"
        "6. Participar en la elaboración e implementación de los programas de SST.\n"
        "7. Servir como organismo de coordinación entre el empleador y los "
        "trabajadores en la solución de los problemas de SST."
    ),
    "Conformado el Comité Paritario de Seguridad y Salud en el Trabajo": (
        "En constancia de lo anterior, el Vigía de SST acepta el cargo y se "
        "compromete a ejercer sus funciones durante el período señalado, "
        "conforme a la normativa vigente."
    ),
    "Presidente del VIGÍA DE SST": "Vigía de SST",
    "Secretaria del VIGÍA DE SST": "Representante Legal",
    "Suplente del presidente": "",
    "Suplente de la secretaria": "",
    "presidenta del comité y como secretario se nom": "",
}

# FMT-016: Informe trimestral – párrafos de introducción y objetivo
FMT016_PARRAFOS = {
    "gestión adelantada por el Vigía de Seguridad y Salud en el Trabajo (VIGÍA DE SST)": (
        "gestión adelantada por el Vigía de Seguridad y Salud en el Trabajo de "
        "{{RAZON_SOCIAL}} durante el período comprendido entre el "
        "_________________ y el _________________"
    ),
    "aportar a la mejora continua del Sistema de Gestión de la Seguridad y Salud "
    "en el Trabajo": (
        "apoyar al empleador en la mejora continua del Sistema de Gestión de la "
        "Seguridad y Salud en el Trabajo (SG-SST)"
    ),
    "Presentar los resultados de las acciones ejecutadas por el VIGÍA DE SST durante el trimestre": (
        "Presentar los resultados de las acciones ejecutadas por el Vigía de SST "
        "durante el trimestre"
    ),
}


# ─── Funciones de utilidad ───────────────────────────────────────────────────

def _reemplazar_en_texto(texto: str, reemplazos: list[tuple]) -> str:
    """Aplica una lista de (viejo, nuevo) al texto sin usar regex."""
    for viejo, nuevo in reemplazos:
        texto = texto.replace(viejo, nuevo)
    return texto


def _procesar_parrafos(paragraphs, reemplazos: list[tuple]) -> int:
    """Reemplaza texto run a run conservando el formato. Retorna cambios."""
    cambios = 0
    for p in paragraphs:
        # Reconstruir el texto completo del párrafo
        texto_completo = "".join(r.text for r in p.runs)
        texto_nuevo = _reemplazar_en_texto(texto_completo, reemplazos)
        if texto_nuevo != texto_completo and p.runs:
            # Poner todo el texto nuevo en el primer run y vaciar el resto
            p.runs[0].text = texto_nuevo
            for r in p.runs[1:]:
                r.text = ""
            cambios += 1
    return cambios


def _procesar_documento(doc: Document, reemplazos: list[tuple]) -> int:
    """Aplica reemplazos a párrafos, tablas, encabezados y pies. Retorna cambios."""
    cambios = 0
    cambios += _procesar_parrafos(doc.paragraphs, reemplazos)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                cambios += _procesar_parrafos(celda.paragraphs, reemplazos)
    for section in doc.sections:
        for elem in [section.header, section.footer]:
            cambios += _procesar_parrafos(elem.paragraphs, reemplazos)
            for t in elem.tables:
                for f in t.rows:
                    for c in f.cells:
                        cambios += _procesar_parrafos(c.paragraphs, reemplazos)
    return cambios


def _reemplazar_parrafo_si_contiene(doc: Document, fragmento: str, nuevo_texto: str) -> bool:
    """Reemplaza el texto completo del párrafo que contenga `fragmento`."""
    for p in doc.paragraphs:
        texto = "".join(r.text for r in p.runs)
        if fragmento in texto and p.runs:
            p.runs[0].text = nuevo_texto
            for r in p.runs[1:]:
                r.text = ""
            return True
    # Buscar también en tablas
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    texto = "".join(r.text for r in p.runs)
                    if fragmento in texto and p.runs:
                        p.runs[0].text = nuevo_texto
                        for r in p.runs[1:]:
                            r.text = ""
                        return True
    return False


def _reemplazar_celda_si_contiene(doc: Document, fragmento: str, nuevo_texto: str) -> bool:
    """Reemplaza el texto de la celda que contenga `fragmento`."""
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                texto = celda.text
                if fragmento in texto:
                    for p in celda.paragraphs:
                        if p.runs:
                            p.runs[0].text = nuevo_texto
                            for r in p.runs[1:]:
                                r.text = ""
                            nuevo_texto = ""  # Los demás párrafos de la celda se vacían
                    return True
    return False


# ─── Adaptaciones específicas por documento ──────────────────────────────────

def _adaptar_pro001(doc: Document):
    """Adapta el Procedimiento de Designación del Vigía."""
    # Párrafos generales
    for viejo, nuevo in PRO001_REEMPLAZOS:
        _reemplazar_parrafo_si_contiene(doc, viejo[:50], nuevo)
        _reemplazar_parrafo_si_contiene(doc, viejo[:80], nuevo)

    # Tabla de actividades: identificar celdas por su contenido
    for tabla in doc.tables:
        for fila in tabla.rows:
            celdas = fila.cells
            if len(celdas) < 4:
                continue
            desc_col3 = celdas[2].text.strip()[:60]  # columna "UNIDAD FUNCIONAL"
            desc_col4 = celdas[3].text.strip()[:60]  # columna "DESCRIPCIÓN"
            for clave, (resp, unidad, desc) in PRO001_ACTIVIDADES.items():
                if clave[:40] in desc_col3 or clave[:40] in desc_col4:
                    # Reemplazar responsable
                    for p in celdas[1].paragraphs:
                        if p.runs:
                            p.runs[0].text = resp
                            for r in p.runs[1:]:
                                r.text = ""
                    # Reemplazar unidad funcional
                    for p in celdas[2].paragraphs:
                        if p.runs:
                            p.runs[0].text = unidad
                            for r in p.runs[1:]:
                                r.text = ""
                    # Reemplazar descripción
                    for p in celdas[3].paragraphs:
                        if p.runs:
                            p.runs[0].text = desc
                            for r in p.runs[1:]:
                                r.text = ""
                    break


def _adaptar_ftm013(doc: Document):
    """Adapta el Acta de Designación (era Acta de Resultados de Votación)."""
    for fragmento, nuevo in FTM013_PARRAFOS.items():
        _reemplazar_parrafo_si_contiene(doc, fragmento[:50], nuevo)

    # Vaciar la tabla de votos (NOMBRE DEL CANDIDATO | N° DE VOTOS)
    for tabla in doc.tables:
        encabezados = [c.text.strip() for c in tabla.rows[0].cells] if tabla.rows else []
        if "NOMBRE DEL CANDIDATO" in encabezados or "N° DE VOTOS" in encabezados:
            # Convertir la tabla en un campo de firma
            for fila in tabla.rows:
                for celda in fila.cells:
                    for p in celda.paragraphs:
                        for r in p.runs:
                            r.text = ""


def _adaptar_ftm014(doc: Document):
    """Adapta el Acta de Conformación del Vigía (era Acta COPASST con presidente/secretario)."""
    for fragmento, nuevo in FTM014_PARRAFOS.items():
        _reemplazar_parrafo_si_contiene(doc, fragmento[:50], nuevo)


def _adaptar_fmt016(doc: Document):
    """Adapta el Informe Trimestral del Vigía."""
    for fragmento, nuevo in FMT016_PARRAFOS.items():
        _reemplazar_parrafo_si_contiene(doc, fragmento[:50], nuevo)

    # Tabla de indicadores: cambiar "COPASST" / "Comité" por "Vigía de SST"
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                _procesar_parrafos(celda.paragraphs, REEMPLAZOS)


# ─── Main ─────────────────────────────────────────────────────────────────────

ADAPTACIONES_ESPECIFICAS = {
    "SST-PRO-001": _adaptar_pro001,
    "SST-FTM-013": _adaptar_ftm013,
    "SST-FTM-014": _adaptar_ftm014,
    "SST-FMT-016": _adaptar_fmt016,
}


def main():
    if not VIGIA_DIR.exists():
        print(f"ERROR: No se encontró la carpeta: {VIGIA_DIR}")
        return

    docx_files = [f for f in VIGIA_DIR.rglob("*.docx") if not f.name.startswith("~$")]
    print(f"Encontrados {len(docx_files)} documentos en la carpeta Vigía SST\n")

    for archivo in sorted(docx_files):
        print(f"  Procesando: {archivo.name}")
        try:
            doc = Document(str(archivo))

            # 1. Aplicar reemplazos básicos a todo el documento
            cambios = _procesar_documento(doc, REEMPLAZOS)

            # 2. Aplicar adaptaciones específicas según el código del documento
            codigo = archivo.stem[:11]  # ej. "SST-PRO-001"
            for clave, fn in ADAPTACIONES_ESPECIFICAS.items():
                if archivo.name.startswith(clave):
                    fn(doc)
                    break

            doc.save(str(archivo))
            print(f"    {cambios} reemplazos básicos + adaptación específica aplicada")

        except Exception as e:
            print(f"    ERROR: {e}")

    print("\nListo. Verifica los documentos en Word antes de usar la plantilla.")


if __name__ == "__main__":
    main()
