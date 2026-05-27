#!/usr/bin/env python3
"""
Adapta el contenido de los archivos [VIGÍA] recién creados en la plantilla.
Ejecución única:  python adaptar_vigia_dual.py
"""
from pathlib import Path
from docx import Document

PLANTILLA = Path(r"Plantilla SG-SST 2026")

# ─── Utilidad ────────────────────────────────────────────────────────────────

def reemplazar(doc: Document, viejo: str, nuevo: str) -> int:
    cambios = 0
    todos = list(doc.paragraphs)
    for t in doc.tables:
        for f in t.rows:
            for c in f.cells:
                todos += list(c.paragraphs)
    for section in doc.sections:
        for elem in [section.header, section.footer]:
            todos += list(elem.paragraphs)
    for p in todos:
        texto = "".join(r.text for r in p.runs)
        if viejo in texto and p.runs:
            p.runs[0].text = texto.replace(viejo, nuevo)
            for r in p.runs[1:]:
                r.text = ""
            cambios += 1
    return cambios


def reemplazar_parrafo(doc: Document, fragmento: str, nuevo: str) -> bool:
    todos = list(doc.paragraphs)
    for t in doc.tables:
        for f in t.rows:
            for c in f.cells:
                todos += list(c.paragraphs)
    for p in todos:
        texto = "".join(r.text for r in p.runs)
        if fragmento in texto and p.runs:
            p.runs[0].text = nuevo
            for r in p.runs[1:]:
                r.text = ""
            return True
    return False


REEMPLAZOS_BASE = [
    ("Comité Paritario de Seguridad y Salud en el Trabajo (COPASST)", "Vigía de Seguridad y Salud en el Trabajo"),
    ("Comité Paritario de Seguridad y Salud en el Trabajo", "Vigía de Seguridad y Salud en el Trabajo"),
    ("Comité Paritario de la Seguridad y Salud en el Trabajo (COPASST)", "Vigía de Seguridad y Salud en el Trabajo"),
    ("Comité Paritario de la Seguridad y Salud en el Trabajo", "Vigía de Seguridad y Salud en el Trabajo"),
    ("COMITÉ PARITARIO DE SEGURIDAD Y SALUD EN EL TRABAJO", "VIGÍA DE SEGURIDAD Y SALUD EN EL TRABAJO"),
    ("Comité Paritario", "Vigía"),
    ("COPASST", "VIGÍA DE SST"),
    ("del Comité", "del Vigía de SST"),
    ("el Comité", "el Vigía de SST"),
    ("al Comité", "al Vigía de SST"),
    ("Presidente del VIGÍA DE SST", "Vigía de SST"),
    ("presidente del VIGÍA DE SST", "Vigía de SST"),
]


# ─── FMT-005: Responsabilidades del Vigía ────────────────────────────────────

def adaptar_fmt005(ruta: Path):
    doc = Document(str(ruta))

    for viejo, nuevo in REEMPLAZOS_BASE:
        reemplazar(doc, viejo, nuevo)

    # Ajustes contextuales específicos
    reemplazar(doc,
        "se detallan las responsabilidades asignadas al Vigía de Seguridad y Salud en el Trabajo",
        "se detallan las responsabilidades asignadas al Vigía de Seguridad y Salud en el Trabajo "
        "conforme al artículo 2.2.4.6.20 del Decreto 1072 de 2015",
    )
    reemplazar(doc,
        "Responsabilidades del Vigía de Seguridad y Salud en el Trabajo:",
        "Responsabilidades del Vigía de Seguridad y Salud en el Trabajo "
        "(Art. 2.2.4.6.20 Decreto 1072 de 2015):",
    )
    # Responsabilidad de conformación del comité → designación
    reemplazar(doc,
        "Planificación de Auditorías: Participar activamente en la planificación de auditorías de seguridad y salud en el trabajo, contribuyendo a la evaluación y mejora de las condiciones laborales.",
        "Planificación de Auditorías: Participar activamente en la planificación de auditorías de seguridad y salud en el trabajo junto con el empleador, contribuyendo a la evaluación y mejora de las condiciones laborales.",
    )
    reemplazar(doc,
        "Coordinación en la Solución de Problemas SST: Servir como un órgano de coordinación entre el empleador y los colaboradores",
        "Coordinación en la Solución de Problemas SST: Servir como enlace de coordinación entre el empleador y los colaboradores",
    )
    # Firma final
    reemplazar(doc,
        "en mi calidad de _____________________,",
        "en mi calidad de Vigía de Seguridad y Salud en el Trabajo de {{RAZON_SOCIAL}},",
    )

    doc.save(str(ruta))
    print(f"  FMT-005 Vigía adaptado: {ruta.name}")


# ─── FMT-081: Informe de auditoría del Vigía ─────────────────────────────────

def adaptar_fmt081(ruta: Path):
    doc = Document(str(ruta))

    for viejo, nuevo in REEMPLAZOS_BASE:
        reemplazar(doc, viejo, nuevo)

    # Título
    reemplazar(doc,
        "INFORME ANUAL DE GESTIÓN FRENTE AL SG-SST POR EL VIGÍA DE SST (VIGÍA DE SST)",
        "INFORME ANUAL DE GESTIÓN FRENTE AL SG-SST POR EL VIGÍA DE SEGURIDAD Y SALUD EN EL TRABAJO",
    )
    reemplazar(doc,
        "INFORME ANUAL DE GESTIÓN FRENTE AL SG-SST POR EL COMITÉ PARITARIO DE SEGURIDAD Y SALUD EN EL TRABAJO (VIGÍA DE SST)",
        "INFORME ANUAL DE GESTIÓN FRENTE AL SG-SST POR EL VIGÍA DE SEGURIDAD Y SALUD EN EL TRABAJO",
    )

    # Introducción
    reemplazar_parrafo(doc,
        "El Vigía de Seguridad y Salud en el Trabajo presenta su informe anual",
        "El Vigía de Seguridad y Salud en el Trabajo de {{RAZON_SOCIAL}} presenta su "
        "informe anual de gestión frente al Sistema de Gestión de Seguridad y Salud en el "
        "Trabajo (SG-SST) correspondiente al año {{AÑO}}, en cumplimiento del artículo "
        "2.2.4.6.20 del Decreto 1072 de 2015 y la Resolución 0312 de 2019. Este informe "
        "tiene como objetivo evaluar el cumplimiento de las actividades de SST y formular "
        "recomendaciones de mejora.",
    )

    # Objetivo
    reemplazar_parrafo(doc,
        "Evaluar de manera integral el cumplimiento, funcionamiento y gestión del Vigía de SST",
        "Evaluar de manera integral el cumplimiento, funcionamiento y gestión del Vigía de "
        "Seguridad y Salud en el Trabajo, con el fin de identificar fortalezas, oportunidades "
        "de mejora y generar recomendaciones que contribuyan al fortalecimiento del SG-SST.",
    )

    # Sección CONFORMACIÓN → DATOS DEL VIGÍA
    reemplazar(doc, "CONFORMACIÓN DEL COMITÉ", "DATOS DEL VIGÍA DE SST")
    reemplazar(doc, "Designados por el representante legal", "Nombre del Vigía:")
    reemplazar(doc, "Elegidos por parte de los colaboradores:", "Número de identificación:")

    # Metodología
    reemplazar(doc,
        "Revisión documental: Se realizó la verificación de documentos relacionados con el VIGÍA DE SST, como las actas de conformación, actas de reuniones, informes trimestrales, registros de capacitaciones.",
        "Revisión documental: Se realizó la verificación de documentos del SG-SST, como las actas de designación, actas de reuniones, informes trimestrales, registros de capacitaciones e inspecciones realizadas por el Vigía de SST.",
    )
    reemplazar(doc,
        "Participación del VIGÍA DE SST: Se incluyeron los resultados de las reuniones del VIGÍA DE SST y su gestión para fomentar el cumplimiento normativo y la mejora continua.",
        "Gestión del Vigía de SST: Se incluyeron los resultados de las reuniones mensuales del Vigía con el empleador y su gestión para fomentar el cumplimiento normativo y la mejora continua.",
    )

    # Resultados
    reemplazar(doc, "Reuniones del VIGÍA DE SST:", "Reuniones del Vigía de SST con el empleador:")
    reemplazar(doc, "Gestión del VIGÍA DE SST:", "Gestión del Vigía de SST:")

    # Firma revisión
    reemplazar(doc, "Revisado por:", "Elaborado por:")
    reemplazar(doc, "Presidente del VIGÍA DE SST", "Vigía de SST")

    doc.save(str(ruta))
    print(f"  FMT-081 Vigía adaptado: {ruta.name}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    fmt005 = next((f for f in PLANTILLA.rglob("*.docx") if "FMT-005" in f.name and "[VIGÍA]" in f.name), None)
    fmt081 = next((f for f in PLANTILLA.rglob("*.docx") if "FMT-081" in f.name and "[VIGÍA]" in f.name), None)

    if not fmt005:
        print("ERROR: No se encontró SST-FMT-005 [VIGÍA]")
    else:
        adaptar_fmt005(fmt005)

    if not fmt081:
        print("ERROR: No se encontró SST-FMT-081 [VIGÍA]")
    else:
        adaptar_fmt081(fmt081)

    print("\nListo. Verifica los documentos en Word.")


if __name__ == "__main__":
    main()
