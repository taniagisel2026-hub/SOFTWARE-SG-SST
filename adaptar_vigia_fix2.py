#!/usr/bin/env python3
"""
Segunda pasada de correcciones sobre los documentos VIGÍA SST ya adaptados.
Corrige los problemas que quedaron en la primera pasada.
    python adaptar_vigia_fix2.py
"""
from pathlib import Path
from docx import Document

VIGIA_DIR = Path(
    r"Plantilla SG-SST 2026\I. PLANEAR\1.1.6 DESIGNACIÓN DEL VIGÍA SST"
)


def _reemplazar_parrafo(doc: Document, fragmento: str, nuevo: str) -> bool:
    for p in doc.paragraphs:
        texto = "".join(r.text for r in p.runs)
        if fragmento in texto and p.runs:
            p.runs[0].text = nuevo
            for r in p.runs[1:]:
                r.text = ""
            return True
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    texto = "".join(r.text for r in p.runs)
                    if fragmento in texto and p.runs:
                        p.runs[0].text = nuevo
                        for r in p.runs[1:]:
                            r.text = ""
                        return True
    return False


def _reemplazar_texto_simple(doc: Document, viejo: str, nuevo: str) -> int:
    """Reemplaza texto run a run en todo el documento. Retorna cantidad de cambios."""
    cambios = 0
    todos = list(doc.paragraphs)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                todos += list(celda.paragraphs)
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


# ─── Correcciones por documento ──────────────────────────────────────────────

def fix_pro001(doc: Document):
    # "elección directa" → "designación directa" en responsabilidades del Gerente
    _reemplazar_texto_simple(doc, "mediante elección directa", "mediante designación directa")

    # Fila 04: quedó con el mismo contenido que 02 (Designación formal).
    # Buscar la fila 04 por el texto que tiene y corregir a Aceptación.
    for tabla in doc.tables:
        for fila in tabla.rows:
            celdas = fila.cells
            if len(celdas) < 4:
                continue
            # La fila que dice "04" en la primera columna
            if celdas[0].text.strip() == "04":
                for p in celdas[1].paragraphs:
                    if p.runs:
                        p.runs[0].text = "Vigía de SST designado"
                        for r in p.runs[1:]:
                            r.text = ""
                for p in celdas[2].paragraphs:
                    if p.runs:
                        p.runs[0].text = "Aceptación"
                        for r in p.runs[1:]:
                            r.text = ""
                for p in celdas[3].paragraphs:
                    if p.runs:
                        p.runs[0].text = (
                            "El trabajador designado firma el Acta de Conformación y "
                            "Aceptación (SST-FTM-014), manifestando su aceptación del "
                            "cargo y su compromiso de cumplir las funciones establecidas "
                            "en la normativa."
                        )
                        for r in p.runs[1:]:
                            r.text = ""

    # Fila 06: el texto tiene restos de "Integrantes del VIGÍA DE SST" y la mención
    # a "artículo sexto" de la resolución. Limpiar la celda de descripción.
    for tabla in doc.tables:
        for fila in tabla.rows:
            celdas = fila.cells
            if len(celdas) < 4:
                continue
            if celdas[0].text.strip() == "06":
                for p in celdas[1].paragraphs:
                    if p.runs:
                        p.runs[0].text = "Vigía de SST / Encargado SG-SST"
                        for r in p.runs[1:]:
                            r.text = ""
                for p in celdas[3].paragraphs:
                    if p.runs:
                        p.runs[0].text = (
                            "El Vigía de SST se reúne mensualmente con el empleador para "
                            "revisar el avance del SG-SST, las condiciones de trabajo, los "
                            "accidentes e incidentes y las recomendaciones de mejora. Al "
                            "finalizar el año se evalúa el desempeño del Vigía y se decide "
                            "la redesignación."
                        )
                        for r in p.runs[1:]:
                            r.text = ""

    # DOCUMENTOS RELACIONADOS: eliminar el "ANEXO SST-FMT-015 ACTA DE REUNIÓN." duplicado
    _reemplazar_texto_simple(
        doc,
        "ANEXO SST-FMT-015 ACTA DE REUNIÓN.",
        "",
    )

    # BIBLIOGRAFÍA: Resolución 2013 aún presente
    _reemplazar_texto_simple(
        doc,
        "Resolución 2013 (06 de 006 de 1986)",
        "Decreto 1072 (26 de 05 de 2015), artículo 2.2.4.6.8",
    )
    _reemplazar_texto_simple(
        doc,
        "Por la cual se reglamenta la organización y funcionamiento de los Comités Paritarios"
        " de Seguridad y Salud en el Trabajo».",
        "Designación del Vigía de SST en empresas con menos de diez (10) trabajadores.",
    )
    _reemplazar_texto_simple(
        doc,
        "Ministerio de Trabajo y Protección social,",
        "Ministerio de Trabajo.",
    )


def fix_ftm014(doc: Document):
    # Primer párrafo: reemplazar "se conformó el Vigía..."
    _reemplazar_parrafo(
        doc,
        "se conformó el Vigía de La Seguridad",
        "En cumplimiento del artículo 2.2.4.6.8 del Decreto 1072 de 2015 y la "
        "Resolución 0312 de 2019, el señor(a) {{REPRESENTANTE_LEGAL}}, en calidad de "
        "{{CARGO_RL}} de {{RAZON_SOCIAL}}, con NIT {{NIT}}, designó como Vigía de "
        "Seguridad y Salud en el Trabajo a:",
    )

    # Párrafo "El periodo de los miembros..."
    _reemplazar_texto_simple(
        doc,
        "El periodo de los miembros es de un (1) año",
        "El período de gestión del Vigía de SST es de un (1) año",
    )

    # Párrafo de conformación que quedó mal
    _reemplazar_parrafo(
        doc,
        "Conformado el Vigía de Seguridad y Salud en el Trabajo, se procede a nombrar",
        "En constancia de lo anterior, el Vigía de SST acepta el cargo y se compromete "
        "a ejercer sus funciones durante el período señalado, conforme a la normativa "
        "vigente.",
    )

    # Eliminar suplentes que quedaron en los bloques de firma
    _reemplazar_texto_simple(doc, "Suplente del Vigía de SST", "")
    _reemplazar_texto_simple(doc, "Suplente de la Representante Legal", "")
    _reemplazar_texto_simple(doc, "Suplente del presidente", "")
    _reemplazar_texto_simple(doc, "Suplente de la secretaria", "")


def fix_fmt015(doc: Document):
    # Asistentes en tabla: "Se verifica asistencia de todos los Vigía de SST."
    _reemplazar_texto_simple(
        doc,
        "Se verifica asistencia de todos los Vigía de SST.",
        "Se verifica la asistencia del Vigía de SST y del Representante Legal.",
    )
    # Eliminar suplentes
    _reemplazar_texto_simple(doc, "Suplente del Vigía de SST", "")
    _reemplazar_texto_simple(doc, "Suplente de la Representante Legal", "")


def fix_fmt016(doc: Document):
    # Introducción: la primera frase del párrafo se perdió, reconstruir
    _reemplazar_parrafo(
        doc,
        "gestión adelantada por el Vigía de Seguridad y Salud en el Trabajo de",
        "El presente informe consolida la gestión adelantada por el Vigía de "
        "Seguridad y Salud en el Trabajo de {{RAZON_SOCIAL}} durante el período "
        "comprendido entre el _________________ y el _________________.",
    )

    # Párrafo con "desarrolladas por el Comité"
    _reemplazar_texto_simple(
        doc,
        "desarrolladas por el Comité,",
        "desarrolladas por el Vigía de SST,",
    )

    # "evaluar la eficacia del Comité"
    _reemplazar_texto_simple(
        doc,
        "evaluar la eficacia del Comité,",
        "evaluar la gestión del Vigía de SST,",
    )
    _reemplazar_texto_simple(
        doc,
        "eficacia del Comité,",
        "gestión del Vigía de SST,",
    )

    # Tabla indicadores: limpiar "VIGÍA DE SST" en mayúsculas sostenidas → "Vigía de SST"
    _reemplazar_texto_simple(doc, "VIGÍA DE SST", "Vigía de SST")

    # Eliminar suplentes en firmas
    _reemplazar_texto_simple(doc, "Suplente del Vigía de SST", "")
    _reemplazar_texto_simple(doc, "Suplente de la Representante Legal", "")


# ─── Main ─────────────────────────────────────────────────────────────────────

FIXES = {
    "SST-PRO-001": fix_pro001,
    "SST-FTM-014": fix_ftm014,
    "SST-FMT-015": fix_fmt015,
    "SST-FMT-016": fix_fmt016,
}


def main():
    for archivo in sorted(VIGIA_DIR.rglob("*.docx")):
        if archivo.name.startswith("~$"):
            continue
        codigo = archivo.name[:11]
        if codigo not in FIXES:
            continue
        print(f"  Corrigiendo: {archivo.name}")
        try:
            doc = Document(str(archivo))
            FIXES[codigo](doc)
            doc.save(str(archivo))
            print("    OK")
        except Exception as e:
            print(f"    ERROR: {e}")

    print("\nSegunda pasada completada.")


if __name__ == "__main__":
    main()
