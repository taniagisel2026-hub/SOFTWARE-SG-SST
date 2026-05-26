"""Generador del Listado Maestro de Documentos del SG-SST."""
import logging
import re
from pathlib import Path
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

logger = logging.getLogger(__name__)

# Codigo tipo SST-FMT-001, PESV-POL-007, PESV-MTZ-002, etc.
PATRON_CODIGO_GENERAL = re.compile(r"(?:SST|PESV)-[A-Z]{3}-\d{3}")

# Equivalencias de prefijos (typos historicos en plantilla):
# FMT y FTM se usan indistintamente para "Formato"
_PREFIJOS_EQUIVALENTES = {
    "FMT": "FMT",
    "FTM": "FMT",  # FTM se normaliza a FMT
}


def _normalizar_codigo(codigo: str) -> str:
    """Normaliza un codigo SST/PESV-XXX-NNN aplicando equivalencias de prefijos.
    Ejemplo: SST-FTM-010 -> SST-FMT-010
    """
    m = re.match(r"((?:SST|PESV))-([A-Z]{3})-(\d{3})", codigo)
    if not m:
        return codigo
    sistema, prefijo, num = m.groups()
    prefijo_norm = _PREFIJOS_EQUIVALENTES.get(prefijo, prefijo)
    return f"{sistema}-{prefijo_norm}-{num}"

# Ubicacion de la matriz de gestion documental dentro del output
RUTA_MATRIZ_GESTION = Path("I. PLANEAR") / "2.5.1 ARCHIVO Y RETENCIÓN DOCUMENTAL" / "SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL.xlsx"

# Configuracion por hoja:
# - col_codigo: columna con el codigo del documento (ej. SST-FMT-001)
# - col_nombre: columna con el NOMBRE/TITULO del documento (para desambiguar
#   cuando un mismo codigo se usa para varios documentos)
# - col_link: columna donde se inserta el hipervinculo
# - fila_inicio: primera fila con datos
HOJAS_MATRIZ = {
    "LISTADO MAESTRO PHVA": {
        "col_codigo": 4, "col_nombre": 6, "col_link": 22,
        "col_seccion": 3,  # Numeral + nombre de seccion (ej: '4.1.2 Identificacion de peligros')
        "fila_inicio": 7,
    },
    "PESV": {
        "col_codigo": 3, "col_nombre": 5, "col_link": 15,
        "col_seccion": 2,  # Numeral en hoja PESV
        "fila_inicio": 5,
    },
}

# Numeral tipo 1.1.6, 4.1.2, 5.1.1, etc.
PATRON_NUMERAL = re.compile(r"\b(\d+(?:\.\d+){1,3})\b")

FUENTE_HIPERVINCULO = Font(size=10, name="Arial", color="0563C1", underline="single")


def _normalizar_nombre(texto: str) -> str:
    """Normaliza un texto para comparacion fuzzy: minusculas, sin acentos,
    sin puntuacion ni dobles espacios, sin el codigo SST-XXX-NNN."""
    import unicodedata
    t = str(texto or "").strip().lower()
    # Quitar extension
    t = re.sub(r"\.(docx?|xlsx?|pdf)$", "", t)
    # Quitar el codigo
    t = PATRON_CODIGO_GENERAL.sub("", t)
    # Quitar acentos
    t = "".join(
        c for c in unicodedata.normalize("NFD", t)
        if unicodedata.category(c) != "Mn"
    )
    # Quitar puntuacion, dejar solo letras/numeros/espacios
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    # Colapsar espacios
    t = re.sub(r"\s+", " ", t).strip()
    return t


_STOPWORDS = {
    "de", "del", "la", "el", "los", "las", "y", "o", "en", "para",
    "con", "por", "a", "al", "un", "una", "sst", "sgsst", "sg",
}


def _similitud_palabras(a: str, b: str, umbral: float = 0.85) -> bool:
    """Considera dos palabras 'iguales' si difieren en pocas letras (typos).
    Usa SequenceMatcher de difflib para tolerar errores ortograficos comunes.
    Ejemplos: ADQUISICION ~ ADQUISISION, INSPECION ~ INSPECCION.
    """
    if a == b:
        return True
    if abs(len(a) - len(b)) > 3:
        return False
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a, b).ratio() >= umbral


def _similitud(a: str, b: str) -> float:
    """Calcula similitud entre dos nombres normalizados.

    Combina:
      - Jaccard sobre palabras significativas (ignorando stopwords)
      - Containment (que el set menor este contenido en el mayor)
      - Tolerancia a typos: palabras casi iguales (ADQUISICION~ADQUISISION) cuentan
    """
    pa = set(_normalizar_nombre(a).split()) - _STOPWORDS
    pb = set(_normalizar_nombre(b).split()) - _STOPWORDS
    if not pa or not pb:
        return 0.0

    # Calcular interseccion FUZZY: cada palabra de pa que tenga su "casi igual" en pb
    interseccion_fuzzy = 0
    pb_libre = set(pb)
    for pal_a in pa:
        for pal_b in list(pb_libre):
            if _similitud_palabras(pal_a, pal_b):
                interseccion_fuzzy += 1
                pb_libre.discard(pal_b)
                break

    union = len(pa | pb)
    jaccard = interseccion_fuzzy / union if union else 0.0
    menor = min(len(pa), len(pb))
    containment = interseccion_fuzzy / menor if menor else 0.0
    return max(jaccard, containment)


def _indexar_drive_links_por_codigo(drive_links: dict) -> tuple[dict, dict]:
    """Convierte el mapeo {ruta_relativa: {link, nombre, id}} en dos indices:
      1. {codigo_exacto: [(link, nombre_archivo, ruta), ...]}
      2. {prefijo: [(link, nombre_archivo, ruta), ...]} para fallback por nombre

    El indice por prefijo permite que si la matriz dice SST-PRG-005 pero el
    archivo en Drive esta numerado SST-PRG-006, se pueda igualmente encontrar
    el archivo correcto buscando por NOMBRE dentro de la familia SST-PRG-*.

    IMPORTANTE: el indice exacto devuelve una LISTA por codigo (no un solo
    valor) porque en el sistema existen codigos duplicados (ej. SST-RGL-001
    se usa para Reglamento COPASST y Reglamento de Higiene).
    """
    indice = {}
    indice_por_prefijo = {}
    if not drive_links:
        return indice, indice_por_prefijo
    for ruta_relativa, info in drive_links.items():
        nombre = info.get("nombre") or Path(ruta_relativa).name
        link = info.get("link", "")
        if not link:
            continue
        m = PATRON_CODIGO_GENERAL.search(nombre)
        if not m:
            continue
        codigo = _normalizar_codigo(m.group(0))  # FTM -> FMT
        entrada = (link, Path(nombre).stem, ruta_relativa)
        indice.setdefault(codigo, []).append(entrada)
        # Indexar tambien por prefijo (SST-PRG, SST-FMT, etc.) para fallback
        m2 = re.match(r"((?:SST|PESV)-[A-Z]{3})-\d{3}", codigo)
        if m2:
            indice_por_prefijo.setdefault(m2.group(1), []).append(entrada)
    return indice, indice_por_prefijo


def _elegir_mejor_match(
    codigo: str,
    nombre_esperado: str,
    candidatos: list,
    seccion_esperada: str = "",
    umbral_minimo: float = 0.40,
) -> tuple[str, str] | None:
    """Dado un codigo + nombre esperado + seccion (numeral) de la matriz,
    elige el candidato (link, nombre) que mejor coincide.

    Reglas:
      - Si hay un solo candidato pero el nombre es muy distinto al esperado,
        retorna None (no asigna un link incorrecto). El nombre debe pasar
        un umbral minimo de similitud.
      - Si hay varios candidatos, elige el de mayor similitud por nombre o ruta.
      - El numeral de la seccion (ej: '4.1.2') se usa para filtrar a la carpeta
        correcta cuando hay duplicados.
    """
    if not candidatos:
        return None

    # Filtrar candidatos por seccion (numeral en la ruta) si aplica
    candidatos_efectivos = candidatos
    numeral = ""
    if seccion_esperada:
        m = PATRON_NUMERAL.search(seccion_esperada)
        if m:
            numeral = m.group(1)
            con_numeral = [c for c in candidatos if numeral in (c[2] or "")]
            if con_numeral:
                candidatos_efectivos = con_numeral

    # Calcular score de cada candidato (nombre vs nombre_archivo y vs ruta)
    scores = []
    for link, nombre_archivo, ruta in candidatos_efectivos:
        if nombre_esperado:
            s = _similitud(nombre_esperado, nombre_archivo)
            s_ruta = _similitud(nombre_esperado, ruta or "")
            score = max(s, s_ruta)
        else:
            score = 0.0
        if numeral and numeral in (ruta or ""):
            score += 0.05
        scores.append((score, link, nombre_archivo, ruta))

    scores.sort(key=lambda x: x[0], reverse=True)
    mejor_score, mejor_link, mejor_nombre, mejor_ruta = scores[0]

    # Si NO tenemos nombre esperado, aceptar el primer candidato
    if not nombre_esperado:
        return (mejor_link, mejor_nombre)

    # Validar similitud minima: si el mejor es muy bajo, no asignar
    if mejor_score < umbral_minimo:
        logger.info(
            "Codigo %s: mejor candidato '%s' (sim=%.2f) bajo umbral %.2f. "
            "No se asigna link. Esperado='%s'",
            codigo, mejor_nombre[:50], mejor_score, umbral_minimo,
            nombre_esperado[:50],
        )
        return None

    return (mejor_link, mejor_nombre)


def actualizar_matriz_gestion_documental(
    output_dir: Path,
    drive_links: dict,
    generar_pesv: bool = True,
) -> Path | None:
    """Inyecta los links de Drive en la matriz SST-MTZ-005 existente.

    - Recorre las hojas 'LISTADO MAESTRO PHVA' y 'PESV'.
    - Por cada fila lee el codigo del documento (col D en PHVA, col C en PESV)
      y escribe un hipervinculo hacia Drive en la columna LINK (V en PHVA, O en PESV).
    - Si el codigo no tiene un archivo subido, limpia la celda (queda en blanco).
    - Respeta las filas sin codigo.
    - Si `generar_pesv=False`, elimina COMPLETAMENTE la hoja 'PESV' del archivo
      (la empresa no solicito Plan Estrategico de Seguridad Vial).

    Args:
        output_dir: Raiz del SG-SST generado del cliente.
        drive_links: Mapeo {ruta_relativa: {link, nombre, id}} retornado por Drive.
        generar_pesv: True si la empresa solicita PESV; si es False, la hoja PESV
            se elimina y no se procesa.

    Returns:
        Ruta a la matriz actualizada, o None si no se encontro la matriz.
    """
    matriz_path = output_dir / RUTA_MATRIZ_GESTION
    if not matriz_path.exists():
        logger.warning("No se encontro la matriz de gestion documental: %s", matriz_path)
        return None

    indice, indice_por_prefijo = _indexar_drive_links_por_codigo(drive_links or {})
    # Umbral minimo de similitud para aceptar un fallback por prefijo+nombre
    UMBRAL_FALLBACK = 0.75

    wb = load_workbook(str(matriz_path))

    # Si la empresa NO solicita PESV, eliminar la hoja PESV por completo
    if not generar_pesv and "PESV" in wb.sheetnames:
        del wb["PESV"]
        logger.info("Hoja PESV eliminada de SST-MTZ-005 (empresa no solicito PESV).")

    total_links = 0
    total_limpiadas = 0

    for nombre_hoja, cfg in HOJAS_MATRIZ.items():
        # Saltar PESV si no se solicita
        if nombre_hoja == "PESV" and not generar_pesv:
            continue
        if nombre_hoja not in wb.sheetnames:
            logger.info("Hoja '%s' no presente en la matriz, se omite.", nombre_hoja)
            continue
        ws = wb[nombre_hoja]
        col_codigo = cfg["col_codigo"]
        col_nombre = cfg["col_nombre"]
        col_link = cfg["col_link"]
        col_seccion = cfg.get("col_seccion")
        seccion_actual = ""  # se hereda hacia abajo cuando esta vacia
        for row in range(cfg["fila_inicio"], ws.max_row + 1):
            # Rastrear seccion/numeral actual (heredado si la celda esta vacia)
            if col_seccion:
                v_sec = ws.cell(row=row, column=col_seccion).value
                if v_sec and str(v_sec).strip():
                    seccion_actual = str(v_sec).strip()
            valor_codigo = ws.cell(row=row, column=col_codigo).value
            if not valor_codigo:
                continue
            m = PATRON_CODIGO_GENERAL.search(str(valor_codigo))
            if not m:
                # Filas sin codigo (ej: "SIN CÓDIGO", certificados) - limpiar link
                celda_link = ws.cell(row=row, column=col_link)
                if celda_link.value or celda_link.hyperlink:
                    celda_link.value = None
                    celda_link.hyperlink = None
                    total_limpiadas += 1
                continue
            codigo = _normalizar_codigo(m.group(0))  # FTM -> FMT

            # Leer el NOMBRE esperado del documento desde la columna NOMBRE de la matriz.
            # Esto permite desambiguar codigos duplicados (ej. SST-RGL-001 se usa
            # para Reglamento COPASST y para Reglamento de Higiene).
            nombre_esperado = ws.cell(row=row, column=col_nombre).value or ""
            nombre_esperado = str(nombre_esperado).strip()

            celda_link = ws.cell(row=row, column=col_link)
            candidatos = indice.get(codigo) or []
            match = _elegir_mejor_match(codigo, nombre_esperado, candidatos, seccion_actual)

            # FALLBACK: si no hay match por codigo exacto, buscar por NOMBRE
            # dentro de la familia del prefijo (ej: SST-PRG-005 en matriz vs
            # SST-PRG-006 en el archivo real - mismo programa, codigo distinto).
            # Tambien filtra por carpeta usando la seccion/numeral.
            if not match and nombre_esperado:
                m_pref = re.match(r"((?:SST|PESV)-[A-Z]{3})-\d{3}", codigo)
                if m_pref:
                    candidatos_prefijo = indice_por_prefijo.get(m_pref.group(1)) or []
                    # Filtrar por numeral de la seccion si esta presente
                    numeral = ""
                    if seccion_actual:
                        mn = PATRON_NUMERAL.search(seccion_actual)
                        if mn:
                            numeral = mn.group(1)
                            cand_filtrados = [c for c in candidatos_prefijo if numeral in (c[2] or "")]
                            if cand_filtrados:
                                candidatos_prefijo = cand_filtrados
                    mejor_score = 0.0
                    mejor_cand = None
                    for cand in candidatos_prefijo:
                        score = _similitud(nombre_esperado, cand[1])
                        score_ruta = _similitud(nombre_esperado, cand[2] or "")
                        score = max(score, score_ruta)
                        if numeral and numeral in (cand[2] or ""):
                            score += 0.05
                        if score > mejor_score:
                            mejor_score = score
                            mejor_cand = cand
                    if mejor_cand and mejor_score >= UMBRAL_FALLBACK:
                        match = (mejor_cand[0], mejor_cand[1])
                        logger.info(
                            "Codigo %s (R%d, hoja '%s', seccion '%s') no encontrado por "
                            "codigo exacto; fallback por nombre+carpeta encontro '%s' (sim=%.2f).",
                            codigo, row, nombre_hoja, seccion_actual[:40],
                            mejor_cand[1][:60], mejor_score,
                        )

            if match:
                link, nombre_archivo = match
                # Escapar comillas dobles en el texto para la formula
                texto_escapado = nombre_archivo.replace('"', '""')
                link_escapado = link.replace('"', '""')
                # Usar formula HYPERLINK (mas compatible con Excel y Google Sheets)
                celda_link.value = f'=HYPERLINK("{link_escapado}","{texto_escapado}")'
                celda_link.hyperlink = link
                celda_link.font = FUENTE_HIPERVINCULO
                total_links += 1
                if len(candidatos) > 1:
                    logger.info(
                        "Codigo duplicado %s en R%d (hoja '%s'): nombre esperado='%s' "
                        "-> elegido archivo '%s' entre %d candidatos.",
                        codigo, row, nombre_hoja, nombre_esperado[:60],
                        nombre_archivo[:60], len(candidatos),
                    )
            else:
                # Codigo valido pero sin archivo generado - dejar en blanco
                if celda_link.value or celda_link.hyperlink:
                    celda_link.value = None
                    celda_link.hyperlink = None
                    total_limpiadas += 1

    # Generar pestañas adicionales por tipo de documento
    try:
        _generar_pestanas_por_tipo(wb)
    except Exception as e:
        logger.warning("No se pudieron generar pestañas por tipo: %s", e)

    wb.save(str(matriz_path))
    logger.info(
        "Matriz SST-MTZ-005 actualizada: %d links insertados, %d celdas limpiadas.",
        total_links, total_limpiadas,
    )
    return matriz_path


# ---------------------------------------------------------------------------
# Pestañas por tipo de documento (PROGRAMA, FORMATO, MATRIZ, etc.)
# ---------------------------------------------------------------------------

# Tipos a generar como pestaña independiente, en el orden deseado
TIPOS_PESTANA = [
    "PROGRAMA", "PLAN", "PROCEDIMIENTO", "INSTRUCTIVO",
    "REGLAMENTO", "MANUAL", "POLITICA", "INDICADOR", "MATRIZ", "FORMATO",
]


def _normalizar_tipo(t: str) -> str:
    """Normaliza el tipo: mayusculas sin tildes."""
    import unicodedata
    s = unicodedata.normalize("NFD", str(t or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.strip().upper()


def _generar_pestanas_por_tipo(wb) -> None:
    """Genera una hoja por cada TIPO de documento con todas las filas
    correspondientes del LISTADO MAESTRO PHVA (codigo, nombre, link)."""
    if "LISTADO MAESTRO PHVA" not in wb.sheetnames:
        return

    from openpyxl.styles import Font as _Font, Alignment as _Align, PatternFill as _Fill, Border as _Border, Side as _Side

    ws_origen = wb["LISTADO MAESTRO PHVA"]
    cfg = HOJAS_MATRIZ["LISTADO MAESTRO PHVA"]
    col_codigo = cfg["col_codigo"]
    col_nombre = cfg["col_nombre"]
    col_tipo = 5  # E
    col_link = cfg["col_link"]

    # Recolectar filas por tipo
    por_tipo: dict[str, list[dict]] = {}
    for row in range(cfg["fila_inicio"], ws_origen.max_row + 1):
        tipo_v = ws_origen.cell(row=row, column=col_tipo).value
        if not tipo_v:
            continue
        tipo_norm = _normalizar_tipo(tipo_v)
        if tipo_norm not in [_normalizar_tipo(t) for t in TIPOS_PESTANA]:
            continue
        codigo_v = ws_origen.cell(row=row, column=col_codigo).value
        nombre_v = ws_origen.cell(row=row, column=col_nombre).value
        link_v = ws_origen.cell(row=row, column=col_link).value
        por_tipo.setdefault(tipo_norm, []).append({
            "codigo": str(codigo_v or "").strip(),
            "nombre": str(nombre_v or "").strip(),
            "link": link_v,
        })

    # Estilos
    fuente_header = _Font(name="Arial", size=11, bold=True, color="FFFFFF")
    fuente_celda = _Font(name="Arial", size=10, color="000000")
    fill_header = _Fill(start_color="305496", end_color="305496", fill_type="solid")
    fill_alt1 = _Fill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    fill_alt2 = _Fill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    align_center = _Align(horizontal="center", vertical="center", wrap_text=True)
    align_left = _Align(horizontal="left", vertical="center", wrap_text=True)
    borde = _Border(
        left=_Side(style="thin"), right=_Side(style="thin"),
        top=_Side(style="thin"), bottom=_Side(style="thin"),
    )

    for tipo in TIPOS_PESTANA:
        tipo_norm = _normalizar_tipo(tipo)
        # Borrar la hoja previa si existe (para regenerar limpio)
        if tipo in wb.sheetnames:
            del wb[tipo]
        filas = por_tipo.get(tipo_norm, [])
        if not filas:
            continue
        ws = wb.create_sheet(title=tipo)
        # Cabecera
        ws.cell(row=1, column=1, value="CÓDIGO")
        ws.cell(row=1, column=2, value="NOMBRE DEL DOCUMENTO")
        ws.cell(row=1, column=3, value="LINK")
        for c in range(1, 4):
            cel = ws.cell(row=1, column=c)
            cel.font = fuente_header
            cel.fill = fill_header
            cel.alignment = align_center
            cel.border = borde
        ws.row_dimensions[1].height = 28

        # Anchos de columna
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 60
        ws.column_dimensions["C"].width = 80

        # Filas de datos
        for i, f in enumerate(filas, start=2):
            cel_cod = ws.cell(row=i, column=1, value=f["codigo"])
            cel_cod.alignment = align_center
            cel_nom = ws.cell(row=i, column=2, value=f["nombre"])
            cel_nom.alignment = align_left
            cel_link = ws.cell(row=i, column=3, value=f["link"])
            cel_link.alignment = align_left
            if isinstance(f["link"], str) and f["link"].upper().startswith("=HYPERLINK"):
                cel_link.font = _Font(name="Arial", size=10, color="0563C1", underline="single")
            else:
                cel_link.font = fuente_celda
            for c in (1, 2):
                ws.cell(row=i, column=c).font = fuente_celda
            # Color alternado
            fill = fill_alt1 if (i % 2 == 0) else fill_alt2
            for c in range(1, 4):
                ws.cell(row=i, column=c).fill = fill
                ws.cell(row=i, column=c).border = borde
            ws.row_dimensions[i].height = 30

        # Congelar primera fila
        ws.freeze_panes = "A2"
        logger.info("Pestaña '%s' creada con %d filas en SST-MTZ-005.", tipo, len(filas))

