"""Actualizador legal de normas SG-SST - correcciones deterministas.

Basado en el trabajo de análisis de Medina SST Empresarial S.A.S.
Aplica correcciones verificadas a normas derogadas o modificadas
en los documentos del SG-SST.

Para agregar una nueva corrección, solo agrega una entrada a la
lista CORRECCIONES (las más específicas van primero):

    ("NORMA VIEJA", "Norma nueva correcta"),

IMPORTANTE sobre runs fragmentados en Word:
    Word puede dividir "RESOLUCIÓN 652 DE 2012" en varios runs:
    ["RESOLUCIÓN 652 DE", " 2012"] o ["RESOL", "UCIÓN 652 DE 2012"]
    Este módulo une los runs del párrafo para buscar coincidencias
    y luego reconstruye los runs preservando el formato original.
"""
import re
import logging
from dataclasses import dataclass, field

from docx import Document


logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════
#  MAPA DE CORRECCIONES LEGALES VERIFICADAS
#  ⚠️  ORDEN IMPORTA: las más específicas van PRIMERO.
#  ⚠️  NO reordenar sin revisar que las reglas largas precedan a las cortas.
#  Se aplican con re.IGNORECASE y \b (word boundaries).
# ══════════════════════════════════════════════════════════════════════

CORRECCIONES = [
    # [1] Cita conjunta (la más específica, va primero)
    ("RESOLUCIÓN NO. 652 Y 1356 DE 2012",   "Resolución 3461 de 2025"),
    ("Resoluciones No. 652 y 1356 de 2012",  "Resolución 3461 de 2025"),

    # [2] Error tipográfico: no es Decreto, es Ley (específico antes que genérico)
    ("DECRETO 1010 DE 2006", "Ley 1010 de 2006"),
    ("DECRETO 1010",         "Ley 1010"),

    # [3] Decreto 614 de 1984 - derogado, compilado en DUR
    ("DECRETO 614 DE 1984",  "Decreto 1072 de 2015"),

    # [4] Decreto 1127 de 1991 - incorporado en DUR
    ("DECRETO 1127 DE 1991", "Decreto 1072 de 2015"),

    # [5] Resolución 2669 de 2012 - no existe, ir directo a 3461
    ("RESOLUCIÓN 2669 DE 2012", "Resolución 3461 de 2025"),

    # [6] Resolución 652 de 2012 - derogada por 3461 de 2025 (específica antes)
    ("RESOLUCIÓN 652 DE 2012",  "Resolución 3461 de 2025"),
    ("RESOLUCIÓN 652",          "Resolución 3461 de 2025"),

    # [7] Resolución 1356 de 2012 - derogada por 3461 de 2025
    ("RESOLUCIÓN 1356 DE 2012",         "Resolución 3461 de 2025"),
    ("RESOLUCIÓN\xa01356 DE 2012",      "Resolución 3461 de 2025"),

    # [8] Auditoría normativa 2026-07-08 - errores de cita (número/año que no existe)
    ("LEY 9 DE 1997",           "Ley 9 de 1979"),
    ("RESOLUCIÓN 2400 DE 1997", "Resolución 2400 de 1979"),
    ("DECRETO 1972 DE 1990",    "Resolución 1792 de 1990"),
    ("LEY 1562 DE 2013",        "Ley 1562 de 2012"),
    ("RESOLUCIÓN 2013 DE 2017", "Resolución 2013 de 1986"),

    # [9] Auditoría normativa 2026-07-08 - compiladas/derogadas por el Decreto 1072 de 2015
    ("DECRETO 1443 DE 2014", "Decreto 1072 de 2015"),
    ("DECRETO 1530 DE 1996", "Decreto 1072 de 2015"),

    # [10] Auditoría normativa 2026-07-08 - evaluaciones médicas ocupacionales,
    # derogadas por la Resolución 1843 de 2025 (Mintrabajo, 29-abr-2025)
    ("RESOLUCIÓN 2346 DE 2007", "Resolución 1843 de 2025"),
    ("RESOLUCIÓN 1918 DE 2009", "Resolución 1843 de 2025"),
    ("RESOLUCIÓN 1075 DE 1992", "Resolución 1843 de 2025"),

    # [11] Auditoría normativa 2026-07-08 - trabajo seguro en alturas,
    # derogadas por la Resolución 4272 de 2021
    ("RESOLUCIÓN 3673 DE 2008", "Resolución 4272 de 2021"),
    ("RESOLUCIÓN 1409 DE 2012", "Resolución 4272 de 2021"),

    # [12] Auditoría normativa 2026-07-08 - licencias SST,
    # derogadas por la Resolución 908 de 2025 (vigente desde 15-sep-2025)
    ("RESOLUCIÓN 754 DE 2021",  "Resolución 908 de 2025"),
    ("RESOLUCIÓN 1151 DE 2022", "Resolución 908 de 2025"),

    # [13] Auditoría normativa 2026-07-08 - uso racional de bolsas plásticas,
    # derogadas por la Resolución 803 de 2024 (Minambiente)
    ("RESOLUCIÓN 2184 DE 2019", "Resolución 803 de 2024"),
    ("RESOLUCIÓN 668 DE 2016",  "Resolución 803 de 2024"),

    # [14] Auditoría normativa 2026-07-08 - gestión del riesgo de desastres,
    # derogadas por la Ley 1523 de 2012
    ("DECRETO 919 DE 1989", "Ley 1523 de 2012"),
    ("LEY 46 DE 1988",      "Ley 1523 de 2012"),

    # [15] Auditoría normativa 2026-07-08 - circulares COVID/despido con fuero,
    # derogadas expresamente por la Circular Externa 0071 de 2022 (Mintrabajo)
    ("CIRCULAR 33 DE 2020", "Circular Externa 0071 de 2022"),
    ("CIRCULAR 49 DE 2019", "Circular Externa 0071 de 2022"),

    # [16] Auditoría normativa 2026-07-08 - otras derogatorias puntuales confirmadas
    ("DECRETO 1607 DE 2002",    "Decreto 768 de 2022"),
    ("RESOLUCIÓN 1111 DE 2017", "Resolución 0312 de 2019"),
    ("RESOLUCIÓN 881 DE 2021",  "Resolución 978 de 2023"),
    ("RESOLUCIÓN 3032 DE 2022", "Resolución 2851 de 2023"),
    ("RESOLUCIÓN 90708 DE 2013", "Resolución 40117 de 2024"),
    ("LEY 392 DE 1997",         "Ley 842 de 2003"),
    ("RESOLUCIÓN 2566 DE 2009", "Decreto 1477 de 2014"),
    ("LEY 322 DE 1996",         "Ley 1575 de 2012"),
    ("RESOLUCIÓN 1565 DE 2014", "Resolución 40595 de 2022"),
    ("RESOLUCIÓN 2404 DE 2019", "Resolución 2764 de 2022"),
    ("RESOLUCIÓN 1045 DE 2003", "Resolución 754 de 2014"),
    ("DECRETO 423 DE 2006",     "Decreto 646 de 2025"),
]

# Compilar patrones una sola vez (con word boundaries para evitar parciales)
_PATRONES_COMPILADOS = []
for _viejo, _nuevo in CORRECCIONES:
    # Normalizar espacios especiales para el regex
    _patron_str = re.escape(_viejo).replace(r"\ ", r"\s+").replace(re.escape("\xa0"), r"\s+")
    _patron = re.compile(r"\b" + _patron_str + r"\b", re.IGNORECASE)
    _PATRONES_COMPILADOS.append((_patron, _nuevo))


# ══════════════════════════════════════════════════════════════════════
#  REGISTRO DE CAMBIOS (trazabilidad / auditoría)
# ══════════════════════════════════════════════════════════════════════

@dataclass
class CambioRegistrado:
    """Un cambio individual aplicado a un documento."""
    ubicacion: str          # "párrafo 12", "tabla 3, fila 2, celda 1", etc.
    texto_original: str     # Fragmento encontrado
    texto_nuevo: str        # Reemplazo aplicado
    regla: str              # Patrón que hizo match


@dataclass
class ReporteCorrecciones:
    """Reporte completo de correcciones aplicadas a un documento."""
    archivo: str
    total_cambios: int = 0
    detalle: list[CambioRegistrado] = field(default_factory=list)

    def agregar(self, cambio: CambioRegistrado):
        self.detalle.append(cambio)
        self.total_cambios += 1

    def resumen(self) -> str:
        """Resumen legible del reporte."""
        if not self.detalle:
            return f"  {self.archivo}: sin correcciones necesarias."
        lineas = [f"  {self.archivo}: {self.total_cambios} corrección(es)"]
        for c in self.detalle:
            lineas.append(f"    [{c.ubicacion}] «{c.texto_original}» → «{c.texto_nuevo}»")
        return "\n".join(lineas)


# ══════════════════════════════════════════════════════════════════════
#  FUNCIONES PRINCIPALES
# ══════════════════════════════════════════════════════════════════════

def _reparar_content_types_si_falta_watermark(docx_path: str) -> bool:
    """Repara [Content_Types].xml si tiene imágenes de marca de agua sin declarar.

    Algunos archivos .docx de plantillas vienen con referencias a imágenes
    (watermark_*.png) que el [Content_Types].xml no declara, lo que hace que
    python-docx falle al abrir el archivo. Este helper detecta el problema
    y agrega las declaraciones faltantes sin alterar el contenido del documento.

    Returns:
        True si se hizo una reparación, False si no fue necesario.
    """
    import zipfile
    import shutil
    import re as _re
    import tempfile
    from pathlib import Path as _Path

    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            nombres = z.namelist()
            content_types = z.read("[Content_Types].xml").decode("utf-8")
    except Exception:
        return False

    # Encontrar imagenes en word/media/ que NO esten declaradas
    imagenes = [n for n in nombres if n.startswith("word/media/") and "." in n]
    extensiones_imagenes = set()
    for img in imagenes:
        ext = img.rsplit(".", 1)[-1].lower()
        if ext in ("png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "wmf", "emf"):
            extensiones_imagenes.add(ext)

    # Verificar que cada extension este declarada con Default Extension="ext"
    faltantes = []
    for ext in extensiones_imagenes:
        if f'Extension="{ext}"' not in content_types:
            faltantes.append(ext)

    if not faltantes:
        return False

    # Mapeo extension -> ContentType
    mime_map = {
        "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "gif": "image/gif", "bmp": "image/bmp", "tif": "image/tiff",
        "tiff": "image/tiff", "wmf": "image/x-wmf", "emf": "image/x-emf",
    }

    # Inyectar las declaraciones Default faltantes
    nuevos = "".join(
        f'<Default Extension="{ext}" ContentType="{mime_map.get(ext, "application/octet-stream")}"/>'
        for ext in faltantes
    )
    content_types_reparado = _re.sub(
        r"(<Types[^>]*>)", r"\1" + nuevos, content_types, count=1,
    )

    # Reescribir el ZIP con el [Content_Types].xml corregido preservando TODO
    # el metadato original (compresion, fechas, atributos, comentarios).
    tmp_path = _Path(docx_path).with_suffix(".tmp.docx")
    try:
        with zipfile.ZipFile(docx_path, "r") as z_in:
            with zipfile.ZipFile(str(tmp_path), "w") as z_out:
                for item in z_in.infolist():
                    # Leer bytes originales
                    if item.filename == "[Content_Types].xml":
                        data = content_types_reparado.encode("utf-8")
                    else:
                        data = z_in.read(item.filename)
                    # Crear nuevo ZipInfo preservando compresion y atributos
                    nuevo = zipfile.ZipInfo(filename=item.filename, date_time=item.date_time)
                    nuevo.compress_type = item.compress_type
                    nuevo.external_attr = item.external_attr
                    nuevo.internal_attr = item.internal_attr
                    nuevo.create_system = item.create_system
                    nuevo.create_version = item.create_version
                    nuevo.extract_version = item.extract_version
                    nuevo.flag_bits = item.flag_bits
                    z_out.writestr(nuevo, data)
        shutil.move(str(tmp_path), docx_path)
        logger.info("Reparado [Content_Types].xml de %s (extensiones: %s)",
                    docx_path, ", ".join(faltantes))
        return True
    except Exception as e:
        logger.warning("No se pudo reparar Content_Types de %s: %s", docx_path, e)
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except Exception:
                pass
        return False


def aplicar_correcciones_docx(docx_path: str, detallado: bool = False):
    """Aplica correcciones legales verificadas a un documento Word.

    Usa el texto completo de cada párrafo (uniendo runs) para detectar
    normas que Word haya fragmentado en múltiples runs.
    Preserva el formato original de cada run lo mejor posible.

    Args:
        docx_path: Ruta al archivo .docx (se modifica in-place,
                   solo en archivos de SALIDA, nunca plantillas)
        detallado: Si True, retorna ReporteCorrecciones con detalle
                   de cada cambio. Si False, retorna int (compatibilidad).

    Returns:
        int o ReporteCorrecciones según el parámetro detallado
    """
    # Intentar abrir; si falla por Content_Types mal armado, repararlo
    try:
        doc = Document(docx_path)
    except Exception as e:
        if "content type" in str(e).lower() or "partname" in str(e).lower():
            if _reparar_content_types_si_falta_watermark(docx_path):
                doc = Document(docx_path)
            else:
                raise
        else:
            raise
    reporte = ReporteCorrecciones(archivo=docx_path)

    # Párrafos del cuerpo
    _corregir_parrafos(doc.paragraphs, reporte, "cuerpo")

    # Párrafos dentro de tablas
    for t_idx, tabla in enumerate(doc.tables, 1):
        for f_idx, fila in enumerate(tabla.rows, 1):
            for c_idx, celda in enumerate(fila.cells, 1):
                ubicacion = f"tabla {t_idx}, fila {f_idx}, celda {c_idx}"
                _corregir_parrafos(celda.paragraphs, reporte, ubicacion)

    # Encabezados y pies de página
    for s_idx, section in enumerate(doc.sections, 1):
        for nombre, element in [("encabezado", section.header), ("pie", section.footer)]:
            ubi_base = f"{nombre} sección {s_idx}"
            _corregir_parrafos(element.paragraphs, reporte, ubi_base)
            for t_idx, tabla in enumerate(element.tables, 1):
                for f_idx, fila in enumerate(tabla.rows, 1):
                    for c_idx, celda in enumerate(fila.cells, 1):
                        ubicacion = f"{ubi_base}, tabla {t_idx}, fila {f_idx}, celda {c_idx}"
                        _corregir_parrafos(celda.paragraphs, reporte, ubicacion)

    if reporte.total_cambios > 0:
        doc.save(docx_path)
        logger.info("Guardado %s con %d corrección(es)", docx_path, reporte.total_cambios)

    if detallado:
        return reporte
    return reporte.total_cambios


def _corregir_parrafos(parrafos, reporte: ReporteCorrecciones, ubicacion: str) -> None:
    """Aplica correcciones a una lista de párrafos.

    Para cada párrafo:
    1. Une el texto de todos los runs para obtener el texto completo
    2. Busca coincidencias en el texto ORIGINAL (no el ya modificado)
    3. Aplica TODAS las reglas de una sola pasada sobre el original
    4. Si hay cambios, reconstruye los runs preservando formato

    Esto evita encadenamientos silenciosos (A->B, B->C = A->C) y
    garantiza que el reporte refleje lo que realmente estaba en el documento.
    """
    for p_idx, p in enumerate(parrafos, 1):
        if not p.runs:
            continue

        texto_original = "".join(run.text for run in p.runs)
        if not texto_original.strip():
            continue

        # Fase 1: Detectar TODOS los matches sobre el texto ORIGINAL
        # Cada match se registra con su posición para evitar solapamientos
        reemplazos = []  # [(inicio, fin, texto_nuevo, regla_pattern)]
        for patron, nuevo in _PATRONES_COMPILADOS:
            for m in patron.finditer(texto_original):
                # Verificar que no solape con un reemplazo ya registrado
                solapa = False
                for r_ini, r_fin, _, _ in reemplazos:
                    if m.start() < r_fin and m.end() > r_ini:
                        solapa = True
                        break
                if not solapa:
                    reemplazos.append((m.start(), m.end(), nuevo, patron.pattern))
                    reporte.agregar(CambioRegistrado(
                        ubicacion=f"{ubicacion}, parrafo {p_idx}",
                        texto_original=m.group(),
                        texto_nuevo=nuevo,
                        regla=patron.pattern,
                    ))
                    logger.debug(
                        "[%s, parr %d] '%s' -> '%s'",
                        ubicacion, p_idx, m.group(), nuevo,
                    )

        if not reemplazos:
            continue

        # Fase 2: Aplicar reemplazos de derecha a izquierda (para no mover offsets)
        reemplazos.sort(key=lambda r: r[0], reverse=True)
        texto_corregido = texto_original
        for inicio, fin, nuevo_txt, _ in reemplazos:
            texto_corregido = texto_corregido[:inicio] + nuevo_txt + texto_corregido[fin:]

        # Fase 3: Reconstruir runs preservando formato
        _reconstruir_runs(p.runs, texto_original, texto_corregido)


def _reconstruir_runs(runs, texto_viejo: str, texto_nuevo: str) -> None:
    """Reconstruye los runs de un párrafo con texto nuevo preservando formato.

    Estrategia blindada:
    1. Encuentra la zona de diferencia entre texto viejo y nuevo
    2. Clasifica cada run: ANTES / INICIO / MEDIO / FIN / DESPUÉS
    3. Runs ANTES y DESPUÉS: no se tocan (preservan formato)
    4. Run de INICIO: recibe prefijo intacto + texto nuevo
    5. Runs MEDIO: se vacían SIEMPRE (sin excepción)
    6. Run de FIN: conserva solo el sufijo posterior al cambio

    Garantía: al terminar, "".join(r.text for r in runs) == texto_nuevo
    """
    if not runs:
        return

    # --- Paso 1: Encontrar zona de diferencia ---
    start_diff = 0
    min_len = min(len(texto_viejo), len(texto_nuevo))
    while start_diff < min_len and texto_viejo[start_diff] == texto_nuevo[start_diff]:
        start_diff += 1

    end_diff_viejo = len(texto_viejo)
    end_diff_nuevo = len(texto_nuevo)
    while (end_diff_viejo > start_diff and end_diff_nuevo > start_diff
           and texto_viejo[end_diff_viejo - 1] == texto_nuevo[end_diff_nuevo - 1]):
        end_diff_viejo -= 1
        end_diff_nuevo -= 1

    fragmento_nuevo = texto_nuevo[start_diff:end_diff_nuevo]

    # --- Paso 2: Mapa de runs (posición inicio/fin en texto original) ---
    run_map = []  # [(inicio, fin)]
    pos = 0
    for run in runs:
        end = pos + len(run.text)
        run_map.append((pos, end))
        pos = end

    # --- Paso 3: Clasificar y reconstruir cada run ---
    insercion_hecha = False

    for idx, (inicio_run, fin_run) in enumerate(run_map):
        run = runs[idx]

        # ANTES de la zona de cambio: no tocar
        if fin_run <= start_diff:
            continue

        # DESPUÉS de la zona de cambio: no tocar
        if inicio_run >= end_diff_viejo:
            continue

        # --- Este run está DENTRO o CRUZA la zona de cambio ---

        if inicio_run <= start_diff and fin_run >= end_diff_viejo:
            # Caso especial: UN SOLO run contiene todo el cambio
            prefijo = run.text[:start_diff - inicio_run]
            sufijo = run.text[end_diff_viejo - inicio_run:]
            run.text = prefijo + fragmento_nuevo + sufijo
            insercion_hecha = True

        elif inicio_run <= start_diff:
            # Este run contiene el INICIO del cambio
            prefijo = run.text[:start_diff - inicio_run]
            run.text = prefijo + fragmento_nuevo
            insercion_hecha = True

        elif fin_run >= end_diff_viejo:
            # Este run contiene el FINAL del cambio
            sufijo = run.text[end_diff_viejo - inicio_run:]
            run.text = sufijo

        else:
            # Run completamente DENTRO de la zona: vaciar SIEMPRE
            run.text = ""

    # --- Paso 4: Verificación final ---
    resultado = "".join(r.text for r in runs)
    if resultado != texto_nuevo:
        # Algo salió mal, fallback seguro
        logger.warning(
            "Reconstruccion de runs inconsistente, usando fallback. "
            "Esperado %d chars, obtenido %d chars",
            len(texto_nuevo), len(resultado),
        )
        runs[0].text = texto_nuevo
        for run in runs[1:]:
            run.text = ""


def listar_correcciones() -> list[dict]:
    """Retorna la lista de correcciones configuradas (para mostrar en CLI)."""
    return [
        {"vieja": vieja, "nueva": nueva}
        for vieja, nueva in CORRECCIONES
    ]
