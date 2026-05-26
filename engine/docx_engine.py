"""Motor de procesamiento de archivos .docx - FUSIONADO v7.6 + motor nuevo.

Usa docxtpl (Jinja2 templates) del v7 para reemplazo robusto de marcadores,
más marca de agua, color negro forzado, verificación de residuos,
y soporte de humanización por secciones con Claude API.
"""
import os
import re
import shutil
from pathlib import Path
from datetime import datetime
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.shared import Mm, RGBColor
from docxtpl import DocxTemplate, InlineImage

from .utils import (
    get_long_path, norm_key, looks_like_image_path,
    download_image_if_url, resolve_image_path,
)


# ══════════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN DE TAMAÑOS
# ══════════════════════════════════════════════════════════════════════

ANCHO_LOGO_WORD_MM = 12
ALTO_LOGO_WORD_MM = 12
ANCHO_FIRMA_MM = 28
ALTO_FIRMA_MM = 12
OPACIDAD_MARCA_AGUA = 0.15
ANCHO_MARCA_AGUA_CM = 16


# ══════════════════════════════════════════════════════════════════════
#  MÓDULO: FORZAR COLOR NEGRO EN WORD (del v7)
# ══════════════════════════════════════════════════════════════════════

# Colores que SI se convierten a negro. Todos los demas se PRESERVAN.
# Solo el rojo corporativo C00000 de la plantilla debe forzarse a negro;
# los azules de titulos, tablas e hipervinculos se mantienen intactos.
_COLORES_FORZAR_NEGRO = {
    "C00000",  # rojo corporativo de plantilla (pediste que se vea negro)
}


def _set_runs_black(paragraph):
    """Convierte SOLO los colores listados en _COLORES_FORZAR_NEGRO a negro.

    - Si run.font.color.rgb es None: color heredado del estilo (heading azul,
      etc.) -> NO tocar, preservar color del estilo.
    - Si tiene color explicito en la lista negra (rojo C00000) -> forzar negro.
    - Cualquier otro color explicito (azul titulos, hipervinculos, etc.) -> preservar.
    """
    for run in paragraph.runs:
        try:
            color_actual = run.font.color.rgb
            if color_actual is None:
                continue  # color del estilo -> preservar
            hex_actual = str(color_actual).upper()
            if hex_actual in _COLORES_FORZAR_NEGRO:
                run.font.color.rgb = RGBColor(0, 0, 0)
            # cualquier otro color explicito -> preservar (no tocar)
        except Exception:
            pass


def _force_text_black_docx(doc_obj):
    """Fuerza texto a color negro SOLO en párrafos del cuerpo,
    respetando colores corporativos (rojo C00000, blanco, azules).

    NO toca tablas para no destruir fuentes blancas en celdas
    con fondo oscuro. NO toca encabezados/pies de página.
    """
    for p in doc_obj.paragraphs:
        _set_runs_black(p)


def _force_xml_colors_black(docx_path: str):
    """Fuerza color negro en todo el texto EXCEPTO en celdas con fondo oscuro.

    Logica:
    - Texto del cuerpo (fuera de tablas) -> fuerza negro
    - Celdas con fondo oscuro (w:fill oscuro) -> intactas (letra blanca)
    - Celdas con fondo claro o sin fondo -> fuerza colores a negro
    """
    import re as _re
    temp_path = docx_path + ".tmp_color"

    # Regex
    tc_pattern = _re.compile(r'<w:tc\b.*?</w:tc>', _re.DOTALL)
    fill_pattern = _re.compile(r'w:fill="([A-Fa-f0-9]{6})"')
    # Solo los colores en _COLORES_FORZAR_NEGRO (C00000 rojo) se convierten a negro.
    # Todo lo demas se preserva (azules de titulos, tablas, hipervinculos, etc.).
    _blacklist_re = "|".join(_COLORES_FORZAR_NEGRO)
    color_re = _re.compile(
        r'(<w:color\b[^>]*\bw:val=")(' + _blacklist_re + r')(")',
        flags=_re.IGNORECASE,
    )

    def _es_oscuro(hex_color: str) -> bool:
        try:
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            # Luminancia estandar; < 128 = oscuro
            return (0.299 * r + 0.587 * g + 0.114 * b) < 128
        except Exception:
            return False

    def _procesar_celda(tc_xml: str) -> str:
        """Si la celda tiene fondo oscuro -> intacta.
        Si fondo claro o ausente -> forzar colores a negro.
        """
        # Buscar fill en los primeros 500 chars (dentro de <w:tcPr>)
        fill_match = fill_pattern.search(tc_xml[:500])
        if fill_match and _es_oscuro(fill_match.group(1)):
            return tc_xml  # preservar letra blanca
        # Fondo claro/ausente: forzar colores de fuente a negro
        return color_re.sub(r"\g<1>000000\3", tc_xml)

    # XMLs a procesar: cuerpo + encabezados + pies de pagina
    # (los encabezados contienen el FORMATO/CODIGO/VERSION/FECHA que tambien
    # tiene texto rojo que debe volverse negro sobre fondos claros)
    xml_targets_re = _re.compile(
        r'^word/(document\.xml|header\d*\.xml|footer\d*\.xml)$'
    )

    try:
        with ZipFile(docx_path, "r") as zin, ZipFile(temp_path, "w", compression=ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if xml_targets_re.match(item.filename):
                    try:
                        text = data.decode("utf-8")
                        # Extraer celdas; cada una se procesa segun su fondo
                        # Padding de 6 digitos: evita colision __CELDA_1__ vs __CELDA_10__
                        celdas_procesadas = []
                        def _guardar_celda(m):
                            celdas_procesadas.append(_procesar_celda(m.group(0)))
                            return f"__CELDA_{len(celdas_procesadas) - 1:06d}__"
                        text_sin_tablas = tc_pattern.sub(_guardar_celda, text)
                        # Forzar negro fuera de tablas
                        text_sin_tablas = color_re.sub(r"\g<1>000000\3", text_sin_tablas)
                        # Restaurar celdas (ya procesadas)
                        for i, celda in enumerate(celdas_procesadas):
                            text_sin_tablas = text_sin_tablas.replace(f"__CELDA_{i:06d}__", celda)
                        data = text_sin_tablas.encode("utf-8")
                    except Exception as e:
                        from rich.console import Console as _C
                        _C().print(f"[yellow]Aviso en _force_xml_colors_black ({item.filename}): {e}[/yellow]")
                zout.writestr(item, data)
        os.replace(temp_path, docx_path)
    except Exception as e:
        from rich.console import Console as _C
        _C().print(f"[yellow]Aviso en _force_xml_colors_black ({docx_path}): {e}[/yellow]")
        if os.path.exists(temp_path):
            os.remove(temp_path)


# ══════════════════════════════════════════════════════════════════════
#  MÓDULO: MARCA DE AGUA CON LOGO (del v7)
# ══════════════════════════════════════════════════════════════════════

def _agregar_marca_agua_logo(docx_path: str, logo_path: str) -> bool:
    """Agrega marca de agua con logo semitransparente en todas las páginas."""
    if not logo_path or not os.path.exists(logo_path):
        return False
    try:
        from PIL import Image as PILImage
        PIL_DISPONIBLE = True
    except ImportError:
        PIL_DISPONIBLE = False

    logo_tmp = logo_path + "_wm_tmp.png"
    try:
        if PIL_DISPONIBLE:
            img_pil = PILImage.open(logo_path).convert("RGBA")
            orig_w, orig_h = img_pil.size
            r, g, b, a = img_pil.split()
            a = a.point(lambda x: int(x * OPACIDAD_MARCA_AGUA))
            PILImage.merge("RGBA", (r, g, b, a)).save(logo_tmp, "PNG")
            logo_efectivo = logo_tmp
        else:
            logo_efectivo = logo_path
            orig_w, orig_h = 1, 1
    except Exception:
        return False

    EMU_POR_CM, EMU_POR_PT = 360000, 12700
    ancho_emu = int(ANCHO_MARCA_AGUA_CM * EMU_POR_CM)
    alto_emu = int(ancho_emu * orig_h / orig_w) if orig_w > 0 else ancho_emu
    ancho_pt = round(ancho_emu / EMU_POR_PT, 2)
    alto_pt = round(alto_emu / EMU_POR_PT, 2)

    with open(logo_efectivo, "rb") as f:
        img_data = f.read()

    NS_REQUERIDOS = {
        "xmlns:v": "urn:schemas-microsoft-com:vml",
        "xmlns:o": "urn:schemas-microsoft-com:office:office",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    }

    def _asegurar_ns_en_root(xml_text):
        for prefijo, uri in NS_REQUERIDOS.items():
            if prefijo not in xml_text:
                xml_text = re.sub(r"(<w:hdr\b)", rf'\1 {prefijo}="{uri}"', xml_text, count=1)
        return xml_text

    def _parrafo_wm(rid):
        return (
            f'<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr/><w:pict>'
            f'<v:shape id="WaterMark" style="position:absolute;margin-left:0;margin-top:0;'
            f"width:{ancho_pt}pt;height:{alto_pt}pt;z-index:-251657216;"
            f'mso-position-horizontal:center;mso-position-horizontal-relative:page;'
            f'mso-position-vertical:center;mso-position-vertical-relative:page" '
            f'o:allowincell="f" filled="f" stroked="f">'
            f'<v:imagedata r:id="{rid}" o:title="watermark"/></v:shape></w:pict></w:r></w:p>'
        )

    def _xml_header_nuevo(rid):
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">'
            + _parrafo_wm(rid)
            + '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p></w:hdr>'
        )

    def _xml_rels_con_imagen(rid, img_filename):
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="media/{img_filename}"/></Relationships>'
        )

    temp_path = docx_path + ".tmp_watermark"
    try:
        with ZipFile(docx_path, "r") as zin, ZipFile(temp_path, "w", compression=ZIP_DEFLATED) as zout:
            archivos_en_zip = set(zin.namelist())
            archivos_ya_escritos = set()
            headers_encontrados = sorted(
                n for n in archivos_en_zip if re.match(r"word/header\d*\.xml$", n)
            )
            rid_counter = [1]

            def _write_once(nombre, data_bytes):
                if nombre not in archivos_ya_escritos:
                    zout.writestr(nombre, data_bytes)
                    archivos_ya_escritos.add(nombre)

            for item in zin.infolist():
                nombre = item.filename
                if nombre in archivos_ya_escritos:
                    continue
                data = zin.read(nombre)

                if nombre in headers_encontrados:
                    try:
                        xml_text = data.decode("utf-8")

                        # IDEMPOTENTE: si este header ya tiene WaterMark,
                        # no agregar otro (evita shape IDs duplicados que
                        # hacen que Word muestre "contenido no legible").
                        if 'id="WaterMark"' in xml_text or "id='WaterMark'" in xml_text:
                            pass  # conservar header tal cual
                        else:
                            hdr_base = nombre.split("/")[-1]
                            hdr_rels = f"word/_rels/{hdr_base}.rels"
                            rid = f"rIdWM{rid_counter[0]}"
                            rid_counter[0] += 1
                            img_fname = f"watermark_{rid}.png"
                            img_part = f"word/media/{img_fname}"

                            xml_text = _asegurar_ns_en_root(xml_text)
                            xml_text = re.sub(
                                r"(<w:hdr\b[^>]*>)", r"\1" + _parrafo_wm(rid), xml_text, count=1
                            )

                            if hdr_rels in archivos_en_zip:
                                rels_txt = zin.read(hdr_rels).decode("utf-8")
                                # Por si acaso: no agregar si ya existe un rIdWM
                                if "rIdWM" not in rels_txt:
                                    nueva_rel = (
                                        f'<Relationship Id="{rid}" '
                                        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                                        f'Target="media/{img_fname}"/>'
                                    )
                                    rels_txt = rels_txt.replace("</Relationships>", nueva_rel + "</Relationships>")
                            else:
                                rels_txt = _xml_rels_con_imagen(rid, img_fname)

                            _write_once(hdr_rels, rels_txt.encode("utf-8"))
                            _write_once(img_part, img_data)
                            data = xml_text.encode("utf-8")
                    except Exception as e:
                        from rich.console import Console as _C
                        _C().print(f"[yellow]Aviso en _agregar_marca_agua ({nombre}): {e}[/yellow]")

                _write_once(nombre, data)

            if not headers_encontrados:
                rid = f"rIdWM{rid_counter[0]}"
                img_fname = f"watermark_{rid}.png"
                _write_once("word/header1.xml", _xml_header_nuevo(rid).encode("utf-8"))
                _write_once(
                    "word/_rels/header1.xml.rels",
                    _xml_rels_con_imagen(rid, img_fname).encode("utf-8"),
                )
                _write_once(f"word/media/{img_fname}", img_data)

        os.replace(temp_path, docx_path)
        if PIL_DISPONIBLE and os.path.exists(logo_tmp):
            try:
                os.remove(logo_tmp)
            except Exception:
                pass
        return True
    except Exception:
        for p in (temp_path, logo_tmp):
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass
        return False


# ══════════════════════════════════════════════════════════════════════
#  MÓDULO: RESOLUCIÓN DE IMÁGENES EN CONTEXTO (del v7 adaptado)
# ══════════════════════════════════════════════════════════════════════

def _to_inline_image(doc: DocxTemplate, key: str, img_path: str):
    """Convierte una imagen a InlineImage de docxtpl con tamaño según tipo."""
    k = norm_key(key).lower()
    if "logo" in k:
        return InlineImage(doc, img_path, width=Mm(ANCHO_LOGO_WORD_MM), height=Mm(ALTO_LOGO_WORD_MM))
    return InlineImage(doc, img_path, width=Mm(ANCHO_FIRMA_MM), height=Mm(ALTO_FIRMA_MM))


def _resolver_imagenes_en_contexto(doc: DocxTemplate, contexto: dict,
                                     fallback_dirs: list, temp_dir: str) -> list:
    """Resuelve imágenes (logo, firmas) en el contexto y las convierte a InlineImage."""
    log = []
    for k in list(contexto.keys()):
        v = contexto.get(k)
        if not isinstance(v, str) or not looks_like_image_path(v):
            continue

        img_path = download_image_if_url(v, temp_dir)
        if not img_path:
            img_path = resolve_image_path(v, fallback_dirs)

        log.append((k, v, img_path, bool(img_path)))
        if img_path:
            contexto[k] = _to_inline_image(doc, k, img_path)
    return log


# ══════════════════════════════════════════════════════════════════════
#  MÓDULO: VERIFICACIÓN DE RESIDUOS (del v7)
# ══════════════════════════════════════════════════════════════════════

def verificar_residuos_docx(docx_path: str) -> list:
    """Escanea el XML del .docx buscando marcadores {{}} sin reemplazar."""
    residuos = []
    patron = re.compile(r"\{\{[^}]+\}\}")
    try:
        with ZipFile(docx_path, "r") as z:
            for nombre in z.namelist():
                if nombre.startswith("word/") and nombre.endswith(".xml"):
                    try:
                        texto = z.read(nombre).decode("utf-8")
                        for match in patron.finditer(texto):
                            residuos.append(f"  [{nombre}] -> '{match.group()}'")
                    except Exception:
                        pass
    except Exception:
        pass
    return list(dict.fromkeys(residuos))


# ══════════════════════════════════════════════════════════════════════
#  FUNCIÓN PRINCIPAL: PROCESAR DOCX (fusión v7 + nuevo)
# ══════════════════════════════════════════════════════════════════════

def procesar_docx(
    plantilla_path: Path,
    output_path: Path,
    contexto: dict,
    temp_dir: str,
    logo_path: str = None,
    secciones_humanizadas: list = None,
) -> dict:
    """Procesa un archivo .docx usando docxtpl (Jinja2) del v7 + mejoras.

    Args:
        plantilla_path: Ruta al archivo .docx plantilla (NO se modifica)
        output_path: Ruta donde se guardará el documento procesado
        contexto: Dict con TODAS las variantes de marcadores (ya expandido)
        temp_dir: Directorio temporal para imágenes
        logo_path: Ruta al logo para marca de agua (opcional)
        secciones_humanizadas: Secciones humanizadas por Claude (opcional)

    Returns:
        Dict con template_vars, faltantes, residuos
    """
    ruta_origen = get_long_path(str(plantilla_path))
    ruta_destino = get_long_path(str(output_path))

    # Cargar con docxtpl
    doc = DocxTemplate(ruta_origen)

    try:
        template_vars = sorted(doc.get_undeclared_template_variables() or [])
    except Exception:
        template_vars = []

    # Resolver imágenes en el contexto
    base_dir = os.path.dirname(ruta_origen)
    contexto_copia = dict(contexto)
    img_log = _resolver_imagenes_en_contexto(doc, contexto_copia, [base_dir, temp_dir], temp_dir)

    # Determinar faltantes
    faltantes = [v for v in template_vars if v not in contexto_copia]

    # Crear directorio de salida
    os.makedirs(os.path.dirname(ruta_destino), exist_ok=True)

    # Renderizar con docxtpl
    try:
        doc.render(contexto_copia, autoescape=False)
    except Exception as e_render:
        # Si falla el render (llaves mal cerradas), copiar original
        shutil.copy2(ruta_origen, ruta_destino)
        return {
            "template_vars": template_vars,
            "faltantes": faltantes,
            "residuos": ["Documento copiado intacto por error de sintaxis en llaves"],
            "error_render": str(e_render),
        }

    # Forzar texto negro
    _force_text_black_docx(doc.docx)

    # Guardar
    try:
        doc.save(ruta_destino)
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base, ext = os.path.splitext(ruta_destino)
        ruta_destino = f"{base}_{ts}{ext}"
        doc.save(ruta_destino)

    # Forzar negro en XML (más profundo)
    _force_xml_colors_black(ruta_destino)

    # Aplicar humanización si hay secciones de Claude
    if secciones_humanizadas:
        _aplicar_humanizacion_post(ruta_destino, secciones_humanizadas)

    # Marca de agua con logo
    if logo_path:
        _agregar_marca_agua_logo(ruta_destino, logo_path)

    # Verificar residuos
    residuos = verificar_residuos_docx(ruta_destino)

    return {
        "template_vars": template_vars,
        "faltantes": faltantes,
        "residuos": residuos,
    }


def _aplicar_humanizacion_post(docx_path: str, secciones_humanizadas: list):
    """Aplica el texto humanizado por Claude al documento ya renderizado.

    Abre el documento con python-docx (no docxtpl) y reemplaza
    los párrafos de contenido humanizable.

    Preserva el formato del run principal (el más largo / no-negrita)
    para que el texto humanizado no herede negrita de runs cortos
    como el nombre de la empresa.
    """
    if not secciones_humanizadas:
        return

    doc = Document(docx_path)

    for seccion in secciones_humanizadas:
        parrafos_info = seccion.get("parrafos", [])
        texto_humanizado = seccion.get("texto_humanizado", "")

        if not texto_humanizado or not parrafos_info:
            continue

        lineas_nuevas = [l for l in texto_humanizado.split("\n") if l.strip()]

        # Si Claude devolvio menos lineas que las esperadas, intentar
        # re-dividir por marcadores comunes que Claude pudo haber pegado en
        # la misma linea: numeros (1., 2., 3.), items de definicion
        # (TERMINO: ...), letras (a), b), c)).
        if len(lineas_nuevas) < len(parrafos_info):
            import re as _re_split
            # Patron que divide ANTES de:
            # - Numero + punto/parentesis: "1. ", "2) "
            # - Letra + punto/parentesis: "a) ", "b. "
            # - Palabra(s) en MAYUSCULAS seguida de ":" (definiciones).
            #   Ej: "ACTIVIDAD FISICA:", "CONDICION FISICA:", "SG-SST:"
            patron_divisor = _re_split.compile(
                r"(?<=[\.\:\s])\s+(?="
                r"\d{1,2}[\.\)]\s"                    # 1. / 2)
                r"|[a-z][\.\)]\s"                      # a) / b.
                r"|[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s\-]{2,40}:"  # TERMINO EN MAYUS:
                r")"
            )
            expandidas = []
            for ln in lineas_nuevas:
                partes = patron_divisor.split(ln)
                expandidas.extend(p.strip() for p in partes if p.strip())
            if len(expandidas) > len(lineas_nuevas):
                lineas_nuevas = expandidas

        # Asignar líneas a párrafos (1:1 si coinciden, o hasta donde alcancen)
        # Si sigue sin coincidir, los parrafos sin match conservan su texto
        # ORIGINAL (no se borran ni quedan vacios).
        pares = []
        if len(lineas_nuevas) == len(parrafos_info):
            pares = list(zip(parrafos_info, lineas_nuevas))
        elif len(lineas_nuevas) < len(parrafos_info):
            # Solo reemplazar hasta donde alcancen las lineas; el resto
            # se deja intacto (original del template)
            pares = [(parrafos_info[i], lineas_nuevas[i]) for i in range(len(lineas_nuevas))]
        else:
            # Claude devolvio mas lineas; tomar solo las que caben
            pares = [(parrafos_info[i], lineas_nuevas[i]) for i in range(len(parrafos_info))]

        for p_info, linea in pares:
            idx = p_info["idx"]
            if idx < len(doc.paragraphs):
                paragraph = doc.paragraphs[idx]
                linea_limpia = linea.strip()
                # Guard: si por error quedo una linea con ** impares
                # (abierto sin cerrar), removerlos para evitar basura visible
                if linea_limpia.count("**") % 2 != 0:
                    linea_limpia = linea_limpia.replace("**", "")
                if linea_limpia:
                    _reemplazar_texto_preservando_formato(paragraph, linea_limpia)

    doc.save(docx_path)


def _reemplazar_texto_preservando_formato(paragraph, texto_nuevo: str):
    """Reemplaza el texto de un párrafo con texto humanizado.

    Preserva negrillas originales: si el texto original tenía palabras
    en negrita (como razón social, normas legales), las detecta en el
    texto nuevo y mantiene esas palabras en negrita.

    Usa paragraph.add_run() (API segura) en vez de deepcopy (que corrompe).
    """
    if not paragraph.runs:
        return

    runs = paragraph.runs

    # Paso 1: Recopilar frases que estaban en negrita
    frases_bold = []
    for run in runs:
        es_bold = run.bold or (run.font and run.font.bold)
        texto = run.text.strip()
        if es_bold and texto and len(texto) > 2:
            frases_bold.append(texto)

    # Paso 2: Capturar formato base del run normal (fuente, tamaño, etc.)
    run_base = runs[0]
    for r in runs:
        if not (r.bold or (r.font and r.font.bold)):
            run_base = r
            break
    font_name = run_base.font.name
    font_size = run_base.font.size
    font_italic = run_base.font.italic

    # Paso 3a: PRIORIDAD 1 - Si Claude marco negrillas con **texto**,
    # usar eso directamente (mas confiable que detectar por frases)
    if "**" in texto_nuevo:
        segmentos = _parsear_markdown_bold(texto_nuevo)
        # Vaciar runs existentes
        for run in runs:
            run.text = ""
        # Aplicar segmentos
        if segmentos:
            texto_seg, es_bold = segmentos[0]
            runs[0].text = texto_seg
            runs[0].bold = es_bold
            runs[0].font.italic = False
            if font_name:
                runs[0].font.name = font_name
            if font_size:
                runs[0].font.size = font_size
        for i in range(1, len(segmentos)):
            texto_seg, es_bold = segmentos[i]
            nuevo_run = paragraph.add_run(texto_seg)
            nuevo_run.bold = es_bold
            nuevo_run.font.italic = False
            if font_name:
                nuevo_run.font.name = font_name
            if font_size:
                nuevo_run.font.size = font_size
        return

    # Paso 3b: Si no había negrillas, simplemente reemplazar
    if not frases_bold:
        runs[0].text = texto_nuevo
        runs[0].bold = False
        runs[0].font.italic = False
        for run in runs[1:]:
            run.text = ""
        return

    # Paso 4: Dividir texto nuevo en segmentos (bold / normal)
    segmentos = _dividir_por_negrillas(texto_nuevo, frases_bold)

    # Paso 5: Vaciar todos los runs existentes
    for run in runs:
        run.text = ""

    # Paso 6: Usar el primer run para el primer segmento
    if segmentos:
        texto_seg, es_bold = segmentos[0]
        runs[0].text = texto_seg
        runs[0].bold = es_bold
        runs[0].font.italic = False
        if font_name:
            runs[0].font.name = font_name
        if font_size:
            runs[0].font.size = font_size

    # Paso 7: Agregar runs nuevos para los demás segmentos (API segura)
    for i in range(1, len(segmentos)):
        texto_seg, es_bold = segmentos[i]
        nuevo_run = paragraph.add_run(texto_seg)
        nuevo_run.bold = es_bold
        nuevo_run.font.italic = False
        if font_name:
            nuevo_run.font.name = font_name
        if font_size:
            nuevo_run.font.size = font_size


def _autonegrilla_termino_definicion(texto: str) -> str:
    """Si la linea tiene formato de definicion ('Termino: descripcion'),
    envuelve el termino (hasta los dos puntos) en **...** si aun no lo esta.

    Esto garantiza consistencia de negrillas en listas de definiciones aunque
    Claude haya olvidado marcar alguna. Solo aplica cuando:
    - La linea tiene la forma "Palabra(s): descripcion" (primer ':' en las
      primeras ~60 chars).
    - El termino NO esta ya envuelto en **.
    - El termino empieza con letra (no es solo numeros/simbolos).

    No aplica a:
    - Lineas muy cortas (probablemente titulos, no definiciones).
    - Lineas que empiezan con numero + punto (listas numeradas).
    - Lineas sin dos puntos en las primeras chars.
    """
    import re
    stripped = texto.lstrip()
    if not stripped or len(stripped) < 15:
        return texto
    # Saltar listas numeradas "1. ..." o letras "a) ..."
    if re.match(r"^\d{1,2}[\.\)]\s", stripped) or re.match(r"^[a-z][\.\)]\s", stripped):
        return texto
    # Saltar si ya empieza con **
    if stripped.startswith("**"):
        return texto
    # Buscar el primer ':' en los primeros 70 chars
    m = re.match(r"^([A-Za-zÁÉÍÓÚÑáéíóúñ][A-Za-zÁÉÍÓÚÑáéíóúñ0-9\s\-]{1,68}):\s", stripped)
    if not m:
        return texto
    termino = m.group(1)
    # Debe tener al menos una letra y maximo ~8 palabras (para que no matchee
    # oraciones largas con ":" accidental)
    if len(termino.split()) > 8:
        return texto
    # Preservar el leading whitespace que stripped() elimino
    leading = texto[: len(texto) - len(stripped)]
    resto = stripped[len(termino) + 1 :]  # despues de los dos puntos
    return f"{leading}**{termino}:**{resto}"


def _parsear_markdown_bold(texto: str) -> list:
    """Parsea texto con marcadores **bold** y lo divide en segmentos.

    Antes de parsear, aplica auto-negrilla a terminos de definicion
    ("TERMINO: descripcion") para garantizar consistencia.

    Ejemplo:
        '**FUNDACION** cumple con **Decreto 1072 de 2015**'
        -> [('FUNDACION', True), (' cumple con ', False), ('Decreto 1072 de 2015', True)]

    Args:
        texto: Texto con marcadores Markdown-style **texto**

    Returns:
        Lista de tuplas (segmento, es_bold)
    """
    import re
    # Auto-bold del termino antes de ':' si es una linea de definicion
    texto = _autonegrilla_termino_definicion(texto)
    # Regex: **contenido** (no greedy, SIN saltos de linea - un ** nunca debe
    # abarcar mas de una linea; si Claude lo hizo, lo tratamos como invalido)
    patron = re.compile(r'\*\*([^\n]+?)\*\*')
    segmentos = []
    pos = 0
    for m in patron.finditer(texto):
        # Texto antes del match (normal)
        if m.start() > pos:
            segmentos.append((texto[pos:m.start()], False))
        # Contenido entre ** ** (bold)
        segmentos.append((m.group(1), True))
        pos = m.end()
    # Texto restante al final (normal)
    if pos < len(texto):
        segmentos.append((texto[pos:], False))
    # Limpiar asteriscos sueltos (unbalanceados) que sobrevivieron en
    # segmentos no-bold; si quedan **, son basura que romperia el render
    segmentos = [
        (t.replace("**", "") if not b else t, b)
        for t, b in segmentos
    ]
    # Filtrar segmentos vacios
    segmentos = [(t, b) for t, b in segmentos if t]
    return segmentos if segmentos else [(texto.replace("**", ""), False)]


def _dividir_por_negrillas(texto: str, frases_bold: list) -> list:
    """Divide texto en segmentos marcando cuáles deben ir en negrita.

    SOLO bold-ea frases que estaban EN NEGRITA en el original (match flexible).
    NO aplica patrones automáticos de normas/siglas SG-SST para evitar
    sobregeneración de negrillas respecto a la plantilla.

    Args:
        texto: Texto completo del párrafo humanizado
        frases_bold: Lista de frases que estaban en negrita originalmente

    Returns:
        Lista de (texto, es_bold) en orden
    """
    import re

    todos_patrones = []

    # Solo frases del original (con match flexible)
    frases_ordenadas = sorted(frases_bold, key=len, reverse=True)
    for frase in frases_ordenadas:
        # Flexibilizar: permitir variaciones de espacios y puntuación
        patron = re.escape(frase)
        patron = patron.replace(r'\ ', r'[\s,\.]*\s*')  # espacios flexibles
        patron = patron.replace(r'\.', r'\.?')  # puntos opcionales
        todos_patrones.append(patron)

    if not todos_patrones:
        return [(texto, False)]

    # Compilar patrón combinado
    patron_combinado = re.compile(
        '(' + '|'.join(todos_patrones) + ')',
        re.IGNORECASE,
    )

    # Dividir el texto usando el patrón
    partes = patron_combinado.split(texto)

    segmentos = []
    for parte in partes:
        if not parte:
            continue
        es_bold = bool(patron_combinado.fullmatch(parte))
        segmentos.append((parte, es_bold))

    return segmentos if segmentos else [(texto, False)]


# ══════════════════════════════════════════════════════════════════════
#  FUNCIONES DE EXTRACCIÓN (para Claude API y escaneo)
# ══════════════════════════════════════════════════════════════════════

_PATRON_MARCADOR = re.compile(r"\{\{[A-ZÑ_]+\}\}")
MIN_LONGITUD_HUMANIZAR = 30


def _es_titulo_por_patron(texto: str) -> bool:
    """Detecta titulos por patron visual sin depender del estilo de parrafo.

    Un titulo se caracteriza por:
    - Ser relativamente corto (< 120 chars).
    - Empezar con numeracion ('7.', '7.6', '7.6.1', 'I.', 'II.') o ser
      completamente en MAYUSCULAS.
    - NO terminar con punto final (los titulos suelen no tener punto).
    """
    t = texto.strip()
    if not t or len(t) > 120:
        return False
    # Empezar con numeracion tipo "7.", "7.6", "7.6.1", "10.", "I.", "II."
    patron_numeracion = re.compile(
        r"^(?:\d+(?:\.\d+)*\.?\s+|[IVX]+\.\s+)"
    )
    m = patron_numeracion.match(t)
    if m:
        resto = t[m.end():].strip()
        # Si lo que sigue es corto (tipico de titulo) y no termina con punto,
        # es un titulo
        if resto and len(resto) < 100 and not resto.endswith(('.', ':')):
            return True
        # Mayusculas indica titulo fuerte
        if resto and resto.upper() == resto and len(resto) >= 3:
            return True
    # Todo en mayusculas, sin punto final, corto -> titulo
    if t.upper() == t and len(t) >= 6 and not t.endswith('.'):
        # Debe tener al menos una letra
        if re.search(r"[A-ZÁÉÍÓÚÑ]", t):
            return True
    return False


def _es_parrafo_humanizable(texto: str, style_name: str | None = None) -> bool:
    """Determina si un párrafo tiene contenido sustancial para humanizar.

    Excluye titulos (por estilo o por patron), campos vacios, encabezados
    de metadata (Version/Codigo) y parrafos solo con placeholders.
    """
    texto = texto.strip()
    if not texto or len(texto) < MIN_LONGITUD_HUMANIZAR:
        return False
    if re.match(r"^[\s{}\w]*(\{\{[A-ZÑ_]+\}\}[\s{}\w]*)+$", texto):
        return False
    if texto.startswith("_____"):
        return False
    if texto.startswith("Version") or texto.startswith("Codigo"):
        return False
    # Excluir por estilo de parrafo (Heading, Titulo, Ttulo, Subtitulo, TOC)
    if style_name:
        sn = style_name.lower()
        titulo_keywords = ("heading", "titulo", "ttulo", "título",
                           "subtitulo", "subtítulo", "toc", "caption")
        if any(kw in sn for kw in titulo_keywords):
            return False
    # Excluir por patron visual (numeracion + mayusculas + corto, sin punto final)
    if _es_titulo_por_patron(texto):
        return False
    return True


# Titulos de secciones que NO deben humanizarse. Todo el contenido que
# aparezca DESPUES de uno de estos titulos (hasta el final del documento o
# hasta otro titulo de nivel similar) se ignora por Claude.
_SECCIONES_NO_HUMANIZAR = (
    "bibliografia", "bibliografía",
    "referencias bibliograficas", "referencias bibliográficas",
    "documentos relacionados", "documentos de referencia",
    "referencias", "webgrafia", "webgrafía",
    "normograma",  # el normograma se actualiza con su propio flujo
)

# Titulos que marcan el INICIO de una nueva seccion (rompen la zona "no humanizar").
# Si despues de "BIBLIOGRAFIA" aparece p.ej. "CONTROL DE CAMBIOS" o un numeral
# como "11.", significa que la zona no humanizable termino.
_PATRON_TITULO_NUEVO = re.compile(
    r"^(?:\d+\.?\s|[IVX]+\.\s|[A-ZÁÉÍÓÚÑ\s]{6,}$)"
)


def _es_titulo_no_humanizar(texto: str) -> bool:
    """Detecta si un parrafo es un titulo de seccion no humanizable."""
    import unicodedata
    t = texto.strip().lower()
    # Normalizar: quitar acentos y puntuacion/numeros al inicio
    t_norm = "".join(
        c for c in unicodedata.normalize("NFD", t)
        if unicodedata.category(c) != "Mn"
    )
    # Quitar numeracion tipo "10." "10.1" "X." al inicio
    t_norm = re.sub(r"^[\divx]+[\.\s]+", "", t_norm).strip()
    # Debe ser relativamente corto (es un titulo)
    if len(t_norm) > 80:
        return False
    for clave in _SECCIONES_NO_HUMANIZAR:
        if t_norm == clave or t_norm.startswith(clave + " ") or t_norm.endswith(" " + clave):
            return True
    return False


def extraer_secciones_contenido(docx_path: Path, marcadores: dict = None) -> list:
    """Extrae secciones de contenido humanizable para Claude API.

    IMPORTANTE: No expande marcadores de imagen (rutas .jpg/.png/.jpeg) al
    texto que se manda a Claude, porque Claude luego reproduciria la ruta
    como texto visible en el documento humanizado.

    Tampoco humaniza secciones de BIBLIOGRAFIA, REFERENCIAS BIBLIOGRAFICAS
    ni DOCUMENTOS RELACIONADOS: esos contienen listas de normas y referencias
    que deben quedar EXACTAS como estan en la plantilla.
    """
    doc = Document(str(docx_path))
    secciones = []
    seccion_actual = []

    # Pre-filtrar marcadores: saltar los que son rutas de imagen o InlineImage
    marcadores_texto = {}
    if marcadores:
        for m, v in marcadores.items():
            if not isinstance(v, str):
                # Probablemente InlineImage o algo no-texto; saltar
                continue
            # Si parece ruta de imagen, saltar (deja el marcador {{...}} en el
            # texto para que _es_parrafo_humanizable lo descarte como no-humanizable)
            v_low = v.lower().strip().strip('"').strip("'")
            if v_low.endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')):
                continue
            if v_low.startswith(('http://', 'https://')) and any(
                ext in v_low for ext in ('.jpg', '.jpeg', '.png', '.gif')
            ):
                continue
            marcadores_texto[m] = v

    # Estado: True mientras estemos dentro de una seccion no humanizable
    en_zona_no_humanizar = False

    for idx, paragraph in enumerate(doc.paragraphs):
        texto = paragraph.text.strip()
        for m, v in marcadores_texto.items():
            texto = texto.replace("{{" + m + "}}", v)

        # Detectar entrada a seccion no humanizable (bibliografia, etc.)
        if _es_titulo_no_humanizar(texto):
            # Cerrar seccion abierta antes de entrar a la zona
            if seccion_actual:
                secciones.append({
                    "id": len(secciones),
                    "parrafos": seccion_actual,
                    "texto_completo": "\n".join(p["texto"] for p in seccion_actual),
                })
                seccion_actual = []
            en_zona_no_humanizar = True
            continue

        # Si estamos en zona no humanizable: saltar hasta encontrar otra seccion
        # con titulo tipo "11.", "CONTROL DE CAMBIOS", etc.
        if en_zona_no_humanizar:
            es_nuevo_titulo = (
                bool(_PATRON_TITULO_NUEVO.match(texto))
                and len(texto) < 120
                and not _es_titulo_no_humanizar(texto)
            )
            if es_nuevo_titulo:
                en_zona_no_humanizar = False
                # No incluir el titulo mismo en la seccion a humanizar
                continue
            else:
                # Seguir saltando contenido de bibliografia/referencias
                continue

        # Obtener nombre del estilo del parrafo para detectar headings
        try:
            style_name = paragraph.style.name if paragraph.style else None
        except Exception:
            style_name = None

        if _es_parrafo_humanizable(texto, style_name=style_name):
            seccion_actual.append({"idx": idx, "texto": texto})
        else:
            if seccion_actual:
                secciones.append({
                    "id": len(secciones),
                    "parrafos": seccion_actual,
                    "texto_completo": "\n".join(p["texto"] for p in seccion_actual),
                })
                seccion_actual = []

    if seccion_actual:
        secciones.append({
            "id": len(secciones),
            "parrafos": seccion_actual,
            "texto_completo": "\n".join(p["texto"] for p in seccion_actual),
        })

    return secciones


def extraer_texto_docx(docx_path: Path) -> str:
    """Extrae todo el texto de un .docx."""
    doc = Document(str(docx_path))
    textos = []
    for paragraph in doc.paragraphs:
        texto = paragraph.text.strip()
        if texto:
            textos.append(texto)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texto = cell.text.strip()
                if texto:
                    textos.append(texto)
    return "\n".join(textos)


def listar_marcadores_en_docx(docx_path: Path) -> list:
    """Lista todos los marcadores {{...}} encontrados en un documento."""
    doc = Document(str(docx_path))
    marcadores = set()

    for paragraph in doc.paragraphs:
        texto = "".join(run.text for run in paragraph.runs)
        marcadores.update(_PATRON_MARCADOR.findall(texto))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    texto = "".join(run.text for run in paragraph.runs)
                    marcadores.update(_PATRON_MARCADOR.findall(texto))

    try:
        for section in doc.sections:
            for element in [section.header, section.footer]:
                for paragraph in element.paragraphs:
                    texto = "".join(run.text for run in paragraph.runs)
                    marcadores.update(_PATRON_MARCADOR.findall(texto))
                for table in element.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                texto = "".join(run.text for run in paragraph.runs)
                                marcadores.update(_PATRON_MARCADOR.findall(texto))
    except Exception:
        pass

    return sorted(marcadores)
