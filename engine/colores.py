"""Personalización de colores corporativos en documentos Word y Excel.

Reemplaza la paleta vino tinto de las plantillas base por el color
corporativo del cliente. Se configura en el YAML del cliente con:

    color_corporativo: "azul"       # nombre predefinido
    color_corporativo: "#2E75B6"    # o hex directo

Paletas predefinidas: azul, verde, morado, naranja, gris
"""
import re
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from rich.console import Console

console = Console(highlight=False)


# ══════════════════════════════════════════════════════════════════════
#  PALETA BASE (vino tinto de las plantillas originales)
# ══════════════════════════════════════════════════════════════════════

# Colores que se encuentran en las plantillas y deben cambiarse
_COLORES_VINOTINTO = {
    # Fondos oscuros (shading fill)
    "660000",   # Fondo principal de encabezados (1045 apariciones)
    "632423",   # Fondo alterno oscuro
    "990000",   # Fondo alterno
    "4A1A1A",   # Fondo muy oscuro
    # Fondos claros (filas alternas)
    "F4CCCC",   # Rosado claro
    "E6BABA",   # Rosado claro 2
    "F5E3E3",   # Rosado claro 3
    "EBC7C7",   # Rosado claro 4
    "E6B8AF",   # Rosado claro 5 (Excel, 10 archivos)
    # Intermedios
    "CE7574",   # Rojo medio
    "DD7E6B",   # Salmón intermedio (Excel, 2 archivos)
    "C00000",   # Rojo corporativo (fuente + fondos)
}

# Nota: los colores de fuente vinotinto (C00000, 632423, 990000) se cambian
# automáticamente al reemplazar fondos, ya que comparten los mismos hex.
# No se necesita un set separado para fuentes.


# ══════════════════════════════════════════════════════════════════════
#  PALETAS PREDEFINIDAS
#  Cada paleta define: oscuro, medio, claro, muy_claro, intermedio
# ══════════════════════════════════════════════════════════════════════

PALETAS = {
    "azul": {
        "nombre": "Azul Corporativo",
        "oscuro":     "1F3864",   # Navy oscuro (encabezados)
        "oscuro2":    "2E4A7A",   # Azul oscuro medio
        "muy_oscuro": "1B2D4F",   # Navy muy oscuro
        "medio":      "2E75B6",   # Azul medio (títulos)
        "intermedio": "6B9BD2",   # Azul medio claro
        "claro1":     "D6E4F0",   # Azul claro
        "claro2":     "C5D9F1",   # Azul claro 2
        "claro3":     "E9F0F8",   # Azul muy claro
        "claro4":     "D0E0F0",   # Azul claro 4
    },
    "verde": {
        "nombre": "Verde Corporativo",
        "oscuro":     "1D4D2C",
        "oscuro2":    "2E7A3E",
        "muy_oscuro": "153D20",
        "medio":      "2E8B57",
        "intermedio": "6BBF8A",
        "claro1":     "D6F0E0",
        "claro2":     "C5E8D1",
        "claro3":     "E9F8EE",
        "claro4":     "D0F0DC",
    },
    "morado": {
        "nombre": "Morado Corporativo",
        "oscuro":     "3B1F64",
        "oscuro2":    "4A2E7A",
        "muy_oscuro": "2D1B4F",
        "medio":      "7B4FB6",
        "intermedio": "9B7BD2",
        "claro1":     "E4D6F0",
        "claro2":     "D9C5F1",
        "claro3":     "F0E9F8",
        "claro4":     "E0D0F0",
    },
    "naranja": {
        "nombre": "Naranja Corporativo",
        "oscuro":     "7A3D00",
        "oscuro2":    "995200",
        "muy_oscuro": "5C2E00",
        "medio":      "E67E22",
        "intermedio": "F0A050",
        "claro1":     "FDE8D0",
        "claro2":     "FCE0C0",
        "claro3":     "FEF4E8",
        "claro4":     "FDEBD6",
    },
    "gris": {
        "nombre": "Gris Corporativo",
        "oscuro":     "2D2D2D",
        "oscuro2":    "404040",
        "muy_oscuro": "1A1A1A",
        "medio":      "5B5B5B",
        "intermedio": "8C8C8C",
        "claro1":     "E0E0E0",
        "claro2":     "D5D5D5",
        "claro3":     "F0F0F0",
        "claro4":     "E8E8E8",
    },
}


def _construir_mapa_colores(paleta: dict) -> dict[str, str]:
    """Construye el mapa de reemplazo vinotinto → nuevo color.

    Returns:
        Dict de {color_viejo_hex: color_nuevo_hex} (sin #, uppercase)
    """
    return {
        # Fondos oscuros (encabezados de tablas)
        "660000": paleta["oscuro"],
        "632423": paleta["oscuro2"],
        "990000": paleta["oscuro2"],
        "4A1A1A": paleta["muy_oscuro"],
        # Fondos claros (filas alternas)
        "F4CCCC": paleta["claro1"],
        "E6BABA": paleta["claro2"],
        "F5E3E3": paleta["claro3"],
        "EBC7C7": paleta["claro4"],
        "E6B8AF": paleta["claro2"],   # Rosado claro Excel
        # Fondos intermedios
        "CE7574": paleta["intermedio"],
        "DD7E6B": paleta["intermedio"],  # Salmón intermedio Excel
        "C00000": paleta["medio"],
    }


def resolver_paleta(color_corporativo: str) -> dict | None:
    """Resuelve la paleta a usar según el valor del YAML.

    Args:
        color_corporativo: nombre predefinido ("azul") o hex ("#2E75B6")

    Returns:
        Paleta dict o None si no aplica cambio de color
    """
    if not color_corporativo:
        return None

    valor = color_corporativo.strip().lower().lstrip('#')

    # Nombre predefinido
    if valor in PALETAS:
        return PALETAS[valor]

    # Hex directo: generar paleta automática basada en ese color
    hex_match = re.match(r"^#?([0-9A-Fa-f]{6})$", color_corporativo.strip())
    if hex_match:
        hex_base = hex_match.group(1).upper()
        return _generar_paleta_desde_hex(hex_base)

    return None


def _generar_paleta_desde_hex(hex_color: str) -> dict:
    """Genera una paleta completa a partir de un solo color hex.

    Si el color del usuario YA es oscuro (luminancia < 100), se usa tal cual
    como color principal (oscuro/medio). Solo se generan variantes claras.
    Si es claro/medio, se oscurecen las variantes para los fondos de tablas.
    """
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)

    def _clamp(v):
        return max(0, min(255, int(v)))

    def _hex(r_, g_, b_):
        return f"{_clamp(r_):02X}{_clamp(g_):02X}{_clamp(b_):02X}"

    luminancia = 0.299 * r + 0.587 * g + 0.114 * b

    if luminancia < 100:
        # Color YA oscuro: usar tal cual en encabezados de tablas y titulos.
        return {
            "nombre":     f"Personalizado #{hex_color}",
            "oscuro":     hex_color,                                # ← tal cual
            "oscuro2":    hex_color,                                # ← tal cual
            "muy_oscuro": _hex(r * 0.75, g * 0.75, b * 0.75),       # un poco mas oscuro
            "medio":      hex_color,                                # ← tal cual
            "intermedio": _hex(r * 0.7 + 77, g * 0.7 + 77, b * 0.7 + 77),
            "claro1":     _hex(r * 0.15 + 217, g * 0.15 + 217, b * 0.15 + 217),
            "claro2":     _hex(r * 0.20 + 200, g * 0.20 + 200, b * 0.20 + 200),
            "claro3":     _hex(r * 0.08 + 232, g * 0.08 + 232, b * 0.08 + 232),
            "claro4":     _hex(r * 0.12 + 210, g * 0.12 + 210, b * 0.12 + 210),
        }
    else:
        # Color claro/medio: oscurecer para los encabezados de tablas
        return {
            "nombre":     f"Personalizado #{hex_color}",
            "oscuro":     _hex(r * 0.45, g * 0.45, b * 0.45),
            "oscuro2":    _hex(r * 0.60, g * 0.60, b * 0.60),
            "muy_oscuro": _hex(r * 0.35, g * 0.35, b * 0.35),
            "medio":      hex_color,
            "intermedio": _hex(r * 0.7 + 77, g * 0.7 + 77, b * 0.7 + 77),
            "claro1":     _hex(r * 0.15 + 217, g * 0.15 + 217, b * 0.15 + 217),
            "claro2":     _hex(r * 0.20 + 200, g * 0.20 + 200, b * 0.20 + 200),
            "claro3":     _hex(r * 0.08 + 232, g * 0.08 + 232, b * 0.08 + 232),
            "claro4":     _hex(r * 0.12 + 210, g * 0.12 + 210, b * 0.12 + 210),
        }


# ══════════════════════════════════════════════════════════════════════
#  APLICAR COLORES A DOCUMENTOS WORD (.docx)
# ══════════════════════════════════════════════════════════════════════

def _es_color_oscuro(hex_color: str) -> bool:
    """Determina si un color hex es oscuro (necesita fuente blanca)."""
    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        # Luminancia relativa: <128 = oscuro
        luminancia = 0.299 * r + 0.587 * g + 0.114 * b
        return luminancia < 128
    except (ValueError, IndexError):
        return False


def aplicar_colores_docx(docx_path: str, mapa_colores: dict[str, str]) -> int:
    """Reemplaza colores de FONDO en un .docx y ajusta fuentes en tablas.

    Paso 1 (XML): Cambia fondos de celdas (w:shd fill) al nuevo color.
    Paso 2 (python-docx): En celdas de tabla con fondo oscuro pone fuente
    blanca; en celdas con fondo claro pone fuente negra.
    NO toca el color de fuente fuera de tablas (títulos, marcadores, texto).

    Args:
        docx_path: Ruta al archivo .docx (se modifica in-place)
        mapa_colores: Dict {viejo_hex: nuevo_hex} sin #

    Returns:
        Número de reemplazos realizados
    """
    cambios = [0]
    temp_path = docx_path + ".tmp_color_corp"

    colores_regex = {}
    for viejo, nuevo in mapa_colores.items():
        viejo_up = viejo.upper()
        nuevo_up = nuevo.upper()
        if viejo_up != nuevo_up:
            colores_regex[viejo_up] = nuevo_up

    if not colores_regex:
        return 0

    # --- Paso 1: Reemplazo de fondos en XML ---

    # Pre-compilar regex una sola vez (no dentro del loop de archivos XML)
    _patrones_docx = []
    for viejo, nuevo in colores_regex.items():
        _patrones_docx.append((
            re.compile(r'(w:fill\s*=\s*")' + viejo + r'(")', re.IGNORECASE),
            re.compile(r'(<a:srgbClr\s+val\s*=\s*")' + viejo + r'(")', re.IGNORECASE),
            nuevo,
        ))

    def _reemplazar_colores_xml(xml_text: str) -> str:
        resultado = xml_text
        for patron_fill, patron_srgb, nuevo in _patrones_docx:
            resultado, n = patron_fill.subn(r'\g<1>' + nuevo + r'\2', resultado)
            cambios[0] += n
            resultado, n = patron_srgb.subn(r'\g<1>' + nuevo + r'\2', resultado)
            cambios[0] += n
        return resultado

    try:
        with ZipFile(docx_path, "r") as zin, ZipFile(temp_path, "w", compression=ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    try:
                        text = data.decode("utf-8", errors="replace")
                        text = _reemplazar_colores_xml(text)
                        data = text.encode("utf-8")
                    except Exception as e:
                        console.print(f"   [yellow]Aviso colores DOCX ({item.filename}): {e}[/yellow]")
                zout.writestr(item, data)

        import os
        import time
        # Reintentos para tolerar OneDrive sincronizando en ese momento
        for intento in range(5):
            try:
                os.replace(temp_path, docx_path)
                break
            except PermissionError:
                if intento < 4:
                    time.sleep(0.5 * (intento + 1))  # 0.5s, 1s, 1.5s, 2s
                else:
                    raise
    except Exception as e:
        import os
        if os.path.exists(temp_path):
            os.remove(temp_path)
        console.print(f"   [yellow]Aviso procesando DOCX: {e}[/yellow]")
        return cambios[0]

    # --- Paso 2: Ajustar fuentes en celdas de tabla con fondo oscuro ---
    # Se hace directamente en XML para no depender de python-docx
    # (python-docx no siempre ve el shd en celdas fusionadas/complejas)
    _ajustar_fuentes_xml(docx_path, cambios)

    return cambios[0]


def _ajustar_fuentes_xml(docx_path: str, cambios: list):
    """Ajusta color de fuente en celdas con fondo oscuro, via XML directo.

    Busca cada <w:tc> que contiene <w:shd ... w:fill="XXXXXX"/>.
    Si el fill es oscuro, pone w:color val="FFFFFF" en todos los <w:rPr>.
    Opera sobre el XML crudo, no python-docx (evita bugs de parsing).
    """
    import re as _re
    temp_path = docx_path + ".tmp_font_fix"

    # Regex para encontrar celdas de tabla con su contenido completo
    # Cada celda es <w:tc>...</w:tc>
    tc_pattern = _re.compile(r'(<w:tc\b[^>]*>)(.*?)(</w:tc>)', _re.DOTALL)
    # Fill dentro de shd
    fill_pattern = _re.compile(r'w:fill="([A-Fa-f0-9]{6})"')
    # Run properties existentes
    rpr_pattern = _re.compile(r'(<w:rPr>)(.*?)(</w:rPr>)', _re.DOTALL)
    # Color de fuente existente dentro de rPr
    color_pattern = _re.compile(r'<w:color\s+w:val="[^"]*"/>')
    # Runs sin rPr: <w:r><w:t> o <w:r ><w:t>
    run_sin_rpr = _re.compile(r'(<w:r\b[^>]*>)(?!<w:rPr>)(<w:t[ >])')

    try:
        with ZipFile(docx_path, "r") as zin, ZipFile(temp_path, "w", compression=ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    text = data.decode("utf-8", errors="replace")

                    def _procesar_celda(match):
                        tc_open = match.group(1)
                        tc_body = match.group(2)
                        tc_close = match.group(3)

                        # Buscar fill en esta celda
                        fill_match = fill_pattern.search(tc_body[:500])  # shd está al inicio
                        if not fill_match:
                            return match.group(0)

                        fill_hex = fill_match.group(1)
                        if not _es_color_oscuro(fill_hex):
                            return match.group(0)

                        # Fondo oscuro: poner fuente blanca
                        color_tag = '<w:color w:val="FFFFFF"/>'
                        nuevo_body = tc_body

                        # En rPr existentes: reemplazar o agregar w:color
                        def _fix_rpr(m):
                            rpr_open = m.group(1)
                            rpr_content = m.group(2)
                            rpr_close = m.group(3)
                            # Quitar color existente
                            rpr_content = color_pattern.sub('', rpr_content)
                            cambios[0] += 1
                            return rpr_open + color_tag + rpr_content + rpr_close

                        nuevo_body = rpr_pattern.sub(_fix_rpr, nuevo_body)

                        # Runs sin rPr: agregar rPr con color
                        def _add_rpr(m):
                            cambios[0] += 1
                            return m.group(1) + '<w:rPr>' + color_tag + '</w:rPr>' + m.group(2)

                        nuevo_body = run_sin_rpr.sub(_add_rpr, nuevo_body)

                        return tc_open + nuevo_body + tc_close

                    text = tc_pattern.sub(_procesar_celda, text)
                    data = text.encode("utf-8")

                zout.writestr(item, data)

        import os
        import time
        # Reintentos para tolerar OneDrive sincronizando en ese momento
        for intento in range(5):
            try:
                os.replace(temp_path, docx_path)
                break
            except PermissionError:
                if intento < 4:
                    time.sleep(0.5 * (intento + 1))  # 0.5s, 1s, 1.5s, 2s
                else:
                    raise
    except Exception as e:
        import os
        if os.path.exists(temp_path):
            os.remove(temp_path)
        console.print(f"   [yellow]Aviso ajustando fuentes: {e}[/yellow]")


# ══════════════════════════════════════════════════════════════════════
#  APLICAR COLORES A DOCUMENTOS EXCEL (.xlsx)
# ══════════════════════════════════════════════════════════════════════

def aplicar_colores_xlsx(xlsx_path: str, mapa_colores: dict[str, str]) -> int:
    """Reemplaza colores en un .xlsx directamente en el XML.

    Los .xlsx también son ZIPs con XML. Modifica estilos y celdas.

    Args:
        xlsx_path: Ruta al archivo .xlsx (se modifica in-place)
        mapa_colores: Dict {viejo_hex: nuevo_hex} sin #

    Returns:
        Número de reemplazos realizados
    """
    cambios = [0]
    temp_path = xlsx_path + ".tmp_color_corp"

    colores_regex = {}
    for viejo, nuevo in mapa_colores.items():
        viejo_up = viejo.upper()
        nuevo_up = nuevo.upper()
        if viejo_up != nuevo_up:
            colores_regex[viejo_up] = nuevo_up

    if not colores_regex:
        return 0

    # Archivos de Excel donde SÍ hay colores de celdas/fuentes
    # Excluimos drawings/ (contienen GUIDs que parecen colores hex)
    _XLSX_XML_RELEVANTES = {"xl/styles.xml", "xl/sharedStrings.xml"}

    def _es_xml_con_colores(filename: str) -> bool:
        """Determina si un archivo XML de xlsx puede contener colores."""
        if filename in _XLSX_XML_RELEVANTES:
            return True
        # Hojas de cálculo: xl/worksheets/sheet1.xml, etc.
        if filename.startswith("xl/worksheets/") and filename.endswith(".xml"):
            return True
        # Tema: xl/theme/theme1.xml
        if filename.startswith("xl/theme/") and filename.endswith(".xml"):
            return True
        return False

    # Pre-compilar regex una sola vez para XLSX
    _patrones_xlsx = []
    for viejo, nuevo in colores_regex.items():
        _patrones_xlsx.append((
            re.compile(r'(<(?:fgColor|bgColor)\s+[^>]*rgb\s*=\s*"FF)' + viejo + r'(")', re.IGNORECASE),
            re.compile(r'(<(?:fgColor|bgColor)\s+[^>]*rgb\s*=\s*")' + viejo + r'(")', re.IGNORECASE),
            nuevo,
        ))

    def _reemplazar_colores_xlsx_xml(xml_text: str) -> str:
        """Reemplaza SOLO colores de fondo de celdas en Excel.

        Tags que SÍ cambia:
          - <fgColor rgb="FFXXXXXX"/>  (color de relleno de celda)
          - <bgColor rgb="FFXXXXXX"/>  (color de fondo de celda)

        Tags que NO toca:
          - <color rgb="FFXXXXXX"/>    (color de fuente)
        """
        resultado = xml_text
        for patron_con_ff, patron_sin_ff, nuevo in _patrones_xlsx:
            resultado, n = patron_con_ff.subn(r'\g<1>' + nuevo + r'\2', resultado)
            cambios[0] += n
            resultado, n = patron_sin_ff.subn(r'\g<1>' + nuevo + r'\2', resultado)
            cambios[0] += n
        return resultado

    try:
        with ZipFile(xlsx_path, "r") as zin, ZipFile(temp_path, "w", compression=ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                # Solo procesar XMLs con colores (no drawings con GUIDs)
                if _es_xml_con_colores(item.filename):
                    try:
                        text = data.decode("utf-8", errors="replace")
                        text = _reemplazar_colores_xlsx_xml(text)
                        data = text.encode("utf-8")
                    except Exception as e:
                        console.print(f"   [yellow]Aviso colores XLSX ({item.filename}): {e}[/yellow]")
                zout.writestr(item, data)

        import os
        import time
        # Reintentos para tolerar OneDrive sincronizando
        for intento in range(5):
            try:
                os.replace(temp_path, xlsx_path)
                break
            except PermissionError:
                if intento < 4:
                    time.sleep(0.5 * (intento + 1))
                else:
                    raise
    except Exception as e:
        import os
        if os.path.exists(temp_path):
            os.remove(temp_path)
        console.print(f"   [yellow]Aviso procesando XLSX: {e}[/yellow]")

    return cambios[0]


# ══════════════════════════════════════════════════════════════════════
#  FUNCIÓN PRINCIPAL: aplicar a todos los documentos generados
# ══════════════════════════════════════════════════════════════════════

def aplicar_colores_corporativos(output_base: Path, color_corporativo: str) -> dict:
    """Aplica colores corporativos a TODOS los documentos generados.

    Args:
        output_base: Directorio raíz del SG-SST generado
        color_corporativo: nombre ("azul") o hex ("#2E75B6")

    Returns:
        Dict con estadísticas {docx_cambios, xlsx_cambios, archivos_procesados}
    """
    paleta = resolver_paleta(color_corporativo)
    if paleta is None:
        return {"docx_cambios": 0, "xlsx_cambios": 0, "archivos": 0, "paleta": None}

    mapa = _construir_mapa_colores(paleta)
    stats = {"docx_cambios": 0, "xlsx_cambios": 0, "archivos": 0, "paleta": paleta["nombre"]}

    for archivo in output_base.rglob("*"):
        if not archivo.is_file() or archivo.name.startswith("~$"):
            continue

        ext = archivo.suffix.lower()
        try:
            if ext == ".docx":
                cambios = aplicar_colores_docx(str(archivo), mapa)
                if cambios > 0:
                    stats["docx_cambios"] += cambios
                    stats["archivos"] += 1
            elif ext == ".xlsx":
                cambios = aplicar_colores_xlsx(str(archivo), mapa)
                if cambios > 0:
                    stats["xlsx_cambios"] += cambios
                    stats["archivos"] += 1
        except Exception as e:
            console.print(f"   [red]Error aplicando colores en {archivo.name}: {e}[/red]")

    return stats


def listar_paletas() -> list[dict]:
    """Retorna las paletas disponibles (para CLI)."""
    return [
        {"nombre": key, "descripcion": val["nombre"], "color_medio": val["medio"]}
        for key, val in PALETAS.items()
    ]
