"""Integración con Claude API para humanización de documentos SG-SST."""
import hashlib
import logging
import re
import time

import anthropic
import diskcache

from .config import (
    ANTHROPIC_API_KEY, CACHE_DIR, CACHE_EXPIRY_DAYS,
    CLAUDE_PRECIO_INPUT_MTOK, CLAUDE_PRECIO_OUTPUT_MTOK,
    CLAUDE_MODEL,
)

logger = logging.getLogger(__name__)


# Cache persistente en disco
_cache = diskcache.Cache(str(CACHE_DIR / "claude_responses"), expire=CACHE_EXPIRY_DAYS * 86400)

# Límites para batching
MAX_TOKENS_POR_BATCH = 3000  # ~3000 chars de input por batch para no exceder contexto
MAX_SECCIONES_POR_BATCH = 8


# ---------------------------------------------------------------------------
# Tracking de uso de tokens por sesión
# ---------------------------------------------------------------------------

_uso_sesion = {
    "llamadas": 0,
    "tokens_input": 0,
    "tokens_output": 0,
    "desde_cache": 0,
}


def _registrar_uso(response) -> None:
    """Registra el uso de tokens de una respuesta de Claude API."""
    usage = getattr(response, "usage", None)
    if usage:
        _uso_sesion["llamadas"] += 1
        _uso_sesion["tokens_input"] += getattr(usage, "input_tokens", 0)
        _uso_sesion["tokens_output"] += getattr(usage, "output_tokens", 0)
        logger.debug(
            "API call: %d in / %d out tokens",
            getattr(usage, "input_tokens", 0),
            getattr(usage, "output_tokens", 0),
        )


def obtener_uso_sesion() -> dict:
    """Retorna estadísticas de uso de Claude API en la sesión actual."""
    costo_input = _uso_sesion["tokens_input"] * CLAUDE_PRECIO_INPUT_MTOK / 1_000_000
    costo_output = _uso_sesion["tokens_output"] * CLAUDE_PRECIO_OUTPUT_MTOK / 1_000_000
    return {
        **_uso_sesion,
        "costo_input_usd": round(costo_input, 6),
        "costo_output_usd": round(costo_output, 6),
        "costo_total_usd": round(costo_input + costo_output, 6),
    }


def resetear_uso_sesion() -> None:
    """Resetea los contadores de uso de la sesión."""
    _uso_sesion["llamadas"] = 0
    _uso_sesion["tokens_input"] = 0
    _uso_sesion["tokens_output"] = 0
    _uso_sesion["desde_cache"] = 0


def _cache_key(texto: str, contexto: str, nivel: str) -> str:
    """Genera una clave de caché basada en el contenido y contexto."""
    contenido = f"{nivel}|{contexto}|{texto}"
    return hashlib.sha256(contenido.encode("utf-8")).hexdigest()


def _construir_contexto(cliente: dict) -> str:
    """Construye el contexto del cliente para el prompt de Claude."""
    sedes_info = ""
    for sede in cliente.get("sedes", []):
        sedes_info += f"\n  - {sede.get('nombre', 'N/A')}: {sede.get('direccion', '')}, {sede.get('ciudad', '')}"

    # Determinar COPASST o Vigía
    num_trab = 0
    try:
        num_trab = int(str(
            cliente.get("numero_trabajadores_empresa")
            or cliente.get("num_trabajadores", 0)
        ).strip())
    except ValueError:
        pass
    if num_trab >= 10:
        organo_participacion = (
            "COPASST (Comité Paritario de Seguridad y Salud en el Trabajo) "
            f"- la empresa tiene {num_trab} trabajadores (>= 10). "
            "USAR SIEMPRE 'COPASST' en el texto. "
            "Si el texto original dice 'Vigía de SST' o 'Vigía', REEMPLAZARLO por 'COPASST'."
        )
    else:
        organo_participacion = (
            "Vigía de Seguridad y Salud en el Trabajo "
            f"- la empresa tiene {num_trab} trabajadores (< 10). "
            "USAR SIEMPRE 'Vigía de SST' o 'Vigía de Seguridad y Salud en el Trabajo' en el texto. "
            "Si el texto original dice 'COPASST' o 'Comité Paritario de SST', REEMPLAZARLO por 'Vigía de SST'. "
            "Si menciona 'los miembros del COPASST', cambiar a 'el Vigía de SST' (singular). "
            "Si menciona 'elección de miembros del COPASST', cambiar a 'designación del Vigía de SST'."
        )

    return f"""EMPRESA: {cliente.get('razon_social', 'N/A')}
ACTIVIDAD ECONOMICA: {cliente.get('actividad_economica', 'N/A')}
NIVEL DE RIESGO ARL: {cliente.get('nivel_riesgo', 'N/A')}
NUMERO DE TRABAJADORES: {cliente.get('numero_trabajadores_empresa', 'N/A')}
ORGANO DE PARTICIPACION: {organo_participacion}
ARL: {cliente.get('arl', 'N/A')}
CIUDAD: {cliente.get('ciudad_empresa', 'N/A')}
DESCRIPCION DE LA EMPRESA: {cliente.get('descripcion_empresa', 'N/A')}
SEDES ({len(cliente.get('sedes', []))}): {sedes_info if sedes_info else 'N/A'}"""


def _crear_client():
    """Crea una instancia del cliente Anthropic."""
    return anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


def _llamar_claude(prompt: str, max_tokens: int = 8192, max_reintentos: int = 3):
    """Llama a Claude API con reintentos y backoff exponencial ante rate limits.

    Reintenta hasta max_reintentos veces con esperas 30s → 60s → 120s.
    Relanza la excepción si se agotan los reintentos.
    """
    espera = 30
    for intento in range(max_reintentos):
        try:
            client = _crear_client()
            return client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=max_tokens,
                messages=[{"role": "user", "content": prompt}],
            )
        except anthropic.RateLimitError:
            if intento < max_reintentos - 1:
                logger.warning(
                    "Rate limit. Esperando %ds (intento %d/%d)...",
                    espera, intento + 1, max_reintentos,
                )
                time.sleep(espera)
                espera *= 2
            else:
                raise


# ---------------------------------------------------------------------------
# Humanización por secciones (batch)
# ---------------------------------------------------------------------------

def humanizar_secciones(
    secciones: list[dict],
    cliente: dict,
    nivel: str = "completa",
    tipo_documento: str = "documento",
    nombre_archivo: str = "",
    usar_cache: bool = True,
) -> list[dict]:
    """Humaniza múltiples secciones de un documento agrupándolas en batches.

    Envía grupos de secciones en una sola llamada a Claude para reducir
    el número de API calls. Usa delimitadores para poder separar las respuestas.

    Args:
        secciones: Lista de dicts con "id", "parrafos", "texto_completo"
        cliente: Datos del cliente
        nivel: "completa" o "media"
        tipo_documento: Nombre del tipo (Política, Procedimiento, etc.)
        nombre_archivo: Nombre del archivo para contexto
        usar_cache: Si True, verifica caché antes de llamar

    Returns:
        Las mismas secciones con campo "texto_humanizado" agregado
    """
    if not ANTHROPIC_API_KEY:
        logger.warning("ANTHROPIC_API_KEY no configurada. Saltando humanizacion.")
        return secciones

    if not secciones:
        return secciones

    contexto = _construir_contexto(cliente)

    # Agrupar secciones en batches por tamaño
    batches = _crear_batches(secciones)
    resultados = []

    for batch_idx, batch in enumerate(batches):
        # Verificar caché para cada sección del batch
        batch_sin_cache = []
        for sec in batch:
            key = _cache_key(sec["texto_completo"], contexto, nivel)
            cached = _cache.get(key) if usar_cache else None
            if cached:
                sec["texto_humanizado"] = cached
                resultados.append(sec)
                _uso_sesion["desde_cache"] += 1
            else:
                batch_sin_cache.append(sec)

        if not batch_sin_cache:
            continue

        # Construir prompt con delimitadores para el batch
        prompt = _construir_prompt_batch(
            batch_sin_cache, contexto, nivel, tipo_documento, nombre_archivo
        )

        try:
            response = _llamar_claude(prompt)
            _registrar_uso(response)
            respuesta = response.content[0].text

            # Parsear respuesta separada por delimitadores
            textos_humanizados = _parsear_respuesta_batch(respuesta, len(batch_sin_cache))

            for sec, texto_h in zip(batch_sin_cache, textos_humanizados):
                sec["texto_humanizado"] = texto_h
                # Guardar en caché
                if usar_cache and texto_h:
                    key = _cache_key(sec["texto_completo"], contexto, nivel)
                    _cache.set(key, texto_h)
                resultados.append(sec)

        except Exception as e:
            logger.error("Error en Claude API: %s. Secciones sin humanizar.", e)
            for sec in batch_sin_cache:
                sec["texto_humanizado"] = sec["texto_completo"]
                resultados.append(sec)

    # Reordenar por id original
    resultados.sort(key=lambda s: s["id"])
    return resultados


def _crear_batches(secciones: list[dict]) -> list[list[dict]]:
    """Agrupa secciones en batches respetando límites de tokens."""
    batches = []
    batch_actual = []
    chars_actual = 0

    for sec in secciones:
        largo = len(sec["texto_completo"])
        if (chars_actual + largo > MAX_TOKENS_POR_BATCH
                or len(batch_actual) >= MAX_SECCIONES_POR_BATCH) and batch_actual:
            batches.append(batch_actual)
            batch_actual = []
            chars_actual = 0

        batch_actual.append(sec)
        chars_actual += largo

    if batch_actual:
        batches.append(batch_actual)

    return batches


def _detectar_tema_documento(nombre_archivo: str) -> str:
    """Detecta el tema/comite principal del documento a partir del nombre.

    Acepta nombres simples ("SST-PRO-003.docx") o rutas con carpetas padre
    ("1.1.8 COMITE DE CONVIVENCIA LABORAL/3. PROCEDIMIENTOS/archivo.docx").
    La carpeta padre suele ser la fuente de verdad del tema, asi que se
    analiza TODA la cadena (no solo el nombre del archivo).

    Esto evita que Claude confunda contextos (ej. hablar de COPASST en un
    documento de CCL). Devuelve una descripcion del tema que se usara en
    el prompt para forzar coherencia.
    """
    # Normalizar separadores y pasar a minusculas (incluye carpetas padre)
    texto = nombre_archivo.replace("\\", "/").lower()
    # Eliminar acentos comunes para matching robusto
    import unicodedata
    texto_norm = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )

    # --- Comite de Convivencia Laboral (CCL) ---
    # Palabras que indican acoso laboral/CCL
    ccl_keywords = [
        "ccl", "convivencia", "convivencia laboral",
        "acoso", "acoso laboral",
        "atencion de quejas", "atencion-de-quejas",
        "queja frente al ccl", "quejas frente al ccl",
        "ley 1010", "ley1010",
    ]
    if any(kw in texto_norm for kw in ccl_keywords):
        return ("COMITE DE CONVIVENCIA LABORAL (CCL) - Prevencion de acoso laboral. "
                "Normas aplicables: Ley 1010 de 2006, Resolucion 652 de 2012, "
                "Resolucion 1356 de 2012. NO hablar del COPASST en este documento.")

    # --- COPASST ---
    copasst_keywords = [
        "copasst", "comite paritario",
        "eleccion copasst", "conformacion copasst",
        "vigia de sst", "vigia sst",
    ]
    if any(kw in texto_norm for kw in copasst_keywords):
        return ("COPASST o Vigia de SST - Vigilancia del SG-SST. "
                "Normas aplicables: Decreto 1072 de 2015, Resolucion 0312 de 2019. "
                "NO hablar del Comite de Convivencia Laboral (CCL) en este documento.")

    # --- PESV (Plan Estrategico de Seguridad Vial) ---
    pesv_keywords = ["pesv", "seguridad vial", "plan estrategico"]
    if any(kw in texto_norm for kw in pesv_keywords):
        return ("PLAN ESTRATEGICO DE SEGURIDAD VIAL (PESV). "
                "Normas aplicables: Ley 1503 de 2011, Resolucion 40595 de 2022.")

    # --- Emergencias / Brigada ---
    emerg_keywords = ["brigada", "emergencia", "evacuacion", "incendio"]
    if any(kw in texto_norm for kw in emerg_keywords):
        return ("PLAN DE EMERGENCIAS Y BRIGADA. "
                "Normas aplicables: Decreto 1072 de 2015 articulo 2.2.4.6.25, "
                "Resolucion 0312 de 2019.")

    # --- Medico / Examenes ocupacionales ---
    med_keywords = ["medico", "examenes", "ocupacional", "epidemiologia"]
    if any(kw in texto_norm for kw in med_keywords):
        return ("MEDICINA LABORAL / EXAMENES OCUPACIONALES. "
                "Normas aplicables: Resolucion 2346 de 2007, Decreto 1072 de 2015.")

    # --- Riesgos / Peligros ---
    riesgo_keywords = ["peligro", "riesgo", "matriz", "iper", "identificacion"]
    if any(kw in texto_norm for kw in riesgo_keywords):
        return ("IDENTIFICACION DE PELIGROS Y EVALUACION DE RIESGOS. "
                "Normas aplicables: GTC 45, Decreto 1072 de 2015.")

    # Por defecto: SG-SST general
    return ("SG-SST general. Normas aplicables: Decreto 1072 de 2015, "
            "Resolucion 0312 de 2019.")


def _construir_prompt_batch(
    secciones: list[dict],
    contexto: str,
    nivel: str,
    tipo_documento: str,
    nombre_archivo: str,
) -> str:
    """Construye el prompt para un batch de secciones."""

    # Detectar el tema/comite del documento a partir del nombre del archivo
    tema_especifico = _detectar_tema_documento(nombre_archivo)

    reglas_comites = """REGLAS CRITICAS SOBRE COMITES (NO CONFUNDIR):
- COPASST = Comite Paritario de SST. Se encarga de vigilancia del SG-SST,
  riesgos laborales, accidentes, enfermedades laborales. Normas: Resolucion
  2013 de 1986 (derogada en 2022, ver abajo), Decreto 1072 de 2015, Res 0312 de 2019.
- CCL = Comite de Convivencia Laboral. Se encarga EXCLUSIVAMENTE de prevenir
  y atender acoso laboral (Ley 1010 de 2006). Normas especificas: Resolucion
  652 de 2012, Resolucion 1356 de 2012, Ley 1010 de 2006.
- VIGIA de SST = Equivalente al COPASST para empresas con menos de 10 trabajadores.
- NUNCA mezcles funciones del CCL con el COPASST ni viceversa.
- Si el documento es sobre CCL: habla SOLO de convivencia laboral, acoso, Ley 1010.
- Si el documento es sobre COPASST: habla SOLO de SG-SST, riesgos, Decreto 1072.
- Si el documento es sobre Vigia de SST: usa Vigia, no COPASST."""

    reglas_redaccion = """INSTRUCCIONES DE REDACCION:
- Lenguaje tecnico-profesional, formal pero HUMANO (no robotico ni generico)
- Parrafos fluidos, bien estructurados, con transiciones naturales
- Evitar redundancias, frases vacias y muletillas
- Mantener el tono institucional y el rigor tecnico
- Cada parrafo debe sonar como si lo hubiera escrito un especialista SST con experiencia real

ORTOGRAFIA ESPAÑOL COLOMBIA (OBLIGATORIO — USAR TILDES SIEMPRE):
El texto de salida DEBE estar en español de Colombia con TODAS las tildes
correctas. Aunque el texto original venga sin tildes, tu salida SIEMPRE las
lleva. La falta de tildes es INACEPTABLE.

Palabras con tilde que DEBEN llevarla siempre (ejemplos de uso SG-SST):
- **Resolución** (no "Resolucion")
- **Decreto / Artículo** (no "Articulo")
- **Seguridad y Salud en el Trabajo** (ok sin tilde)
- **Implementación / Aplicación / Identificación / Evaluación / Prevención**
- **Capacitación / Formación / Inducción / Reinducción**
- **Información / Comunicación / Divulgación / Notificación**
- **Organización / Actualización / Planificación / Verificación**
- **Gestión / Revisión / Supervisión / Administración**
- **Técnico / Técnica / Metodología / Epidemiológico / Biomecánico**
- **Análisis / Diagnóstico / Estadístico / Periódico**
- **Ergonomía / Higiene / Psicológico / Médico / Clínico**
- **Vigía / Comité / Reunión / Sesión**
- **Electrónico / Automático / Número**
- Monosilabos con tilde diacritica: **sí** (afirmativo), **él** (pronombre),
  **tú** (pronombre), **más** (cantidad), **té**, **sé**, **dé**, **mí**
- Interrogativos/exclamativos con tilde: **qué, cómo, cuándo, dónde, quién,
  cuánto, cuál, por qué**

Palabras que NO llevan tilde (para que no las pongas de mas):
- este, esta, estos, estas (demostrativos sin tilde desde 2010)
- solo (adverbio o adjetivo, sin tilde desde 2010)
- fue, fui, dio, vio (monosilabos sin tilde)
- ti (pronombre, sin tilde)

Otras reglas ortograficas:
- Usar comillas españolas «» o dobles rectas "", no simples '
- Signos de apertura en preguntas/exclamaciones: ¿que haces? / ¡cuidado!
- Guion largo — para incisos (no doble guion -- ni guion corto -)
- Numeros: usar punto para miles (1.000) y coma para decimales (1,5)"""

    reglas_negrillas = """NEGRILLAS - FORMATO OBLIGATORIO **texto**:
Usa DOS asteriscos **asi** para marcar en negrilla estos terminos cada vez que aparezcan:
- El nombre/razon social de la empresa (ej: **FUNDACION CRISTOVISION**)
- El nombre de la Consultora SST (ej: **MEDINA SST EMPRESARIAL S.A.S**)
- El nombre del Representante Legal (ej: **Rolfi Rozo Villamizar**)
- El nombre del Responsable del SG-SST (ej: **Pablo Alexander Medina Pedraza**)
- "Vigia de Seguridad y Salud en el Trabajo" / "COPASST" / "CCL"
- "SG-SST" (siempre que aparezca esta sigla)
- Numeros de NIT, cedulas, licencias y resoluciones (ej: **830114869-4**, **8866 del 30 de abril de 2021**)
- Titulos de articulos y secciones normativas clave
- Referencias normativas: **Decreto 1072 de 2015**, **Resolucion 0312 de 2019**, **Ley 1562 de 2012**, etc.

REGLA CRITICA - LISTAS DE DEFINICIONES:
Cuando una linea tenga el formato "TERMINO: descripcion" (termino seguido de
dos puntos y luego la definicion), DEBES marcar el TERMINO COMPLETO en
negrilla, incluyendo los dos puntos. SIN EXCEPCIONES. TODOS los terminos
de una lista de definiciones llevan negrilla — no puede haber unos con y
otros sin negrilla en la misma lista.

EJEMPLOS CORRECTOS:
  **Movimiento repetitivo:** Se caracteriza por ciclos de trabajo cortos...
  **Peligro:** Se define como toda fuente, situacion o acto con potencial...
  **Postura:** Se define como la relacion de las diferentes partes del cuerpo...
  **Posturas forzadas:** Cuando se adoptan posturas por fuera de los angulos...
  **Puesto de trabajo:** Emplazamiento fisico donde el trabajador desarrolla...
  **Riesgo:** Combinacion de la probabilidad de que ocurra un evento...

EJEMPLO INCORRECTO (mezcla de negrillas, NO HACER):
  **Movimiento repetitivo:** ciclos cortos...   <-- bien
  Postura: relacion de partes del cuerpo...     <-- MAL, falta la negrilla

IMPORTANTE: NO pongas en negrilla parrafos enteros; solo los terminos clave
dentro de cada oracion. El formato **texto** se convertira automaticamente
en negrilla real al insertar el texto en el Word.
IMPORTANTISIMO: Los marcadores **texto** NUNCA pueden abrir en una linea y
cerrar en otra. Si necesitas negrilla multi-linea, cierra al final de cada
linea y vuelve a abrir al inicio de la siguiente."""

    restricciones_absolutas = """RESTRICCIONES ABSOLUTAS (INVIOLABLES):
- NO resumir ni eliminar informacion: cada idea del original debe estar en tu respuesta.
- NO fusionar parrafos NI definiciones: si recibes N lineas separadas por \\n,
  DEBES devolver EXACTAMENTE N lineas separadas por \\n. Ni una mas, ni una menos.
- CRITICO para listas de DEFINICIONES (items tipo "TERMINO EN MAYUSCULAS: descripcion"):
  cada definicion DEBE quedar en su propia linea, SEPARADA por \\n de la siguiente.
  NUNCA unir dos definiciones en un solo parrafo continuo.
  Ejemplo CORRECTO (3 definiciones, 3 lineas):
    ACTIVIDAD FISICA: Todo movimiento corporal voluntario...\\n
    EJERCICIO FISICO: Actividad fisica programada y estructurada...\\n
    ERGONOMIA: Ciencia que estudia la adaptacion del entorno...
  Ejemplo INCORRECTO (fusionadas en un solo parrafo):
    "La actividad fisica se define como... El ejercicio fisico es... La ergonomia..."
- Devolver EXACTAMENTE el mismo numero de lineas (\\n) que recibas en cada seccion.
- NO agregues lineas extra ni quites lineas; respeta la numeracion (1., 2., 3., ...).
- Si un item es numerado (1., 2., 3., ...), mantenlo como un item separado en su propia linea.
- Un marcador **negrilla** NUNCA puede abarcar mas de una linea. Cierra antes del \\n.
- NO uses listas con viñetas nuevas; conserva los numeros y estructura original.
- Si no puedes mejorar una linea, devuelvela IGUAL (no la borres)."""

    if nivel == "completa":
        instrucciones = f"""Eres un experto en Seguridad y Salud en el Trabajo (SST) en Colombia,
con conocimiento profundo del Decreto 1072 de 2015, la Resolucion 0312 de 2019,
la Ley 1010 de 2006 (acoso laboral) y la Resolucion 652 de 2012 (CCL).

CONTEXTO DE LA EMPRESA:
{contexto}

DOCUMENTO: {nombre_archivo}
TIPO: {tipo_documento}
TEMA ESPECIFICO DETECTADO: {tema_especifico}

{reglas_comites}

{reglas_redaccion}

{reglas_negrillas}

{restricciones_absolutas}

INSTRUCCIONES:
Vas a recibir {len(secciones)} secciones de texto de este documento.
Para CADA seccion, reescribe el contenido para que sea:
1. Especifico para la actividad economica y nivel de riesgo de esta empresa
2. Con redaccion profesional pero humanizada, no generica ni robotica
3. Adaptado al contexto real de la empresa (tipo de trabajo, riesgos del sector)
4. COHERENTE con el TEMA ESPECIFICO DEL DOCUMENTO ({tema_especifico})
5. Cumpliendo con la normatividad colombiana vigente aplicable al tema
6. Manteniendo la misma cantidad de parrafos (lineas) que el original
7. Aplicando negrillas **asi** a los terminos clave indicados arriba

REGLAS IMPORTANTES:
- NO cambies datos del encabezado (razon social, NIT, version, fechas)
- NO inventes datos especificos (nombres, cifras, fechas)
- NO mezcles temas: si es CCL NO hables del COPASST ni viceversa
- SI mejora la redaccion para que sea clara, fluida y contextualizada
- SI adapta riesgos, medidas de control y actividades al sector de la empresa
- MANTIENE la misma cantidad de lineas/parrafos en cada seccion
- REGLA COPASST vs VIGIA (CRITICA - aplica aunque el texto original diga lo contrario):
  * Si la empresa tiene COPASST (>= 10 trabajadores): USA "COPASST" en todo el texto.
    Si el original dice "Vigia" -> REEMPLAZALO por "COPASST".
  * Si la empresa tiene Vigia (< 10 trabajadores): USA "Vigia de SST" en todo el texto.
    Si el original dice "COPASST" o "Comite Paritario" -> REEMPLAZALO por "Vigia de SST".
    Si dice "miembros del COPASST" (plural) -> usa "el Vigia de SST" (singular unipersonal).
    Si dice "eleccion/votacion del COPASST" -> usa "designacion del Vigia de SST".
  * NUNCA uses ambos terminos mezclados en el mismo documento.
  * El contexto de la empresa en la seccion ORGANO DE PARTICIPACION arriba indica
    cual usar; esa decision IGNORA lo que diga el texto original.
- Cada seccion se separa con el delimitador ---SECCION---

Responde con SOLO el texto mejorado de cada seccion, separado por ---SECCION---
No incluyas explicaciones, titulos de seccion ni comentarios."""

    else:  # media
        instrucciones = f"""Eres un experto en Seguridad y Salud en el Trabajo (SST) en Colombia.

CONTEXTO DE LA EMPRESA:
{contexto}

DOCUMENTO: {nombre_archivo}
TIPO: {tipo_documento}
TEMA ESPECIFICO DETECTADO: {tema_especifico}

{reglas_comites}

{reglas_redaccion}

{reglas_negrillas}

{restricciones_absolutas}

INSTRUCCIONES:
Vas a recibir {len(secciones)} secciones de texto. Para CADA seccion:
1. Corrige errores gramaticales o de redaccion
2. Mejora la claridad y profesionalismo del texto
3. Adapta minimamente al contexto de la empresa si es muy generico
4. MANTEN coherencia con el TEMA ESPECIFICO del documento ({tema_especifico})
5. NO cambies la estructura, datos del encabezado ni contenido sustancial
6. NO mezcles temas (si es CCL no hables del COPASST ni viceversa)
7. Aplica negrillas **asi** a los terminos clave indicados arriba
8. MANTIENE la misma cantidad de lineas/parrafos
8. Usa CORRECTAMENTE el organo de participacion: COPASST o Vigia de SST

Responde con SOLO el texto mejorado de cada seccion, separado por ---SECCION---"""

    # Agregar las secciones al prompt
    secciones_texto = ""
    for i, sec in enumerate(secciones):
        secciones_texto += f"\n---SECCION---\n{sec['texto_completo']}"

    return f"""{instrucciones}

TEXTO ORIGINAL:
{secciones_texto}
---FIN---"""


# ---------------------------------------------------------------------------
# Post-proceso ortografico: corrige tildes que Claude pudo omitir
# ---------------------------------------------------------------------------

# Diccionario de correcciones comunes: sin_tilde -> con_tilde.
# Solo palabras de alta frecuencia en documentos SG-SST. Orden importa
# (las mas largas primero para no romper palabras mas cortas que contienen
# una mas larga).
_CORRECCIONES_TILDES = [
    # Raices largas primero (para no chocar con sufijos)
    ("Epidemiologico", "Epidemiológico"),
    ("epidemiologico", "epidemiológico"),
    ("Biomecanico", "Biomecánico"),
    ("biomecanico", "biomecánico"),
    ("Implementacion", "Implementación"),
    ("implementacion", "implementación"),
    ("Identificacion", "Identificación"),
    ("identificacion", "identificación"),
    ("Planificacion", "Planificación"),
    ("planificacion", "planificación"),
    ("Administracion", "Administración"),
    ("administracion", "administración"),
    ("Organizacion", "Organización"),
    ("organizacion", "organización"),
    ("Actualizacion", "Actualización"),
    ("actualizacion", "actualización"),
    ("Verificacion", "Verificación"),
    ("verificacion", "verificación"),
    ("Comunicacion", "Comunicación"),
    ("comunicacion", "comunicación"),
    ("Divulgacion", "Divulgación"),
    ("divulgacion", "divulgación"),
    ("Notificacion", "Notificación"),
    ("notificacion", "notificación"),
    ("Informacion", "Información"),
    ("informacion", "información"),
    ("Capacitacion", "Capacitación"),
    ("capacitacion", "capacitación"),
    ("Reinduccion", "Reinducción"),
    ("reinduccion", "reinducción"),
    ("Induccion", "Inducción"),
    ("induccion", "inducción"),
    ("Formacion", "Formación"),
    ("formacion", "formación"),
    ("Evaluacion", "Evaluación"),
    ("evaluacion", "evaluación"),
    ("Prevencion", "Prevención"),
    ("prevencion", "prevención"),
    ("Aplicacion", "Aplicación"),
    ("aplicacion", "aplicación"),
    ("Investigacion", "Investigación"),
    ("investigacion", "investigación"),
    ("Supervision", "Supervisión"),
    ("supervision", "supervisión"),
    ("Resolucion", "Resolución"),
    ("resolucion", "resolución"),
    ("Revision", "Revisión"),
    ("revision", "revisión"),
    ("Gestion", "Gestión"),
    ("gestion", "gestión"),
    ("Articulo", "Artículo"),
    ("articulo", "artículo"),
    ("Metodologia", "Metodología"),
    ("metodologia", "metodología"),
    ("Ergonomia", "Ergonomía"),
    ("ergonomia", "ergonomía"),
    ("Periodica", "Periódica"),
    ("periodica", "periódica"),
    ("Periodico", "Periódico"),
    ("periodico", "periódico"),
    ("Diagnostico", "Diagnóstico"),
    ("diagnostico", "diagnóstico"),
    ("Estadistico", "Estadístico"),
    ("estadistico", "estadístico"),
    ("Analisis", "Análisis"),
    ("analisis", "análisis"),
    ("Tecnico", "Técnico"),
    ("tecnico", "técnico"),
    ("Tecnica", "Técnica"),
    ("tecnica", "técnica"),
    ("Medico", "Médico"),
    ("medico", "médico"),
    ("Clinico", "Clínico"),
    ("clinico", "clínico"),
    ("Psicologico", "Psicológico"),
    ("psicologico", "psicológico"),
    ("Automatico", "Automático"),
    ("automatico", "automático"),
    ("Electronico", "Electrónico"),
    ("electronico", "electrónico"),
    ("Comite", "Comité"),
    ("comite", "comité"),
    ("Vigia", "Vigía"),
    ("vigia", "vigía"),
    ("Reunion", "Reunión"),
    ("reunion", "reunión"),
    ("Sesion", "Sesión"),
    ("sesion", "sesión"),
    ("Numero", "Número"),
    ("numero", "número"),
    ("Decreto Unico", "Decreto Único"),
    ("decreto unico", "decreto único"),
]

# Compilar patrones una sola vez al cargar el módulo (evita recompilar en cada llamada)
_TILDES_COMPILED = [
    (re.compile(r"\b" + re.escape(sin) + r"\b"), con)
    for sin, con in _CORRECCIONES_TILDES
]


def _corregir_tildes(texto: str) -> str:
    """Aplica correcciones ortograficas automaticas sobre palabras comunes
    del vocabulario SG-SST que Claude pudo devolver sin tilde.

    Usa word-boundary para no afectar partes de palabras mas largas.
    """
    resultado = texto
    for patron, con in _TILDES_COMPILED:
        resultado = patron.sub(con, resultado)
    return resultado


def _parsear_respuesta_batch(respuesta: str, num_secciones: int) -> list[str]:
    """Parsea la respuesta de Claude separada por delimitadores ---SECCION---."""
    # Limpiar respuesta
    respuesta = respuesta.strip()

    # Eliminar ---FIN--- si Claude lo repite
    respuesta = respuesta.replace("---FIN---", "").strip()

    # Separar por delimitador
    partes = respuesta.split("---SECCION---")
    partes = [p.strip() for p in partes if p.strip()]

    # Si Claude devolvió una de más (a veces el primer split genera un vacío)
    if len(partes) == num_secciones + 1 and not partes[0]:
        partes = partes[1:]

    if len(partes) != num_secciones:
        logger.warning(
            "Respuesta Claude: %d secciones recibidas, %d esperadas. "
            "El parseo puede ser impreciso.", len(partes), num_secciones,
        )

    # Si Claude devolvió la cantidad correcta de secciones
    if len(partes) == num_secciones:
        # Aplicar red de seguridad ortografica sobre cada seccion
        return [_corregir_tildes(p) for p in partes]

    # Fallback: si las cantidades no coinciden, devolver lo que hay
    # rellenando con strings vacíos si faltan
    while len(partes) < num_secciones:
        partes.append("")

    return [_corregir_tildes(p) if p else p for p in partes[:num_secciones]]


# ---------------------------------------------------------------------------
# Humanización de texto simple (para un bloque completo)
# ---------------------------------------------------------------------------

def humanizar_texto(
    texto: str,
    cliente: dict,
    nivel: str = "completa",
    tipo_documento: str = "documento",
    usar_cache: bool = True,
) -> str:
    """Humaniza un bloque de texto usando Claude API (para uso directo)."""
    if nivel == "ninguna" or not texto.strip():
        return texto

    if not ANTHROPIC_API_KEY:
        logger.warning("ANTHROPIC_API_KEY no configurada. Saltando humanizacion.")
        return texto

    contexto = _construir_contexto(cliente)

    if usar_cache:
        key = _cache_key(texto, contexto, nivel)
        cached = _cache.get(key)
        if cached:
            return cached

    if nivel == "completa":
        instrucciones = """Reescribe el siguiente texto de un documento de SG-SST para que sea:
1. Especifico para la actividad economica y nivel de riesgo de esta empresa
2. Con redaccion profesional pero humanizada, no generica ni robotica
3. Adaptado al contexto real de la empresa
4. Cumpliendo con la normatividad colombiana vigente
5. Sin inventar datos especificos (nombres, fechas, cifras)
6. Usando correctamente COPASST o Vigia segun el organo de participacion del contexto
Devuelve SOLO el texto mejorado."""
    else:
        instrucciones = """Mejora ligeramente la redaccion de este texto de SG-SST:
1. Corrige errores gramaticales
2. Mejora claridad y profesionalismo
3. NO cambies estructura ni contenido sustancial
Devuelve SOLO el texto mejorado."""

    prompt = f"""Eres un experto en SST en Colombia (Decreto 1072/2015, Res. 0312/2019).

CONTEXTO: {contexto}
TIPO: {tipo_documento}

{instrucciones}

TEXTO:
---
{texto}
---"""

    try:
        response = _llamar_claude(prompt, max_tokens=4096)
        _registrar_uso(response)
        resultado = response.content[0].text

        if usar_cache:
            _cache.set(key, resultado)

        return resultado

    except Exception as e:
        logger.error("Error en Claude API: %s. Usando texto original.", e)
        return texto


# ---------------------------------------------------------------------------
# Normograma
# ---------------------------------------------------------------------------

def actualizar_normograma(normograma_texto: str, cliente: dict, usar_cache: bool = True) -> str:
    """Usa Claude para validar y actualizar el normograma eliminando normas derogadas."""
    if not ANTHROPIC_API_KEY:
        logger.warning("ANTHROPIC_API_KEY no configurada. Saltando actualizacion de normograma.")
        return normograma_texto

    contexto = _construir_contexto(cliente)

    if usar_cache:
        key = _cache_key(normograma_texto, contexto, "normograma")
        cached = _cache.get(key)
        if cached:
            return cached

    prompt = f"""Eres un experto en legislacion colombiana de Seguridad y Salud en el Trabajo.

CONTEXTO DE LA EMPRESA:
{contexto}

TAREA: Revisa el siguiente normograma de SST y:
1. Identifica normas que hayan sido DEROGADAS o MODIFICADAS a la fecha (abril 2026)
2. Para normas derogadas: indica cual norma la derogo y eliminala
3. Para normas modificadas: actualiza la referencia
4. Agrega normas relevantes que falten segun la actividad economica de la empresa
5. Asegurate de que el Decreto 1072 de 2015 y la Resolucion 0312 de 2019 esten incluidos
6. Incluye normas especificas del sector segun la actividad economica
7. Incluye resoluciones recientes relevantes (2023-2026) si aplican

IMPORTANTE: Devuelve el normograma actualizado en el MISMO FORMATO que el original.
Solo el contenido del normograma, sin explicaciones ni comentarios.

NORMOGRAMA ACTUAL:
---
{normograma_texto}
---"""

    try:
        response = _llamar_claude(prompt)
        _registrar_uso(response)
        resultado = response.content[0].text

        if usar_cache:
            _cache.set(key, resultado)

        return resultado

    except Exception as e:
        logger.error("Error actualizando normograma: %s", e)
        return normograma_texto


# ---------------------------------------------------------------------------
# Utilidades de caché
# ---------------------------------------------------------------------------

def limpiar_cache():
    """Limpia toda la caché de respuestas de Claude."""
    _cache.clear()
    logger.info("Cache limpiada.")


def estadisticas_cache() -> dict:
    """Retorna estadísticas de uso de la caché."""
    return {
        "entradas": len(_cache),
        "tamaño_bytes": _cache.volume(),
        "directorio": str(CACHE_DIR / "claude_responses"),
    }
