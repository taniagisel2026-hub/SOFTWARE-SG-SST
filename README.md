# SG-SST Generator

**Herramienta CLI** (y, desde la Fase 4, también **API REST**) para generar Sistemas de Gestión de Seguridad y Salud en el Trabajo (SG-SST) personalizados a partir de una plantilla base, datos del cliente y manuales de funciones. Cumple Decreto 1072/2015 y Resolución 0312/2019 (Colombia).

Diseñada para uso interno de **MEDINA SST EMPRESARIAL S.A.S** y consultores asociados.

---

## ¿Qué hace?

1. Lee un **YAML del cliente** con sus datos (razón social, NIT, ARL, sedes, color corporativo, logo, etc.)
2. Toma la **plantilla SG-SST 2026** y reemplaza los marcadores `{{...}}` con los datos del cliente
3. **Humaniza con Claude API** los documentos según su tipo (políticas, procedimientos, programas)
4. **Adapta automáticamente** COPASST ↔ Vigía según el número de trabajadores
5. **Aplica colores corporativos** de la empresa
6. **Genera la Matriz de Peligros IPVR (SST-MTZ-012)** según los cargos del cliente
7. **Genera la Matriz de Peligros Viales (PESV-MTZ-002)** si el cliente solicita PESV
8. **Sube todo a Google Drive** del cliente y deja el listado maestro con enlaces clicables a cada documento

---

## Requisitos

- **Python 3.12+** (probado en 3.12)
- **Cuenta de Anthropic** con API key (para humanización con Claude)
- **Credenciales de Google Drive** (`credentials.json`) — OAuth 2.0
- **Plantilla SG-SST 2026** completa (carpeta en la raíz del proyecto)
- **Sistema operativo**: Windows (la herramienta se ha probado en Windows; debería funcionar en Linux/Mac con ajustes menores)

---

## Instalación

```cmd
# 1. Clonar o copiar el proyecto a tu máquina
cd "C:\Users\<tu_usuario>\Desktop\SOFTWARE SG-SST"

# 2. (Opcional pero recomendado) Crear entorno virtual
python -m venv .venv
.venv\Scripts\activate

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Configurar variables de entorno
# Crear archivo .env en la raíz con:
#    ANTHROPIC_API_KEY=sk-ant-...
#    CLAUDE_MODEL=claude-haiku-4-5-20251001
```

Coloca tus credenciales de Google en la raíz como `credentials.json` (descárgalo de la consola de Google Cloud).

---

## Quick Start: generar un cliente nuevo en 5 pasos

### 1. Copia y edita el YAML

```cmd
copy clientes_config\ejemplo_cliente.yaml clientes_config\MI_CLIENTE.yaml
notepad clientes_config\MI_CLIENTE.yaml
```

Llena todos los campos del cliente. Detalles en **[docs/GUIA_CONSULTOR.md](docs/GUIA_CONSULTOR.md)**.

### 2. (Opcional) Prepara carpeta con manuales de funciones

```
C:\Users\bts19\Downloads\MANUALES_MI_CLIENTE\
├── 1. GERENTE GENERAL.docx
├── 2. CONTADOR.docx
├── 3. CONDUCTOR.docx
└── ...
```

Y agrega al YAML:

```yaml
manuales_cargos_dir: "C:/Users/bts19/Downloads/MANUALES_MI_CLIENTE"
```

### 3. Generar el SG-SST

```cmd
python cli.py generar clientes_config\MI_CLIENTE.yaml
```

El resultado queda en `clientes/<RAZON_SOCIAL>/` con la estructura PHVA completa.

### 4. Subir a Google Drive del cliente

```cmd
python cli.py subir-drive clientes_config\MI_CLIENTE.yaml
```

Primera vez: abre el navegador para autorizar el acceso a Drive del cliente.

### 5. Verificar los links del listado maestro

```cmd
python verificar_links_listado.py "clientes/<RAZON_SOCIAL>"
```

---

## Comandos CLI principales

| Comando | Para qué |
|---|---|
| `python cli.py generar <yaml>` | Genera el SG-SST completo del cliente |
| `python cli.py subir-drive <yaml>` | Sube el SG-SST a Drive e inyecta links en SST-MTZ-005 |
| `python cli.py verificar <yaml>` | Valida que el YAML tenga todos los campos necesarios |
| `python cli.py probar-doc <documento.docx> <yaml>` | Procesa un solo documento (útil para depurar) |
| `python cli.py preview <documento.docx> <yaml>` | Muestra cómo quedaría un documento sin generarlo |
| `python cli.py auth-drive --cliente-id <NIT>` | Re-autoriza acceso a Drive del cliente |
| `python cli.py limpiar-cache` | Limpia el caché de Claude |
| `python cli.py normas` | Lista las correcciones legales configuradas (norma derogada → norma vigente, y desde qué año aplica cada una) |
| `python cli.py cache` | Muestra estadísticas de la caché de respuestas de Claude |
| `python cli.py colores` | Muestra las paletas de colores disponibles o prueba una paleta |
| `python cli.py escanear <yaml>` | Escanea la plantilla y muestra todos los marcadores `{{...}}` encontrados |
| `python cli.py info` | Muestra información del sistema y configuración actual |
| `python verificar_links_listado.py <ruta_cliente>` | Verifica que cada link de la matriz apunte al documento correcto |
| `python auditar_normativa.py [--json <archivo>]` | Extrae y audita todas las normas citadas en la plantilla maestra |

Para ver todas las opciones de un comando:

```cmd
python cli.py <comando> --help
```

---

## API REST (Fase 4)

Segunda puerta de entrada a las mismas funciones del CLI, expuesta por HTTP — pensada como paso previo al portal de clientes (Fase 5). No reemplaza al CLI: ambos llaman a las mismas funciones de `engine/`.

```cmd
uvicorn api:app --reload --port 8000
```

Documentación interactiva (Swagger) en `http://localhost:8000/docs`.

Todas las rutas salvo `/health` y `/auth/drive/callback` requieren el header `Authorization: Bearer <API_AUTH_TOKEN>` (configurado en `.env`, ver `.env.example`).

| Método | Ruta | Para qué |
|---|---|---|
| `GET` | `/health` | Chequeo de vida, sin autenticación |
| `GET` | `/normas` | Lista las correcciones legales configuradas |
| `POST` | `/clientes/validar` | Valida un YAML de cliente |
| `POST` | `/clientes/verificar` | Equivalente a `cli.py verificar`, responde JSON |
| `POST` | `/clientes/{id}/generar` | Genera el SG-SST en segundo plano, responde `202` con un `job_id` |
| `POST` | `/clientes/{id}/subir-drive` | Sube a Drive en segundo plano, responde `202` con un `job_id` |
| `GET` | `/jobs/{job_id}` | Consulta el estado de un trabajo (`queued`/`running`/`done`/`failed`) |
| `GET` | `/clientes/{id}/jobs` | Lista los trabajos de un cliente |
| `GET` | `/auth/drive/{id}/start` | Inicia el login de Drive por navegador (flujo web) |
| `GET` | `/auth/drive/callback` | Google redirige aquí tras el consentimiento |

Las generaciones (2-40 min) y subidas a Drive corren en segundo plano; su estado se guarda en `jobs.db` (SQLite) y sobrevive un reinicio del servidor — un trabajo interrumpido a medias queda marcado `failed` con un mensaje claro en vez de perderse.

**Login de Drive por navegador**: requiere crear un Client ID de tipo **"Web application"** en Google Cloud Console (distinto del `credentials.json` de escritorio que usa el CLI) con el redirect URI `http://localhost:8000/auth/drive/callback` registrado, y descargarlo como `credentials_web.json` en la raíz del proyecto. Sin ese archivo, `/auth/drive/{id}/start` falla con un mensaje explicando el paso pendiente. El login del CLI (`auth-drive`, `subir-drive`) no depende de esto y sigue funcionando igual que siempre.

---

## Estructura del proyecto

```
SOFTWARE SG-SST/
├── cli.py                          # Punto de entrada CLI
├── api.py                          # Punto de entrada API REST (Fase 4)
├── api_auth.py                     # Autenticación por API key de la API REST
├── verificar_links_listado.py      # Verificador de links del listado maestro
├── limpiar_cache.py                # Limpieza de caché de Claude
├── auditar_normativa.py            # Extrae y audita las normas citadas en la plantilla maestra
├── inventario_normas.json          # Inventario de normas generado por auditar_normativa.py
├── requirements.txt
├── credentials.json                # OAuth Google, flujo de escritorio/CLI (no subir a git)
├── credentials_web.json            # OAuth Google, flujo web/API (no subir a git, paso manual)
├── jobs.db                         # Estado de trabajos en segundo plano de la API (no subir a git)
├── .env                            # API keys y tokens (no subir a git)
│
├── engine/                         # Motor principal
│   ├── generador.py                # Orquestador del pipeline
│   ├── cliente.py                  # Carga YAML y arma contexto
│   ├── config.py                   # Configuración y tipos de documento
│   ├── validacion.py               # Validación/verificación de cliente, compartida por CLI y API
│   ├── docx_engine.py              # Procesa archivos Word
│   ├── xlsx_engine.py              # Procesa archivos Excel
│   ├── claude_service.py           # Integración Claude API
│   ├── drive_service.py            # Autenticación y subida a Google Drive (flujo escritorio + web)
│   ├── drive_upload.py             # Logica de subida a Drive, compartida por CLI y API
│   ├── jobs.py                     # Persistencia SQLite de trabajos en segundo plano (API)
│   ├── listado_maestro.py          # Llena SST-MTZ-005 con links
│   ├── matriz_documental.py        # Matriz documental del SG-SST
│   ├── matriz_peligros.py          # Genera Matriz IPVR (SST-MTZ-012)
│   ├── matriz_peligros_viales.py   # Genera Matriz Vial (PESV-MTZ-002)
│   ├── pesv_cleanup.py             # Limpia hojas PESV si no aplica
│   ├── normativa.py                # Correcciones de normograma, filtradas por el año del SG-SST del cliente
│   ├── colores.py                  # Aplica colores corporativos
│   ├── auditoria.py                # Reporte final
│   └── utils.py
│
├── Plantilla SG-SST 2026/          # Plantilla base (no modificar)
├── clientes_config/                # YAML por cliente
│   └── ejemplo_cliente.yaml
├── clientes/                       # SG-SST generados (output)
├── tokens/                         # Tokens OAuth por cliente
└── .cache/                         # Caché de respuestas Claude
```

---

## Documentación adicional

- **[docs/GUIA_CONSULTOR.md](docs/GUIA_CONSULTOR.md)** — Cómo usar la herramienta paso a paso (para consultores SST sin experiencia técnica)
- **[docs/ARQUITECTURA.md](docs/ARQUITECTURA.md)** — Arquitectura técnica del sistema (para mantenimiento y extensión)
- **[docs/CHECKLIST_PRUEBAS_CLIENTE.md](docs/CHECKLIST_PRUEBAS_CLIENTE.md)** — Lista de verificación al generar un cliente nuevo
- **[docs/BUGS.md](docs/BUGS.md)** — Registro y plantilla para reportar bugs
- **[docs/CHANGELOG.md](docs/CHANGELOG.md)** — Historial de cambios y versiones

## Tests

```cmd
python -m pytest tests/
```

93 tests automatizados que validan los puntos críticos del sistema. Ver `tests/` para detalle.

---

## Soporte

Para reportar bugs o sugerencias, abre un issue interno o contacta al equipo técnico de Medina SST Empresarial.

---

## Licencia

Software propietario de **MEDINA SST EMPRESARIAL S.A.S** — NIT 900.988.068-5. Uso interno y bajo contrato con clientes autorizados.
