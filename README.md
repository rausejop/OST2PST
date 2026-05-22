# ost2pst

[![tests](https://github.com/rausejop/OST2PST/actions/workflows/test.yml/badge.svg)](https://github.com/rausejop/OST2PST/actions/workflows/test.yml)
[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-lightgrey.svg)]()

Herramienta CLI para Windows que procesa ficheros **Outlook OST** en dos
modos independientes y completos:

| Modo            | Salida                            | Outlook | Perfil OST | Adjuntos | Recipientes | RTF | Inline cid: | Filtros | Formatos |
|-----------------|-----------------------------------|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| **PST** (default) | un `.pst` Unicode                 | sí | sí | sí | sí | sí | sí | — | — |
| **`--extract`** | árbol de `.eml` o `.mbox` por carpeta | no | no | sí | sí (MAPI) | sí (striprtf) | sí (matching estricto) | folder · exclude · since · until · max-att-mb · clase MAPI · require-date · **limit** · **codepage override** | `eml` / `mbox` / `eml.gz` |

Distribución: un único `dist\ost2pst.exe` (~10 MB) con `VersionInfo`
embebido y `SHA-256` publicado. Sin Python ni dependencias en destino.

---

## Instalación rápida

### Opción A — Descargar el `.exe` compilado

1. Ve a [Releases](https://github.com/rausejop/OST2PST/releases) y descarga
   `ost2pst.exe` (+ `ost2pst.exe.sha256` opcional para verificar).
2. Copia el `.exe` a cualquier carpeta del Windows destino. Listo.

### Opción B — Clonar y compilar

Requiere Python 3.9 (`winget install Python.Python.3.9`).

```powershell
git clone https://github.com/rausejop/OST2PST.git
cd OST2PST
.\build.ps1 -Clean -Test
```

El binario queda en `dist\ost2pst.exe`. El script verifica Python 3.9,
instala dependencias desde el wheel cacheado en `vendor/`, ejecuta la
suite de tests y compila con `PyInstaller`.

---

## Novedades v3.1

- **Formato `--format mbox`**: un fichero `messages.mbox` por carpeta del
  OST (formato mbox estándar, ideal para Thunderbird / mutt / scripts).
- **`--limit N`**: procesa solo los primeros N mensajes — imprescindible
  para iterar rápido sobre OSTs grandes.
- **`--codepage CP`**: override del codepage detectado (1252, 1251, 932…)
  para mailboxes donde la autodetección falla.
- **`--quiet / -q`**: nivel ERROR (silencia warnings, útil para scripts).
- **`chardet`** como último recurso de decodificación.
- **Recorrido iterativo** del árbol de carpetas (inmune a profundidades
  patológicas).
- **Inline cid: matching estricto** (sin falsos positivos por substring).
- **Validación PST antes de Outlook**: ya no espera 30 s al COM Dispatch
  para fallar con "falta --input".
- **Bugfix crítico (recipientes)**: los Cc/Bcc múltiples se cortaban en
  el primer `;`. Detectado por la nueva test suite MAPI.
- **76 tests** (subió de 51) — la suite MAPI cubre los lectores de
  propiedades con mock realista de `record_sets`.
- **Ruff** integrado en CI.

## Funcionalidades clave de la v3.0

- **Cabeceras originales fieles** vía `transport_headers` preservando
  `Received` / `DKIM-Signature` / `Authentication-Results`
  multi-instancia (el chain de relays se mantiene íntegro).
- **Recipientes To/Cc/Bcc** leídos directamente de propiedades MAPI
  (`PR_DISPLAY_TO/CC/BCC`) vía `pypff.record_sets`. Sobreviven cuando
  el `transport_headers` no los trae (típico en OST de Exchange).
- **Cuerpos**: text/plain + text/html + fallback a **RTF** (`striprtf`)
  para correos antiguos. Decodificación con el **codepage del OST**
  como hint.
- **Adjuntos embebidos**, con **dedup de nombres** (`doc.pdf`, `doc (2).pdf`),
  **Content-Type adivinado** por extensión, **inline images** detectadas
  heurísticamente por `cid:`.
- **Adjuntos también a disco** opcionalmente con `--attachments-dir`.
- **Filtros**: `--folder` / `--exclude` (repetibles), `--since` / `--until`
  (ISO 8601), `--max-attachment-mb`, `--require-date`, `--include-non-email`.
- **Clase MAPI**: por defecto solo se procesan `IPM.Note` (emails);
  appointments, contacts, tasks se saltan a no ser que pases
  `--include-non-email`.
- **Salidas**: `--dry-run` (sin escribir), `--compress` (`.eml.gz`),
  `--json` (resumen JSON por stdout), `--errors-log` (JSON Lines).
- **Forense / cadena de custodia**: SHA-256 del OST origen en cada
  ejecución, SHA-256 del `.exe` publicado junto al binario.
- **Rutas largas** Windows (`\\?\` automático >240 chars).
- **UTF-8** en stdout/stderr (acentos correctos en cualquier consola).
- **Modo `--debug`** con `sys.settrace`: timestamp ms, LINE / CALL / RET / EXC.

---

## Estructura del proyecto

```
51_OST2PST/
├── ost2pst.py                  # ~1100 líneas, un único módulo bien seccionado
├── requirements.txt            # libpff-python==20211114, pywin32, pyinstaller, striprtf
├── build.ps1                   # Compila reproducible: -Clean, -Test, -Upx
├── version_info.txt            # AUTO-GENERADO desde __version__ por build.ps1
├── LICENSE                     # MIT
├── CHANGELOG.md                # Historial de versiones
├── .gitignore                  # PyInstaller + Python + *.ost/*.pst
├── .github/workflows/test.yml  # CI: tests + smoke build en Windows
├── vendor/
│   └── libpff_python-…cp39-win_amd64.whl
├── tests/
│   ├── test_helpers.py         # 40 tests unitarios (helpers puros)
│   └── test_build_eml.py       # 11 tests funcionales con mocks de pypff
├── skills/ost-extract/SKILL.md # Agent Skill (spec Anthropic)
├── dist/
│   ├── ost2pst.exe             # Ejecutable portable
│   └── ost2pst.exe.sha256      # Hash para verificación
└── README.md
```

Organización interna de `ost2pst.py`:

```
módulo + __version__
├── dependencias opcionales (pywin32, pypff, striprtf)
├── constantes (MAPI property IDs, exit codes)
├── logging (_MsFormatter con precisión ms)
├── modo --debug (sys.settrace)
├── helpers comunes (sanitize, long-path, mime, cid, sha256, decode)
├── lectura de propiedades MAPI desde pypff
├── ExtractStats (contadores + errors-log JSON Lines)
├── Progress (indicador stderr cuando no -v)
├── ExtractConfig (dataclass agrupa args de --extract)
├── Modo 1: PST (Outlook COM)
├── Modo 2: EML (libpff)
│   ├── _apply_headers
│   ├── _apply_body  (con RTF fallback)
│   ├── _apply_attachments  (con inline cid heurístico)
│   ├── build_eml
│   └── _extract_folder  (con filtros)
└── CLI (argparse con grupos)
```

---

## Cómo se construyó

### Dependencias

| Paquete                     | Para qué                                      |
|-----------------------------|-----------------------------------------------|
| `pywin32`                   | Cliente COM de Outlook (modo PST)             |
| `libpff-python==20211114`   | Lectura nativa de OST (modo `--extract`)      |
| `striprtf`                  | Conversión RTF → texto (fallback de cuerpo)   |
| `pyinstaller`               | Empaquetado a `.exe`                          |
| `pytest` (dev)              | Suite de tests                                |

`libpff-python==20211114` se fija porque es la **última versión con wheel
precompilado para Windows** (`cp39-win_amd64`). El wheel queda cacheado
en `vendor/` para sobrevivir a Python 3.9 EOL.

### Entorno

Build atado a **Python 3.9.13** (`winget install Python.Python.3.9`).

### Compilación reproducible

```powershell
.\build.ps1 -Clean -Test
```

El script:

1. Lee `__version__` de `ost2pst.py`, regenera `version_info.txt`.
2. Verifica Python 3.9.
3. Instala desde `vendor/` local + `requirements.txt`.
4. (`-Test`) ejecuta `pytest tests/` y aborta si falla.
5. Compila con PyInstaller (`--onefile --console --add-data --version-file`).
6. Calcula y publica el SHA-256 del `.exe` en `dist\ost2pst.exe.sha256`.
7. (`-Upx`) si UPX está en PATH, lo aplica para reducir tamaño ~50%.

### Tests

```powershell
py -3.9 -m pytest tests/ -v
```

**51 tests** divididos en:
- `test_helpers.py` (40): sanitización, decodificación (incluido
  codepage hint), MIME, cid extraction, dedup, fechas ISO, long-path,
  clase MAPI, SHA-256, ExtractStats, ExtractConfig.
- `test_build_eml.py` (11): cabeceras Received multi-instancia, RTF
  fallback, multipart alternative, adjuntos, dedup en .eml, inline
  cid con Content-ID, skip por tamaño con contadores correctos.

CI en GitHub Actions ejecuta tests + smoke build en cada push.

---

## Uso

### Ayuda

```cmd
ost2pst.exe                 :: sin args → imprime ayuda y exit 0
ost2pst.exe --help
ost2pst.exe --version
```

### Modo PST

```cmd
ost2pst.exe --list
ost2pst.exe -i "C:\…\cuenta.ost" -o "C:\backup\cuenta.pst"
ost2pst.exe -i "Cuenta Exchange" --by-name -o cuenta.pst -v
ost2pst.exe -i x.ost -o y.pst -f --json
```

### Modo `--extract`

```cmd
:: Extracción completa
ost2pst.exe --extract -i "C:\correo.ost" -o "C:\extraidos" -v

:: Solo Bandeja, año 2024, comprimido
ost2pst.exe --extract -i C:\correo.ost -o C:\extraidos ^
            --folder "Bandeja de entrada" ^
            --since 2024-01-01 --until 2024-12-31 --compress

:: Excluir papelera, adjuntos a disco aparte, JSON al final
ost2pst.exe --extract -i C:\correo.ost -o C:\extraidos ^
            --exclude "Eliminados" ^
            --attachments-dir C:\extraidos\_attachments ^
            --json

:: Procesar también calendarios/contactos/tareas
ost2pst.exe --extract -i C:\correo.ost -o C:\extraidos --include-non-email

:: Limitar adjuntos a 25 MB, log estructurado
ost2pst.exe --extract -i C:\correo.ost -o C:\extraidos ^
            --max-attachment-mb 25 --errors-log errores.jsonl

:: Inventariar sin escribir
ost2pst.exe --extract -i C:\correo.ost -o C:\extraidos --dry-run --json
```

Resumen JSON de ejemplo (con `--json`):

```json
{
  "mode": "eml",
  "ost_path": "C:\\correo.ost",
  "ost_sha256": "a1b2c3...",
  "out_dir": "C:\\extraidos",
  "codepage": 1252,
  "dry_run": false,
  "folders": 14,
  "items": 12453,
  "skipped_by_filter": 230,
  "skipped_by_class": 87,
  "attachments_saved": 3201,
  "attachments_skipped": 4,
  "errors": 0
}
```

---

## Argumentos

**modo de operación** (mutuamente excluyentes)

| Flag             | Descripción                                          |
|------------------|------------------------------------------------------|
| `-x, --extract`  | Extracción a `.eml` con libpff                       |
| `-l, --list`     | Lista stores cargados en Outlook y sale              |

**entrada / salida**

| Flag                    | Descripción                                       |
|-------------------------|---------------------------------------------------|
| `-i OST, --input`       | Ruta OST (o nombre store si `--by-name`)          |
| `-o DEST, --output`     | PST (modo PST) o directorio (modo `--extract`)    |

**opciones modo PST**

| Flag                    | Descripción                                       |
|-------------------------|---------------------------------------------------|
| `--by-name`             | `--input` como nombre de store                    |
| `-f, --force`           | Sobrescribe el PST destino                        |

**filtros modo `--extract`**

| Flag                          | Descripción                                  |
|-------------------------------|----------------------------------------------|
| `--folder NOMBRE` (repetible) | Solo carpetas cuyo nombre contenga NOMBRE    |
| `--exclude NOMBRE` (repetible)| Salta carpetas y subárbol                    |
| `--since YYYY-MM-DD`          | Mensajes con fecha ≥ esta                    |
| `--until YYYY-MM-DD`          | Mensajes con fecha ≤ esta (fin de día)       |
| `--max-attachment-mb MB`      | Salta adjuntos >MB                            |
| `--require-date`              | Excluye mensajes sin fecha cuando hay since/until |
| `--include-non-email`         | Incluye appointments/contacts/tasks          |

**salidas modo `--extract`**

| Flag                    | Descripción                                            |
|-------------------------|--------------------------------------------------------|
| `--attachments-dir DIR` | Vuelca adjuntos también como ficheros sueltos          |
| `--compress`            | `.eml.gz` (gzip nivel 6)                               |
| `--json`                | Resumen final como JSON por stdout                     |
| `--dry-run`             | No escribe, solo cuenta                                |
| `--errors-log PATH`     | Errores como JSON Lines                                |

**logging**

| Flag             | Descripción                                            |
|------------------|--------------------------------------------------------|
| `-v, --verbose`  | Nivel INFO                                             |
| `-d, --debug`    | Traza línea-a-línea (implica `--verbose`)              |

## Códigos de salida

| Code | Significado                                              |
|------|----------------------------------------------------------|
| 0    | OK                                                       |
| 1    | Argumentos / destino inválidos                           |
| 2    | No se pudo iniciar Outlook                               |
| 3    | Store no encontrado por `--by-name`                      |
| 4    | El fichero OST no existe                                 |
| 5    | El OST no está cargado en el perfil de Outlook           |
| 6    | No se pudo crear el PST                                  |
| 7    | No se pudo localizar el PST recién creado                |
| 8    | Completado pero con errores en algunos items             |
| 10   | `pywin32` no disponible                                  |
| 11   | `libpff` no disponible                                   |
| 12   | No se pudo abrir el OST con libpff                       |
| 13   | No se pudo leer la raíz del OST                          |

---

## Skill incluido (Agent Skills)

`skills/ost-extract/SKILL.md` cumple la spec de **Anthropic Agent Skills**.
Búsqueda no-hardcoded del binario: `$env:OST2PST_EXE` → PATH → `dist\`.

```powershell
:: Nivel usuario
Copy-Item -Recurse skills\ost-extract "$env:USERPROFILE\.claude\skills\"

:: Indicar al skill dónde está el .exe
[Environment]::SetEnvironmentVariable("OST2PST_EXE", "$PWD\dist\ost2pst.exe", "User")
```

---

## Distribución

Para llevarlo a otro Windows:

1. Copia `dist\ost2pst.exe` (~10 MB) y opcionalmente
   `dist\ost2pst.exe.sha256` para verificar integridad.
2. Para el modo PST, además Outlook + cuenta del OST en su perfil.

Sin Python ni nada más en destino.

**Verificación de integridad:**

```powershell
$expected = (Get-Content .\dist\ost2pst.exe.sha256).Split()[0]
$actual = (Get-FileHash .\dist\ost2pst.exe -Algorithm SHA256).Hash.ToLower()
if ($expected -eq $actual) { "OK" } else { "ALTERADO" }
```

**Code signing** (Authenticode) **no implementado** — requiere certificado
de pago. El `.exe` sin firmar dispara SmartScreen la primera vez ("Windows
protegió tu PC"). Se ejecuta vía *Más información → Ejecutar de todas
formas*. Para distribución a terceros conviene firmar con certificado
EV / OV de una CA reconocida.

---

## Limitaciones conocidas

- **Modo PST**: el OST debe estar cargado en el perfil de Outlook activo
  (limitación inherente al formato OST de Microsoft).
- **Modo `--extract`**: `libpff` puede fallar con OST muy corruptos o con
  `compressible encryption` antiguo. Los mensajes fallidos se reportan en
  `--errors-log`.
- **Inline images**: heurística por nombre vs cid. Cubre el caso común
  (`image001.png` ↔ `cid:image001@01D…`); cids totalmente desligados
  del nombre quedan como adjunto convencional.
- **Recipientes**: leídos de `PR_DISPLAY_TO/CC/BCC` (string semicolon-
  separated, no estructura individual). Suficiente para forense; para
  manipulación programática del Address Type / propiedades extra de cada
  destinatario habría que iterar `sub_items` (no implementado).
- **Calendarios / contactos / tareas**: se saltan por defecto (clase MAPI
  `IPM.Appointment`, `IPM.Contact`, `IPM.Task`). Con `--include-non-email`
  se procesan como `.eml`, pero el formato resultante no es el ideal
  (un `.ics` o `.vcf` sería más apropiado — fuera de alcance).
- **Python 3.9 EOL** (octubre 2025): wheel cacheado en `vendor/`. Cuando
  `libpff-python` publique wheel para 3.12+ basta bumpear
  `requirements.txt`, `build.ps1` y CI.
- **Paralelismo**: single-thread. Para OST muy grandes (>500k mensajes),
  considerar un wrapper externo que llame al `.exe` por subdirectorios.
- **Resume**: si la extracción falla a mitad, no hay checkpoint para
  reanudar — empieza desde cero.

---

## Contribuir

Issues y pull requests son bienvenidos en
[github.com/rausejop/OST2PST](https://github.com/rausejop/OST2PST).

Antes de abrir un PR:

```powershell
# Lint + tests deben pasar
py -3.9 -m ruff check ost2pst.py tests/
py -3.9 -m pytest tests/ -v

# Build smoke
.\build.ps1 -Clean -Test
```

El workflow `.github/workflows/test.yml` ejecuta los mismos pasos en
cada push. Mantén el coverage: cualquier helper público nuevo debería
tener su test correspondiente en `tests/`.

## Licencia

[MIT](LICENSE) © ost2pst project
