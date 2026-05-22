# Changelog

Todas las versiones publicadas y sus cambios. Sigue el formato [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/).

## [3.1.0] - 2026-05-18

### Añadido
- `--format mbox`: alternativa a `.eml`, un fichero `messages.mbox` por
  carpeta del OST (formato mbox estándar con `mailbox.mboxMessage`).
- `--limit N`: procesa solo los primeros N mensajes (iteración rápida en
  OSTs grandes).
- `--codepage CP`: override del codepage detectado (`get_ascii_codepage`)
  para mailboxes en idiomas con codepage erróneo.
- `--quiet / -q`: nivel ERROR solo (silencia warnings, útil para scripting).
- `chardet` como último recurso en `_decode_body` antes del fallback
  `latin-1`.
- Aviso en `LOG.info` cuando el modo `--extract` salta items no-email
  por defecto (transparencia del cambio de comportamiento v3.0).
- Test suite con 76 tests (subió de 51):
  - `tests/test_mapi.py` con 16 tests: mock realista de `record_sets`
    y `record_entry` cubriendo formato corto y full-tag de `entry_type`,
    propiedades faltantes, errores en pypff (robustez), recipientes,
    sender, message-id, clase MAPI.
  - Tests adicionales en `test_helpers.py` y `test_build_eml.py` para
    los nuevos contadores, decoding por codepage, matching estricto
    de cid, etc.
- `pyproject.toml` con metadata, deps, ruff config y pytest config.
- `ruff check` integrado en CI (`.github/workflows/test.yml`).
- Smoke test del `.exe` compilado en CI (valida exit codes esperados).

### Cambiado
- **Recorrido del árbol iterativo** (stack-based) en lugar de recursivo;
  inmune a `RecursionError` aunque la jerarquía del OST sea patológica.
- **Matching de inline `cid:` estricto**: solo igualdad exacta del
  `basename` del adjunto con la parte local del cid. Antes el substring
  matching producía falsos positivos (`cid:abc` → `abcdef.png`).
- **Conteo de carpetas correcto**: `stats.folders` solo incrementa
  cuando la carpeta se procesa (antes inflaba con `--folder`).
- **Validación PST antes de Outlook**: `run_pst_export()` valida `--input`
  / `--output` / `--force` / destino existente ANTES de arrancar Outlook.
  Antes esperaba 30 s al Dispatch para fallar con "falta --input".
- **`--attachments-dir` salta inline images** por defecto (no duplica las
  imágenes de firma como ficheros sueltos).
- **Contadores de stats granulares**: `skipped_date_range`,
  `skipped_no_date`, `skipped_class`, `skipped_limit` (antes todo iba a
  `skipped_filter`).
- **`errors-log` incluye timestamp** (`"ts"`) en cada entrada JSON.

### Corregido
- **Bug en recipientes Cc/Bcc múltiples**: las cadenas `PR_DISPLAY_*`
  separan con `;` pero `EmailMessage` con policy estricta interpretaba
  el primer `;` como fin de la dirección. Se convierte `;` → `,` al
  asignar la cabecera. **Detectado por los nuevos tests MAPI**.

### Limitaciones documentadas
- Adjuntos siguen cargándose en memoria; `--max-attachment-mb` es el
  workaround para casos extremos.
- Code signing (Authenticode) sigue sin implementar.
- Sin resume desde checkpoint.

## [3.0.0] - 2026-05-18

### Añadido
- Lectura de recipientes (To/Cc/Bcc) desde propiedades MAPI vía pypff
  record sets (PR_DISPLAY_TO/CC/BCC).
- Fallback a cuerpo RTF (vía `striprtf`) si el mensaje no tiene plain ni html.
- Detección de la clase MAPI (`IPM.Note`, `IPM.Appointment`, …): por defecto
  el modo `--extract` salta calendarios, contactos y tareas. Override con
  `--include-non-email`.
- Detección automática del codepage del OST (`pst.get_ascii_codepage()`) y
  uso como hint en la decodificación de cuerpos.
- Flag `--require-date` para excluir mensajes sin fecha cuando se usan
  `--since`/`--until`.
- Flag `--attachments-dir` para volcar adjuntos también como ficheros sueltos.
- Flag `--compress` que escribe `.eml.gz` en lugar de `.eml` (gzip nivel 6).
- Flag `--json` que emite el resumen final como JSON por stdout.
- SHA-256 del OST origen calculado y reportado en el resumen
  (chain of custody forense).
- SHA-256 del `.exe` generado y guardado en `dist\ost2pst.exe.sha256`.
- LICENSE (MIT), CHANGELOG.md, .gitignore.
- GitHub Actions workflow (`.github/workflows/test.yml`) para ejecutar tests
  en cada push.
- Tests funcionales de `build_eml()` con mocks de pypff (11 tests nuevos,
  total 51).

### Cambiado
- `build_eml()` descompuesto en `_apply_headers`, `_apply_body`,
  `_apply_attachments` (cada uno con una responsabilidad clara).
- Parámetros de `_extract_folder` agrupados en una dataclass `ExtractConfig`.
- `ExtractStats` con contadores separados: `attachments_saved`,
  `attachments_skipped`, `skipped_filter`, `skipped_class`.
- Logging con precisión de milisegundos (`_MsFormatter`).
- `version_info.txt` se auto-genera en `build.ps1` desde `__version__` para
  evitar drift de versión.

### Corregido
- Código muerto: línea `stats.bytes_attachments += ... * 0 + 0`.
- Doble check redundante de `max_attachment_bytes` (uno en build_eml, otro en
  `_attachment_bytes`).
- Conteo de adjuntos confundía "leídos" con "guardados" cuando había skips
  por tamaño.

### Limitaciones documentadas
- Code signing del `.exe` no implementado (requiere certificado Authenticode).
- Sin paralelismo (single-thread). Para OST muy grandes (>500k msgs)
  considerar un wrapper externo.
- Sin resume desde checkpoint: si falla a mitad, hay que reiniciar.

## [2.1.0] - 2026-05-18

### Añadido
- `Received`, `DKIM-Signature`, `Authentication-Results` multi-instancia
  preservados en el `.eml`.
- Detección heurística de imágenes inline (`cid:`) — nombre del adjunto vs
  referencias en el HTML.
- Dedup de nombres de adjuntos (`doc.pdf`, `doc (2).pdf`).
- Mutex `--extract` vs `--list` con `add_mutually_exclusive_group`.
- `--max-attachment-mb` para saltar adjuntos grandes.
- `--folder`, `--exclude` (repetibles), `--since`, `--until`.
- Indicador de progreso en consola cuando `-v` está desactivado.
- `--dry-run` y `--errors-log` (JSON Lines).
- Wheel de libpff cacheado en `vendor/`.
- `VersionInfo` embebido en el `.exe`.
- Tests unitarios (27).
- Soporte de rutas largas Windows (`\\?\`).
- Reconfiguración UTF-8 de stdout/stderr.
- HTML-only ya no genera un `text/plain` vacío adicional.
- Date faithfulness: `format_datetime` para datetimes aware.

### Cambiado
- SKILL.md ya no hardcodea la ruta del `.exe` — busca `$OST2PST_EXE`, PATH,
  `dist\`.

## [2.0.0] - 2026-05-18

### Añadido
- Adjuntos embebidos en el `.eml` (modo `--extract`) vía
  `email.message.EmailMessage`.
- `requirements.txt`, `build.ps1`, `tests/`.
- Logging estructurado con `logging.Logger`.
- Argparse argument groups.
- Exit codes nombrados.

### Cambiado
- `--alobruto` renombrado a `--extract` / `-x`.
- Easter egg `--Mike` eliminado.

## [1.x] - inicial

- Modo PST (Outlook COM).
- Modo `--alobruto` con libpff (sin adjuntos).
- Modo `--debug` con `sys.settrace`.
