---
name: ost-extract
description: Extracts emails from a Microsoft Outlook OST file using libpff (no Outlook required, no credentials of the original user needed). Output as individual .eml files OR a single .mbox per folder. Includes attachments (with dedup), inline images (cid: matching), recipients To/Cc/Bcc read directly from MAPI properties, RTF body fallback for legacy messages, charset detection via the OST codepage + chardet. Filters by folder include/exclude, date range, attachment size, message class (default emails only; opt-in for appointments/contacts/tasks via --include-non-email), and N-message limit. Side outputs include attachments-to-disk, JSON summary, JSON-lines error log with timestamps, dry-run, gzip compression. SHA-256 of source OST reported for chain of custody. Works on orphan OSTs recovered from backups or other users. Use this skill when the user provides an OST file and asks to extract its messages, dump emails to disk, read OST contents without Outlook, recover mail from an orphan OST, convert OST to .eml/.mbox, archive an OST, or inventory an OST. Trigger phrases include "extraer emails de OST", "leer fichero OST", "OST a EML", "OST a MBOX", "extract OST", "dump OST messages", "OST without Outlook", "OST sin Outlook", "archivar OST", "inventariar OST".
---

# OST Email Extraction (libpff, no Outlook required)

Drives the `ost2pst.exe` binary in its `--extract` mode to read a Microsoft
Outlook OST file directly with `libpff` and dump each message as either
individual `.eml` files (default) or one `.mbox` per folder. Attachments
are embedded; inline images are linked back to the HTML body via
`Content-ID` when the heuristic finds an exact match.

The tool does **not** require Outlook and does **not** require the OST to
belong to the current Outlook profile.

## When to invoke

Invoke when the user wants:

- Extract emails (with attachments) from an `.ost` into individual `.eml`
  or per-folder `.mbox` files.
- Read OST contents without configuring Outlook.
- Recover messages from an OST that belongs to a different user / profile.
- Filter the extraction by folder, date range, attachment size, or
  message class.
- Inventory an OST without writing (`--dry-run`).
- Process only a sample with `--limit N` for quick iteration.

Do **NOT** invoke for:

- Converting OST to a single `.pst` file — use `ost2pst.exe` without
  `--extract` (uses Outlook COM, requires the profile).
- Reading `.msg` or `.pst` files.

## Inputs needed

1. **OST file path** (`-i`).
2. **Output directory** (`-o`). Created if missing.
3. (optional) verbose progress (`-v`), folder filters, date filters,
   message limit, output format, attachment size cap, dry-run, errors log.

Ask the user for `-i` and `-o` if they didn't provide both.

## How to locate the binary

Look up `ost2pst.exe` in this order:

1. Environment variable `$env:OST2PST_EXE` if set.
2. On `PATH` (`Get-Command ost2pst.exe`).
3. In the project's `dist\` folder.

If none of those work, ask the user for the path or tell them to run
`.\build.ps1 -Clean` from the project root.

## How to run

```
ost2pst.exe --extract -i "<ruta_al_ost>" -o "<directorio_salida>" [opciones]
```

### Common invocations

```powershell
# Extracción completa con progreso por carpeta
ost2pst.exe --extract -i "C:\correo.ost" -o "C:\extraidos" -v

# Formato mbox (un fichero por carpeta del OST)
ost2pst.exe --extract -i "C:\correo.ost" -o "C:\extraidos" --format mbox

# Iteración rápida: solo los primeros 100 mensajes con resumen JSON
ost2pst.exe --extract -i "C:\correo.ost" -o "C:\extraidos" --limit 100 --json

# Solo "Bandeja de entrada", solo 2024, comprimido
ost2pst.exe --extract -i C:\correo.ost -o C:\extraidos `
    --folder "Bandeja de entrada" --since 2024-01-01 --until 2024-12-31 --compress

# Mailbox ruso/japonés con codepage override
ost2pst.exe --extract -i C:\correo.ost -o C:\extraidos --codepage 1251

# Excluir papelera, limitar adjuntos a 25 MB, errores estructurados
ost2pst.exe --extract -i C:\correo.ost -o C:\extraidos `
    --exclude "Eliminados" --max-attachment-mb 25 --errors-log errores.jsonl

# Incluir calendarios/contactos/tareas (por defecto se saltan)
ost2pst.exe --extract -i C:\correo.ost -o C:\extraidos --include-non-email

# Inventariar sin escribir (rápido, no toca disco)
ost2pst.exe --extract -i C:\correo.ost -o C:\extraidos --dry-run --json
```

Run via the `Bash` or `PowerShell` tool. Capture stdout/stderr. Surface
the final summary back to the user verbatim — includes counts of folders,
emails, omitidos (rango fecha / sin fecha / no-email / limit), adjuntos
(guardados / saltados), y errores.

## Output layout

**`--format eml` (default)**: replica el árbol de carpetas, un `.eml`
por mensaje. Adjuntos embebidos.

```
<output_dir>/
  Bandeja de entrada/
    00001-RE Presupuesto.eml
    00002-Factura abril.eml
    Subcarpeta/
      00001-...eml
  Elementos enviados/
  Elementos eliminados/
```

**`--format mbox`**: replica el árbol, un `messages.mbox` por carpeta
(formato mbox estándar, ideal para Thunderbird / mutt).

```
<output_dir>/
  Bandeja de entrada/
    messages.mbox
    Subcarpeta/
      messages.mbox
  Elementos enviados/
    messages.mbox
```

Cada `.eml` o entrada mbox incluye:
- Cabeceras `Received` / `DKIM-Signature` / `Authentication-Results` originales (multi-instancia).
- `To`/`Cc`/`Bcc` leídos de propiedades MAPI (`PR_DISPLAY_*`).
- `Message-ID` original.
- Cuerpo `text/plain` y/o `text/html`, fallback a RTF.
- Adjuntos embebidos como `multipart/mixed`, con `Content-ID` para inline images.

## Flags resumen

| Categoría | Flags |
|---|---|
| Filtros carpeta/fecha | `--folder NOMBRE` · `--exclude NOMBRE` · `--since` · `--until` · `--require-date` |
| Filtros mensaje | `--max-attachment-mb MB` · `--limit N` · `--include-non-email` |
| Salida | `--format eml\|mbox` · `--compress` · `--attachments-dir DIR` · `--json` |
| Decoding | `--codepage CP` |
| Logs | `-v` · `-q` · `-d` · `--errors-log PATH` |
| Control | `--dry-run` |

## Exit codes

| Code | Meaning                                            |
|------|----------------------------------------------------|
| 0    | OK                                                 |
| 1    | Argumentos inválidos                               |
| 4    | El fichero OST no existe                           |
| 8    | Completado con errores en algunos mensajes         |
| 11   | libpff no disponible en este build                 |
| 12   | No se pudo abrir el OST (formato malo o corrupto)  |
| 13   | No se pudo leer la raíz del OST                    |

Surface stderr verbatim on non-zero exit. Exit 8 → suggest re-running with
`--errors-log errores.jsonl` to capture per-message detail.

## Notes for the agent

- `--extract` and `--list` son mutuamente excluyentes.
- `--by-name`, `--force` aplican solo al modo PST; el modo `--extract`
  los ignora.
- `--compress` aplica solo a `--format eml` (mbox tiene su propia
  compactación implícita).
- Por defecto se saltan IPM.Appointment/Contact/Task. Si el usuario
  pregunta "¿por qué hay menos correos que en Outlook?", proponer
  `--include-non-email`.
- CLI tool only — never launch interactively / via double-click.
- Imágenes inline se detectan por igualdad estricta del basename del
  adjunto con la parte local del `cid:` (sin substring matching).
