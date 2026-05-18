---
name: ost-extract
description: Extracts all emails (including attachments and inline images detected via cid:) from a Microsoft Outlook OST file into individual .eml files using libpff. Does NOT require Outlook installed and does NOT require the original mailbox profile or credentials — works on orphan OSTs recovered from backups or other users. Supports folder include/exclude filters, date range filtering, dry-run mode, attachment size limits, and structured JSON error logging. Use this skill whenever the user provides an OST file and asks to extract its messages, dump emails to disk, read OST contents without Outlook, recover mail from an orphan OST, or convert OST to .eml. Trigger phrases include "extraer emails de OST", "leer fichero OST", "OST a EML", "extract OST", "dump OST messages", "OST without Outlook", "OST sin Outlook".
---

# OST Email Extraction (libpff, no Outlook required)

Drives the `ost2pst.exe` binary in its `--extract` mode to read a Microsoft
Outlook OST file directly with `libpff` and dump each message as a standalone
`.eml` file. Attachments are embedded inside the `.eml` (`multipart/mixed`)
and inline images are linked back to the HTML body via `Content-ID` when
the heuristic finds a match.

The tool does **not** require Outlook and does **not** require the OST to
belong to the current Outlook profile.

## When to invoke

Invoke when the user wants:

- Extract emails (with attachments) from an `.ost` into individual `.eml` files.
- Read contents of an OST without configuring Outlook.
- Recover messages from an OST that belongs to a different user / profile.
- Convert OST to `.eml` (one message per file, RFC 822 format).
- Filter the extraction by folder name, date range, or attachment size.
- Inventory an OST without writing (`--dry-run`).

Do **NOT** invoke for:

- Converting OST to a single `.pst` file — use `ost2pst.exe` without
  `--extract` (uses Outlook COM, requires the profile).
- Reading `.msg` or `.pst` files.

## Inputs needed

1. **OST file path** (`-i`).
2. **Output directory** (`-o`). Created if missing.
3. (optional) verbose progress (`-v`), folder filters, date filters,
   attachment size cap, dry-run, errors log.

Ask the user for `-i` and `-o` if they didn't provide both.

## How to locate the binary

The binary is named `ost2pst.exe`. Locate it in this order:

1. Environment variable `$env:OST2PST_EXE` (if set, use it).
2. On `PATH` (try `Get-Command ost2pst.exe`).
3. In the project's `dist\` folder if invoked from the project root.

If none of those work, ask the user for the path or tell them to run
`.\build.ps1 -Clean` from the project root to build it.

## How to run

```
ost2pst.exe --extract -i "<ruta_al_ost>" -o "<directorio_salida>" [opciones]
```

Common invocations:

```powershell
# Extracción completa con progreso por carpeta
ost2pst.exe --extract -i "C:\correo.ost" -o "C:\extraidos" -v

# Solo "Bandeja de entrada", solo 2024
ost2pst.exe --extract -i "C:\correo.ost" -o "C:\extraidos" `
    --folder "Bandeja de entrada" --since 2024-01-01 --until 2024-12-31

# Excluir papelera, limitar adjuntos a 25 MB, log de errores
ost2pst.exe --extract -i "C:\correo.ost" -o "C:\extraidos" `
    --exclude "Eliminados" --max-attachment-mb 25 --errors-log errores.jsonl

# Inventariar sin escribir nada
ost2pst.exe --extract -i "C:\correo.ost" -o "C:\extraidos" --dry-run
```

Capture stdout/stderr. Surface the final summary line back to the user
verbatim — incluye folders, emails, omitidos por filtro, adjuntos, errores.

## Output layout

Mirrors the OST folder hierarchy. Each message is one `.eml` named
`NNNNN-<asunto_saneado>.eml`. Attachments are embedded inside the `.eml`.

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

## Filters (modo --extract)

| Flag                        | Significado                                            |
|-----------------------------|--------------------------------------------------------|
| `--folder NOMBRE`           | Incluir solo carpetas cuyo nombre contenga NOMBRE (repetible) |
| `--exclude NOMBRE`          | Saltar carpetas (y subárbol) (repetible)               |
| `--since YYYY-MM-DD`        | Mensajes con fecha >= esta                              |
| `--until YYYY-MM-DD`        | Mensajes con fecha <= esta (inclusivo hasta fin de día) |
| `--max-attachment-mb MB`    | Salta adjuntos mayores que MB (se loguean)              |
| `--dry-run`                 | No escribe nada; solo cuenta                            |
| `--errors-log PATH`         | Vuelca cada error como JSON línea                       |

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
- `--by-name`, `--force` aplican solo al modo PST; el modo `--extract` los ignora.
- CLI tool only — never launch interactively / via double-click.
- Imágenes inline se detectan por heurística (nombre del adjunto vs `cid:`
  del HTML). Casos raros pueden quedar como adjunto convencional.
