"""ost2pst - herramienta CLI para procesar ficheros Outlook OST en Windows.

Dos modos:

  1) Exportación a PST (Outlook COM + pywin32).
  2) Extracción a .eml / .mbox  (libpff/pypff, sin Outlook).

Detalle de funcionalidades en README.md y CHANGELOG.md.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import linecache
import logging
import mailbox
import mimetypes
import os
import re
import sys
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from email import message_from_string as _parse_headers
from email.message import EmailMessage
from email.utils import format_datetime, formatdate

__version__ = "3.1.0"


# =========================================================================
# Dependencias opcionales
# =========================================================================

try:
    import pywintypes
    import win32com.client
    _HAS_WIN32 = True
except ImportError:
    win32com = None  # type: ignore
    pywintypes = None  # type: ignore
    _HAS_WIN32 = False

try:
    import pypff
    _HAS_PYPFF = True
except ImportError:
    pypff = None  # type: ignore
    _HAS_PYPFF = False

try:
    from striprtf.striprtf import rtf_to_text
    _HAS_STRIPRTF = True
except ImportError:
    rtf_to_text = None  # type: ignore
    _HAS_STRIPRTF = False

try:
    import chardet
    _HAS_CHARDET = True
except ImportError:
    chardet = None  # type: ignore
    _HAS_CHARDET = False


# =========================================================================
# Constantes y códigos de salida
# =========================================================================

OL_STORE_UNICODE = 3
SAFE_NAME_MAXLEN = 80
_SAFE_NAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_CID_RE = re.compile(r'''cid:([^"'\s>]+)''', re.IGNORECASE)
_RECEIVED_LIKE = {"received", "x-received", "authentication-results",
                  "arc-authentication-results", "arc-message-signature",
                  "arc-seal", "dkim-signature"}
_HEADERS_REBUILD = {"content-type", "content-transfer-encoding",
                    "mime-version", "content-disposition"}

# MAPI property IDs (16-bit, los 16 bits altos del entry_type)
MAPI_DISPLAY_TO = 0x0E04
MAPI_DISPLAY_CC = 0x0E03
MAPI_DISPLAY_BCC = 0x0E02
MAPI_MESSAGE_CLASS = 0x001A
MAPI_SENDER_EMAIL = 0x0C1F
MAPI_SENDER_NAME = 0x0C1A
MAPI_INTERNET_MESSAGE_ID = 0x1035

EXIT_OK = 0
EXIT_ARGS = 1
EXIT_OUTLOOK_FAILED = 2
EXIT_STORE_NOT_FOUND_BY_NAME = 3
EXIT_INPUT_NOT_FOUND = 4
EXIT_STORE_NOT_LOADED = 5
EXIT_PST_CREATE_FAILED = 6
EXIT_PST_LOCATE_FAILED = 7
EXIT_COMPLETED_WITH_ERRORS = 8
EXIT_WIN32_MISSING = 10
EXIT_LIBPFF_MISSING = 11
EXIT_OST_OPEN_FAILED = 12
EXIT_OST_ROOT_FAILED = 13


# =========================================================================
# Logging
# =========================================================================

LOG = logging.getLogger("ost2pst")


class _MsFormatter(logging.Formatter):
    def formatTime(self, record, datefmt=None):
        ct = datetime.fromtimestamp(record.created)
        return ct.strftime("%Y-%m-%d %H:%M:%S.") + f"{int(record.msecs):03d}"


def _configure_logging(verbose: bool, quiet: bool = False) -> None:
    if quiet:
        level = logging.ERROR
    elif verbose:
        level = logging.INFO
    else:
        level = logging.WARNING

    if LOG.handlers:
        LOG.setLevel(level)
        return
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(_MsFormatter(
        fmt="[%(asctime)s] %(levelname)-5s %(message)s",
    ))
    LOG.addHandler(handler)
    LOG.setLevel(level)


def _configure_console_utf8() -> None:
    for stream_name in ("stdout", "stderr"):
        s = getattr(sys, stream_name, None)
        if s and hasattr(s, "reconfigure"):
            try:
                s.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


# =========================================================================
# Modo --debug: traza línea-a-línea
# =========================================================================

_DEBUG_SRC_PATH: str | None = None


def _resolve_source_path() -> str | None:
    candidates = []
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidates.append(os.path.join(meipass, "ost2pst.py"))
        candidates.append(os.path.join(os.path.dirname(sys.executable), "ost2pst.py"))
    candidates.append(os.path.abspath(__file__))
    return next((c for c in candidates if os.path.isfile(c)), None)


def _now_ms() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]


def _debug_tracer(frame, event, arg):
    code = frame.f_code
    bn = os.path.basename(code.co_filename).lower()
    if bn != "ost2pst.py":
        return None

    ts = _now_ms()
    lineno = frame.f_lineno
    func = code.co_name
    try:
        if event == "line":
            src = linecache.getline(_DEBUG_SRC_PATH or code.co_filename, lineno).rstrip()
            sys.stderr.write(f"[{ts}] LINE  {bn}:{lineno:<4} {func}(): {src}\n")
        elif event == "call":
            sys.stderr.write(f"[{ts}] CALL  {bn}:{lineno:<4} -> {func}()\n")
        elif event == "return":
            sys.stderr.write(f"[{ts}] RET   {bn}:{lineno:<4} <- {func}() = {arg!r}\n")
        elif event == "exception":
            exc_type, exc_value, _ = arg
            sys.stderr.write(
                f"[{ts}] EXC   {bn}:{lineno:<4} {func}(): "
                f"{exc_type.__name__}: {exc_value}\n"
            )
        sys.stderr.flush()
    except Exception:
        pass
    return _debug_tracer


def _install_debug_tracer() -> None:
    global _DEBUG_SRC_PATH
    _DEBUG_SRC_PATH = _resolve_source_path()
    if _DEBUG_SRC_PATH:
        linecache.checkcache(_DEBUG_SRC_PATH)
        linecache.getlines(_DEBUG_SRC_PATH)
    sys.stderr.write(
        f"[{_now_ms()}] DEBUG modo activado "
        f"(fuente: {_DEBUG_SRC_PATH or '<no encontrado>'})\n"
    )
    sys.stderr.flush()
    sys.settrace(_debug_tracer)
    frame = sys._getframe(1)
    while frame is not None:
        frame.f_trace = _debug_tracer
        frame = frame.f_back


# =========================================================================
# Helpers comunes
# =========================================================================

def _safe_filename(s: str, maxlen: int = SAFE_NAME_MAXLEN) -> str:
    s = (s or "").strip()
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = _SAFE_NAME_RE.sub("_", s)
    s = re.sub(r"\s+", " ", s).strip(" .")
    return (s or "sin_asunto")[:maxlen]


def _decode_body(body, codepage: int = 0) -> str:
    """Decodifica probando: utf-8, codepage del OST, utf-16-le, chardet, latin-1."""
    if body is None:
        return ""
    if isinstance(body, str):
        return body
    if not isinstance(body, (bytes, bytearray)):
        return str(body)

    encodings = ["utf-8"]
    if codepage:
        encodings.append(f"cp{codepage}")
    encodings += ["utf-16-le", "latin-1"]

    for enc in encodings[:-1]:
        try:
            return body.decode(enc, errors="strict")
        except (UnicodeDecodeError, LookupError):
            continue

    # chardet como último recurso antes del fallback agresivo
    if _HAS_CHARDET and len(body) > 0:
        try:
            detected = chardet.detect(body[:16384])
            enc = (detected or {}).get("encoding")
            conf = (detected or {}).get("confidence", 0)
            if enc and conf > 0.7:
                try:
                    return body.decode(enc, errors="strict")
                except (UnicodeDecodeError, LookupError):
                    pass
        except Exception:
            pass

    return body.decode("latin-1", errors="replace")


def _guess_mime_type(filename: str) -> tuple:
    if filename:
        guess, _ = mimetypes.guess_type(filename)
        if guess and "/" in guess:
            main, sub = guess.split("/", 1)
            return main, sub
    return "application", "octet-stream"


def _ensure_long_path(p: str) -> str:
    if os.name != "nt" or not p:
        return p
    if p.startswith("\\\\?\\"):
        return p
    abs_p = os.path.abspath(p)
    if len(abs_p) <= 240:
        return p
    if abs_p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + abs_p[2:]
    return "\\\\?\\" + abs_p


def _extract_cids(html: str) -> set:
    if not html:
        return set()
    return set(_CID_RE.findall(html))


def _attachment_inline_cid(attachment_name: str, cids: set) -> str | None:
    """Devuelve el cid: si el adjunto matchea exactamente uno; None en otro caso.

    Matching estricto: el basename del adjunto (sin extensión) tiene que
    ser igual a la parte local del cid (antes de '@'). Sin substrings,
    para evitar falsos positivos tipo 'cid:abc' contra 'abcdef.png'.
    """
    if not attachment_name or not cids:
        return None
    name_base = os.path.splitext(attachment_name)[0].strip().lower()
    if not name_base:
        return None
    for cid in cids:
        cid_base = cid.split("@", 1)[0].strip().lower()
        if cid_base and name_base == cid_base:
            return cid
    return None


def _dedupe_filename(name: str, used: set) -> str:
    if name not in used:
        used.add(name)
        return name
    base, ext = os.path.splitext(name)
    i = 2
    while True:
        candidate = f"{base} ({i}){ext}"
        if candidate not in used:
            used.add(candidate)
            return candidate
        i += 1


def _sha256_file(path: str, chunk_size: int = 1024 * 1024) -> str:
    h = hashlib.sha256()
    with open(_ensure_long_path(path), "rb") as f:
        while True:
            buf = f.read(chunk_size)
            if not buf:
                break
            h.update(buf)
    return h.hexdigest()


def _parse_iso_dt(s: str | None, end_of_day: bool = False) -> datetime | None:
    if not s:
        return None
    dt = datetime.fromisoformat(s)
    if end_of_day and dt.hour == 0 and dt.minute == 0 and dt.second == 0 and dt.microsecond == 0:
        dt = dt.replace(hour=23, minute=59, second=59, microsecond=999999)
    return dt


# =========================================================================
# Lectura de propiedades MAPI desde pypff
# =========================================================================

def _entry_property_id(entry_type: int) -> int:
    """Devuelve el property ID (16 bits) de un entry_type pypff.

    pypff puede devolver el tag completo (id<<16 | type) o solo el id;
    normalizamos a property ID.
    """
    if entry_type > 0xFFFF:
        return entry_type >> 16
    return entry_type


def _read_mapi_string(pf_item, prop_id: int) -> str | None:
    """Lee propiedad MAPI string por ID (16 bits). None si no existe."""
    try:
        rs_iter = pf_item.record_sets
    except Exception:
        return None
    try:
        for rs in rs_iter:
            try:
                entries = rs.entries
            except Exception:
                continue
            for entry in entries:
                try:
                    et = entry.entry_type
                except Exception:
                    continue
                if _entry_property_id(et) != prop_id:
                    continue
                try:
                    val = entry.data_as_string
                    if val:
                        return val
                except Exception:
                    try:
                        raw = entry.data
                        if raw:
                            if isinstance(raw, bytes):
                                return raw.decode("utf-16-le", errors="replace").rstrip("\x00")
                            return str(raw)
                    except Exception:
                        pass
    except Exception:
        pass
    return None


def _read_message_class(msg) -> str:
    return _read_mapi_string(msg, MAPI_MESSAGE_CLASS) or ""


def _read_recipients(msg) -> dict:
    return {
        "To": _read_mapi_string(msg, MAPI_DISPLAY_TO) or "",
        "Cc": _read_mapi_string(msg, MAPI_DISPLAY_CC) or "",
        "Bcc": _read_mapi_string(msg, MAPI_DISPLAY_BCC) or "",
    }


def _read_sender_email(msg) -> str | None:
    return _read_mapi_string(msg, MAPI_SENDER_EMAIL)


def _read_internet_message_id(msg) -> str | None:
    return _read_mapi_string(msg, MAPI_INTERNET_MESSAGE_ID)


def _is_email_class(message_class: str) -> bool:
    mc = (message_class or "").lower()
    if not mc:
        return True
    return mc == "ipm.note" or mc.startswith("ipm.note.")


# =========================================================================
# ExtractStats
# =========================================================================

class ExtractStats:
    def __init__(self, errors_log_path: str | None = None):
        self.folders = 0
        self.items = 0
        self.skipped_date_range = 0
        self.skipped_no_date = 0
        self.skipped_class = 0
        self.skipped_limit = 0
        self.attachments_saved = 0
        self.attachments_skipped = 0
        self.errors = 0
        self._errors_path = errors_log_path
        self._errors_fp = None

    def _open_errors(self):
        # Mantenemos el fp abierto durante toda la extracción para escribir
        # entradas incrementalmente; close() libera al final.
        if self._errors_fp is None and self._errors_path:
            self._errors_fp = open(  # noqa: SIM115
                _ensure_long_path(self._errors_path),
                "w", encoding="utf-8",
            )

    def record_error(self, **kwargs) -> None:
        self.errors += 1
        if self._errors_path:
            self._open_errors()
            kwargs.setdefault("ts", _now_ms())
            try:
                self._errors_fp.write(json.dumps(kwargs, default=str) + "\n")
                self._errors_fp.flush()
            except Exception:
                pass

    def to_dict(self) -> dict:
        return {
            "folders": self.folders,
            "items": self.items,
            "skipped_date_range": self.skipped_date_range,
            "skipped_no_date": self.skipped_no_date,
            "skipped_class": self.skipped_class,
            "skipped_limit": self.skipped_limit,
            "attachments_saved": self.attachments_saved,
            "attachments_skipped": self.attachments_skipped,
            "errors": self.errors,
        }

    def close(self) -> None:
        if self._errors_fp is not None:
            try:
                self._errors_fp.close()
            except Exception:
                pass
            self._errors_fp = None


# =========================================================================
# Progress
# =========================================================================

class Progress:
    def __init__(self, enabled: bool):
        self.enabled = enabled and sys.stderr.isatty()
        self._last_emit = 0.0
        self._messages = 0

    def tick(self, label: str = "") -> None:
        if not self.enabled:
            return
        self._messages += 1
        now = time.monotonic()
        if now - self._last_emit < 0.25:
            return
        self._last_emit = now
        msg = f"\r[ost2pst] procesados: {self._messages} mensajes"
        if label:
            msg += f"  | {label[:60]}"
        sys.stderr.write(msg.ljust(120))
        sys.stderr.flush()

    def finish(self) -> None:
        if self.enabled and self._messages:
            sys.stderr.write("\r" + " " * 120 + "\r")
            sys.stderr.flush()


# =========================================================================
# Configuración de extracción
# =========================================================================

@dataclass
class ExtractConfig:
    ost_path: str
    out_dir: str
    folders: list = field(default_factory=list)
    excludes: list = field(default_factory=list)
    since_dt: datetime | None = None
    until_dt: datetime | None = None
    max_att_bytes: int = 0
    require_date: bool = False
    include_non_email: bool = False
    dry_run: bool = False
    compress: bool = False
    attachments_dir: str | None = None
    codepage: int = 0
    verbose: bool = False
    limit: int = 0
    out_format: str = "eml"   # "eml" o "mbox"


# =========================================================================
# Modo 1: Exportación a PST (Outlook COM)
# =========================================================================

def _validate_pst_args(args: argparse.Namespace) -> int | None:
    """Valida args del modo PST ANTES de arrancar Outlook. Devuelve exit code o None."""
    if args.list:
        return None  # --list no requiere input/output

    if not args.input or not args.output:
        LOG.error("El modo PST requiere --input y --output.")
        return EXIT_ARGS

    out_path = os.path.abspath(args.output)
    if not out_path.lower().endswith(".pst"):
        LOG.error("El fichero destino debe tener extensión .pst")
        return EXIT_ARGS

    out_dir = os.path.dirname(out_path)
    if out_dir and not os.path.isdir(out_dir):
        LOG.error(f"El directorio destino no existe: {out_dir}")
        return EXIT_ARGS

    if os.path.exists(out_path):
        if args.force:
            try:
                os.remove(out_path)
            except OSError as e:
                LOG.error(f"No se pudo eliminar {out_path}: {e}")
                return EXIT_ARGS
        else:
            LOG.error(f"El fichero destino ya existe: {out_path}. Usa --force.")
            return EXIT_ARGS

    if not args.by_name:
        in_path = os.path.abspath(args.input)
        if not os.path.exists(in_path):
            LOG.error(f"El fichero OST no existe: {in_path}")
            return EXIT_INPUT_NOT_FOUND

    return None


def _list_stores(namespace):
    out = []
    for store in namespace.Stores:
        try:
            out.append((store.DisplayName, store.FilePath))
        except Exception:
            pass
    return out


def _print_stores(namespace, stream=sys.stdout) -> None:
    stream.write("Stores cargados en Outlook:\n")
    for name, path in _list_stores(namespace):
        stream.write(f"  - {name!r}  ->  {path}\n")


def _find_store_by_path(namespace, target_path: str):
    target = os.path.normcase(os.path.abspath(target_path))
    for store in namespace.Stores:
        try:
            sp = store.FilePath
            if sp and os.path.normcase(os.path.abspath(sp)) == target:
                return store
        except Exception:
            continue
    return None


def _find_store_by_name(namespace, name: str):
    for store in namespace.Stores:
        try:
            if store.DisplayName == name:
                return store
        except Exception:
            continue
    return None


def _copy_folder_recursive(src_folder, dst_parent, depth: int = 0,
                           stats: dict | None = None) -> dict:
    if stats is None:
        stats = {"folders": 0, "items": 0, "errors": 0}

    name = src_folder.Name
    try:
        dst_folder = dst_parent.Folders[name]
    except Exception:
        try:
            dst_folder = dst_parent.Folders.Add(name)
        except Exception as e:
            stats["errors"] += 1
            LOG.error(f"No se pudo crear carpeta destino '{name}': {e}")
            return stats
    stats["folders"] += 1

    items = src_folder.Items
    total = items.Count
    LOG.info("  " * depth + f"[+] {name}  ({total} items)")

    copied = 0
    for i in range(1, total + 1):
        try:
            items.Item(i).Copy().Move(dst_folder)
            copied += 1
        except Exception as e:
            stats["errors"] += 1
            LOG.info("  " * depth + f"    ! item {i}: {e}")
    stats["items"] += copied

    for sub in src_folder.Folders:
        try:
            _copy_folder_recursive(sub, dst_folder, depth=depth + 1, stats=stats)
        except Exception as e:
            stats["errors"] += 1
            LOG.error(f"Error procesando subcarpeta '{getattr(sub, 'Name', '?')}': {e}")
    return stats


def run_pst_export(args: argparse.Namespace) -> int:
    # Validar args ANTES de tocar Outlook
    err = _validate_pst_args(args)
    if err is not None:
        return err

    if not _HAS_WIN32:
        LOG.error("pywin32 no disponible. El modo PST requiere Outlook + pywin32.")
        LOG.error("Usa --extract para extracción cruda con libpff.")
        return EXIT_WIN32_MISSING

    LOG.info("Conectando con Outlook...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except pywintypes.com_error as e:
        LOG.error(f"No se pudo iniciar Outlook: {e}")
        return EXIT_OUTLOOK_FAILED

    if args.list:
        _print_stores(namespace, stream=sys.stdout)
        return EXIT_OK

    out_path = os.path.abspath(args.output)

    if args.by_name:
        src_store = _find_store_by_name(namespace, args.input)
        if not src_store:
            LOG.error(f"No se encontró un store con el nombre '{args.input}'")
            _print_stores(namespace, stream=sys.stderr)
            return EXIT_STORE_NOT_FOUND_BY_NAME
    else:
        in_path = os.path.abspath(args.input)
        src_store = _find_store_by_path(namespace, in_path)
        if not src_store:
            LOG.error(f"El OST '{in_path}' no está cargado en el perfil de Outlook.")
            LOG.error("Configura la cuenta en Outlook o usa --by-name.")
            _print_stores(namespace, stream=sys.stderr)
            return EXIT_STORE_NOT_LOADED

    LOG.info(f"Store origen: {src_store.DisplayName}")
    LOG.info(f"PST destino:  {out_path}")

    try:
        namespace.AddStoreEx(out_path, OL_STORE_UNICODE)
    except pywintypes.com_error as e:
        LOG.error(f"No se pudo crear el PST: {e}")
        return EXIT_PST_CREATE_FAILED

    dst_store = _find_store_by_path(namespace, out_path)
    if not dst_store:
        LOG.error("No se pudo localizar el PST recién creado.")
        return EXIT_PST_LOCATE_FAILED

    src_root = src_store.GetRootFolder()
    dst_root = dst_store.GetRootFolder()

    LOG.info("Iniciando copia de carpetas...")
    stats = {"folders": 0, "items": 0, "errors": 0}
    for folder in src_root.Folders:
        try:
            _copy_folder_recursive(folder, dst_root, stats=stats)
        except Exception as e:
            stats["errors"] += 1
            LOG.error(f"Error procesando '{folder.Name}': {e}")

    try:
        namespace.RemoveStore(dst_root)
    except Exception:
        pass

    summary = {
        "mode": "pst",
        "output": out_path,
        "folders": stats["folders"],
        "items": stats["items"],
        "errors": stats["errors"],
    }
    if getattr(args, "json", False):
        print(json.dumps(summary, indent=2, ensure_ascii=False))
    else:
        print(
            f"OK: exportación a PST completada -> {out_path}\n"
            f"    carpetas: {stats['folders']}  items: {stats['items']}  "
            f"errores: {stats['errors']}"
        )
    return EXIT_OK if stats["errors"] == 0 else EXIT_COMPLETED_WITH_ERRORS


# =========================================================================
# Modo 2: Extracción a .eml / .mbox  (libpff)
# =========================================================================

_FILETIME_GETTERS = (
    "get_delivery_time",
    "get_client_submit_time",
    "get_creation_time",
)


def _msg_datetime(pypff_msg) -> datetime | None:
    for getter_name in _FILETIME_GETTERS:
        getter = getattr(pypff_msg, getter_name, None)
        if getter is None:
            continue
        try:
            dt = getter()
            if dt:
                return dt
        except Exception:
            continue
    return None


def _attachment_bytes(att) -> bytes:
    try:
        size = att.get_size()
    except Exception:
        return b""
    if size <= 0:
        return b""
    try:
        return att.read_buffer(size) or b""
    except Exception:
        return b""


def _attachment_name(att, index: int) -> str:
    name = None
    try:
        name = att.get_name()
    except Exception:
        name = None
    return name or f"attachment_{index:03d}.bin"


# --- build_eml descompuesto ----------------------------------------------

def _apply_headers(em: EmailMessage, msg, transport: str) -> None:
    if transport.strip():
        try:
            parsed = _parse_headers(transport)
            single_seen = set()
            for key, value in parsed.items():
                k_lower = key.lower()
                if k_lower in _HEADERS_REBUILD:
                    continue
                if k_lower in _RECEIVED_LIKE:
                    try:
                        em[key] = value
                    except Exception:
                        pass
                    continue
                if k_lower in single_seen:
                    continue
                single_seen.add(k_lower)
                try:
                    em[key] = value
                except Exception:
                    pass
        except Exception:
            pass

    if "Subject" not in em:
        try:
            em["Subject"] = msg.get_subject() or ""
        except Exception:
            em["Subject"] = ""

    if "From" not in em:
        sender_name = ""
        sender_email = ""
        try:
            sender_name = msg.get_sender_name() or ""
        except Exception:
            pass
        try:
            sender_email = _read_sender_email(msg) or ""
        except Exception:
            pass
        if sender_name and sender_email:
            em["From"] = f"{sender_name} <{sender_email}>"
        else:
            em["From"] = sender_name or sender_email or ""

    if "Date" not in em:
        dt = _msg_datetime(msg)
        if dt:
            try:
                if dt.tzinfo is None:
                    em["Date"] = formatdate(dt.timestamp(), localtime=False)
                else:
                    em["Date"] = format_datetime(dt)
            except Exception:
                pass

    if "Message-ID" not in em:
        mid = _read_internet_message_id(msg)
        if mid:
            em["Message-ID"] = mid

    recipients = _read_recipients(msg)
    for hdr_name in ("To", "Cc", "Bcc"):
        if hdr_name not in em and recipients.get(hdr_name):
            # PR_DISPLAY_* usa ';' como separador; RFC 2822 espera ','
            value = recipients[hdr_name].replace(";", ",")
            try:
                em[hdr_name] = value
            except Exception:
                pass


def _apply_body(em: EmailMessage, msg, codepage: int) -> str:
    try:
        plain = _decode_body(msg.get_plain_text_body(), codepage)
    except Exception:
        plain = ""
    try:
        html = _decode_body(msg.get_html_body(), codepage)
    except Exception:
        html = ""

    if not plain and not html and _HAS_STRIPRTF:
        try:
            rtf = msg.get_rtf_body()
            if rtf:
                rtf_text = rtf_to_text(_decode_body(rtf, codepage), errors="ignore")
                if rtf_text and rtf_text.strip():
                    plain = rtf_text
        except Exception:
            pass

    if plain and html:
        em.set_content(plain)
        em.add_alternative(html, subtype="html")
    elif html:
        em.set_content(html, subtype="html")
    else:
        em.set_content(plain or "")
    return html


def _apply_attachments(em: EmailMessage, msg, html: str,
                       config: ExtractConfig,
                       stats: ExtractStats,
                       attachments_out_dir: str | None = None,
                       folder_path_label: str = "",
                       message_index: int = 0) -> int:
    try:
        n_att = msg.get_number_of_attachments()
    except Exception:
        return 0

    cids = _extract_cids(html)
    used_names = set()
    skipped_count = 0

    for i in range(n_att):
        try:
            att = msg.get_attachment(i)
        except Exception:
            continue
        name = _dedupe_filename(_attachment_name(att, i + 1), used_names)

        try:
            size = att.get_size()
        except Exception:
            size = 0

        if config.max_att_bytes and size > config.max_att_bytes:
            skipped_count += 1
            stats.attachments_skipped += 1
            stats.record_error(
                stage="attachment_too_big",
                folder=folder_path_label,
                message_index=message_index,
                attachment=name,
                size=size,
                limit=config.max_att_bytes,
            )
            continue

        data = _attachment_bytes(att)
        if not data:
            continue

        maintype, subtype = _guess_mime_type(name)
        inline_cid = _attachment_inline_cid(name, cids)
        kwargs = {"maintype": maintype, "subtype": subtype, "filename": name}
        if inline_cid:
            kwargs["disposition"] = "inline"
            kwargs["cid"] = f"<{inline_cid}>"

        try:
            em.add_attachment(data, **kwargs)
        except Exception:
            try:
                em.add_attachment(data, maintype="application",
                                  subtype="octet-stream", filename=name)
            except Exception:
                continue

        stats.attachments_saved += 1

        # --attachments-dir: vuelca a disco SOLO si no es inline
        if attachments_out_dir and not inline_cid:
            try:
                os.makedirs(_ensure_long_path(attachments_out_dir), exist_ok=True)
                att_path = os.path.join(attachments_out_dir, name)
                with open(_ensure_long_path(att_path), "wb") as f:
                    f.write(data)
            except OSError as e:
                stats.record_error(
                    stage="attachment_write",
                    path=attachments_out_dir,
                    attachment=name,
                    error=str(e),
                )

    return skipped_count


def build_eml(msg, *, config: ExtractConfig | None = None,
              stats: ExtractStats | None = None,
              attachments_out_dir: str | None = None,
              folder_path_label: str = "",
              message_index: int = 0) -> bytes:
    if config is None:
        config = ExtractConfig(ost_path="", out_dir="")
    if stats is None:
        stats = ExtractStats()

    em = EmailMessage()
    try:
        transport = msg.get_transport_headers() or ""
    except Exception:
        transport = ""

    _apply_headers(em, msg, transport)
    html = _apply_body(em, msg, config.codepage)
    _apply_attachments(em, msg, html, config, stats,
                       attachments_out_dir=attachments_out_dir,
                       folder_path_label=folder_path_label,
                       message_index=message_index)
    return bytes(em)


# --- escritura por mensaje (eml o mbox) ----------------------------------

@contextmanager
def _folder_writer(target_dir: str, out_format: str, compress: bool):
    """Yields callable add(data, eml_filename) que escribe un mensaje."""
    if out_format == "mbox":
        mbox_path = os.path.join(target_dir, "messages.mbox")
        mb = mailbox.mbox(_ensure_long_path(mbox_path))
        mb.lock()
        try:
            def _add(data: bytes, _eml_filename: str) -> None:
                mb.add(mailbox.mboxMessage(data))
            yield _add
        finally:
            mb.flush()
            mb.unlock()
            mb.close()
    else:
        def _add(data: bytes, eml_filename: str) -> None:
            path = os.path.join(target_dir, eml_filename)
            if compress:
                if not path.endswith(".gz"):
                    path += ".gz"
                with gzip.open(_ensure_long_path(path), "wb", compresslevel=6) as f:
                    f.write(data)
            else:
                with open(_ensure_long_path(path), "wb") as f:
                    f.write(data)
        yield _add


# --- recorrido iterativo del árbol ---------------------------------------

def _folder_path_str(stack: list) -> str:
    return "/".join(stack) if stack else "<root>"


def _folder_matches(name: str, patterns: list) -> bool:
    return any(p.lower() in name.lower() for p in patterns)


def _extract_folder_iter(root_folder, out_dir: str, config: ExtractConfig,
                         stats: ExtractStats, progress: Progress) -> None:
    """Recorre el árbol iterativamente (sin recursión) y procesa cada carpeta."""
    # Stack: (folder, target_dir, depth, path_stack, in_included_subtree)
    stack = [(root_folder, out_dir, 0, [], False)]

    while stack:
        folder, target, depth, path_stack, in_inc = stack.pop()

        if config.limit and stats.items >= config.limit:
            break

        try:
            name = folder.get_name() or ""
        except Exception:
            name = ""

        if name and config.excludes and _folder_matches(name, config.excludes):
            LOG.info("  " * depth + f"[-] skip excluida: {name}")
            continue

        if name and config.folders and _folder_matches(name, config.folders):
            in_inc = True

        current_path = path_stack + [name] if name else path_stack
        folder_label = _folder_path_str(current_path)
        process_messages = (not config.folders) or in_inc

        if process_messages and not config.dry_run:
            try:
                os.makedirs(_ensure_long_path(target), exist_ok=True)
            except OSError as e:
                stats.record_error(stage="mkdir", path=target, error=str(e))
                continue

        try:
            n_msg = folder.get_number_of_sub_messages()
        except Exception:
            n_msg = 0

        LOG.info("  " * depth + f"[+] {name or '<root>'}  ({n_msg} mensajes)"
                 + ("  [solo conteo]" if not process_messages else ""))

        # Conteo de carpetas: solo si realmente procesamos
        if process_messages:
            stats.folders += 1
            _process_messages_in_folder(folder, target, n_msg, config, stats,
                                        progress, folder_label, current_path)

        # Apilar subcarpetas en orden inverso para mantener DFS visual
        try:
            n_sub = folder.get_number_of_sub_folders()
        except Exception:
            n_sub = 0

        children = []
        for i in range(n_sub):
            try:
                sub = folder.get_sub_folder(i)
            except Exception as e:
                stats.record_error(stage="subfolder", folder=folder_label,
                                   subfolder_index=i, error=str(e))
                LOG.error(f"Error procesando subcarpeta {i}: {e}")
                continue
            try:
                sub_name = sub.get_name() or ""
            except Exception:
                sub_name = ""
            sub_target = (
                os.path.join(target, _safe_filename(sub_name))
                if sub_name else target
            )
            children.append((sub, sub_target, depth + 1, current_path, in_inc))

        for child in reversed(children):
            stack.append(child)


def _process_messages_in_folder(folder, target, n_msg, config, stats,
                                progress, folder_label, current_path) -> None:
    with _folder_writer(target, config.out_format, config.compress) as writer:
        for i in range(n_msg):
            if config.limit and stats.items >= config.limit:
                stats.skipped_limit += (n_msg - i)
                return

            try:
                msg = folder.get_sub_message(i)

                # filtro clase MAPI
                if not config.include_non_email:
                    mc = _read_message_class(msg)
                    if mc and not _is_email_class(mc):
                        stats.skipped_class += 1
                        continue

                # filtro fecha
                dt = _msg_datetime(msg)
                if dt is None and config.require_date and (config.since_dt or config.until_dt):
                    stats.skipped_no_date += 1
                    continue
                if dt:
                    if config.since_dt and dt < config.since_dt:
                        stats.skipped_date_range += 1
                        continue
                    if config.until_dt and dt > config.until_dt:
                        stats.skipped_date_range += 1
                        continue

                subj = ""
                try:
                    subj = msg.get_subject() or ""
                except Exception:
                    pass
                eml_name = f"{i+1:05d}-{_safe_filename(subj)}.eml"

                if config.dry_run:
                    try:
                        stats.attachments_saved += msg.get_number_of_attachments()
                    except Exception:
                        pass
                    stats.items += 1
                else:
                    attach_dir = None
                    if config.attachments_dir:
                        attach_dir = os.path.join(
                            config.attachments_dir,
                            *[_safe_filename(p) for p in current_path if p],
                            f"{i+1:05d}",
                        )
                    data = build_eml(
                        msg, config=config, stats=stats,
                        attachments_out_dir=attach_dir,
                        folder_path_label=folder_label,
                        message_index=i,
                    )
                    writer(data, eml_name)
                    stats.items += 1

                progress.tick(label=folder_label)
            except Exception as e:
                stats.record_error(
                    stage="message",
                    folder=folder_label,
                    message_index=i,
                    error=str(e),
                )
                LOG.warning(f"  mensaje {i}: {e}")


def run_eml_extract(args: argparse.Namespace) -> int:
    if not _HAS_PYPFF:
        LOG.error("libpff (pypff) no disponible en este build.")
        LOG.error("Compila con Python 3.9 + libpff-python==20211114.")
        return EXIT_LIBPFF_MISSING

    if not args.input or not args.output:
        LOG.error("--extract requiere --input (fichero OST) y --output (directorio).")
        return EXIT_ARGS

    ost_path = os.path.abspath(args.input)
    out_dir = os.path.abspath(args.output)

    if not os.path.isfile(ost_path):
        LOG.error(f"El fichero OST no existe: {ost_path}")
        return EXIT_INPUT_NOT_FOUND

    try:
        since_dt = _parse_iso_dt(args.since)
        until_dt = _parse_iso_dt(args.until, end_of_day=True)
    except ValueError as e:
        LOG.error(f"Fecha inválida: {e}")
        return EXIT_ARGS

    if args.format not in ("eml", "mbox"):
        LOG.error(f"--format inválido: {args.format} (eml|mbox)")
        return EXIT_ARGS

    if not args.dry_run:
        try:
            os.makedirs(_ensure_long_path(out_dir), exist_ok=True)
        except OSError as e:
            LOG.error(f"No se pudo crear el directorio de salida: {e}")
            return EXIT_ARGS

    max_att_bytes = (args.max_attachment_mb or 0) * 1024 * 1024

    LOG.info(f"Calculando SHA-256 del OST: {ost_path}")
    try:
        ost_sha256 = _sha256_file(ost_path)
        LOG.info(f"SHA-256: {ost_sha256}")
    except Exception as e:
        ost_sha256 = None
        LOG.warning(f"No se pudo calcular SHA-256 del OST: {e}")

    LOG.info(f"Abriendo OST con libpff: {ost_path}")
    if args.dry_run:
        LOG.info("--dry-run: no se escribirán ficheros, solo se contabilizan.")

    pst = pypff.file()
    try:
        pst.open(ost_path)
    except Exception as e:
        LOG.error(f"No se pudo abrir el OST con libpff: {e}")
        return EXIT_OST_OPEN_FAILED

    codepage = args.codepage or 0
    if not codepage:
        try:
            codepage = pst.get_ascii_codepage() or 0
        except Exception:
            codepage = 0
    if codepage:
        LOG.info(f"Codepage del OST: cp{codepage}"
                 + ("  [override por --codepage]" if args.codepage else ""))

    try:
        root = pst.get_root_folder()
    except Exception as e:
        LOG.error(f"No se pudo leer la raíz del OST: {e}")
        try:
            pst.close()
        except Exception:
            pass
        return EXIT_OST_ROOT_FAILED

    config = ExtractConfig(
        ost_path=ost_path,
        out_dir=out_dir,
        folders=args.folder or [],
        excludes=args.exclude or [],
        since_dt=since_dt,
        until_dt=until_dt,
        max_att_bytes=max_att_bytes,
        require_date=args.require_date,
        include_non_email=args.include_non_email,
        dry_run=args.dry_run,
        compress=args.compress,
        attachments_dir=os.path.abspath(args.attachments_dir) if args.attachments_dir else None,
        codepage=codepage,
        verbose=args.verbose,
        limit=args.limit or 0,
        out_format=args.format,
    )

    if not config.include_non_email:
        LOG.info("Saltando appointments/contacts/tasks por defecto. "
                 "Usa --include-non-email para procesarlos.")

    stats = ExtractStats(errors_log_path=args.errors_log)
    progress = Progress(enabled=not args.verbose)

    try:
        _extract_folder_iter(root, out_dir, config, stats, progress)
    finally:
        progress.finish()
        try:
            pst.close()
        except Exception:
            pass
        stats.close()

    summary = {
        "mode": "eml" if args.format == "eml" else "mbox",
        "ost_path": ost_path,
        "ost_sha256": ost_sha256,
        "out_dir": out_dir,
        "codepage": codepage,
        "dry_run": args.dry_run,
        "compress": args.compress,
        "limit": args.limit or 0,
        **stats.to_dict(),
    }
    if args.errors_log and stats.errors:
        summary["errors_log"] = args.errors_log

    if args.json:
        print(json.dumps(summary, indent=2, ensure_ascii=False))
    else:
        prefix = "(dry-run) " if args.dry_run else ""
        print(
            f"OK: extracción {prefix}completada -> {out_dir}\n"
            f"    SHA-256 OST: {ost_sha256 or 'n/a'}\n"
            f"    formato: {args.format}  codepage: {codepage or 'desconocido'}\n"
            f"    carpetas: {stats.folders}  emails: {stats.items}\n"
            f"    omitidos: rango fecha={stats.skipped_date_range}  "
            f"sin fecha={stats.skipped_no_date}  "
            f"no-email={stats.skipped_class}  "
            f"limit={stats.skipped_limit}\n"
            f"    adjuntos: guardados={stats.attachments_saved}  "
            f"saltados (tamaño)={stats.attachments_skipped}\n"
            f"    errores: {stats.errors}"
        )
        if args.errors_log and stats.errors:
            print(f"    detalle de errores en: {args.errors_log}")

    return EXIT_OK if stats.errors == 0 else EXIT_COMPLETED_WITH_ERRORS


# =========================================================================
# CLI
# =========================================================================

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ost2pst",
        description=(
            "Herramienta para procesar ficheros Outlook OST. Dos modos:\n"
            "  - por defecto: convierte OST a PST con Outlook COM.\n"
            "  - --extract:   extracción a .eml o .mbox con libpff "
            "(sin Outlook)."
        ),
        epilog=(
            "Ejemplos:\n"
            "  ost2pst --list\n"
            "  ost2pst -i C:\\...\\cuenta.ost -o C:\\backup\\cuenta.pst\n"
            "  ost2pst --extract -i C:\\correo.ost -o C:\\extraidos -v\n"
            "  ost2pst --extract -i x.ost -o out --format mbox\n"
            "  ost2pst --extract -i x.ost -o out --limit 100 --json\n"
            "  ost2pst --extract -i x.ost -o out --codepage 1252\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    g_mode = parser.add_argument_group("modo de operación")
    mx = g_mode.add_mutually_exclusive_group()
    mx.add_argument("-x", "--extract", action="store_true",
                    help="Extrae cada email a .eml/.mbox con libpff.")
    mx.add_argument("-l", "--list", action="store_true",
                    help="Lista los stores cargados en Outlook y sale.")

    g_io = parser.add_argument_group("entrada / salida")
    g_io.add_argument("-i", "--input", metavar="OST",
                      help="Ruta al fichero OST (o nombre del store si --by-name).")
    g_io.add_argument("-o", "--output", metavar="DEST",
                      help="PST destino (modo PST) o directorio (modo --extract).")

    g_pst = parser.add_argument_group("opciones modo PST")
    g_pst.add_argument("--by-name", action="store_true",
                       help="Interpreta --input como nombre de store.")
    g_pst.add_argument("-f", "--force", action="store_true",
                       help="Sobrescribe el PST destino si existe.")

    g_filt = parser.add_argument_group("filtros modo --extract")
    g_filt.add_argument("--folder", action="append", metavar="NOMBRE",
                        help="Procesa solo carpetas cuyo nombre contenga este texto "
                             "(repetible).")
    g_filt.add_argument("--exclude", action="append", metavar="NOMBRE",
                        help="Salta carpetas y subárbol (repetible).")
    g_filt.add_argument("--since", metavar="YYYY-MM-DD",
                        help="Mensajes con fecha >= esta (ISO 8601).")
    g_filt.add_argument("--until", metavar="YYYY-MM-DD",
                        help="Mensajes con fecha <= esta (inclusivo).")
    g_filt.add_argument("--max-attachment-mb", type=int, metavar="MB",
                        help="Salta adjuntos mayores de MB megabytes.")
    g_filt.add_argument("--require-date", action="store_true",
                        help="Excluye mensajes sin fecha cuando se usan --since/--until.")
    g_filt.add_argument("--include-non-email", action="store_true",
                        help="Incluye appointments/contacts/tasks (default: solo emails).")
    g_filt.add_argument("--limit", type=int, metavar="N",
                        help="Procesa solo los primeros N mensajes (útil para iterar).")
    g_filt.add_argument("--codepage", type=int, metavar="CP",
                        help="Override del codepage del OST (p.ej. 1252, 1251, 932).")

    g_out = parser.add_argument_group("salidas modo --extract")
    g_out.add_argument("--format", choices=("eml", "mbox"), default="eml",
                       help="Formato de salida: eml (uno por mensaje) o mbox "
                            "(uno por carpeta). Default: eml.")
    g_out.add_argument("--attachments-dir", metavar="DIR",
                       help="Adjuntos también a disco (excluye inline images).")
    g_out.add_argument("--compress", action="store_true",
                       help="Solo --format eml: escribe .eml.gz (gzip nivel 6).")
    g_out.add_argument("--json", action="store_true",
                       help="Resumen final como JSON por stdout.")
    g_out.add_argument("--dry-run", action="store_true",
                       help="No escribe; solo cuenta y reporta.")
    g_out.add_argument("--errors-log", metavar="PATH",
                       help="Errores como JSON Lines (incluye timestamp).")

    g_log = parser.add_argument_group("logging")
    g_log.add_argument("-v", "--verbose", action="store_true",
                       help="Nivel INFO.")
    g_log.add_argument("-q", "--quiet", action="store_true",
                       help="Nivel ERROR (silencia warnings).")
    g_log.add_argument("-d", "--debug", action="store_true",
                       help="Traza línea-a-línea por stderr. Implica --verbose.")

    parser.add_argument("--version", action="version",
                        version=f"%(prog)s {__version__}")
    return parser


def main(argv: list | None = None) -> int:
    _configure_console_utf8()
    # Defensa contra OST con jerarquías profundas (uso iterativo, pero
    # algunos callbacks de pypff pueden anidar en C).
    try:
        sys.setrecursionlimit(max(sys.getrecursionlimit(), 5000))
    except Exception:
        pass

    parser = build_parser()
    effective_argv = sys.argv[1:] if argv is None else list(argv)
    if not effective_argv:
        parser.print_help()
        return EXIT_OK

    args = parser.parse_args(argv)

    if args.debug:
        args.verbose = True
        args.quiet = False

    _configure_logging(verbose=args.verbose, quiet=args.quiet)

    if args.debug:
        _install_debug_tracer()

    if args.extract:
        return run_eml_extract(args)
    return run_pst_export(args)


if __name__ == "__main__":
    sys.exit(main())
