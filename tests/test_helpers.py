"""Tests unitarios de los helpers puros (no requieren OST real)."""

import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ost2pst import (
    ExtractConfig,
    ExtractStats,
    _attachment_inline_cid,
    _decode_body,
    _dedupe_filename,
    _ensure_long_path,
    _entry_property_id,
    _extract_cids,
    _guess_mime_type,
    _is_email_class,
    _parse_iso_dt,
    _safe_filename,
    _sha256_file,
)

# --- _safe_filename ------------------------------------------------------

def test_safe_filename_basic():
    assert _safe_filename("hola") == "hola"


def test_safe_filename_forbidden_chars():
    assert _safe_filename("a/b\\c:d") == "a_b_c_d"
    assert _safe_filename('"<>|?*') == "______"


def test_safe_filename_empty_falls_back():
    assert _safe_filename("") == "sin_asunto"
    assert _safe_filename("   ") == "sin_asunto"


def test_safe_filename_truncates():
    assert _safe_filename("x" * 200, maxlen=10) == "x" * 10


def test_safe_filename_collapses_whitespace():
    assert _safe_filename("foo\nbar\r\n  baz") == "foo bar baz"


def test_safe_filename_strips_dots():
    assert _safe_filename(" .archivo. ") == "archivo"


# --- _decode_body --------------------------------------------------------

def test_decode_body_none():
    assert _decode_body(None) == ""


def test_decode_body_str_passthrough():
    assert _decode_body("hola") == "hola"


def test_decode_body_utf8():
    assert _decode_body("hóla".encode()) == "hóla"


def test_decode_body_codepage_hint():
    raw = "café".encode("cp1252")
    assert _decode_body(raw, codepage=1252) == "café"


def test_decode_body_cyrillic_via_codepage():
    raw = "Привет".encode("cp1251")
    assert _decode_body(raw, codepage=1251) == "Привет"


def test_decode_body_garbage_fallback_no_crash():
    # bytes que no decodifican limpiamente en utf-8/utf-16-le; chardet o
    # latin-1 deben dar algo no-vacío sin crashear
    assert _decode_body(b"\xe1\xe9\xff\xfe random bytes \x80") != ""


# --- _guess_mime_type ----------------------------------------------------

def test_guess_mime_pdf():
    assert _guess_mime_type("doc.pdf") == ("application", "pdf")


def test_guess_mime_png():
    assert _guess_mime_type("img.png") == ("image", "png")


def test_guess_mime_unknown():
    assert _guess_mime_type("file.xyz123") == ("application", "octet-stream")


def test_guess_mime_empty():
    assert _guess_mime_type("") == ("application", "octet-stream")


# --- _extract_cids -------------------------------------------------------

def test_extract_cids_basic():
    html = '<img src="cid:image001@01D2A5F4">'
    assert "image001@01D2A5F4" in _extract_cids(html)


def test_extract_cids_multiple():
    html = "<img src='cid:a'><img src=\"cid:b@x\">"
    assert _extract_cids(html) == {"a", "b@x"}


def test_extract_cids_empty():
    assert _extract_cids("") == set()
    assert _extract_cids(None) == set()


# --- _attachment_inline_cid ----------------------------------------------

def test_inline_cid_exact_basename():
    cids = {"image001@01D2"}
    assert _attachment_inline_cid("image001.png", cids) == "image001@01D2"


def test_inline_cid_no_match():
    cids = {"foo@bar"}
    assert _attachment_inline_cid("unrelated.pdf", cids) is None


def test_inline_cid_strict_no_substring_match():
    """v3.1: matching estricto. 'abcdef.png' NO debe matchear con 'cid:abc'."""
    cids = {"abc@bar"}
    assert _attachment_inline_cid("abcdef.png", cids) is None


def test_inline_cid_strict_no_partial_in_either_direction():
    cids = {"longname@xx"}
    assert _attachment_inline_cid("long.png", cids) is None


def test_inline_cid_empty_inputs():
    assert _attachment_inline_cid("", {"a"}) is None
    assert _attachment_inline_cid("foo.png", set()) is None


# --- _dedupe_filename ----------------------------------------------------

def test_dedupe_first_use():
    used = set()
    assert _dedupe_filename("doc.pdf", used) == "doc.pdf"
    assert "doc.pdf" in used


def test_dedupe_collision():
    used = {"doc.pdf"}
    assert _dedupe_filename("doc.pdf", used) == "doc (2).pdf"
    assert _dedupe_filename("doc.pdf", used) == "doc (3).pdf"


def test_dedupe_no_extension():
    used = {"README"}
    assert _dedupe_filename("README", used) == "README (2)"


# --- _parse_iso_dt -------------------------------------------------------

def test_parse_iso_dt_date_only():
    dt = _parse_iso_dt("2024-01-15")
    assert dt.year == 2024 and dt.month == 1 and dt.day == 15
    assert dt.hour == 0


def test_parse_iso_dt_end_of_day():
    dt = _parse_iso_dt("2024-01-15", end_of_day=True)
    assert (dt.hour, dt.minute, dt.second) == (23, 59, 59)


def test_parse_iso_dt_none():
    assert _parse_iso_dt(None) is None
    assert _parse_iso_dt("") is None


def test_parse_iso_dt_full():
    dt = _parse_iso_dt("2024-06-01T15:30:00")
    assert dt.hour == 15 and dt.minute == 30


# --- _ensure_long_path ---------------------------------------------------

def test_ensure_long_path_short_unchanged():
    p = "C:\\short\\path.txt"
    assert _ensure_long_path(p) == p


def test_ensure_long_path_long_wrapped():
    if os.name != "nt":
        return
    long_path = "C:\\" + ("a" * 50 + "\\") * 6 + "file.txt"
    assert len(os.path.abspath(long_path)) > 240
    assert _ensure_long_path(long_path).startswith("\\\\?\\")


def test_ensure_long_path_empty():
    assert _ensure_long_path("") == ""


# --- _is_email_class -----------------------------------------------------

def test_is_email_class_ipm_note():
    assert _is_email_class("IPM.Note") is True
    assert _is_email_class("ipm.note") is True
    assert _is_email_class("IPM.Note.Receipt") is True


def test_is_email_class_appointment():
    assert _is_email_class("IPM.Appointment") is False
    assert _is_email_class("IPM.Contact") is False
    assert _is_email_class("IPM.Task") is False


def test_is_email_class_empty_default_true():
    assert _is_email_class("") is True
    assert _is_email_class(None) is True


# --- _sha256_file --------------------------------------------------------

def test_sha256_file_known_value():
    with tempfile.NamedTemporaryFile(delete=False) as f:
        f.write(b"hola mundo")
        path = f.name
    try:
        expected = "0b894166d3336435c800bea36ff21b29eaa801a52f584c006c49289a0dcf6e2f"
        assert _sha256_file(path) == expected
    finally:
        os.unlink(path)


# --- _entry_property_id --------------------------------------------------

def test_entry_property_id_short():
    # Si entry_type ya es 16 bits, devuelve tal cual
    assert _entry_property_id(0x0E04) == 0x0E04


def test_entry_property_id_full_tag():
    # Si es el tag completo (id<<16 | type), extrae el id
    assert _entry_property_id(0x0E04001F) == 0x0E04  # PT_UNICODE
    assert _entry_property_id(0x001A001E) == 0x001A  # PT_STRING8


# --- ExtractStats --------------------------------------------------------

def test_stats_counters_start_zero():
    s = ExtractStats()
    assert s.folders == 0 and s.items == 0 and s.errors == 0
    assert s.attachments_saved == 0 and s.attachments_skipped == 0
    assert s.skipped_date_range == 0 and s.skipped_no_date == 0
    assert s.skipped_class == 0 and s.skipped_limit == 0


def test_stats_record_error_increments():
    s = ExtractStats()
    s.record_error(stage="test", error="boom")
    assert s.errors == 1


def test_stats_to_dict_has_expected_keys():
    s = ExtractStats()
    d = s.to_dict()
    assert set(d.keys()) >= {
        "folders", "items", "skipped_date_range", "skipped_no_date",
        "skipped_class", "skipped_limit",
        "attachments_saved", "attachments_skipped", "errors",
    }


def test_stats_errors_log_writes_jsonl_with_timestamp():
    with tempfile.NamedTemporaryFile(suffix=".jsonl", delete=False) as f:
        path = f.name
    try:
        s = ExtractStats(errors_log_path=path)
        s.record_error(stage="message", folder="x", error="boom")
        s.record_error(stage="attachment", attachment="y.pdf")
        s.close()
        with open(path, encoding="utf-8") as f:
            lines = [json.loads(line) for line in f if line.strip()]
        assert len(lines) == 2
        assert lines[0]["stage"] == "message"
        assert "ts" in lines[0]  # v3.1: timestamp añadido
        assert "ts" in lines[1]
        assert lines[1]["attachment"] == "y.pdf"
    finally:
        os.unlink(path)


# --- ExtractConfig -------------------------------------------------------

def test_extract_config_defaults():
    c = ExtractConfig(ost_path="x.ost", out_dir="out")
    assert c.folders == [] and c.excludes == []
    assert c.since_dt is None and c.until_dt is None
    assert c.max_att_bytes == 0
    assert c.require_date is False
    assert c.include_non_email is False
    assert c.dry_run is False
    assert c.compress is False
    assert c.attachments_dir is None
    assert c.limit == 0
    assert c.out_format == "eml"
