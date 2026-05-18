"""Tests funcionales de build_eml() con mocks de pypff."""

import email
import email.policy
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ost2pst import (
    MAPI_DISPLAY_BCC,
    MAPI_DISPLAY_CC,
    MAPI_DISPLAY_TO,
    MAPI_MESSAGE_CLASS,
    MAPI_SENDER_EMAIL,
    ExtractConfig,
    ExtractStats,
    build_eml,
)


class FakeAttachment:
    def __init__(self, name, data, size=None):
        self._name = name
        self._data = data
        self._size = size if size is not None else len(data)

    def get_name(self):
        return self._name

    def get_size(self):
        return self._size

    def read_buffer(self, n):
        return self._data[:n]


class FakeRecordEntry:
    def __init__(self, entry_type, value):
        self._entry_type = entry_type
        self._value = value

    @property
    def entry_type(self):
        return self._entry_type

    @property
    def data_as_string(self):
        return self._value

    @property
    def data(self):
        return self._value


class FakeRecordSet:
    def __init__(self, entries):
        self._entries = entries

    @property
    def entries(self):
        return self._entries


class FakeMessage:
    def __init__(self, subject="", sender="", plain="", html="", rtf=None,
                 transport_headers="", attachments=None, mapi_props=None):
        self._subject = subject
        self._sender = sender
        self._plain = plain
        self._html = html
        self._rtf = rtf
        self._transport = transport_headers
        self._attachments = attachments or []
        entries = [FakeRecordEntry(et, v) for et, v in (mapi_props or {}).items()]
        self._record_sets = [FakeRecordSet(entries)]

    def get_subject(self):
        return self._subject

    def get_sender_name(self):
        return self._sender

    def get_plain_text_body(self):
        return self._plain

    def get_html_body(self):
        return self._html

    def get_rtf_body(self):
        return self._rtf

    def get_transport_headers(self):
        return self._transport

    def get_number_of_attachments(self):
        return len(self._attachments)

    def get_attachment(self, i):
        return self._attachments[i]

    @property
    def record_sets(self):
        return self._record_sets


def _parse(raw_bytes):
    return email.message_from_bytes(raw_bytes, policy=email.policy.default)


# --- Cabeceras básicas ----------------------------------------------------

def test_build_eml_subject_from_pypff():
    msg = FakeMessage(subject="Hola mundo", sender="Pepe")
    parsed = _parse(build_eml(msg))
    assert parsed["Subject"] == "Hola mundo"
    assert "Pepe" in parsed["From"]


def test_build_eml_sender_with_mapi_email():
    msg = FakeMessage(
        subject="x",
        sender="Pepe Pérez",
        mapi_props={MAPI_SENDER_EMAIL: "pepe@example.com"},
    )
    parsed = _parse(build_eml(msg))
    assert "Pepe" in parsed["From"]
    assert "pepe@example.com" in parsed["From"]


def test_build_eml_preserves_transport_headers():
    transport = (
        "Received: from a by b ; Mon, 1 Jan 2024 00:00:00 +0000\r\n"
        "Received: from c by d ; Mon, 1 Jan 2024 00:00:01 +0000\r\n"
        "Subject: Asunto original\r\n"
        "Message-ID: <abc@example.com>\r\n"
    )
    msg = FakeMessage(transport_headers=transport, subject="ignored")
    parsed = _parse(build_eml(msg))
    assert parsed["Subject"] == "Asunto original"
    assert parsed["Message-ID"] == "<abc@example.com>"
    assert len(parsed.get_all("Received")) == 2


def test_build_eml_received_multi_instance():
    transport = (
        "Received: from a\r\nReceived: from b\r\nReceived: from c\r\nSubject: x\r\n"
    )
    msg = FakeMessage(transport_headers=transport)
    parsed = _parse(build_eml(msg))
    assert len(parsed.get_all("Received")) == 3


# --- Recipientes desde MAPI (v3.0) ---------------------------------------

def test_build_eml_recipients_from_mapi_when_no_transport():
    msg = FakeMessage(
        subject="x",
        mapi_props={
            MAPI_DISPLAY_TO: "alice@example.com",
            MAPI_DISPLAY_CC: "bob@example.com; carol@example.com",
            MAPI_DISPLAY_BCC: "secret@example.com",
        },
    )
    parsed = _parse(build_eml(msg))
    assert parsed["To"] == "alice@example.com"
    assert "bob@example.com" not in (parsed["To"] or "")
    assert "bob@example.com" in (parsed["Cc"] or "")
    assert "carol@example.com" in (parsed["Cc"] or "")
    assert parsed["Bcc"] == "secret@example.com"


def test_build_eml_transport_to_takes_precedence_over_mapi():
    transport = "To: from-transport@x.com\r\nSubject: x\r\n"
    msg = FakeMessage(
        transport_headers=transport,
        mapi_props={MAPI_DISPLAY_TO: "from-mapi@x.com"},
    )
    parsed = _parse(build_eml(msg))
    assert "from-transport" in parsed["To"]
    assert "from-mapi" not in parsed["To"]


# --- Cuerpo ---------------------------------------------------------------

def test_build_eml_text_plain_only():
    msg = FakeMessage(subject="x", plain="Hola texto plano")
    parsed = _parse(build_eml(msg))
    body = parsed.get_body(preferencelist=("plain",))
    assert "Hola texto plano" in body.get_content()


def test_build_eml_html_only_no_empty_plain():
    msg = FakeMessage(subject="x", html="<p>Hola HTML</p>")
    parsed = _parse(build_eml(msg))
    body = parsed.get_body(preferencelist=("html",))
    assert "Hola HTML" in body.get_content()
    plain = parsed.get_body(preferencelist=("plain",))
    assert plain is None


def test_build_eml_multipart_alternative():
    msg = FakeMessage(subject="x", plain="texto", html="<p>html</p>")
    parsed = _parse(build_eml(msg))
    assert parsed.is_multipart()
    types = [p.get_content_type() for p in parsed.iter_parts()]
    assert "text/plain" in types and "text/html" in types


def test_build_eml_rtf_fallback():
    rtf = (r"{\rtf1\ansi solo cuerpo RTF aqui }").encode("ascii")
    msg = FakeMessage(subject="x", rtf=rtf)
    parsed = _parse(build_eml(msg))
    body = parsed.get_body(preferencelist=("plain",))
    assert body is not None
    assert "RTF" in body.get_content()


# --- Adjuntos -------------------------------------------------------------

def test_build_eml_attaches_pdf():
    att = FakeAttachment("factura.pdf", b"%PDF-1.4 fake")
    msg = FakeMessage(subject="x", plain="hi", attachments=[att])
    parsed = _parse(build_eml(msg))
    atts = list(parsed.iter_attachments())
    assert len(atts) == 1
    assert atts[0].get_filename() == "factura.pdf"
    assert atts[0].get_content_type() == "application/pdf"


def test_build_eml_dedupes_attachment_names():
    a1 = FakeAttachment("doc.pdf", b"first")
    a2 = FakeAttachment("doc.pdf", b"second")
    msg = FakeMessage(subject="x", plain="hi", attachments=[a1, a2])
    parsed = _parse(build_eml(msg))
    names = [p.get_filename() for p in parsed.iter_attachments()]
    assert names == ["doc.pdf", "doc (2).pdf"]


def test_build_eml_skips_too_big_attachments():
    big = FakeAttachment("video.mp4", b"\0" * 50)
    small = FakeAttachment("doc.pdf", b"hi")
    msg = FakeMessage(subject="x", plain="hi", attachments=[big, small])
    config = ExtractConfig(ost_path="", out_dir="", max_att_bytes=10)
    stats = ExtractStats()
    parsed = _parse(build_eml(msg, config=config, stats=stats))
    names = [p.get_filename() for p in parsed.iter_attachments()]
    assert "doc.pdf" in names
    assert "video.mp4" not in names
    assert stats.attachments_saved == 1
    assert stats.attachments_skipped == 1
    assert stats.errors == 1


def test_build_eml_inline_image_cid_strict_match():
    html = '<p>Logo: <img src="cid:image001@xx"></p>'
    att = FakeAttachment("image001.png", b"\x89PNG fake")
    msg = FakeMessage(subject="x", html=html, attachments=[att])
    parsed = _parse(build_eml(msg))
    related = [p for p in parsed.walk() if p.get_content_type() == "image/png"]
    assert len(related) == 1
    assert "inline" in (related[0].get("Content-Disposition") or "")
    assert "image001@xx" in (related[0].get("Content-ID", ""))


def test_build_eml_inline_no_false_positive_substring():
    """v3.1: 'cid:abc' NO debe matchear con 'abcdef.png'."""
    html = '<p><img src="cid:abc@xx"></p>'
    att = FakeAttachment("abcdef.png", b"\x89PNG fake")
    msg = FakeMessage(subject="x", html=html, attachments=[att])
    parsed = _parse(build_eml(msg))
    img = [p for p in parsed.walk() if p.get_content_type() == "image/png"]
    assert len(img) == 1
    # Debería ser attachment (no inline), porque el match estricto no aplica
    cd = img[0].get("Content-Disposition") or ""
    assert "inline" not in cd
