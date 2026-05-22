"""Tests de los lectores de propiedades MAPI con mock de pypff record_sets.

Estructura del mock:
    FakeRecordEntry  → expone .entry_type, .data_as_string, .data
    FakeRecordSet    → expone .entries (lista de FakeRecordEntry)
    MapiOnlyMessage  → expone .record_sets (lista de FakeRecordSet)

Verificamos las dos formas de entry_type que pypff puede devolver:
  - sólo property ID (16 bits)
  - tag completo  (id<<16 | type)
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ost2pst import (
    MAPI_DISPLAY_BCC,
    MAPI_DISPLAY_CC,
    MAPI_DISPLAY_TO,
    MAPI_INTERNET_MESSAGE_ID,
    MAPI_MESSAGE_CLASS,
    MAPI_SENDER_EMAIL,
    _is_email_class,
    _read_internet_message_id,
    _read_mapi_string,
    _read_message_class,
    _read_recipients,
    _read_sender_email,
)


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
        if isinstance(self._value, str):
            return self._value.encode("utf-16-le") + b"\x00\x00"
        return self._value


class FakeRecordSet:
    def __init__(self, entries):
        self._entries = entries

    @property
    def entries(self):
        return self._entries


class MapiOnlyMessage:
    """Sólo expone record_sets (lo que necesitan _read_mapi_*)."""
    def __init__(self, props):
        # props: dict {entry_type_or_propid: string_value}
        entries = [FakeRecordEntry(et, val) for et, val in props.items()]
        self._record_sets = [FakeRecordSet(entries)]

    @property
    def record_sets(self):
        return self._record_sets


# --- _read_mapi_string con entry_type en formato corto (16 bits) ---------

def test_read_mapi_string_short_form():
    msg = MapiOnlyMessage({MAPI_DISPLAY_TO: "alice@x.com; bob@y.com"})
    assert _read_mapi_string(msg, MAPI_DISPLAY_TO) == "alice@x.com; bob@y.com"


def test_read_mapi_string_missing_property():
    msg = MapiOnlyMessage({MAPI_DISPLAY_TO: "alice"})
    assert _read_mapi_string(msg, MAPI_DISPLAY_CC) is None


def test_read_mapi_string_empty_value():
    msg = MapiOnlyMessage({MAPI_DISPLAY_TO: ""})
    # data_as_string vacío → caemos al fallback de data, también vacío → None
    val = _read_mapi_string(msg, MAPI_DISPLAY_TO)
    assert val in (None, "")


# --- _read_mapi_string con entry_type en formato tag completo ------------

def test_read_mapi_string_full_tag_form():
    # PR_DISPLAY_TO con tipo PT_UNICODE (0x001F)
    full_tag = (MAPI_DISPLAY_TO << 16) | 0x001F
    msg = MapiOnlyMessage({full_tag: "alice@x.com"})
    assert _read_mapi_string(msg, MAPI_DISPLAY_TO) == "alice@x.com"


def test_read_mapi_string_full_tag_message_class():
    full_tag = (MAPI_MESSAGE_CLASS << 16) | 0x001E  # PT_STRING8
    msg = MapiOnlyMessage({full_tag: "IPM.Note"})
    assert _read_mapi_string(msg, MAPI_MESSAGE_CLASS) == "IPM.Note"


# --- _read_recipients ----------------------------------------------------

def test_read_recipients_all_three():
    msg = MapiOnlyMessage({
        MAPI_DISPLAY_TO: "to@x.com",
        MAPI_DISPLAY_CC: "cc1@y.com; cc2@y.com",
        MAPI_DISPLAY_BCC: "bcc@z.com",
    })
    r = _read_recipients(msg)
    assert r["To"] == "to@x.com"
    assert r["Cc"] == "cc1@y.com; cc2@y.com"
    assert r["Bcc"] == "bcc@z.com"


def test_read_recipients_only_to():
    msg = MapiOnlyMessage({MAPI_DISPLAY_TO: "to@x.com"})
    r = _read_recipients(msg)
    assert r["To"] == "to@x.com"
    assert r["Cc"] == ""
    assert r["Bcc"] == ""


def test_read_recipients_empty():
    msg = MapiOnlyMessage({})
    r = _read_recipients(msg)
    assert r == {"To": "", "Cc": "", "Bcc": ""}


# --- _read_message_class -------------------------------------------------

def test_read_message_class_ipm_note():
    msg = MapiOnlyMessage({MAPI_MESSAGE_CLASS: "IPM.Note"})
    assert _read_message_class(msg) == "IPM.Note"
    assert _is_email_class(_read_message_class(msg))


def test_read_message_class_appointment():
    msg = MapiOnlyMessage({MAPI_MESSAGE_CLASS: "IPM.Appointment"})
    assert _read_message_class(msg) == "IPM.Appointment"
    assert not _is_email_class(_read_message_class(msg))


def test_read_message_class_missing():
    msg = MapiOnlyMessage({})
    assert _read_message_class(msg) == ""


# --- _read_sender_email --------------------------------------------------

def test_read_sender_email_present():
    msg = MapiOnlyMessage({MAPI_SENDER_EMAIL: "sender@example.com"})
    assert _read_sender_email(msg) == "sender@example.com"


def test_read_sender_email_absent():
    msg = MapiOnlyMessage({})
    assert _read_sender_email(msg) is None


# --- _read_internet_message_id -------------------------------------------

def test_read_internet_message_id_present():
    msg = MapiOnlyMessage({MAPI_INTERNET_MESSAGE_ID: "<abc@example.com>"})
    assert _read_internet_message_id(msg) == "<abc@example.com>"


# --- Robustez frente a excepciones en pypff ------------------------------

class BrokenMessage:
    @property
    def record_sets(self):
        raise RuntimeError("pypff internal error")


def test_read_mapi_string_handles_broken_message():
    assert _read_mapi_string(BrokenMessage(), MAPI_DISPLAY_TO) is None


class MessageWithBrokenEntry:
    def __init__(self):
        class BrokenEntry:
            @property
            def entry_type(self):
                raise OSError("bad data")
        self._record_sets = [FakeRecordSet([BrokenEntry()])]

    @property
    def record_sets(self):
        return self._record_sets


def test_read_mapi_string_skips_broken_entry():
    # No debe crashear; simplemente devuelve None
    msg = MessageWithBrokenEntry()
    assert _read_mapi_string(msg, MAPI_DISPLAY_TO) is None
