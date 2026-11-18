"""msgfile.py: the Outlook .msg reader, on messages built by msgbuilder."""
import datetime as dt

import pytest

import msgfile
from msgbuilder import build_cfb, build_msg

SENT = dt.datetime(2026, 3, 1, 9, 30, tzinfo=dt.timezone.utc)
BIG = bytes(range(256)) * 40          # 10 KB: stored outside the mini stream

# The compressed-RTF example from [MS-OXRTFCP] section 3.1.1, which
# decompresses to {\rtf1\ansi\ansicpg1252\pard hello world}\r\n
SPEC_RTF = bytes.fromhex(
    "2d000000 2b000000 4c5a4675 f1c5c7a7 03000a00 72637067 31323542 320af320"
    "68656c09 00206277 05b06c64 7d0a800f a0")


def sample(**overrides):
    kwargs = dict(
        subject="Árvíztűrő tükörfúrógép",
        body="Szia,\nez a törzs.\n",
        sender=("Kovács Anna", "anna@example.com"),
        to=[("Bea", "bea@example.com")],
        cc=[("Cili", "cili@example.com")],
        sent=SENT,
        attachments=[
            {"name": "logo.png", "mime": "image/png", "content_id": "logo@x", "data": b"\x89PNG" + b"0" * 50},
            {"name": "melléklet.pdf", "mime": "application/pdf", "content_id": "x@y", "data": b"%PDF-1.4 x"},
            {"name": "big.bin", "data": BIG},
            {"name": "továbbított", "message": {"subject": "Belső levél", "body": "belső törzs"}},
            {"name": "photo", "ole": True},
        ],
    )
    kwargs.update(overrides)
    return build_msg(**kwargs)


def test_headers_and_body():
    s = msgfile.open_msg(sample()).summary()
    assert s["subject"] == "Árvíztűrő tükörfúrógép"
    assert s["from"] == "Kovács Anna <anna@example.com>"
    assert s["to"] == ["Bea <bea@example.com>"]
    assert s["cc"] == ["Cili <cili@example.com>"]
    assert dt.datetime.fromisoformat(s["date"]) == SENT
    assert s["body"] == "Szia,\nez a törzs.\n" and s["body_source"] == "plain"
    assert s["message_class"] == "IPM.Note"


def test_attachments_listing():
    rows = {a["name"]: a for a in msgfile.open_msg(sample()).attachments()}
    assert rows["logo.png"]["inline"] is True             # picture with a content id
    assert rows["melléklet.pdf"]["inline"] is False       # a content id alone is not enough
    assert rows["big.bin"]["size"] == len(BIG)
    assert rows["továbbított"]["embedded_email"] is True
    assert rows["photo"].get("ole_object") is True


def test_attachment_payloads():
    msg = msgfile.open_msg(sample())
    assert msg.attachment(3) == ("big.bin", BIG)
    name, inner = msg.attachment(4)
    assert isinstance(inner, msgfile.Msg)
    assert inner.summary()["subject"] == "Belső levél"
    with pytest.raises(msgfile.MsgError, match="OLE object"):
        msg.attachment(5)
    with pytest.raises(msgfile.MsgError, match="no attachment 9"):
        msg.attachment(9)


def test_8bit_strings_use_the_message_codepage():
    data = build_msg(subject="Őszi ülés – űrlap", body="Tárgyalás", unicode=False, codepage=1250)
    s = msgfile.open_msg(data).summary()
    assert s["subject"] == "Őszi ülés – űrlap"


def test_html_body_when_there_is_no_plain_text():
    data = build_msg(subject="h", html="<html><body><p>Hello <b>world</b></p></body></html>")
    text, source = msgfile.open_msg(data).body()
    assert source == "html" and "<b>world</b>" in text


def test_compressed_rtf_body():
    assert msgfile.decompress_rtf(SPEC_RTF) == b"{\\rtf1\\ansi\\ansicpg1252\\pard hello world}\r\n"
    data = build_msg(subject="rtf", rtf=SPEC_RTF)
    assert msgfile.open_msg(data).body() == ("hello world", "rtf")


def test_rtf_to_text_handles_unicode_and_html_in_rtf():
    rtf = (rb"{\rtf1\ansi\ansicpg1250\uc1 {\fonttbl{\f0 Arial;}}"
           rb"\pard Gy\'f5r \u337?r\par m\'e1sodik\tab sor"
           # A space after a control word is its delimiter, not text; Outlook
           # writes the visible spaces as plain characters, as here.
           rb" {\*\htmltag84 <b>}\htmlrtf {\b\htmlrtf0 vastag\htmlrtf }\htmlrtf0  v\'e9ge}")
    assert msgfile.rtf_to_text(rtf) == "Győr őr\nmásodik\tsor vastag vége"


def test_uncompressed_rtf_container():
    raw = b"{\\rtf1 plain}"
    container = (len(raw) + 12).to_bytes(4, "little") + len(raw).to_bytes(4, "little") \
        + bytes.fromhex("4d454c41") + b"\0\0\0\0" + raw
    assert msgfile.decompress_rtf(container) == raw


def test_is_msg():
    assert msgfile.is_msg(sample())
    assert not msgfile.is_msg(build_cfb({"WordDocument": b"x" * 10}))   # an OLE file, not a message
    assert not msgfile.is_msg(b"%PDF-1.4")
    with pytest.raises(msgfile.MsgError):
        msgfile.open_msg(b"not an ole file at all")


def test_non_mail_items_are_flagged():
    import files
    result = files._email_result(msgfile.open_msg(build_msg(subject="Anna", body="notes",
                                                            message_class="IPM.Contact")).summary(),
                                 "msg", 1000)
    assert "IPM.Contact" in result["notes"][0]
