"""Windows OCR as the first OCR step, and the way up from it: ocr='server'
when the local text is not good enough, ocr='cloud' only after a yes."""
import sys

import pytest

import pdfread
from conftest import UNREACHABLE
from pdfbuilder import build_pdf

SCANS = build_pdf(["Egy szöveges oldal, amelynek van rendes szövegrétege.", None, None])


def test_auto_reads_a_scan_on_this_machine(fake_local_ocr, fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    r = pdfread.read_pages(SCANS)
    assert r["sources"] == {0: "text-layer", 1: "windows-ocr", 2: "windows-ocr"}
    assert fake_ocr.requests == []                       # the server was not needed
    assert r["ocr_upgrade"]["pages"] == "2-3"            # ...but is offered as the better reading
    assert "ocr='server'" in r["ocr_upgrade"]["how"]
    assert r["cloud_offer"] is None                      # nothing is empty, so nothing to offer


def test_no_upgrade_hint_without_a_server(fake_local_ocr):
    r = pdfread.read_pages(SCANS)
    assert r["sources"][1] == "windows-ocr" and r["ocr_upgrade"] is None


def test_server_on_request(fake_local_ocr, fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    r = pdfread.read_pages(SCANS, ocr="server", pages="2")
    assert r["sources"] == {1: "ocr-server"}
    assert fake_local_ocr.calls == []
    assert "Tétel\tÖsszeg" in r["texts"][1]              # the server keeps table structure


def test_server_unreachable_falls_back_to_windows(fake_local_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", UNREACHABLE)
    r = pdfread.read_pages(SCANS, ocr="server")
    assert r["sources"][1] == "windows-ocr"
    assert any("unreachable" in n for n in r["notes"])


def test_windows_without_the_language_falls_back_to_the_server(fake_local_ocr, fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    monkeypatch.setenv("OCR_LOCAL_LANGUAGE", "xx")
    fake_local_ocr.error = pdfread.EngineUnavailable("Windows OCR has no 'xx' language here")
    r = pdfread.read_pages(SCANS)
    assert fake_local_ocr.calls == ["xx"]                 # tried once, then left alone
    assert r["sources"][1] == "ocr-server" and r["sources"][2] == "ocr-server"
    assert any("no 'xx' language" in n for n in r["notes"])


def test_windows_empty_page_goes_on_to_the_server(fake_local_ocr, fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    fake_local_ocr.reply = "  "                          # a blank or unreadable page
    r = pdfread.read_pages(SCANS)
    assert r["sources"][1] == "ocr-server"


def test_cloud_offer_only_when_nothing_local_could_read(fake_local_ocr, fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_CLOUD_URL", fake_ocr.url)
    monkeypatch.setenv("OCR_CLOUD_MODEL", "vision")
    monkeypatch.setenv("OCR_CLOUD_API_KEY", "k")
    fake_local_ocr.reply = ""
    r = pdfread.read_pages(SCANS)
    assert r["cloud_offer"]["pages"] == "2-3"
    assert fake_ocr.calls("/responses") == []


def test_local_off_and_never(fake_local_ocr, monkeypatch):
    assert pdfread.read_pages(SCANS, ocr="never")["sources"][1] == ""
    monkeypatch.setenv("OCR_LOCAL", "off")
    monkeypatch.setattr(pdfread, "local_settings", lambda: None)   # what "off" makes it return
    r = pdfread.read_pages(SCANS)
    assert r["sources"][1] == "" and any("No OCR engine" in n for n in r["notes"])
    assert fake_local_ocr.calls == []


def test_local_settings_respects_off(monkeypatch):
    monkeypatch.setenv("OCR_LOCAL", "off")
    assert pdfread.local_settings() is None


def test_describe_settings(fake_local_ocr, monkeypatch):
    monkeypatch.setenv("OCR_LOCAL_LANGUAGE", "hu")
    assert pdfread.describe_settings()["ocr_local"] == {"available": True, "engine": "Windows OCR",
                                                        "language": "hu"}


def _windows_ocr_here():
    if sys.platform != "win32":
        return False
    import os
    os.environ.pop("OCR_LOCAL", None)
    try:
        return pdfread.local_settings() is not None
    finally:
        os.environ["OCR_LOCAL"] = "off"


@pytest.mark.skipif(not _windows_ocr_here(), reason="needs Windows with pywinrt")
def test_real_windows_ocr_reads_a_rendered_page(monkeypatch):
    """The real engine, on a page this test draws and renders: no fixture file."""
    monkeypatch.delenv("OCR_LOCAL", raising=False)
    doc = pdfread.open_pdf(build_pdf(["Integrity Authority quarterly report\n"
                                      "Procurement review of the second quarter"]))   # one page, two lines
    try:
        png = pdfread.render_png(doc, 0, pdfread.LOCAL_DPI, pdfread.LOCAL_MAX_PIXELS)
    finally:
        doc.close()
    text = pdfread.local_ocr_page(png, "")
    assert "quarterly report" in text.lower() and "procurement" in text.lower()
