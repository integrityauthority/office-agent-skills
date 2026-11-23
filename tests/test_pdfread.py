"""shared/pdfread.py: the PDF reading order both skills rely on."""
import time

import pytest

import microsoft_graph
import pdfread
from conftest import UNREACHABLE
from pdfbuilder import build_pdf

TEXT_PAGE = "The résumé screening rules\nand the performance review method\nsecond line"
TWO_SCANS = build_pdf([TEXT_PAGE, None, None])


def test_ranges_and_pages():
    assert pdfread.ranges([0, 1, 2, 5, 7, 8]) == "1-3,6,8-9"
    assert pdfread.parse_pages("1-3,7", 10) == [0, 1, 2, 6]
    assert pdfread.parse_pages("8-", 10) == [7, 8, 9]
    assert pdfread.parse_pages(None, 3) == [0, 1, 2]
    assert pdfread.parse_pages("99", 3) == []


def test_text_layer_stays_local(fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    r = pdfread.read_pages(build_pdf([TEXT_PAGE]))
    assert r["sources"] == {0: "text-layer"}
    assert "résumé screening" in r["texts"][0]
    assert fake_ocr.requests == []          # a text page never goes to OCR


def test_scan_without_any_ocr_configured():
    r = pdfread.read_pages(TWO_SCANS)
    assert r["sources"] == {0: "text-layer", 1: "", 2: ""}
    assert r["cloud_offer"] is None         # nothing could read them, so nothing to offer
    assert any("OCR_URL is empty" in n for n in r["notes"])


def test_scan_goes_to_the_ocr_server_and_is_cleaned(fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    r = pdfread.read_pages(TWO_SCANS)
    assert r["sources"] == {0: "text-layer", 1: "ocr-server", 2: "ocr-server"}
    text = r["texts"][1]
    assert "<|" not in text and "<table" not in text
    assert "Item\tAmount" in text and "A\t100" in text
    assert len(fake_ocr.calls("/chat/completions")) == 2


def test_deepseek_prompt_only_for_deepseek(fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    pdfread.read_pages(build_pdf([None]))
    body = fake_ocr.requests[-1]["body"]
    assert body["model"] == pdfread.OCR_MODEL_DEFAULT
    assert body["messages"][0]["content"][0]["text"] == pdfread.DEEPSEEK_OCR_PROMPT
    assert body["skip_special_tokens"] is False

    monkeypatch.setenv("OCR_MODEL", "some/other-vision-model")
    pdfread.read_pages(build_pdf([None]))
    body = fake_ocr.requests[-1]["body"]
    assert body["messages"][0]["content"][0]["text"] == pdfread.TRANSCRIBE_PROMPT
    assert "skip_special_tokens" not in body


def test_cloud_is_offered_not_used(fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", UNREACHABLE)
    monkeypatch.setenv("OCR_CLOUD_URL", fake_ocr.url)
    monkeypatch.setenv("OCR_CLOUD_MODEL", "vision")
    monkeypatch.setenv("OCR_CLOUD_API_KEY", "k")
    started = time.monotonic()
    r = pdfread.read_pages(TWO_SCANS)
    assert time.monotonic() - started < 10       # unreachable server: no timeout per page
    assert r["cloud_offer"]["pages"] == "2-3"
    assert fake_ocr.calls("/responses") == []    # offered, never sent
    assert any("unreachable" in n for n in r["notes"])


def test_cloud_after_approval_goes_to_the_cloud(fake_ocr, monkeypatch):
    """ocr='cloud' is the user's answer to "the local text is not good enough"."""
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    monkeypatch.setenv("OCR_CLOUD_URL", fake_ocr.url)
    monkeypatch.setenv("OCR_CLOUD_MODEL", "vision")
    monkeypatch.setenv("OCR_CLOUD_API_KEY", "k")
    r = pdfread.read_pages(TWO_SCANS, ocr="cloud")
    assert set(r["sources"].values()) == {"text-layer", "cloud-vision"}
    assert fake_ocr.calls("/chat/completions") == []


def test_cloud_with_a_key_held_elsewhere(fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", UNREACHABLE)
    monkeypatch.setenv("OCR_CLOUD_URL", fake_ocr.url)
    monkeypatch.setenv("OCR_CLOUD_MODEL", "vision")
    monkeypatch.setenv("SOME_SECRET_HOLDER", "secret")
    monkeypatch.setenv("OCR_CLOUD_API_KEY_ENV", "SOME_SECRET_HOLDER")
    r = pdfread.read_pages(TWO_SCANS, ocr="cloud")
    assert r["sources"] == {0: "text-layer", 1: "cloud-vision", 2: "cloud-vision"}
    assert r["cloud_offer"] is None
    calls = fake_ocr.calls("/responses")
    assert len(calls) == 2
    assert calls[0]["headers"]["Authorization"] == "Bearer secret"   # key via OCR_CLOUD_API_KEY_ENV


def test_never_and_force(fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    assert pdfread.read_pages(TWO_SCANS, ocr="never")["sources"][1] == ""
    assert fake_ocr.requests == []
    r = pdfread.read_pages(TWO_SCANS, ocr="force")
    assert r["sources"][0] == "ocr-server"       # even the text page
    with pytest.raises(pdfread.PdfError):
        pdfread.read_pages(TWO_SCANS, ocr="sometimes")


def test_ocr_is_capped_per_call(fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    r = pdfread.read_pages(build_pdf([None, None, None]), ocr_limit=2)
    assert r["ocr_deferred"] == [2]
    assert r["sources"][2] == ""
    assert any("capped per call" in n and "3" in n for n in r["notes"])


def test_auth_headers():
    assert pdfread.auth_headers("https://example-resource.openai.azure.com/openai/v1", "k") == {"api-key": "k"}
    assert pdfread.auth_headers("https://api.example.com/v1", "k") == {"Authorization": "Bearer k"}
    assert pdfread.auth_headers("http://srv/v1", "") == {}


def test_server_key_is_sent(fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    monkeypatch.setenv("OCR_API_KEY", "server-key")
    pdfread.read_pages(build_pdf([None]))
    assert fake_ocr.requests[-1]["headers"]["Authorization"] == "Bearer server-key"


def test_a_short_text_layer_is_kept_when_ocr_finds_nothing_better():
    """A cover page with one word is under TEXT_MIN and goes to OCR like a
    scan; with no OCR available its own text must still come back."""
    r = pdfread.read_pages(build_pdf(["Appendix"]))
    assert r["sources"] == {0: "text-layer"} and r["texts"][0] == "Appendix"


def test_errors():
    with pytest.raises(pdfread.PdfError):
        pdfread.read_pages(build_pdf([TEXT_PAGE]), pages="5")
    with pytest.raises(pdfread.PdfError):
        pdfread.open_pdf(b"%PDF-1.4 this is not really a pdf")


@pytest.mark.parametrize("env", ["server", "server-down-cloud", "nothing", "windows+server"])
def test_both_skills_read_the_same_way(fake_ocr, monkeypatch, env, request):
    """file-reader and m365-graph must give the same sources and the same offer."""
    import files
    if env == "server":
        monkeypatch.setenv("OCR_URL", fake_ocr.url)
    elif env == "windows+server":
        request.getfixturevalue("fake_local_ocr")
        monkeypatch.setenv("OCR_URL", fake_ocr.url)
    elif env == "server-down-cloud":
        monkeypatch.setenv("OCR_URL", UNREACHABLE)
        monkeypatch.setenv("OCR_CLOUD_URL", fake_ocr.url)
        monkeypatch.setenv("OCR_CLOUD_MODEL", "vision")
        monkeypatch.setenv("OCR_CLOUD_API_KEY", "k")
    via_files = files.read_pdf(TWO_SCANS)
    text, source, _note, offer, upgrade = microsoft_graph._read_pdf(TWO_SCANS)
    used = {s for s in via_files["page_sources"] if s != "no-text"}
    assert set(source.split("+")) - {""} == used
    assert (offer or {}).get("pages") == (via_files.get("cloud_ocr_needs_approval") or {}).get("pages")
    assert ("--- page 1 ---" in text) == ("--- page 1 ---" in via_files["text"])
    assert (upgrade or {}).get("pages") == (via_files.get("ocr_upgrade") or {}).get("pages")
