"""local-web-reader: the keyless web_extract backend for Hermes. No network:
the address check refuses before any request, and every test of the reading
pipeline replaces `fetch` with a stand-in."""
from __future__ import annotations

import importlib.util
import sys
import types
import urllib.error
from pathlib import Path

import pytest

from pdfbuilder import build_pdf

PLUGIN = Path(__file__).resolve().parent.parent / "plugins" / "local-web-reader"
if str(PLUGIN / "scripts") not in sys.path:
    sys.path.insert(0, str(PLUGIN / "scripts"))
import web_reader  # noqa: E402


def load_plugin():
    spec = importlib.util.spec_from_file_location("plugin_local_web_reader", PLUGIN / "__init__.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


PAGE = """<!doctype html><html><head><meta charset="utf-8"><title>Riverside &amp; partners</title>
<style>.x{color:red}</style><script>var a = 1;</script></head>
<body>
<header><nav><a href="/">Home</a> <a href="/news">News</a></nav></header>
<div class="cookie-banner">We use cookies. Accept all</div>
<main>
  <h1>The riverside boardwalk project</h1>
  <p>The town council opened the boardwalk in 2022; the contractor was Example Builders Ltd.</p>
  <p>The project was funded by a regional development grant of 1.2 million euros.</p>
  <ul><li>Opened: 2022</li><li>Length: 80 metres</li></ul>
  <table><tr><th>Year</th><th>Amount</th></tr><tr><td>2022</td><td>EUR 1.2 M</td></tr></table>
  <div hidden>Hidden text</div>
</main>
<aside class="sidebar">Related articles</aside>
<footer>© Example Gazette 2026</footer>
</body></html>"""


# --- HTML -> text ------------------------------------------------------------------------------------
def test_main_content_wins_and_noise_is_dropped():
    title, text = web_reader.html_to_text(PAGE)
    assert title == "Riverside & partners"
    assert "# The riverside boardwalk project" in text
    assert "the contractor was Example Builders Ltd." in text
    assert "- Opened: 2022" in text
    assert "EUR 1.2 M" in text
    for noise in ("Home", "cookies", "Related articles", "Example Gazette", "var a", "color:red", "Hidden text"):
        assert noise not in text


def test_page_without_main_keeps_the_body():
    _, text = web_reader.html_to_text("<html><body><div><p>First paragraph.</p><p>Second.</p></div></body></html>")
    assert text == "First paragraph.\n\nSecond."


def test_broken_nesting_does_not_lose_text():
    _, text = web_reader.html_to_text("<div><p>one <b>two</p> three</div><p>four")
    for word in ("one", "two", "three", "four"):
        assert word in text


def test_javascript_shell_is_recognised():
    shell = '<html><body><div id="root"></div><script src="/app.js"></script></body></html>'
    assert web_reader.needs_javascript(shell)
    assert not web_reader.needs_javascript(PAGE.replace("<script>var a = 1;</script>", "") * 3)


# --- decoding ----------------------------------------------------------------------------------------
def test_charset_from_header_meta_and_fallback():
    # Central European text on purpose: ő and ű exist only in iso-8859-2 /
    # windows-1250, so a wrong code page would show here.
    text = "Árvíztűrő tükörfúrógép"
    assert web_reader.decode(text.encode("iso-8859-2"), "text/html; charset=iso-8859-2") == text
    meta = b'<meta charset="windows-1250">' + text.encode("cp1250")
    assert text in web_reader.decode(meta, "text/html")
    assert web_reader.decode(text.encode("utf-8"), "") == text
    # Nothing declared and not UTF-8: windows-1252, as HTML5 assumes.
    assert web_reader.decode("café crème".encode("cp1252"), "") == "café crème"


# --- address check -----------------------------------------------------------------------------------
# No RFC 1918 literal here: tools/check_public.py flags those in any file. The
# documentation range and an IPv6 unique-local address fail the same is_global test.
@pytest.mark.parametrize("url", [
    "http://127.0.0.1/", "http://192.0.2.10/a", "http://[fd00::1]/", "http://[::1]/", "http://169.254.169.254/latest/meta-data/",
    "http://localhost:8000/", "http://printer.local/", "http://intranet.internal/",
    "file:///etc/passwd", "ftp://example.com/", "https://user:pw@example.com/", "",
])
def test_private_local_and_non_http_are_refused(url):
    result = web_reader.read(url)
    assert "error" in result and "content" not in result


def test_a_redirect_to_a_private_address_is_refused():
    handler = web_reader._Redirects()
    with pytest.raises(urllib.error.HTTPError) as caught:
        handler.redirect_request(None, None, 302, "Found", {}, "http://127.0.0.1/admin")
    assert "private or local" in str(caught.value.reason)


def test_reader_does_not_shell_out():
    source = (PLUGIN / "scripts" / "web_reader.py").read_text(encoding="utf-8")
    for banned in ("subprocess", "os.system", "exec(", "eval("):
        assert banned not in source


# --- one URL, with fetch replaced -------------------------------------------------------------------
@pytest.fixture
def served(monkeypatch):
    """`read` with the address check passed and `fetch` answering from a table.
    Each entry is (body, content type) or an HTTP status to raise. Records the
    user agent of every request."""
    class Served:
        pages = {}
        agents = []

    def fetch(url, agent=web_reader.USER_AGENT):
        Served.agents.append(agent)
        answer = Served.pages[url]
        if isinstance(answer, dict):
            answer = answer.get(agent, answer["default"])
        if isinstance(answer, int):
            raise urllib.error.HTTPError(url, answer, "refused", {}, None)
        body, ctype = answer
        return body, ctype, url

    monkeypatch.setattr(web_reader, "public_url", lambda u: u)
    monkeypatch.setattr(web_reader, "fetch", fetch)
    return Served


def test_read_html(served):
    served.pages["https://example.com/a"] = (PAGE.encode("utf-8"), "text/html; charset=utf-8")
    result = web_reader.read("https://example.com/a")
    assert result["title"] == "Riverside & partners" and result["via"] == "http"
    assert "Example Builders Ltd." in result["content"]
    assert served.agents == [web_reader.USER_AGENT]


def test_browser_agent_only_after_403(served):
    served.pages["https://example.com/b"] = {web_reader.USER_AGENT: 403,
                                              "default": (PAGE.encode("utf-8"), "text/html")}
    result = web_reader.read("https://example.com/b")
    assert "content" in result
    assert served.agents == [web_reader.USER_AGENT, web_reader.BROWSER_AGENT]


def test_refused_twice_points_to_the_browser(served):
    served.pages["https://example.com/c"] = 403
    assert "browser" in web_reader.read("https://example.com/c")["error"]


def test_javascript_page_points_to_the_browser(served):
    shell = b'<html><body><div id="root"></div><script src="/app.js"></script></body></html>'
    served.pages["https://example.com/app"] = (shell, "text/html")
    assert "browser_exec" in web_reader.read("https://example.com/app")["error"]


def test_short_page_with_a_script_is_read_with_a_note(served):
    page = (b"<html><head><title>Example Domain</title></head><body><p>This domain is for use in "
            b"documentation examples without needing permission.</p><script src=/s.js></script></body></html>")
    served.pages["https://example.com/"] = (page, "text/html")
    result = web_reader.read("https://example.com/")
    assert result["content"].startswith("This domain is for use in documentation examples")
    assert "browser tool" in result["content"]


def test_read_pdf(served):
    served.pages["https://example.com/report.pdf"] = (build_pdf(["Annual report 2025", None]), "application/pdf")
    result = web_reader.read("https://example.com/report.pdf")
    assert result["via"] == "pdf"
    assert result["content"].startswith("[page 1]\nAnnual report 2025")


def test_scanned_pdf_and_binary_are_explained(served):
    served.pages["https://example.com/scan.pdf"] = (build_pdf([None]), "application/pdf")
    served.pages["https://example.com/a.zip"] = (b"PK\x03\x04", "application/zip")
    assert "no text layer" in web_reader.read("https://example.com/scan.pdf")["error"]
    assert "not a readable page" in web_reader.read("https://example.com/a.zip")["error"]


def test_long_text_is_cut_with_a_marker():
    done = web_reader._done("u", "u", "t", "x" * 50, "http", 10)
    assert done["content"].startswith("x" * 10)
    assert "truncated at 10 characters" in done["content"]


def test_accept_language_can_be_set(monkeypatch):
    sent = {}

    class Opener:
        def open(self, request, timeout):
            sent.update(request.header_items())
            raise OSError("stop here")

    monkeypatch.setattr(web_reader.urllib.request, "build_opener", lambda *a: Opener())
    monkeypatch.setenv("WEB_READER_ACCEPT_LANGUAGE", "de-DE,de;q=0.9")
    with pytest.raises(OSError):
        web_reader.fetch("https://example.com/")
    assert sent["Accept-language"] == "de-DE,de;q=0.9"
    assert sent["User-agent"] == web_reader.USER_AGENT


# --- the Hermes plugin -------------------------------------------------------------------------------
def test_entries_have_the_web_extract_shape(monkeypatch):
    plugin = load_plugin()
    monkeypatch.setattr(web_reader, "read", lambda u, max_chars=0: (
        {"url": u, "error": "nope"} if "bad" in u
        else {"url": u, "final_url": u + "#v", "title": "T", "content": "text", "via": "http"}))
    ok, bad = plugin.extract_urls(["https://good.example.com/", "https://bad.example.com/"])
    assert ok["content"] == ok["raw_content"] == "text"
    assert ok["metadata"]["sourceURL"] == "https://good.example.com/#v"
    assert bad == {"url": "https://bad.example.com/", "title": "", "content": "", "error": "nope"}


def test_a_reader_bug_becomes_a_per_url_error(monkeypatch):
    plugin = load_plugin()

    def broken(u, max_chars=0):
        raise RuntimeError("boom")

    monkeypatch.setattr(web_reader, "read", broken)
    assert plugin.extract_urls(["https://example.com/"]) == [
        {"url": "https://example.com/", "title": "", "content": "", "error": "RuntimeError: boom"}]


def test_registers_an_extract_only_provider(monkeypatch):
    """The base class lives inside Hermes; a stand-in takes its place here."""
    base = types.ModuleType("agent.web_search_provider")

    class WebSearchProvider:
        pass

    base.WebSearchProvider = WebSearchProvider
    monkeypatch.setitem(sys.modules, "agent", types.ModuleType("agent"))
    monkeypatch.setitem(sys.modules, "agent.web_search_provider", base)

    registered = []
    ctx = types.SimpleNamespace(register_web_search_provider=registered.append)
    plugin = load_plugin()
    plugin.register(ctx)
    (provider,) = registered
    assert isinstance(provider, WebSearchProvider)
    assert provider.name == "local" and provider.is_available()
    assert provider.supports_extract() and not provider.supports_search()
    assert provider.get_setup_schema()["env_vars"] == []


def test_manifest_names_the_provider():
    manifest = (PLUGIN / "plugin.yaml").read_text(encoding="utf-8")
    assert "kind: backend" in manifest
    assert "name: local-web-reader" in manifest
    assert "\n  - " + load_plugin().NAME + "\n" in manifest


def test_command_line_prints_markdown_and_reports_failures(monkeypatch, capsys):
    pages = {
        "https://example.org/a": {"url": "https://example.org/a", "final_url": "https://example.org/a",
                                  "title": "Page A", "content": "First page.", "via": "http"},
        "https://example.org/b": {"url": "https://example.org/b", "error": "HTTP 404 from origin: Not Found"},
        "https://example.org/c": {"url": "https://example.org/c", "final_url": "https://example.org/c",
                                  "title": "", "content": "Third page.", "via": "http"},
    }
    monkeypatch.setattr(web_reader, "read", lambda url, max_chars=0: pages[url])
    code = web_reader.main(list(pages))
    out, err = capsys.readouterr()
    assert code == 1
    assert out.startswith("# Page A\n\nSource: https://example.org/a\n\nFirst page.")
    assert "\n---\n" in out and "# https://example.org/c" in out and "Third page." in out
    assert "example.org/b" not in out and "HTTP 404 from origin" in err


def test_command_line_json_and_success_code(monkeypatch, capsys):
    import json

    monkeypatch.setattr(web_reader, "read", lambda url, max_chars=0: {
        "url": url, "final_url": url, "title": "T", "content": "x" * max_chars, "via": "http"})
    assert web_reader.main(["https://example.org/", "--json", "--max-chars", "5"]) == 0
    row = json.loads(capsys.readouterr().out)
    assert row["content"] == "xxxxx" and row["title"] == "T"
