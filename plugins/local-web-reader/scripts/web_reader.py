"""Read a public web page on this machine, without an API key.

Hermes' `web_extract` tool needs an extract backend, and every hosted reading
service wants a key. This module reads the page itself:

1. HTTP GET with the standard library, only for a public http/https address:
   a private, loopback or link-local address is refused, also after a
   redirect.
2. HTML -> readable text with headings, paragraphs, list items and table
   rows (the standard library HTML parser). Scripts, styles, navigation,
   headers, footers, forms, hidden elements and cookie or share banners are
   dropped; <main> / <article> wins when it carries most of the text.
3. PDF -> text with pypdf, when it is installed.
4. A page that is empty without JavaScript, or that refuses a plain client,
   gets a clear answer: read it with the Hermes browser tool. No headless
   browser is started here and no other program is run.

Nothing leaves the machine except the request to the page itself.
"""

from __future__ import annotations

import html
import io
import ipaddress
import os
import re
import socket
import urllib.error
import urllib.parse
import urllib.request
from html.parser import HTMLParser

MAX_BYTES = 15 * 1024 * 1024
MAX_CHARS = 200_000
TIMEOUT = 25
# An honest client name first (Wikipedia and W3C refuse a browser imitation);
# the browser-like one only for a site that answers the honest one with 403/429.
USER_AGENT = "Mozilla/5.0 (compatible; local-web-reader/1.0)"
BROWSER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/140.0.0.0 Safari/537.36"
)
ACCEPT_LANGUAGE = "en-US,en;q=0.9,*;q=0.5"
_MIN_TEXT = 200  # fewer visible characters than this, plus scripts: part of the page may need JavaScript
_EMPTY_TEXT = 40  # fewer than this, plus scripts: the page is an app shell and needs the browser

_DROP = {
    "script", "style", "noscript", "template", "svg", "canvas", "iframe", "object", "embed",
    "nav", "header", "footer", "aside", "form", "button", "select", "option", "dialog", "menu",
}
_BLOCK = {
    "p", "div", "section", "article", "main", "br", "hr", "tr", "table", "blockquote", "pre",
    "ul", "ol", "dl", "dt", "dd", "figure", "figcaption", "address", "li", "h1", "h2", "h3",
    "h4", "h5", "h6", "td", "th",
}
_VOID = {"br", "hr", "img", "input", "meta", "link", "area", "base", "col", "source", "wbr", "track", "param"}
_NOISE_ATTR = re.compile(
    r"cookie|consent|banner|popup|modal|newsletter|share|social|breadcrumb|sidebar|advert|\bads?\b", re.I)


# --- address check -------------------------------------------------------------------------------
def public_url(url: str):
    """The normalised URL, or {"error": ...} for anything that is not a public http(s) address."""
    if not url or not str(url).strip():
        return {"error": "url is required"}
    parts = urllib.parse.urlsplit(str(url).strip())
    if parts.scheme not in ("http", "https") or not parts.hostname:
        return {"error": "only a public http or https URL is read"}
    if parts.username or parts.password:
        return {"error": "a URL with a password is not read"}
    host = parts.hostname.lower()
    if host == "localhost" or host.endswith((".local", ".localhost", ".internal", ".lan", ".home.arpa")):
        return {"error": "a private or local address is not read"}
    try:
        infos = socket.getaddrinfo(parts.hostname, None)
    except socket.gaierror as exc:
        return {"error": f"the host could not be resolved: {exc}"}
    for info in infos:
        try:
            address = ipaddress.ip_address(info[4][0])
        except ValueError:
            return {"error": "the address could not be checked"}
        if not address.is_global:
            return {"error": "a private or local address is not read"}
    return urllib.parse.urlunsplit((parts.scheme, parts.netloc, parts.path or "/", parts.query, ""))


class _Redirects(urllib.request.HTTPRedirectHandler):
    """Every redirect target goes through the same address check as the first URL."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        checked = public_url(newurl)
        if isinstance(checked, dict):
            raise urllib.error.HTTPError(newurl, code, checked["error"], headers, fp)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _https_handler() -> list:
    """Certificates verified by the operating system when `truststore` is installed
    (Hermes ships it). Python's own store misses roots that Windows fetches on
    demand and roots an organisation deploys; without truststore, Python's default."""
    try:
        import ssl

        import truststore

        return [urllib.request.HTTPSHandler(context=truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT))]
    except Exception:
        return []


def fetch(url: str, agent: str = USER_AGENT):
    """(body bytes, content type, final URL). Raises on a network error or an HTTP error status."""
    request = urllib.request.Request(url, headers={
        "User-Agent": agent,
        "Accept": "text/html,application/xhtml+xml,application/pdf;q=0.9,*/*;q=0.8",
        "Accept-Language": os.environ.get("WEB_READER_ACCEPT_LANGUAGE", "").strip() or ACCEPT_LANGUAGE,
    })
    opener = urllib.request.build_opener(_Redirects(), *_https_handler())
    with opener.open(request, timeout=TIMEOUT) as response:
        chunks, total = [], 0
        while True:
            chunk = response.read(1024 * 1024)
            if not chunk:
                break
            total += len(chunk)
            if total > MAX_BYTES:
                raise ValueError(f"the page is larger than {MAX_BYTES // (1024 * 1024)} MB")
            chunks.append(chunk)
        return b"".join(chunks), response.headers.get("Content-Type", ""), response.geturl()


# --- decoding --------------------------------------------------------------------------------------
def decode(body: bytes, content_type: str) -> str:
    """The charset from the Content-Type header, then from a <meta> tag, then UTF-8,
    then windows-1252 (what HTML5 assumes for a page that declares nothing)."""
    m = re.search(r"charset=([\w.:-]+)", content_type or "", re.I)
    candidates = [m.group(1)] if m else []
    meta = re.search(rb"<meta[^>]+charset=[\"']?([\w.:-]+)", body[:4096], re.I)
    if meta:
        candidates.append(meta.group(1).decode("ascii", "ignore"))
    candidates += ["utf-8", "cp1252"]
    for enc in candidates:
        try:
            return body.decode(enc)
        except (LookupError, UnicodeDecodeError):
            continue
    return body.decode("utf-8", errors="replace")


# --- HTML -> text ----------------------------------------------------------------------------------
class _Reader(HTMLParser):
    """Collects readable text; tracks <article>/<main> separately so the main content can win."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.title = ""
        self._in_title = False
        self._drop_depth = 0
        self._stack: list[tuple[str, bool, bool]] = []  # (tag, dropped here, opened main content)
        self._main_depth = 0
        self.all: list[str] = []
        self.main: list[str] = []

    def _emit(self, text: str):
        self.all.append(text)
        if self._main_depth:
            self.main.append(text)

    def handle_starttag(self, tag, attrs):
        if tag == "title":
            self._in_title = True
        if tag in _VOID:
            if tag in ("br", "hr") and not self._drop_depth:
                self._emit("\n")
            return
        a = dict(attrs)
        marks = " ".join((a.get(k) or "") for k in ("class", "id", "role"))
        dropped = (
            tag in _DROP
            or ("hidden" in a)
            or a.get("aria-hidden") == "true"
            # cookie banners, share bars, sidebars; never the page or its main content itself
            or (tag not in ("html", "body", "main", "article") and bool(_NOISE_ATTR.search(marks)))
        )
        is_main = tag in ("article", "main") or a.get("role") == "main"
        self._stack.append((tag, dropped, is_main and not dropped))
        if dropped:
            self._drop_depth += 1
            return
        if is_main:
            self._main_depth += 1
        if self._drop_depth:
            return
        if tag in _BLOCK:
            self._emit("\n")
        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._emit("#" * int(tag[1]) + " ")
        elif tag == "li":
            self._emit("- ")
        elif tag in ("td", "th"):
            self._emit(" | ")

    def handle_endtag(self, tag):
        if tag == "title":
            self._in_title = False
        if tag in _VOID:
            return
        # Close up to the matching tag (HTML in the wild is not well nested).
        for i in range(len(self._stack) - 1, -1, -1):
            if self._stack[i][0] == tag:
                for _t, dropped, main in reversed(self._stack[i:]):
                    if dropped:
                        self._drop_depth -= 1
                    elif main and self._main_depth:
                        self._main_depth -= 1
                del self._stack[i:]
                break
        if not self._drop_depth and tag in _BLOCK:
            self._emit("\n")

    def handle_data(self, data):
        if self._in_title:
            self.title += data
            return
        if self._drop_depth or not data.strip():
            if data and not data.strip() and not self._drop_depth:
                self._emit(" ")
            return
        self._emit(re.sub(r"\s+", " ", data))


def _tidy(parts: list[str]) -> str:
    text = "".join(parts)
    lines = []
    for line in text.split("\n"):
        line = re.sub(r"[ \t ]+", " ", line).strip()
        line = re.sub(r"^(\|\s*)+", "", line).strip()
        if line in ("-", "#", "|"):
            continue
        lines.append(line)
    text = "\n".join(lines)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def html_to_text(markup: str) -> tuple[str, str]:
    """(title, readable text). The <article>/<main> content wins when it carries most of the text."""
    reader = _Reader()
    try:
        reader.feed(markup)
        reader.close()
    except Exception:  # a broken page still yields what was parsed
        pass
    full = _tidy(reader.all)
    main = _tidy(reader.main)
    text = main if main and len(main) >= max(_MIN_TEXT, 0.3 * len(full)) else full
    return html.unescape(reader.title).strip(), text


def _scripted_and_short(markup: str, text: str) -> bool:
    return len(text) < _MIN_TEXT and "<script" in markup.lower()


def _looks_empty(markup: str, text: str) -> bool:
    """Scripts and next to no text: an app shell that renders in the browser.
    A short page with real text (example.com) is read, with a note."""
    return _scripted_and_short(markup, text) and len(text) < _EMPTY_TEXT


def needs_javascript(markup: str) -> bool:
    """True for a page that shows (almost) nothing until its scripts run."""
    return _looks_empty(markup, html_to_text(markup)[1])


# --- PDF ---------------------------------------------------------------------------------------------
def pdf_to_text(body: bytes) -> tuple[str, str]:
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise ValueError("this is a PDF, and pypdf is not installed: download it and use the file-reader skill") from exc
    reader = PdfReader(io.BytesIO(body))
    title = ""
    try:
        title = (reader.metadata.title or "") if reader.metadata else ""
    except Exception:
        pass
    pages = []
    for i, page in enumerate(reader.pages, 1):
        try:
            txt = (page.extract_text() or "").strip()
        except Exception:
            txt = ""
        if txt:
            pages.append(f"[page {i}]\n{txt}")
    if not pages:
        raise ValueError("the PDF has no text layer (scanned): download it and use the file-reader skill with OCR")
    return title, "\n\n".join(pages)


# --- one URL ---------------------------------------------------------------------------------------
_NEEDS_BROWSER = (
    "the page needs JavaScript (or refuses a plain client): read it with the Hermes browser tool, e.g. "
    "browser_exec: new_tab(url); wait_for_load(); print(js('document.body.innerText'))"
)
_SHORT_PAGE = "read it with the Hermes browser tool (browser_exec)"


def read(url: str, max_chars: int = MAX_CHARS) -> dict:
    """{"url", "final_url", "title", "content", "via"} or {"url", "error"}."""
    checked = public_url(url)
    if isinstance(checked, dict):
        return {"url": url, "error": checked["error"]}
    try:
        try:
            body, ctype, final = fetch(checked)
        except urllib.error.HTTPError as exc:
            if exc.code not in (403, 429):
                raise
            body, ctype, final = fetch(checked, BROWSER_AGENT)
    except urllib.error.HTTPError as exc:
        if exc.code in (401, 403, 429, 503):
            return {"url": url, "error": f"HTTP {exc.code} from origin; {_NEEDS_BROWSER}"}
        return {"url": url, "error": f"HTTP {exc.code} from origin: {exc.reason}"}
    except Exception as exc:
        return {"url": url, "error": f"{type(exc).__name__}: {exc}"}

    lowered = (ctype or "").lower()
    via = "http"
    try:
        if "pdf" in lowered or body[:5] == b"%PDF-":
            title, text = pdf_to_text(body)
            via = "pdf"
        elif "html" in lowered or "xml" in lowered or body.lstrip()[:1] == b"<":
            markup = decode(body, ctype)
            title, text = html_to_text(markup)
            if _looks_empty(markup, text):
                return {"url": url, "error": _NEEDS_BROWSER}
            if _scripted_and_short(markup, text):
                text += f"\n\n[short page that runs scripts: if content seems missing, {_SHORT_PAGE}]"
        elif lowered.startswith("text/") or "json" in lowered:
            title, text = "", decode(body, ctype)
        else:
            return {"url": url, "error": f"not a readable page ({ctype or 'unknown type'}): "
                                         "download it and use the file-reader skill"}
    except ValueError as exc:
        return {"url": url, "error": str(exc)}
    return _done(url, final, title, text, via, max_chars)


def _done(url, final, title, text, via, max_chars):
    if not text.strip():
        return {"url": url, "error": "the page has no readable text"}
    if max_chars and len(text) > max_chars:
        text = text[:max_chars].rstrip() + f"\n\n[… truncated at {max_chars} characters]"
    return {"url": url, "final_url": final, "title": title, "content": text, "via": via}
