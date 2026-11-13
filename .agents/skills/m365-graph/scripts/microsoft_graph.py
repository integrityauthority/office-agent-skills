"""Search and read Microsoft 365 content with the signed-in user's own rights.

Mail, calendar, OneDrive, SharePoint and Teams chats over Microsoft Graph,
using DELEGATED permissions: the user signs in through a normal browser
window and the agent sees exactly what that person can already see, nothing
more.

Every command prints exactly one JSON object and exits 0 on success, 1 on
failure — same contract as office.py:

    {"ok": true,  "data": ...}
    {"ok": false, "error": "...", "hint": "..."}

Design notes for whoever maintains this:

* Reads are open, writes are CLOSED AT THE TRANSPORT. `graph()` refuses any
  non-GET call unless the caller names a write area AND the path matches that
  area's pattern AND the token actually holds the scope. A skill that only
  searches cannot send mail even if the model decides to try, because the
  refusal happens below the model, not in prose. `POST /search/query` is
  whitelisted — it is the search API, it changes nothing.
* Scopes are granted INCREMENTALLY and remembered. Requesting the full read
  set on every silent call would fail for anyone who consented to a subset,
  so `login` records what consent actually returned and later calls ask for
  exactly that. Missing scope => a message naming the `login --add-scope`
  command to fix it, never a bare 403.
* The Search API's entityTypes cannot be freely mixed (mail, Teams messages,
  events and SharePoint items each want their own request). `search` splits
  the requested types into compatible groups, fans out, and merges. Callers
  should not have to know this.
* Search responses are enormous — one hit carries the whole resource. We
  compact per type down to the fields a human would read, and keep --raw for
  when that is not enough. This is the difference between a search that fits
  in a model's context and one that does not.
* A drive `/children` listing is compacted the same way: `count`, `folders`
  and `files`. Word and Excel bytes from `/content` become text here, as PDFs do.
* The token cache is encrypted with Windows DPAPI, so the file is useless to
  any other account on the machine. It is not a secret store; it is a cache.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import re
import sys

GRAPH = "https://graph.microsoft.com/v1.0"
AUTHORITY = "https://login.microsoftonline.com/organizations"

# Microsoft's own "Graph Command Line Tools" public client — the app id the
# official Connect-MgGraph uses. It is registered in practically every tenant
# with an http://localhost redirect, which is why interactive sign-in works
# here with no app registration of your own.
#
# The trade-off is the audit trail: Entra sign-in logs will name that app, not
# Hermes. For anything beyond trying this out, register your own app (two
# minutes, no admin rights in a default tenant) and set MSGRAPH_CLIENT_ID.
DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

RUNTIME_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                           "m365-graph")
CACHE_PATH = os.path.join(RUNTIME_DIR, "token-cache.bin")
STATE_PATH = os.path.join(RUNTIME_DIR, "state.json")

# Read scopes a normal user can consent to themselves in a default tenant.
# Deliberately excludes anything that always needs an administrator — notably
# ChannelMessage.Read.All (Teams *channel* posts). Chat.Read covers the 1:1
# and group chats the user is in, which is what "search Teams" usually means.
READ_SCOPES = [
    "User.Read",
    "Mail.Read",
    "Calendars.Read",
    "Files.Read.All",
    "Sites.Read.All",
    "Chat.Read",
    "People.Read",
]

# The smallest useful set — try this when the full set is refused, to find out
# whether consent is blocked outright or only for the broader scopes.
MINIMAL_SCOPES = ["User.Read", "Mail.Read"]

# Write areas. Each entry: path pattern the write may target, and the scope
# the token must hold. Adding an area here is a deliberate act; nothing is
# writable by default.
WRITE_AREAS = {
    "calendar": (re.compile(r"^/me/(events|calendars?/[^/]+/events)"),
                 "Calendars.ReadWrite"),
    "mail-send": (re.compile(r"^/me/(sendMail$|messages/[^/]+/(reply|replyAll|forward|send)$)"),
                  "Mail.Send"),
    "files": (re.compile(r"^/me/drive/"), "Files.ReadWrite.All"),
}

# Non-GET calls that change nothing. The search API is a POST because the
# query is a document, not because it writes.
SAFE_POSTS = (re.compile(r"^/search/query$"),)

# Search entityTypes that may travel in one request. Graph rejects most
# cross-group combinations, so we never build one.
SEARCH_GROUPS = [
    ["driveItem", "drive", "listItem", "list", "site"],   # OneDrive + SharePoint
    ["message"],                                          # Outlook mail
    ["chatMessage"],                                      # Teams chats
    ["event"],                                            # calendar
    ["person"],
]

# AADSTS codes worth translating. Everything else falls through with its own
# text, which Microsoft writes reasonably well.
AAD_HINTS = {
    "AADSTS65001": ("consent was not granted for these permissions",
                    "Run 'login --minimal' to see whether a smaller set is "
                    "allowed. If even that fails, user consent is disabled in "
                    "the tenant and an administrator has to approve the app."),
    "AADSTS90094": ("this permission needs administrator approval",
                    "Drop the scope that triggered it and retry — "
                    "'login --minimal' is the safe baseline. Teams CHANNEL "
                    "search (ChannelMessage.Read.All) always needs an admin."),
    "AADSTS700016": ("the application is not present in this tenant",
                     "Set MSGRAPH_CLIENT_ID to an app registered in your own "
                     "tenant, or ask for the default client to be enabled."),
    "AADSTS50194": ("the app is registered single-tenant, but sign-in went to "
                    "the shared /organizations endpoint",
                    "Set MSGRAPH_TENANT_ID to the Directory (tenant) ID from "
                    "the app registration, next to MSGRAPH_CLIENT_ID."),
    "AADSTS7000218": ("the app registration does not allow public client flows",
                      "In Entra: the app -> Authentication -> 'Allow public "
                      "client flows' = Yes. No client secret is needed."),
    "AADSTS50076": ("multi-factor authentication is required",
                    "Complete the MFA prompt in the browser window, then "
                    "retry. Do not retry without the user present."),
    "AADSTS530003": ("device compliance policy blocked the sign-in",
                     "Conditional Access requires a managed device. This "
                     "cannot be worked around from here — tell the user."),
    "AADSTS50105": ("the user is not assigned to this application",
                    "An administrator restricted who may use the app."),
}


class GraphError(RuntimeError):
    """Raised by the library API; the CLI turns it into a JSON error object."""

    def __init__(self, message, hint=""):
        super().__init__(message)
        self.hint = hint


def _utf8_stdout():
    """Force UTF-8 on stdout.

    Python on Windows writes stdout in the console codepage, which is cp852
    here: an em-dash in a hint comes out as '?', and a Hungarian accent in a
    document title raises UnicodeEncodeError mid-print, so the caller gets
    half a JSON object and a traceback.
    """
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def out(data):
    _utf8_stdout()
    print(json.dumps({"ok": True, "data": data}, ensure_ascii=False, default=str))
    sys.exit(0)


def fail(error, hint=""):
    _utf8_stdout()
    print(json.dumps({"ok": False, "error": str(error), "hint": hint},
                     ensure_ascii=False))
    sys.exit(1)


# -------------------------------------------------------------------- auth ---

def _msal():
    try:
        import msal
        return msal
    except ImportError:
        raise GraphError(
            "the msal package is not installed",
            "Install it into the interpreter Hermes uses: "
            "'<hermes venv>\\Scripts\\python.exe -m pip install msal', or "
            "re-run Install-Hermes.ps1 which does it for you.")


def client_id():
    return os.environ.get("MSGRAPH_CLIENT_ID", "").strip() or DEFAULT_CLIENT_ID


def authority():
    """Where to sign in. Tenant-specific when MSGRAPH_TENANT_ID is set.

    /organizations works for the borrowed multi-tenant CLI client, but NOT for
    an app registered as single-tenant ("accounts in this organizational
    directory only") — Entra refuses that with AADSTS50194. Single-tenant is
    exactly what we ask IT for, so the tenant id has to travel with the client
    id; setting only MSGRAPH_CLIENT_ID would fail at the first sign-in.
    """
    tenant = os.environ.get("MSGRAPH_TENANT_ID", "").strip()
    return f"https://login.microsoftonline.com/{tenant}" if tenant else AUTHORITY


def _dpapi(data, unprotect=False):
    """Encrypt/decrypt with the current Windows user's DPAPI key."""
    try:
        import win32crypt
    except ImportError:
        return data                     # no pywin32: cache stays plaintext
    if unprotect:
        return win32crypt.CryptUnprotectData(data, None, None, None, 0)[1]
    return win32crypt.CryptProtectData(data, "hermes m365 token cache",
                                       None, None, None, 0)


def _load_cache():
    msal = _msal()
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_PATH):
        try:
            with open(CACHE_PATH, "rb") as fh:
                cache.deserialize(_dpapi(fh.read(), unprotect=True).decode("utf-8"))
        except Exception:
            # A cache from another user, another machine, or a rotated DPAPI
            # key is unreadable, not corrupt-and-fatal: sign in again.
            pass
    return cache


def _save_cache(cache):
    if not cache.has_state_changed:
        return
    os.makedirs(RUNTIME_DIR, exist_ok=True)
    blob = _dpapi(cache.serialize().encode("utf-8"))
    with open(CACHE_PATH, "wb") as fh:
        fh.write(blob)


def _state():
    try:
        with open(STATE_PATH, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _save_state(**kw):
    st = _state()
    st.update(kw)
    os.makedirs(RUNTIME_DIR, exist_ok=True)
    with open(STATE_PATH, "w", encoding="utf-8") as fh:
        json.dump(st, fh, ensure_ascii=False, indent=2)


def granted_scopes():
    """Scopes consent actually returned at the last sign-in."""
    return _state().get("granted_scopes", [])


def _app(cache):
    msal = _msal()
    return msal.PublicClientApplication(client_id(), authority=authority(),
                                        token_cache=cache)


def _aad_error(result):
    desc = result.get("error_description") or result.get("error") or "sign-in failed"
    for code, (msg, hint) in AAD_HINTS.items():
        if code in desc:
            return GraphError(f"{msg} ({code})", hint)
    first = desc.strip().splitlines()[0][:400]
    return GraphError(f"sign-in failed: {first}",
                      "Read the message above — Microsoft usually names the "
                      "exact policy or permission at fault.")


def login(scopes=None, interactive=True):
    """Acquire a token, opening a browser window only when necessary."""
    scopes = list(scopes or granted_scopes() or READ_SCOPES)
    cache = _load_cache()
    app = _app(cache)
    result = None

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        if not interactive:
            raise GraphError(
                "no usable token for the requested scopes",
                "Someone has to sign in interactively first. Ask the user to "
                "run the login step; a browser window will open.")
        result = app.acquire_token_interactive(
            scopes=scopes,
            prompt="select_account" if not accounts else None,
        )

    if "access_token" not in result:
        _save_cache(cache)
        raise _aad_error(result)

    _save_cache(cache)
    got = (result.get("scope") or "").split()
    # Keep the union: consent is incremental, so a later --add-scope must not
    # make us forget what we already hold.
    merged = sorted(set(granted_scopes()) | set(s for s in got
                                                if not s.startswith("openid")
                                                and s not in ("profile", "email",
                                                              "offline_access")))
    account = (result.get("id_token_claims") or {})
    _save_state(granted_scopes=merged,
                account=account.get("preferred_username") or
                        (accounts[0]["username"] if accounts else None),
                tenant=account.get("tid"),
                client_id=client_id())
    return result["access_token"], merged


def _token(need_scope=None):
    """Silent token for the scopes we already hold. Never opens a browser."""
    have = granted_scopes()
    if not have:
        raise GraphError(
            "not signed in to Microsoft 365",
            "Run the login command first: it opens a browser window where the "
            "user signs in with their work account. Nothing is cached until "
            "they do.")
    if need_scope and need_scope not in have:
        raise GraphError(
            f"the current token does not include {need_scope}",
            f"Ask the user to approve it: "
            f"'microsoft_graph.py login --add-scope {need_scope}'. That opens "
            f"a consent window; it is their decision, not yours to assume.")
    token, _ = login(scopes=have, interactive=False)
    return token


# -------------------------------------------------------------------- http ---

_MAX_FILE_BYTES = 30 * 1024 * 1024
_TEXT_LIMIT = 60_000
_PAGE_TEXT_MIN = 20
_OCR_MAX_PAGES = 15
# llama-swap on the office machine. A text layer never goes there.
# M365_OCR_URL= (empty) skips this hop. There is no :11434 fallback.
_OCR_URL = "http://ihai.inhat.hu:11500/v1"
_OCR_MODEL = "deepseek-ai/DeepSeek-OCR-2"
_OCR_DPI = 200
_OCR_PROMPT = "<|grounding|>Convert the document to markdown."
_VISION_DPI = 300
_VISION_URL = "https://sandbox-ai-dev-swedencentral.openai.azure.com/openai/v1"
_VISION_MODEL = "gpt-5.6-luna-datazone"
_VISION_PROMPT = (
    "This is one scanned document page. Transcribe every visible character, "
    "printed and filled in, including numbers and dates. Do not summarise."
)


def _env(name, default=""):
    return (os.environ.get(name) or default).strip()


def _note_source(sources, name):
    if name not in sources:
        sources.append(name)


def _usable(text):
    return len((text or "").strip()) >= _PAGE_TEXT_MIN


def _pymupdf_page_text(page):
    """Text plus any table the page already has. Stays on this machine."""
    text = (page.get_text("text") or "").strip()
    tables = []
    try:
        found = page.find_tables()
        for table in getattr(found, "tables", None) or []:
            rows = []
            for row in table.extract() or []:
                rows.append("\t".join((cell or "").strip() for cell in row))
            block = "\n".join(row for row in rows if row.strip())
            if block:
                tables.append(block)
    except Exception:
        pass
    if tables:
        text = (text + "\n\n" + "\n\n".join(tables)).strip()
    return text


def _page_png(page, dpi):
    return page.get_pixmap(dpi=dpi).tobytes("png")


def _pdfminer_text(data):
    """Character-level fallback used only after Document Intelligence fails."""
    import io

    try:
        from pdfminer.high_level import extract_pages
        from pdfminer.layout import LAParams, LTFigure, LTTextContainer
    except ImportError:
        return ""

    def layout_text(container):
        parts = []
        for element in container:
            if isinstance(element, LTTextContainer):
                parts.append(element.get_text())
            elif isinstance(element, LTFigure):
                parts.append(layout_text(element))
        return "".join(parts)

    laparams = LAParams(
        line_margin=0.5,
        word_margin=0.1,
        char_margin=2.0,
        boxes_flow=0.5,
        all_texts=True,
    )
    pages = []
    try:
        for page in extract_pages(io.BytesIO(data), laparams=laparams):
            pages.append(layout_text(page).replace("\r\n", "\n").replace("\r", "\n").strip())
    except Exception:
        return ""
    return "\n\n".join(pages).strip()


def _ocr_settings():
    if "M365_OCR_URL" in os.environ:
        url = _env("M365_OCR_URL").rstrip("/")
    else:
        url = _OCR_URL
    model = _env("M365_OCR_MODEL", _OCR_MODEL)
    key = _env("M365_OCR_API_KEY", "local")
    return url, model, key


def _chat_text(payload):
    choice = ((payload or {}).get("choices") or [{}])[0]
    content = ((choice.get("message") or {}).get("content")) or ""
    if isinstance(content, list):
        content = "".join(
            (part.get("text") or "") for part in content if isinstance(part, dict)
        )
    return str(content).strip()


def _local_vlm_page(png, url, model, key):
    import base64
    import requests

    response = requests.post(
        f"{url}/chat/completions",
        timeout=180,
        headers={"Authorization": f"Bearer {key}"},
        json={
            "model": model,
            "temperature": 0,
            "skip_special_tokens": False,
            "messages": [{
                "role": "user",
                "content": [
                    {"type": "text", "text": _OCR_PROMPT},
                    {"type": "image_url", "image_url": {
                        "url": "data:image/png;base64," + base64.b64encode(png).decode(),
                    }},
                ],
            }],
        },
    )
    response.raise_for_status()
    return _chat_text(response.json())


def _azure_di(data):
    """Raw PDF to prebuilt-layout. Returns (text, status) with status skip|fail|ok."""
    endpoint = _env("AZURE_DOC_INTEL_ENDPOINT").rstrip("/")
    key = _env("AZURE_DOC_INTEL_KEY")
    if not endpoint or not key:
        return "", "skip"
    try:
        from azure.ai.documentintelligence import DocumentIntelligenceClient
        from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
        from azure.core.credentials import AzureKeyCredential
    except ImportError:
        return "", "fail"
    try:
        client = DocumentIntelligenceClient(endpoint, AzureKeyCredential(key))
        poller = client.begin_analyze_document(
            "prebuilt-layout",
            AnalyzeDocumentRequest(bytes_source=data),
        )
        result = poller.result()
    except Exception:
        return "", "fail"
    pages = []
    for page in getattr(result, "pages", None) or []:
        lines = []
        for line in getattr(page, "lines", None) or []:
            content = (getattr(line, "content", None) or "").strip()
            if content:
                lines.append(content)
        pages.append("\n".join(lines).strip())
    text = "\n\n".join(page for page in pages if page).strip()
    if not text:
        text = (getattr(result, "content", None) or "").strip()
    return text, ("ok" if text else "fail")


def _vision_settings():
    url = _env("M365_VISION_URL", _VISION_URL).rstrip("/")
    model = _env("M365_VISION_MODEL", _VISION_MODEL)
    key = _env("AZURE_SWEDEN_API_KEY")
    return url, model, key


def _response_text(payload):
    chunks = []
    for item in (payload or {}).get("output") or []:
        for part in item.get("content") or []:
            if part.get("type") in ("output_text", "text"):
                chunks.append(part.get("text") or "")
    if chunks:
        return "".join(chunks).strip()
    return _chat_text(payload)


def _azure_vision_page(png, url, model, key):
    import base64
    import requests

    response = requests.post(
        f"{url}/responses",
        timeout=180,
        headers={"api-key": key},
        json={
            "model": model,
            "input": [{
                "role": "user",
                "content": [
                    {"type": "input_text", "text": _VISION_PROMPT},
                    {"type": "input_image", "image_url":
                        "data:image/png;base64," + base64.b64encode(png).decode()},
                ],
            }],
        },
    )
    response.raise_for_status()
    return _response_text(response.json())


def _place(pages, indexes, text):
    """Put one recovered block on the first still-empty page."""
    if not indexes:
        return
    pages[indexes[0]] = text
    for index in indexes[1:]:
        pages[index] = ""


def _read_pdf(data, limit=_TEXT_LIMIT):
    """Text layer, then local VLM, Azure DI, pdfminer, then Azure vision.

    Returns (text, text_source, note). text_source names every layer that
    contributed, joined with '+'. A text layer stays on this machine.
    """
    notes = []
    sources = []
    pages = []
    doc = None
    try:
        import fitz
        doc = fitz.open(stream=data, filetype="pdf")
    except Exception as exc:
        notes.append(f"PyMuPDF could not open the PDF ({type(exc).__name__}).")

    if doc is not None:
        try:
            total = len(doc)
            count = min(total, _OCR_MAX_PAGES)
            if total > count:
                notes.append(f"Read the first {count} of {total} pages.")
            empty = []
            for index in range(count):
                body = _pymupdf_page_text(doc[index])
                if _usable(body):
                    pages.append(body)
                    _note_source(sources, "text-layer")
                else:
                    pages.append(None)
                    empty.append(index)
            url, model, key = _ocr_settings()
            if empty and url and model:
                failed = 0
                for index in empty:
                    try:
                        text = _local_vlm_page(_page_png(doc[index], _OCR_DPI), url, model, key)
                    except Exception:
                        failed += 1
                        text = ""
                    if _usable(text):
                        pages[index] = text.strip()
                        _note_source(sources, "local-vlm")
                if failed:
                    notes.append(f"{failed} page(s) got no answer from {model}.")
            elif empty and not url:
                notes.append("Local OCR is turned off (M365_OCR_URL is empty).")
        finally:
            doc.close()
    else:
        pages = [None]

    empty = [i for i, text in enumerate(pages) if text is None]
    if empty:
        di_text, di_status = _azure_di(data)
        if di_status == "ok":
            _place(pages, empty, di_text)
            _note_source(sources, "azure-di")
        elif di_status == "fail":
            notes.append("Azure Document Intelligence did not return text.")
            miner = _pdfminer_text(data)
            if miner:
                _place(pages, empty, miner)
                _note_source(sources, "pdfminer")
            else:
                notes.append("pdfminer.six found no text layer either.")
        empty = [i for i, text in enumerate(pages) if text is None]

    if empty:
        url, model, key = _vision_settings()
        if not key:
            notes.append("Azure vision is not configured (AZURE_SWEDEN_API_KEY is empty).")
        else:
            try:
                import fitz
                doc = fitz.open(stream=data, filetype="pdf")
            except Exception:
                doc = None
            if doc is None:
                notes.append("Azure vision needs a rendered page, and PyMuPDF could not open the PDF.")
            else:
                try:
                    failed = 0
                    for index in empty:
                        if index >= len(doc):
                            continue
                        try:
                            text = _azure_vision_page(
                                _page_png(doc[index], _VISION_DPI), url, model, key)
                        except Exception:
                            failed += 1
                            text = ""
                        if text.strip():
                            pages[index] = text.strip()
                            _note_source(sources, "azure-vision")
                    if failed:
                        notes.append(f"{failed} page(s) got no answer from {model}.")
                finally:
                    doc.close()

    parts = [text for text in pages if text]
    text = "\n\n".join(parts).strip()[:limit]
    if not text and not notes:
        notes.append("The file is saved locally and no layer returned text.")
    return text, "+".join(sources), " ".join(n for n in notes if n)


def _compact_children(path, payload):
    """A drive listing down to names the model can repeat without counting."""
    items = payload.get("value") if isinstance(payload, dict) else None
    if not isinstance(items, list):
        return payload
    if any(not isinstance(item, dict) or "name" not in item for item in items):
        return payload
    folders, files = [], []
    for item in items:
        row = {"name": item.get("name") or "", "id": item.get("id") or ""}
        folder = item.get("folder")
        if isinstance(folder, dict):
            child = folder.get("childCount")
            if isinstance(child, int):
                row["child_count"] = child
            folders.append({k: v for k, v in row.items() if v != ""})
        else:
            size = item.get("size")
            if isinstance(size, int):
                row["size"] = size
            files.append({k: v for k, v in row.items() if v != ""})
    out = {
        "path": path,
        "count": len(folders) + len(files),
        "folders": folders,
        "files": files,
    }
    if payload.get("@odata.nextLink") or payload.get("truncated_after_pages"):
        out["truncated"] = True
        out["note"] = (
            "Only the first page is here. Open one folder by id, "
            "or call again with a larger $top."
        )
    return out


def _present_graph(path, payload):
    if path.rstrip("/").endswith("/children"):
        return _compact_children(path, payload)
    return payload


def _office_kind(data, content_type):
    """pdf, docx, xlsx, ole (legacy .doc/.xls), or empty when unknown."""
    ct = (content_type or "").split(";")[0].strip().lower()
    if data.startswith(b"%PDF") or ct == "application/pdf":
        return "pdf"
    if "wordprocessingml" in ct:
        return "docx"
    if "spreadsheetml" in ct:
        return "xlsx"
    if data.startswith(b"PK\x03\x04"):
        import io
        import zipfile
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                names = zf.namelist()
        except Exception:
            return ""
        if any(name.startswith("word/") for name in names):
            return "docx"
        if any(name.startswith("xl/") for name in names):
            return "xlsx"
    if data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return "ole"
    return ""


def _read_docx(data, limit):
    import io
    from docx import Document

    doc = Document(io.BytesIO(data))
    parts = []
    for paragraph in doc.paragraphs:
        text = (paragraph.text or "").strip()
        if text:
            parts.append(text)
    for table in doc.tables:
        for row in table.rows:
            cells = [(cell.text or "").strip().replace("\n", " ") for cell in row.cells]
            line = "\t".join(cells).strip()
            if line:
                parts.append(line)
    return "\n".join(parts).strip()[:limit]


_XLSX_MAX_SHEETS = 10
_XLSX_MAX_ROWS = 40
_XLSX_MAX_COLS = 12


def _workbook_text(data, data_only):
    import io
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=data_only)
    try:
        parts = []
        clipped = False
        names = list(wb.sheetnames)
        if len(names) > _XLSX_MAX_SHEETS:
            clipped = True
            names = names[:_XLSX_MAX_SHEETS]
        for name in names:
            lines = []
            for index, row in enumerate(wb[name].iter_rows(values_only=True)):
                if index >= _XLSX_MAX_ROWS:
                    clipped = True
                    break
                if len(row) > _XLSX_MAX_COLS:
                    clipped = True
                cells = [
                    "" if cell is None else str(cell).replace("\n", " ").strip()
                    for cell in row[:_XLSX_MAX_COLS]
                ]
                if any(cells):
                    lines.append("\t".join(cells).rstrip())
            if lines:
                parts.append("# " + name + "\n" + "\n".join(lines))
        return "\n\n".join(parts).strip(), clipped, len(wb.sheetnames)
    finally:
        wb.close()


def _read_xlsx(data, limit):
    text, clipped, sheet_count = _workbook_text(data, True)
    notes = []
    if not text:
        text, clipped, sheet_count = _workbook_text(data, False)
        if text:
            notes.append("Cached cell values were empty, so formulas are shown.")
    if sheet_count > _XLSX_MAX_SHEETS:
        notes.append(f"{sheet_count - _XLSX_MAX_SHEETS} further sheet(s) omitted.")
    if clipped:
        notes.append(
            f"Each sheet is clipped to {_XLSX_MAX_ROWS} rows and {_XLSX_MAX_COLS} columns."
        )
    return text[:limit], " ".join(notes)


def _download_drive_content(path, timeout):
    """GET a drive item's /content.

    Graph answers with a redirect to a pre-authenticated SharePoint URL and
    the body is the file itself, not JSON. The bearer token must not be sent
    to that URL. Opening it in a browser fails: that session is not signed in.
    """
    import requests

    token = _token()
    url = GRAPH + path
    response = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        allow_redirects=False,
        timeout=timeout,
    )
    if response.status_code in (301, 302, 303, 307, 308):
        location = response.headers.get("Location")
        if not location:
            raise GraphError("the content redirect had no Location header",
                             "Read the item metadata again and retry /content.")
        response = requests.get(location, timeout=timeout)
    if response.status_code == 401:
        raise GraphError("Microsoft Graph rejected the token (401)",
                         "The token expired or was revoked. Sign in again.")
    if response.status_code == 403:
        raise GraphError(
            f"Microsoft Graph denied the request (403): {response.text[:300]}",
            "The signed-in user cannot read this file.")
    if not response.ok:
        raise GraphError(f"HTTP {response.status_code}: {response.text[:400]}")
    data = response.content or b""
    if len(data) > _MAX_FILE_BYTES:
        raise GraphError(
            f"the file is {len(data)} bytes, above the {_MAX_FILE_BYTES} byte cap",
            "Read the metadata and ask the user for a smaller file.")

    content_type = response.headers.get("Content-Type", "").split(";")[0].strip()
    item = path[:-len("/content")].rstrip("/").rsplit("/", 1)[-1]
    item = re.sub(r"[^A-Za-z0-9._-]", "_", item)[:80] or "file"
    kind = _office_kind(data, content_type)
    ext = {"pdf": ".pdf", "docx": ".docx", "xlsx": ".xlsx", "ole": ".ole"}.get(kind, "")
    folder = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                          "hermes", "cache", "m365")
    os.makedirs(folder, exist_ok=True)
    saved = os.path.join(folder, item + ext)
    with open(saved, "wb") as fh:
        fh.write(data)

    text = ""
    source = ""
    ocr_note = ""
    if kind == "pdf":
        text, source, ocr_note = _read_pdf(data)
    elif kind == "docx":
        try:
            text = _read_docx(data, _TEXT_LIMIT)
            source = "docx"
        except Exception as exc:
            ocr_note = f"Word text could not be read ({type(exc).__name__})."
    elif kind == "xlsx":
        try:
            text, ocr_note = _read_xlsx(data, _TEXT_LIMIT)
            source = "xlsx"
        except Exception as exc:
            ocr_note = f"Excel text could not be read ({type(exc).__name__})."
    elif kind == "ole":
        ocr_note = (
            "This is a legacy Word or Excel file (.doc or .xls). "
            "Ask for the .docx or .xlsx copy, then call /content on that item."
        )
    result = {
        "saved_to": saved,
        "content_type": content_type or ("application/pdf" if kind == "pdf" else "application/octet-stream"),
        "size": len(data),
    }
    if text:
        result["text"] = text
        result["text_source"] = source
        result["truncated"] = len(text) >= _TEXT_LIMIT
        if ocr_note:
            result["note"] = ocr_note
    else:
        result["text"] = ""
        if source:
            result["text_source"] = source
        result["note"] = ocr_note or (
            "The file is saved. Its text was not extracted here. "
            "Word, Excel and PDF are the formats read from /content. "
            "Do not open a SharePoint or download URL in the browser, "
            "and do not call /preview."
        )
    return result


def graph(method, path, *, params=None, body=None, allow_write=None,
          headers=None, timeout=60, all_pages=False, page_cap=10):
    """Call Microsoft Graph. Reads are open; writes must be named explicitly.

    Import this from execute_code to go beyond the typed commands:

        import sys; sys.path.insert(0, r"<skill>/scripts")
        from microsoft_graph import graph
        graph("GET", "/me/messages", params={"$top": 5})

    Writes need `allow_write=<area>` from WRITE_AREAS, and the path has to
    belong to that area — otherwise the call is refused here, before any
    request goes out.
    """
    import requests

    method = method.upper()
    path = "/" + (path or "").lstrip("/")
    need = None
    if method == "GET" and path == "/":
        raise GraphError(
            "refused: a Graph path is required",
            "An empty path returns the API catalogue, not the user's files. "
            "The signed-in user's OneDrive root is /me/drive/root/children. "
            "A folder from that list is /me/drive/items/<id>/children. "
            "A Word, Excel or PDF file is /me/drive/items/<id>/content.")
    if method == "GET" and path.rstrip("/").endswith("/children"):
        params = dict(params or {})
        if not any(str(key).lower() == "$select" for key in params):
            params["$select"] = "id,name,size,folder,file"
        if not any(str(key).lower() == "$top" for key in params):
            params["$top"] = "50"
    if method == "GET" and path.rstrip("/").endswith("/content"):
        return _download_drive_content(path, timeout)

    if method != "GET":
        if path == "/$batch":
            inner = ((body or {}).get("requests") or [])
            bad = [r.get("method", "?") for r in inner
                   if str(r.get("method", "")).upper() != "GET"]
            if bad:
                raise GraphError(
                    f"refused: $batch may only carry GET sub-requests (found {bad})",
                    "Batching is for fanning out reads. Send writes one by "
                    "one through the area they belong to.")
        elif method == "POST" and any(p.match(path) for p in SAFE_POSTS):
            pass
        elif allow_write is None:
            raise GraphError(
                f"refused: {method} {path} is a write and no area was allowed",
                "This helper is read-only unless the caller passes "
                f"allow_write=<area>, one of {sorted(WRITE_AREAS)}. Tell the "
                "user what you intend to change and let them confirm first.")
        else:
            if allow_write not in WRITE_AREAS:
                raise GraphError(f"unknown write area {allow_write!r}",
                                 f"known areas: {sorted(WRITE_AREAS)}")
            pattern, need = WRITE_AREAS[allow_write]
            if not pattern.match(path):
                raise GraphError(
                    f"refused: {path} is outside the {allow_write!r} area",
                    f"{allow_write!r} only covers paths matching "
                    f"{pattern.pattern}. Pick the right area rather than a "
                    "wider one.")

    token = _token(need_scope=need)
    hdrs = {"Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json"}
    hdrs.update(headers or {})

    url = path if path.startswith("http") else GRAPH + path
    collected, pages = None, 0

    while True:
        for attempt in (1, 2, 3):
            r = requests.request(method, url, headers=hdrs, params=params,
                                 json=body, timeout=timeout)
            if r.status_code == 429 or r.status_code >= 500:
                wait = int(r.headers.get("Retry-After", 2 * attempt))
                if attempt < 3:
                    import time
                    time.sleep(min(wait, 20))
                    continue
            break

        if r.status_code == 401:
            raise GraphError("Microsoft Graph rejected the token (401)",
                             "The token expired or was revoked. Sign in again "
                             "with the login command.")
        if r.status_code == 403:
            raise GraphError(
                f"Microsoft Graph denied the request (403): {r.text[:300]}",
                "Either the token lacks the scope for this path, or the user "
                "genuinely has no access to that mailbox/site. Check "
                "'scopes' before assuming it is a permission bug.")
        if r.status_code == 404:
            raise GraphError(f"not found: {path}",
                             "Check the id or path. A deleted item and a "
                             "mistyped id look the same from here.")
        if not r.ok:
            raise GraphError(f"HTTP {r.status_code}: {r.text[:400]}")

        payload = r.json() if r.content else {}
        params = None                          # nextLink carries its own query

        if not all_pages:
            return _present_graph(path, payload)
        if collected is None:
            collected = payload
        else:
            collected.setdefault("value", []).extend(payload.get("value", []))
        pages += 1
        url = payload.get("@odata.nextLink")
        if not url or pages >= page_cap:
            if url:
                collected["truncated_after_pages"] = pages
            collected.pop("@odata.nextLink", None)
            return _present_graph(path, collected)


# ------------------------------------------------------------------ search ---

def _text(html, limit=600):
    """Crude HTML to text. Good enough for previews, never for reproduction."""
    if not html:
        return ""
    s = re.sub(r"(?is)<(script|style).*?</\1>", " ", html)
    s = re.sub(r"(?s)<[^>]+>", " ", s)
    s = (s.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<")
          .replace("&gt;", ">").replace("&quot;", '"').replace("&#39;", "'"))
    return re.sub(r"\s+", " ", s).strip()[:limit]


def _compact(hit):
    """One search hit down to the fields a person would actually read."""
    res = hit.get("resource") or {}
    kind = (res.get("@odata.type") or "").rsplit(".", 1)[-1]
    rec = {"type": kind or "unknown", "rank": hit.get("rank")}
    summary = _text(hit.get("summary"), 300)

    if kind == "message":
        frm = ((res.get("from") or {}).get("emailAddress") or {})
        rec.update(subject=res.get("subject"),
                   sender=frm.get("address") or frm.get("name"),
                   received=res.get("receivedDateTime"),
                   has_attachments=res.get("hasAttachments"),
                   preview=summary or _text((res.get("body") or {}).get("content"), 300),
                   id=res.get("id"), link=res.get("webLink"))
    elif kind in ("driveItem", "listItem", "drive", "list", "site"):
        parent = (res.get("parentReference") or {})
        rec.update(name=res.get("name") or res.get("displayName"),
                   modified=res.get("lastModifiedDateTime"),
                   size=res.get("size"),
                   location=parent.get("path") or parent.get("siteId"),
                   preview=summary,
                   id=res.get("id"), link=res.get("webUrl"))
    elif kind == "chatMessage":
        frm = ((res.get("from") or {}).get("user") or {})
        rec.update(sender=frm.get("displayName"),
                   created=res.get("createdDateTime"),
                   preview=summary or _text((res.get("body") or {}).get("content"), 300),
                   chat_id=res.get("chatId"), id=res.get("id"),
                   link=res.get("webUrl"))
    elif kind == "event":
        rec.update(subject=res.get("subject"),
                   start=(res.get("start") or {}).get("dateTime"),
                   organizer=(((res.get("organizer") or {}).get("emailAddress")
                               or {}).get("address")),
                   preview=summary, id=res.get("id"), link=res.get("webLink"))
    else:
        rec.update(name=res.get("displayName") or res.get("name"),
                   preview=summary, id=res.get("id"))
    return {k: v for k, v in rec.items() if v not in (None, "", [])}


def search(query, types=None, top=15, start=0, raw=False):
    """Search across M365. Types that Graph refuses to mix are fanned out."""
    wanted = [t.strip() for t in (types or ["message", "driveItem", "chatMessage"])
              if t.strip()]
    known = {t for g in SEARCH_GROUPS for t in g}
    unknown = [t for t in wanted if t not in known]
    if unknown:
        raise GraphError(f"unknown entityTypes: {unknown}",
                         f"pick from {sorted(known)}")

    # Fail fast on auth. Without this the per-group handler below turns one
    # "nobody signed in" into three identical failures wrapped in a hint about
    # SharePoint scopes, which sends the reader looking in the wrong place.
    _token()

    results, errors = [], {}
    for group in SEARCH_GROUPS:
        ask = [t for t in wanted if t in group]
        if not ask:
            continue
        body = {"requests": [{"entityTypes": ask,
                              "query": {"queryString": query},
                              "from": start, "size": top}]}
        try:
            payload = graph("POST", "/search/query", body=body)
        except GraphError as e:
            # One group failing (usually a missing scope) must not sink the
            # rest of the search — report it alongside what did work.
            errors[",".join(ask)] = str(e)
            continue
        for value in payload.get("value", []):
            for container in value.get("hitsContainers", []):
                for hit in container.get("hits", []):
                    results.append(hit if raw else _compact(hit))

    if not results and errors and len(errors) == len([
            g for g in SEARCH_GROUPS if any(t in wanted for t in g)]):
        raise GraphError("every part of the search failed: "
                         + json.dumps(errors, ensure_ascii=False),
                         "Check 'scopes' — SharePoint/OneDrive needs "
                         "Sites.Read.All, Teams chats need Chat.Read.")

    if not raw:
        results.sort(key=lambda r: (r.get("rank") or 999))
    return {"query": query, "types": wanted, "count": len(results),
            "hits": results, **({"partial_failures": errors} if errors else {})}


# -------------------------------------------------------------- commands ----

def cmd_status(a):
    st = _state()
    report = {
        "msal_installed": True,
        "client_id": client_id(),
        "using_default_client": client_id() == DEFAULT_CLIENT_ID,
        "authority": authority(),
        "signed_in": False,
        "account": st.get("account"),
        "granted_scopes": granted_scopes(),
        "cache": CACHE_PATH if os.path.exists(CACHE_PATH) else None,
    }
    try:
        _msal()
    except GraphError as e:
        report["msal_installed"] = False
        report["hint"] = e.hint
        out(report)
    if granted_scopes():
        try:
            me = graph("GET", "/me", params={
                "$select": "displayName,userPrincipalName,mail"})
            report["signed_in"] = True
            report["me"] = me
        except GraphError as e:
            report["token_error"] = str(e)
            report["hint"] = e.hint
    else:
        report["hint"] = ("Nobody has signed in yet. Run the login command — "
                          "a browser window opens for the user.")
    out(report)


def cmd_login(a):
    scopes = MINIMAL_SCOPES if a.minimal else list(READ_SCOPES)
    scopes = sorted(set(scopes) | set(granted_scopes()) | set(a.add_scope or []))
    _, merged = login(scopes=scopes, interactive=True)
    me = graph("GET", "/me", params={"$select": "displayName,userPrincipalName"})
    _save_state(account=me.get("userPrincipalName"))
    out({"signed_in_as": me.get("userPrincipalName"),
         "display_name": me.get("displayName"),
         "granted_scopes": merged,
         "note": "Token cached and encrypted for this Windows user. "
                 "Later runs are silent until it is revoked or expires."})


def cmd_logout(a):
    for path in (CACHE_PATH, STATE_PATH):
        try:
            os.remove(path)
        except FileNotFoundError:
            pass
    out({"removed": [CACHE_PATH, STATE_PATH],
         "note": "Local token cache cleared. The Entra-side consent still "
                 "stands — revoke that in the account's app permissions."})


def cmd_scopes(a):
    have = set(granted_scopes())
    out({"granted": sorted(have),
         "read_set_missing": [s for s in READ_SCOPES if s not in have],
         "write_areas": {k: {"scope": v[1], "paths": v[0].pattern,
                             "granted": v[1] in have}
                         for k, v in WRITE_AREAS.items()},
         "note": "Add one with: login --add-scope <Scope>. Teams CHANNEL "
                 "messages (ChannelMessage.Read.All) always need an admin."})


def cmd_whoami(a):
    out(graph("GET", "/me"))


def cmd_search(a):
    out(search(a.query, types=(a.types or "").split(",") if a.types else None,
               top=a.top, start=a.start, raw=a.raw))


def cmd_mail_list(a):
    folder = f"/me/mailFolders/{a.folder}/messages" if a.folder else "/me/messages"
    payload = graph("GET", folder, params={
        "$top": a.top, "$orderby": "receivedDateTime desc",
        "$select": "subject,from,receivedDateTime,hasAttachments,isRead,webLink"})
    items = []
    for m in payload.get("value", []):
        frm = ((m.get("from") or {}).get("emailAddress") or {})
        items.append({"subject": m.get("subject"),
                      "sender": frm.get("address"),
                      "received": m.get("receivedDateTime"),
                      "unread": not m.get("isRead"),
                      "has_attachments": m.get("hasAttachments"),
                      "id": m.get("id")})
    out({"folder": a.folder or "all", "count": len(items), "messages": items})


def cmd_mail_get(a):
    m = graph("GET", f"/me/messages/{a.id}")
    out({"subject": m.get("subject"),
         "sender": ((m.get("from") or {}).get("emailAddress") or {}).get("address"),
         "to": [((r.get("emailAddress") or {}).get("address"))
                for r in (m.get("toRecipients") or [])],
         "received": m.get("receivedDateTime"),
         "body": _text((m.get("body") or {}).get("content"), a.max_chars),
         "link": m.get("webLink")})


def cmd_calendar_list(a):
    start = _dt.datetime.now().astimezone()
    end = start + _dt.timedelta(days=a.days)
    payload = graph("GET", "/me/calendarView", params={
        "startDateTime": start.isoformat(), "endDateTime": end.isoformat(),
        "$orderby": "start/dateTime", "$top": 100,
        "$select": "subject,start,end,location,organizer,isAllDay,webLink"},
        headers={"Prefer": 'outlook.timezone="Europe/Budapest"'})
    events = [{"subject": e.get("subject"),
               "start": (e.get("start") or {}).get("dateTime"),
               "end": (e.get("end") or {}).get("dateTime"),
               "all_day": e.get("isAllDay"),
               "location": (e.get("location") or {}).get("displayName"),
               "organizer": (((e.get("organizer") or {}).get("emailAddress")
                              or {}).get("address")),
               "id": e.get("id")}
              for e in payload.get("value", [])]
    out({"from": start.date().isoformat(), "days": a.days,
         "count": len(events), "events": events})


def cmd_calendar_add(a):
    """The one write this CLI performs — deliberately narrow, and gated."""
    body = {
        "subject": a.subject,
        "start": {"dateTime": a.start, "timeZone": a.timezone},
        "end": {"dateTime": a.end, "timeZone": a.timezone},
    }
    if a.body:
        body["body"] = {"contentType": "text", "content": a.body}
    if a.location:
        body["location"] = {"displayName": a.location}
    if a.attendees:
        body["attendees"] = [{"emailAddress": {"address": x.strip()},
                              "type": "required"}
                             for x in a.attendees.split(",") if x.strip()]
    created = graph("POST", "/me/events", body=body, allow_write="calendar")
    out({"created": created.get("id"), "subject": created.get("subject"),
         "start": (created.get("start") or {}).get("dateTime"),
         "attendees_invited": bool(a.attendees),
         "link": created.get("webLink"),
         "note": "An invitation was sent to every attendee listed."})


def cmd_get(a):
    params = dict(p.split("=", 1) for p in (a.param or []))
    out(graph("GET", a.path, params=params or None, all_pages=a.all))


# ------------------------------------------------------------------- cli ----

def build_parser():
    p = argparse.ArgumentParser(
        prog="microsoft_graph.py",
        description="Search and read Microsoft 365 with the signed-in user's "
                    "own delegated permissions.")
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("status", help="preflight: signed in? which scopes?"
                   ).set_defaults(func=cmd_status)
    lg = sub.add_parser("login", help="sign in (opens a browser window)")
    lg.add_argument("--minimal", action="store_true",
                    help="ask only for User.Read + Mail.Read")
    lg.add_argument("--add-scope", action="append",
                    help="request one more scope, keeping the current ones")
    lg.set_defaults(func=cmd_login)
    sub.add_parser("logout", help="delete the local token cache"
                   ).set_defaults(func=cmd_logout)
    sub.add_parser("scopes", help="what is granted, what is missing"
                   ).set_defaults(func=cmd_scopes)
    sub.add_parser("whoami", help="GET /me").set_defaults(func=cmd_whoami)

    s = sub.add_parser("search", help="search mail, files, SharePoint, chats")
    s.add_argument("--query", required=True, help="KQL: words, from:, filetype:")
    s.add_argument("--types", help="comma list; default message,driveItem,chatMessage")
    s.add_argument("--top", type=int, default=15)
    s.add_argument("--start", type=int, default=0, help="paging offset")
    s.add_argument("--raw", action="store_true", help="full Graph hits")
    s.set_defaults(func=cmd_search)

    mail = sub.add_parser("mail").add_subparsers(dest="sub", required=True)
    ml = mail.add_parser("list")
    ml.add_argument("--top", type=int, default=20)
    ml.add_argument("--folder", help="inbox, sentitems, drafts, ...")
    ml.set_defaults(func=cmd_mail_list)
    mg = mail.add_parser("get")
    mg.add_argument("--id", required=True)
    mg.add_argument("--max-chars", type=int, default=8000)
    mg.set_defaults(func=cmd_mail_get)

    cal = sub.add_parser("calendar").add_subparsers(dest="sub", required=True)
    cl = cal.add_parser("list")
    cl.add_argument("--days", type=int, default=7)
    cl.set_defaults(func=cmd_calendar_list)
    ca = cal.add_parser("add", help="create an event (needs Calendars.ReadWrite)")
    ca.add_argument("--subject", required=True)
    ca.add_argument("--start", required=True, help="2026-08-25T10:00:00")
    ca.add_argument("--end", required=True)
    ca.add_argument("--timezone", default="Europe/Budapest")
    ca.add_argument("--location")
    ca.add_argument("--body")
    ca.add_argument("--attendees", help="comma-separated addresses; they get invited")
    ca.set_defaults(func=cmd_calendar_add)

    g = sub.add_parser("get", help="any Graph GET, e.g. /me/drive/recent")
    g.add_argument("path")
    g.add_argument("--param", action="append", help='repeatable: $top=5')
    g.add_argument("--all", action="store_true", help="follow nextLink (max 10 pages)")
    g.set_defaults(func=cmd_get)

    return p


def main():
    args = build_parser().parse_args()
    try:
        args.func(args)
    except SystemExit:
        raise
    except GraphError as e:
        fail(e, e.hint)
    except KeyboardInterrupt:
        fail("interrupted")
    except Exception as e:
        fail(f"{type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
