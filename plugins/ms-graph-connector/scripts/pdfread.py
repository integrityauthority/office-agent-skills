"""One way to read a PDF, shared by the file-reader and m365-graph skills.

This file is the source. `tools/sync_plugins.py` copies it into
``.agents/skills/<skill>/scripts/pdfread.py`` for every skill that needs it,
and CI fails if a copy drifts. Edit it here, run the sync, commit both.

The order, for every page asked for:

1. the text layer, on this machine (pypdfium2; tables by pdfplumber when
   installed; pdfminer.six when pypdfium2 is missing);
2. your own OCR server (``OCR_URL``) for pages without one -- always;
3. only with ``ocr="cloud"``, which the agent passes after the user agreed:
   Azure Document Intelligence (when configured), then the cloud vision model
   (``OCR_CLOUD_*``), for pages still empty.

When pages stay empty and a cloud model could read them, the result carries
``cloud_offer`` instead of sending anything.

Configuration, all from the environment:

    OCR_URL, OCR_MODEL, OCR_API_KEY        your OpenAI-compatible OCR server
    OCR_CLOUD_URL, OCR_CLOUD_MODEL         a model behind the OpenAI Responses API
    OCR_CLOUD_API_KEY                      its key, or OCR_CLOUD_API_KEY_ENV:
                                           the name of the variable holding it
    AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT   optional Azure layout model
    AZURE_DOCUMENT_INTELLIGENCE_KEY

Every library here is permissively licensed. PyMuPDF (AGPL-3.0) is not used.
"""
from __future__ import annotations

import base64
import html as _html
import io
import json
import os
import re
import urllib.error
import urllib.parse
import urllib.request

TEXT_MIN = 20                 # fewer characters than this: the page has no text layer
OCR_PAGES_PER_CALL = 10       # one call never waits on more OCR than this
SERVER_DPI = 200
CLOUD_DPI = 250
OCR_MODEL_DEFAULT = "deepseek-ai/DeepSeek-OCR-2"
DEEPSEEK_OCR_PROMPT = "<|grounding|>Convert the document to markdown."
TRANSCRIBE_PROMPT = (
    "This is one scanned document page. Transcribe every visible character, "
    "printed and filled in, including numbers and dates. Do not summarise."
)

# Labels for where a page's text came from. The same in both skills.
TEXT_LAYER = "text-layer"
PDFMINER = "pdfminer"
OCR_SERVER = "ocr-server"
CLOUD_DI = "azure-document-intelligence"
CLOUD_VISION = "cloud-vision"


class PdfError(Exception):
    def __init__(self, message, hint=""):
        super().__init__(message)
        self.hint = hint


# ------------------------------------------------------------------ small ---

def ranges(indexes):
    """[0, 1, 2, 5] -> '1-3,6' (1-based)."""
    out = []
    for i in sorted(indexes):
        if out and i == out[-1][1] + 1:
            out[-1][1] = i
        else:
            out.append([i, i])
    return ",".join(f"{a + 1}" if a == b else f"{a + 1}-{b + 1}" for a, b in out)


def parse_pages(spec, count):
    """'1-3,7' (1-based) -> sorted 0-based indexes inside the document."""
    if spec in (None, "", "all"):
        return list(range(count))
    if isinstance(spec, (list, tuple, set)):
        return sorted(i for i in spec if 0 <= i < count)
    chosen = set()
    for part in str(spec).replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            chosen.update(range((int(a) if a else 1) - 1, min(int(b) if b else count, count)))
        else:
            chosen.add(int(part) - 1)
    return sorted(i for i in chosen if 0 <= i < count)


def _env(name, default=""):
    return (os.environ.get(name) or default).strip()


# --------------------------------------------------------------- settings ---

def settings():
    """{'server': (url, model, key) | None, 'cloud': (url, model, key) | None,
    'di': (endpoint, key) | None}"""
    server_url = _env("OCR_URL").rstrip("/")
    server = (server_url, _env("OCR_MODEL", OCR_MODEL_DEFAULT), _env("OCR_API_KEY")) if server_url else None
    cloud_key = _env("OCR_CLOUD_API_KEY")
    if not cloud_key and _env("OCR_CLOUD_API_KEY_ENV"):
        cloud_key = _env(_env("OCR_CLOUD_API_KEY_ENV"))
    cloud_url, cloud_model = _env("OCR_CLOUD_URL").rstrip("/"), _env("OCR_CLOUD_MODEL")
    cloud = (cloud_url, cloud_model, cloud_key) if cloud_url and cloud_model and cloud_key else None
    di_endpoint = _env("AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT").rstrip("/")
    di_key = _env("AZURE_DOCUMENT_INTELLIGENCE_KEY")
    di = (di_endpoint, di_key) if di_endpoint and di_key else None
    return {"server": server, "cloud": cloud, "di": di}


def describe_settings():
    s = settings()
    return {
        "ocr_server": {"configured": bool(s["server"]), "model": s["server"][1] if s["server"] else None},
        "ocr_cloud": {"configured": bool(s["cloud"]), "model": s["cloud"][1] if s["cloud"] else None,
                      "used_only_when": "ocr='cloud' is passed, after the user agreed"},
        "azure_document_intelligence": {"configured": bool(s["di"])},
    }


def auth_headers(url, key):
    """Azure takes the key in `api-key`; everyone else as a bearer token."""
    if not key:
        return {}
    if (urllib.parse.urlsplit(url).hostname or "").endswith(".azure.com"):
        return {"api-key": key}
    return {"Authorization": f"Bearer {key}"}


# -------------------------------------------------------------- engines ---

def _post_json(url, payload, headers, timeout=180):
    request = urllib.request.Request(
        url, data=json.dumps(payload).encode("utf-8"), method="POST",
        headers={"Content-Type": "application/json", **headers})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def _chat_text(payload):
    choice = ((payload or {}).get("choices") or [{}])[0]
    content = (choice.get("message") or {}).get("content") or ""
    if isinstance(content, list):
        content = "".join(part.get("text") or "" for part in content if isinstance(part, dict))
    return str(content).strip()


def _table_to_tsv(match):
    rows = []
    for row in re.findall(r"<tr\b.*?</tr>", match.group(0), flags=re.S | re.I):
        cells = re.findall(r"<t[dh]\b[^>]*>(.*?)</t[dh]>", row, flags=re.S | re.I)
        rows.append("\t".join(_html.unescape(re.sub(r"<[^>]+>", "", c)).strip() for c in cells))
    return "\n" + "\n".join(r for r in rows if r.strip()) + "\n"


def clean_ocr(text):
    """DeepSeek-OCR's grounding mode tags each block with a label and a box,
    and returns tables as HTML. Keep the words; one table row per line."""
    text = re.sub(r"<\|det\|>.*?<\|/det\|>", "", text, flags=re.S)
    text = re.sub(r"<\|ref\|>.*?<\|/ref\|>", "", text, flags=re.S)
    text = re.sub(r"<\|[^|>]{1,40}\|>", "", text)
    text = re.sub(r"<table\b.*?</table>", _table_to_tsv, text, flags=re.S | re.I)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def ocr_server_page(png, url, model, key):
    deepseek = "deepseek-ocr" in model.lower()
    payload = _post_json(f"{url}/chat/completions", {
        "model": model,
        "temperature": 0,
        **({"skip_special_tokens": False} if deepseek else {}),
        "messages": [{"role": "user", "content": [
            {"type": "text", "text": DEEPSEEK_OCR_PROMPT if deepseek else TRANSCRIBE_PROMPT},
            {"type": "image_url", "image_url": {
                "url": "data:image/png;base64," + base64.b64encode(png).decode()}},
        ]}],
    }, auth_headers(url, key))
    return clean_ocr(_chat_text(payload))


def cloud_vision_page(png, url, model, key):
    payload = _post_json(f"{url}/responses", {
        "model": model,
        "input": [{"role": "user", "content": [
            {"type": "input_text", "text": TRANSCRIBE_PROMPT},
            {"type": "input_image",
             "image_url": "data:image/png;base64," + base64.b64encode(png).decode()},
        ]}],
    }, auth_headers(url, key))
    chunks = [part.get("text") or ""
              for item in payload.get("output") or []
              for part in item.get("content") or []
              if part.get("type") in ("output_text", "text")]
    return ("".join(chunks) or _chat_text(payload)).strip()


def azure_di_pages(data, endpoint, key):
    """{page_index: text} from Azure Document Intelligence prebuilt-layout."""
    from azure.ai.documentintelligence import DocumentIntelligenceClient
    from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
    from azure.core.credentials import AzureKeyCredential

    client = DocumentIntelligenceClient(endpoint, AzureKeyCredential(key))
    result = client.begin_analyze_document(
        "prebuilt-layout", AnalyzeDocumentRequest(bytes_source=data)).result()
    out = {}
    for page in getattr(result, "pages", None) or []:
        lines = [(getattr(line, "content", None) or "").strip() for line in getattr(page, "lines", None) or []]
        text = "\n".join(line for line in lines if line).strip()
        if text:
            out[(getattr(page, "page_number", 1) or 1) - 1] = text
    return out


# ------------------------------------------------------------ text layer ---

def _pdfium():
    try:
        import pypdfium2
        return pypdfium2
    except ImportError:
        return None


def _plumber(data):
    try:
        import pdfplumber
        return pdfplumber.open(io.BytesIO(data))
    except Exception:
        return None


def _tables(plumber, index):
    if plumber is None or index >= len(plumber.pages):
        return ""
    blocks = []
    try:
        for table in plumber.pages[index].extract_tables() or []:
            rows = ["\t".join((cell or "").replace("\n", " ").strip() for cell in row) for row in table or []]
            block = "\n".join(row for row in rows if row.strip())
            if block:
                blocks.append(block)
    except Exception:
        return ""
    return "\n\n".join(blocks)


def open_pdf(data):
    """A pypdfium2 document, or None when pypdfium2 is not installed."""
    pdfium = _pdfium()
    if pdfium is None:
        return None
    try:
        return pdfium.PdfDocument(data)
    except Exception as exc:
        if "password" in str(exc).lower():
            raise PdfError("the PDF is password-protected", "Ask the user for an unprotected copy.") from exc
        raise PdfError(f"the PDF could not be opened ({type(exc).__name__})",
                       "The file may be damaged or not a PDF.") from exc


def page_count(data):
    doc = open_pdf(data)
    if doc is not None:
        try:
            return len(doc)
        finally:
            doc.close()
    try:
        from pdfminer.pdfpage import PDFPage
        return sum(1 for _ in PDFPage.get_pages(io.BytesIO(data)))
    except ImportError as exc:
        raise PdfError("no PDF reader is installed", "pip install pypdfium2") from exc


def render_png(doc, index, dpi):
    page = doc[index]
    try:
        image = page.render(scale=dpi / 72).to_pil()
    finally:
        page.close()
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


def _text_layer(doc, data, indexes, tables):
    texts = {}
    if doc is not None:
        plumber = _plumber(data) if tables else None
        try:
            for index in indexes:
                page = doc[index]
                try:
                    textpage = page.get_textpage()
                    try:
                        texts[index] = (textpage.get_text_bounded() or "").replace("\r\n", "\n").strip()
                    finally:
                        textpage.close()
                finally:
                    page.close()
                if len(texts[index]) >= TEXT_MIN and plumber is not None:
                    found = _tables(plumber, index)
                    if found:
                        texts[index] += "\n\n[tables]\n" + found
        finally:
            if plumber is not None:
                plumber.close()
        return texts, TEXT_LAYER
    from pdfminer.high_level import extract_text
    for index in indexes:
        try:
            texts[index] = extract_text(io.BytesIO(data), page_numbers=[index]).strip()
        except Exception:
            texts[index] = ""
    return texts, PDFMINER


# ----------------------------------------------------------------- read ---

def read_pages(data, pages=None, ocr="auto", tables=True, ocr_limit=OCR_PAGES_PER_CALL):
    """Read the requested pages in the order described at the top.

    ocr: "auto"  text layer, then your OCR server for pages without one
         "never" text layer only
         "force" your OCR server for every requested page
         "cloud" as auto, then the cloud models for pages still empty

    Returns a dict: count, selected, texts {index: text}, sources {index:
    label or ""}, notes [str], ocr_deferred [indexes beyond ocr_limit],
    cloud_offer {pages, how} or None.
    """
    if ocr not in ("auto", "never", "force", "cloud"):
        raise PdfError(f"ocr must be auto, never, force or cloud, not {ocr!r}")
    notes = []
    doc = open_pdf(data)
    try:
        count = len(doc) if doc is not None else page_count(data)
        selected = parse_pages(pages, count)
        if not selected:
            raise PdfError(f"no pages match {pages!r}; the document has {count}")
        texts, layer = _text_layer(doc, data, selected, tables)
        if doc is None:
            notes.append("Read with pdfminer.six: pypdfium2 is not installed, so scanned pages "
                         "cannot be rendered for OCR.")
        sources = {i: (layer if len(texts.get(i, "")) >= TEXT_MIN else "") for i in selected}
        if ocr == "never":
            need = []
        elif ocr == "force":
            need = list(selected)
        else:
            need = [i for i in selected if not sources[i]]

        conf = settings()
        deferred = []
        if need and doc is not None:
            todo, deferred = need[:ocr_limit], need[ocr_limit:]
            engines = []
            if conf["server"]:
                engines.append((OCR_SERVER, conf["server"], ocr_server_page, SERVER_DPI))
            elif ocr != "never":
                notes.append("No OCR server is configured (OCR_URL is empty).")
            done = set()
            for label, config, call, dpi in engines:
                _run_engine(label, config, call, dpi, doc, todo, done, texts, sources, notes,
                            accept_short=(ocr == "force"))
            if ocr == "cloud":
                still = [i for i in todo if i not in done]
                if still and conf["di"]:
                    try:
                        di = azure_di_pages(data, *conf["di"])
                        for i in still:
                            if len(di.get(i, "")) >= TEXT_MIN:
                                texts[i], sources[i] = di[i], CLOUD_DI
                                done.add(i)
                    except ImportError:
                        notes.append("Azure Document Intelligence is configured but its package "
                                     "(azure-ai-documentintelligence) is not installed.")
                    except Exception as exc:
                        notes.append(f"Azure Document Intelligence did not answer ({type(exc).__name__}).")
                still = [i for i in todo if i not in done]
                if still and conf["cloud"]:
                    _run_engine(CLOUD_VISION, conf["cloud"], cloud_vision_page, CLOUD_DPI, doc,
                                still, done, texts, sources, notes, accept_short=False)
                elif still and not conf["di"]:
                    notes.append("ocr='cloud' was asked for but no cloud model is configured "
                                 "(OCR_CLOUD_URL, OCR_CLOUD_MODEL and OCR_CLOUD_API_KEY or "
                                 "OCR_CLOUD_API_KEY_ENV).")
        unrenderable = need if doc is None else []
    finally:
        if doc is not None:
            doc.close()

    empty = [i for i in selected if not sources[i] and i not in deferred and i not in unrenderable]
    # A short text layer (a cover page that says only "Annex") went to OCR as
    # if it were a scan. If OCR found nothing better, keep what the page had
    # rather than reporting it empty.
    for i in selected:
        if not sources[i] and texts.get(i, "").strip():
            sources[i] = layer
    cloud_offer = None
    if empty and ocr in ("auto", "force") and (conf["cloud"] or conf["di"]):
        cloud_offer = {
            "pages": ranges(empty),
            "how": "Ask the user. Only if they agree, call again with ocr='cloud': these pages "
                   "then go to the cloud model, outside the organisation.",
        }
    if deferred:
        notes.append(f"OCR is capped at {ocr_limit} pages per call; page(s) {ranges(deferred)} "
                     "were not read. Call again with those pages.")
    return {"count": count, "selected": selected, "texts": texts, "sources": sources,
            "notes": notes, "ocr_deferred": deferred, "cloud_offer": cloud_offer}


def _run_engine(label, config, call, dpi, doc, indexes, done, texts, sources, notes, accept_short):
    failed = []
    for index in [i for i in indexes if i not in done]:
        try:
            text = call(render_png(doc, index, dpi), *config)
        except urllib.error.HTTPError as exc:
            failed.append(f"{index + 1} (HTTP {exc.code})")
            continue
        except (urllib.error.URLError, OSError) as exc:
            # Unreachable: do not wait out the timeout once per page.
            notes.append(f"{label} is unreachable ({type(exc).__name__}); "
                         "its remaining pages were skipped.")
            break
        except ValueError as exc:
            failed.append(f"{index + 1} ({type(exc).__name__})")
            continue
        if len(text.strip()) >= TEXT_MIN or (accept_short and text.strip()):
            texts[index], sources[index] = text.strip(), label
            done.add(index)
    if failed:
        notes.append(f"{label} gave no answer for page(s) {', '.join(failed)}.")
