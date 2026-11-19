"""One way to read a PDF, shared by the file-reader and m365-graph skills.

This file is the source. `tools/sync_plugins.py` copies it into
``.agents/skills/<skill>/scripts/pdfread.py`` for every skill that needs it,
and CI fails if a copy drifts. Edit it here, run the sync, commit both.

A page with a text layer is read on this machine, always. A scanned page goes
to the engine the ``ocr`` argument names, cheapest and most private first:

    auto    Windows' own OCR, on this machine; your OCR server when there is
            no Windows OCR here
    server  your OCR server (``OCR_URL``), which reads tables and numbers
            better; Windows OCR when the server cannot be reached
    cloud   the cloud vision model (``OCR_CLOUD_*``) -- only after the user
            agreed; then the server, then Windows OCR
    force   like server, for every page, even one with a text layer
    never   text layers only

So the agent starts with the fast local reading, and asks for ``server`` when
that text looks wrong; ``ocr_upgrade`` in the result says which pages, and
``cloud_offer`` lists pages nothing local could read. Neither sends anything.

Configuration, all from the environment:

    OCR_LOCAL_LANGUAGE                     Windows OCR language, e.g. hu; default:
                                           the user's Windows languages
    OCR_LOCAL=off                          do not use Windows OCR
    OCR_URL, OCR_MODEL, OCR_API_KEY        your OpenAI-compatible OCR server
    OCR_CLOUD_URL, OCR_CLOUD_MODEL         a model behind the OpenAI Responses API
    OCR_CLOUD_API_KEY                      its key, or OCR_CLOUD_API_KEY_ENV:
                                           the name of the variable holding it

Every library here is permissively licensed: pypdfium2, pdfplumber,
pdfminer.six, pywinrt (MIT). PyMuPDF (AGPL-3.0) is not used.
"""
from __future__ import annotations

import base64
import html as _html
import io
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

TEXT_MIN = 20                 # fewer characters than this: the page has no text layer
OCR_PAGES_PER_CALL = 10       # one call never waits on more remote OCR than this
LOCAL_PAGES_PER_CALL = 50     # Windows OCR takes well under a second a page
SERVER_DPI = 200
CLOUD_DPI = 250
LOCAL_DPI = 300               # Windows OCR reads small print better from more pixels
LOCAL_MAX_PIXELS = 9_999      # OcrEngine.MaxImageDimension is 10000
OCR_MODEL_DEFAULT = "deepseek-ai/DeepSeek-OCR-2"
DEEPSEEK_OCR_PROMPT = "<|grounding|>Convert the document to markdown."
TRANSCRIBE_PROMPT = (
    "This is one scanned document page. Transcribe every visible character, "
    "printed and filled in, including numbers and dates. Do not summarise."
)
MODES = ("auto", "server", "cloud", "force", "never")

# Labels for where a page's text came from. The same in both skills.
TEXT_LAYER = "text-layer"
PDFMINER = "pdfminer"
LOCAL_OCR = "windows-ocr"
OCR_SERVER = "ocr-server"
CLOUD_VISION = "cloud-vision"


class PdfError(Exception):
    def __init__(self, message, hint=""):
        super().__init__(message)
        self.hint = hint


class EngineUnavailable(Exception):
    """An OCR engine that cannot run here at all; the next one is tried."""


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

def local_settings():
    """(language,) when Windows OCR can run here, else None. language "" means
    the user's own Windows languages."""
    if _env("OCR_LOCAL").lower() in ("0", "off", "false", "no"):
        return None
    if sys.platform != "win32":
        return None
    import importlib.util
    try:
        if importlib.util.find_spec("winrt.windows.media.ocr") is None:
            return None
    except ModuleNotFoundError:  # no winrt package at all
        return None
    return (_env("OCR_LOCAL_LANGUAGE"),)


def settings():
    """{'local': (language,) | None, 'server': (url, model, key) | None,
    'cloud': (url, model, key) | None}"""
    server_url = _env("OCR_URL").rstrip("/")
    server = (server_url, _env("OCR_MODEL", OCR_MODEL_DEFAULT), _env("OCR_API_KEY")) if server_url else None
    cloud_key = _env("OCR_CLOUD_API_KEY")
    if not cloud_key and _env("OCR_CLOUD_API_KEY_ENV"):
        cloud_key = _env(_env("OCR_CLOUD_API_KEY_ENV"))
    cloud_url, cloud_model = _env("OCR_CLOUD_URL").rstrip("/"), _env("OCR_CLOUD_MODEL")
    cloud = (cloud_url, cloud_model, cloud_key) if cloud_url and cloud_model and cloud_key else None
    return {"local": local_settings(), "server": server, "cloud": cloud}


def describe_settings():
    s = settings()
    return {
        "ocr_local": {"available": bool(s["local"]), "engine": "Windows OCR" if s["local"] else None,
                      "language": (s["local"][0] or "user profile") if s["local"] else None},
        "ocr_server": {"configured": bool(s["server"]), "model": s["server"][1] if s["server"] else None},
        "ocr_cloud": {"configured": bool(s["cloud"]), "model": s["cloud"][1] if s["cloud"] else None,
                      "used_only_when": "ocr='cloud' is passed, after the user agreed"},
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


def local_ocr_page(png, language=""):
    """Windows.Media.Ocr on a PNG, on this machine. Runs its own event loop on
    a worker thread, so it works whether or not the caller is inside one."""
    import asyncio
    import concurrent.futures

    async def recognize():
        from winrt.windows.globalization import Language
        from winrt.windows.graphics.imaging import BitmapDecoder
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream

        if language:
            engine = OcrEngine.try_create_from_language(Language(language))
            if engine is None:
                available = [l.language_tag for l in OcrEngine.available_recognizer_languages]
                raise EngineUnavailable(
                    f"Windows OCR has no '{language}' language here (installed: {', '.join(available)}); "
                    "add it under Settings > Time & language > Language, or change OCR_LOCAL_LANGUAGE")
        else:
            engine = OcrEngine.try_create_from_user_profile_languages()
            if engine is None:
                raise EngineUnavailable("Windows OCR has no language for this user's profile")
        stream = InMemoryRandomAccessStream()
        writer = DataWriter(stream)
        writer.write_bytes(png)
        await writer.store_async()
        writer.detach_stream()
        stream.seek(0)
        bitmap = await (await BitmapDecoder.create_async(stream)).get_software_bitmap_async()
        result = await engine.recognize_async(bitmap)
        return "\n".join(line.text for line in result.lines)

    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
        return pool.submit(asyncio.run, recognize()).result()


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


def render_png(doc, index, dpi, max_pixels=None):
    page = doc[index]
    try:
        scale = dpi / 72
        if max_pixels:
            width, height = page.get_size()
            scale = min(scale, max_pixels / max(width, height))
        image = page.render(scale=scale).to_pil()
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

def _ladder(mode, conf):
    """The engines a mode tries, in order: (label, run(png_args), dpi, limit, max_pixels)."""
    local = (LOCAL_OCR, conf["local"], local_ocr_page, LOCAL_DPI, LOCAL_PAGES_PER_CALL, LOCAL_MAX_PIXELS)
    server = (OCR_SERVER, conf["server"], ocr_server_page, SERVER_DPI, OCR_PAGES_PER_CALL, None)
    cloud = (CLOUD_VISION, conf["cloud"], cloud_vision_page, CLOUD_DPI, OCR_PAGES_PER_CALL, None)
    order = {"auto": [local, server], "server": [server, local], "force": [server, local],
             "cloud": [cloud, server, local], "never": []}[mode]
    return [engine for engine in order if engine[1]]


def read_pages(data, pages=None, ocr="auto", tables=True, ocr_limit=None):
    """Read the requested pages in the order described at the top.

    Returns a dict: count, selected, texts {index: text}, sources {index:
    label or ""}, notes [str], ocr_deferred [indexes beyond the per-call cap],
    ocr_upgrade {pages, how} or None, cloud_offer {pages, how} or None.
    """
    if ocr not in MODES:
        raise PdfError(f"ocr must be one of {', '.join(MODES)}, not {ocr!r}")
    notes = []
    conf = settings()
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
        need = list(selected) if ocr == "force" else [i for i in selected if not sources[i]]

        deferred, unrenderable = [], []
        if ocr == "cloud" and not conf["cloud"]:
            notes.append("ocr='cloud' was asked for but no cloud model is configured (OCR_CLOUD_URL, "
                         "OCR_CLOUD_MODEL and OCR_CLOUD_API_KEY or OCR_CLOUD_API_KEY_ENV).")
        if ocr == "server" and not conf["server"]:
            notes.append("ocr='server' was asked for but no OCR server is configured (OCR_URL is empty).")
        engines = _ladder(ocr, conf)
        if need and doc is None:
            unrenderable = need
        elif need and ocr != "never" and not engines:
            notes.append("No OCR engine is available: no Windows OCR here and OCR_URL is empty.")
        elif need:
            done, attempted = set(), set()
            for label, config, call, dpi, limit, max_pixels in engines:
                todo = [i for i in need if i not in done][:ocr_limit or limit]
                attempted.update(todo)
                _run_engine(label, config, call, dpi, max_pixels, doc, todo, done, texts, sources, notes,
                            accept_short=(ocr == "force"))
            deferred = [i for i in need if i not in attempted]
    finally:
        if doc is not None:
            doc.close()

    empty = [i for i in selected if not sources[i] and i not in deferred and i not in unrenderable]
    # A short text layer (a cover page that says only "Annex") went to OCR as
    # if it were a scan. If OCR found nothing better, keep what the page had.
    for i in selected:
        if not sources[i] and texts.get(i, "").strip():
            sources[i] = layer

    local_pages = [i for i in selected if sources[i] == LOCAL_OCR]
    ocr_upgrade = None
    if local_pages and conf["server"]:
        ocr_upgrade = {
            "pages": ranges(local_pages),
            "how": "These pages were read by Windows OCR on this machine: fast, but tables come "
                   "back as loose lines and digits can slip. If the text looks wrong, call again "
                   "with ocr='server' and those pages; the organisation's OCR server reads them "
                   "with table structure.",
        }
    cloud_offer = None
    if empty and ocr in ("auto", "server", "force") and conf["cloud"]:
        cloud_offer = {
            "pages": ranges(empty),
            "how": "Ask the user. Only if they agree, call again with ocr='cloud': these pages "
                   "then go to the cloud model, outside the organisation.",
        }
    if deferred:
        notes.append(f"OCR is capped per call; page(s) {ranges(deferred)} were not read. "
                     "Call again with those pages.")
    return {"count": count, "selected": selected, "texts": texts, "sources": sources, "notes": notes,
            "ocr_deferred": deferred, "ocr_upgrade": ocr_upgrade, "cloud_offer": cloud_offer}


def _run_engine(label, config, call, dpi, max_pixels, doc, indexes, done, texts, sources, notes, accept_short):
    failed = []
    for index in [i for i in indexes if i not in done]:
        try:
            text = call(render_png(doc, index, dpi, max_pixels), *config)
        except EngineUnavailable as exc:
            notes.append(f"{label}: {exc}.")
            break
        except urllib.error.HTTPError as exc:
            failed.append(f"{index + 1} (HTTP {exc.code})")
            continue
        except (urllib.error.URLError, OSError) as exc:
            # Unreachable: do not wait out the timeout once per page.
            notes.append(f"{label} is unreachable ({type(exc).__name__}); "
                         "its remaining pages were skipped.")
            break
        except Exception as exc:
            failed.append(f"{index + 1} ({type(exc).__name__})")
            continue
        if len(text.strip()) >= TEXT_MIN or (accept_short and text.strip()):
            texts[index], sources[index] = text.strip(), label
            done.add(index)
    if failed:
        notes.append(f"{label} gave no answer for page(s) {', '.join(failed)}.")
