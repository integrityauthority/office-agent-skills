"""Read files on disk as text: PDF (with OCR for scanned pages), Outlook .msg
and .eml e-mail with their attachments, Word, Excel, PowerPoint, HTML and
plain text. One call per file instead of a script per file.

    from files import read_file, save_attachments, render_pages, status

    read_file(r"C:\\path\\report.pdf")                 # every page, OCR where needed
    read_file(r"C:\\path\\report.pdf", pages="3-5")
    read_file(r"C:\\path\\mail.msg")                   # headers, body, attachment list
    read_file(r"C:\\path\\mail.msg", attachment="2")   # read attachment 2 in place
    save_attachments(r"C:\\path\\mail.msg")            # write them to a folder
    render_pages(r"C:\\path\\scan.pdf", pages="1")     # PNG to look at

Everything here reads; nothing changes the source file. A PDF page that has a
text layer never leaves this machine. A scanned page goes to your own OCR
server when ``OCR_URL`` is set, and to a cloud vision model (``OCR_CLOUD_*``)
only when the caller passes ``ocr="cloud"``, after the user agreed. The
variables are listed in the repository README under "Configuration".

Libraries, all permissive: pypdfium2 (Apache/BSD) for PDF text and
rendering, pdfminer.six (MIT) as a text fallback, olefile (BSD) for .msg,
openpyxl (MIT) for .xlsx, xlrd (BSD) for Excel 97-2003 .xls. Word and PowerPoint are read from their XML with the
standard library. Each one is optional; a missing one is named in the answer.
"""
from __future__ import annotations

import argparse
import html as _html
import io
import json
import os
import re
import sys
import zipfile
from email import policy
from email.parser import BytesParser
from html.parser import HTMLParser
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))
import archive  # noqa: E402
import docmeta  # noqa: E402
import msgfile  # noqa: E402
import pdfread  # noqa: E402

TEXT_LIMIT = 40_000
MAX_FILE_BYTES = 100 * 1024 * 1024
XLSX_MAX_SHEETS = 20
XLSX_MAX_ROWS = 300
XLSX_MAX_COLS = 30

TEXT_EXTENSIONS = {".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".log", ".ini",
                   ".yaml", ".yml", ".py", ".ps1", ".sql", ".rtf"}
HTML_EXTENSIONS = {".html", ".htm", ".xhtml"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif"}
IMAGE_DPI = 300                 # when the image does not say; rendering at LOCAL_DPI keeps its pixels
IMAGE_MAX_FRAMES = 50


class FileError(Exception):
    def __init__(self, message, hint=""):
        super().__init__(message)
        self.hint = hint


# ------------------------------------------------------------------ helpers ---

def _cache_dir(*parts):
    base = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "hermes" / "cache" / "files"
    folder = base.joinpath(*parts)
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def _safe_name(name, fallback="file"):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", (name or "").strip()).strip(". ")
    return (name or fallback)[:120]


def _unique(path: Path) -> Path:
    if not path.exists():
        return path
    for n in range(2, 1000):
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        if not candidate.exists():
            return candidate
    raise FileError(f"too many files named {path.name} in {path.parent}")


def _resolve(path) -> Path:
    if not path:
        raise FileError("no path given", "Pass the full path of the file.")
    text = os.path.expandvars(os.path.expanduser(str(path).strip().strip('"').strip("'")))
    p = Path(text)
    if not p.is_absolute():
        p = Path.cwd() / p
    if not p.exists():
        raise FileError(f"no such file: {p}",
                        "Check the path; list the folder first if unsure of the name.")
    if p.is_dir():
        raise FileError(f"{p} is a folder, not a file")
    if p.stat().st_size > MAX_FILE_BYTES:
        raise FileError(f"{p.name} is {p.stat().st_size} bytes, above the {MAX_FILE_BYTES} byte cap")
    return p


_parse_pages = pdfread.parse_pages
_ranges = pdfread.ranges


def _decode_text(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-16") if data[:2] in (b"\xff\xfe", b"\xfe\xff") else ("utf-8-sig",):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            pass
    return data.decode("cp1250", "replace")


class _HtmlText(HTMLParser):
    BLOCK = {"p", "div", "br", "tr", "li", "h1", "h2", "h3", "h4", "h5", "h6",
             "table", "blockquote", "pre", "hr", "section", "article"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "head", "title"):
            self.skip += 1
        elif tag in self.BLOCK:
            self.parts.append("\n")
        elif tag in ("td", "th"):
            self.parts.append("\t")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "head", "title"):
            self.skip = max(0, self.skip - 1)
        elif tag in self.BLOCK:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self.skip:
            self.parts.append(data)


def html_to_text(markup: str) -> str:
    parser = _HtmlText()
    try:
        parser.feed(markup)
        parser.close()
        text = "".join(parser.parts)
    except Exception:
        text = _html.unescape(re.sub(r"<[^>]+>", " ", markup))
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t\r\f\v]+\n", "\n", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _clip(text, limit):
    text = text or ""
    if len(text) <= limit:
        return text, False
    return text[:limit], True


# ---------------------------------------------------------------- detection ---

def detect_kind(data: bytes, name: str = "") -> str:
    ext = Path(name).suffix.lower()
    if data.startswith(b"%PDF") or data[:1024].find(b"%PDF-") >= 0:
        return "pdf"
    if data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        if ext == ".msg" or msgfile.is_msg(data):
            return "msg"
        if ext in (".xls", ".xlt") or _ole_has_workbook(data):
            return "xls"
        return "legacy-office"
    if data.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                names = set(zf.namelist())
        except zipfile.BadZipFile:
            return "zip"
        if "word/document.xml" in names:
            return "docx"
        if any(n.startswith("xl/") for n in names):
            return "xlsx"
        if any(n.startswith("ppt/slides/") for n in names):
            return "pptx"
        return "zip"
    if (data.startswith((b"\x89PNG\r\n\x1a\n", b"\xff\xd8\xff", b"II*\x00", b"MM\x00*"))
            or (ext in IMAGE_EXTENSIONS and data.startswith((b"BM", b"GIF8")))):
        return "image"
    if ext == ".eml":
        return "eml"
    if ext in HTML_EXTENSIONS:
        return "html"
    if ext in TEXT_EXTENSIONS:
        return "text"
    head = data[:2048]
    if re.match(rb"(?i)(received|return-path|from|message-id|mime-version|delivered-to|x-[\w-]+):", head):
        return "eml"
    if b"\x00" not in head:
        stripped = head.lstrip().lower()
        if stripped.startswith(b"<!doctype html") or stripped.startswith(b"<html"):
            return "html"
        return "text"
    return "binary"


# -------------------------------------------------------------------- PDF ---
# The reading itself (text layer, OCR server, cloud) is shared/pdfread.py,
# the same module m365-graph uses. Here it is only formatted for the answer.

def read_pdf(data, pages=None, ocr="auto", max_chars=TEXT_LIMIT):
    r = pdfread.read_pages(data, pages=pages, ocr=ocr)
    selected, texts, sources = r["selected"], r["texts"], r["sources"]
    parts = []
    for index in selected:
        body = texts.get(index, "") if sources[index] else ""
        parts.append(f"--- page {index + 1} ---\n{body or '[no text on this page]'}")
    full = "\n\n".join(parts)
    text, truncated = _clip(full, max_chars)
    by_source = {}
    for index in selected:
        by_source.setdefault(sources[index] or "no-text", []).append(index)
    result = {
        "kind": "pdf",
        "page_count": r["count"],
        "pages_read": pdfread.ranges(selected),
        "page_sources": {k: pdfread.ranges(v) for k, v in by_source.items()},
        "text": text,
        "truncated": truncated,
    }
    try:
        details = docmeta.pdf_details(data)
    except Exception:
        details = {}
    if details.get("signed"):
        result["signed"] = True
        result["signers"] = details.get("signers") or []
    if details.get("form_fields"):
        result["form_fields"] = details["form_fields"]
    if r["ocr_upgrade"]:
        result["ocr_upgrade"] = r["ocr_upgrade"]
    if r["cloud_offer"]:
        result["cloud_ocr_needs_approval"] = r["cloud_offer"]
    if truncated:
        shown = full[:max_chars].count("--- page ")
        rest = selected[max(shown - 1, 0):]
        result["next"] = (f"Text cut at {max_chars} characters. Continue with "
                          f"pages='{pdfread.ranges(rest)}', or search with file_find.")
    if r["notes"]:
        result["notes"] = r["notes"]
    return result


def image_as_pdf(data):
    """A picture (a multi-page TIFF too) as a PDF of image pages, at the
    image's own resolution, so OCR goes through the same reader as a scan."""
    try:
        from PIL import Image, ImageSequence
    except ImportError as exc:
        raise FileError("reading an image needs Pillow", "pip install pillow") from exc
    with Image.open(io.BytesIO(data)) as image:
        dpi = image.info.get("dpi") or (IMAGE_DPI, IMAGE_DPI)
        try:
            resolution = float(dpi[0]) if float(dpi[0]) >= 72 else IMAGE_DPI
        except (TypeError, ValueError, IndexError):
            resolution = IMAGE_DPI
        frames = [frame.convert("RGB") for _, frame in zip(range(IMAGE_MAX_FRAMES), ImageSequence.Iterator(image))]
    buffer = io.BytesIO()
    frames[0].save(buffer, format="PDF", resolution=resolution, save_all=True, append_images=frames[1:])
    return buffer.getvalue()


def read_image(data, pages=None, ocr="auto", max_chars=TEXT_LIMIT):
    """OCR a picture: a scanned page, a photo of a form, a screenshot. The
    same order as a scanned PDF page (Windows OCR here, then the OCR server,
    the cloud only with approval); a multi-page TIFF reads page by page."""
    if ocr == "never":
        return {"kind": "image", "text": "", "truncated": False,
                "notes": ["An image has no text layer; read it with ocr='auto' or look at it."]}
    result = read_pdf(image_as_pdf(data), pages=pages, ocr=ocr, max_chars=max_chars)
    result["kind"] = "image"
    for key in ("signed", "signers", "form_fields"):
        result.pop(key, None)
    return result


RENDER_MAX_DPI = 600
RENDER_BOX_DPI = 300


def parse_box(box):
    """A region of the page as (left, top, right, bottom) fractions of its
    width and height, measured from the top-left corner: '0,0.5,1,1' is the
    lower half. None for the whole page."""
    if box in (None, "", []):
        return None
    if isinstance(box, str):
        parts = [p for p in re.split(r"[,;\s]+", box.strip()) if p]
    else:
        parts = list(box)
    try:
        values = [float(p) for p in parts]
    except (TypeError, ValueError) as exc:
        raise FileError(f"box must be four numbers, got {box!r}",
                        "Pass left,top,right,bottom as fractions of the page, e.g. '0,0.5,1,1'.") from exc
    if len(values) != 4:
        raise FileError(f"box must be four numbers, got {box!r}",
                        "Pass left,top,right,bottom as fractions of the page, e.g. '0,0.5,1,1'.")
    left, top, right, bottom = values
    if not (0 <= left < right <= 1 and 0 <= top < bottom <= 1):
        raise FileError(f"box {box!r} is not inside the page",
                        "Each value is between 0 and 1, and left < right, top < bottom.")
    return left, top, right, bottom


def _crop_png(png, box):
    from PIL import Image
    image = Image.open(io.BytesIO(png))
    width, height = image.size
    left, top, right, bottom = box
    region = image.crop((round(left * width), round(top * height), round(right * width), round(bottom * height)))
    buffer = io.BytesIO()
    region.save(buffer, format="PNG")
    return buffer.getvalue()


def render_pdf_pages(data, pages="1", dpi=150, stem="document", box=None):
    """PNG files of PDF pages. With box, only that region of each page, at a
    higher resolution unless dpi says otherwise: a close look at a digit."""
    region = parse_box(box)
    dpi = int(dpi or (RENDER_BOX_DPI if region else 150))
    dpi = max(36, min(dpi, RENDER_MAX_DPI))
    doc = pdfread.open_pdf(data)
    if doc is None:
        raise FileError("rendering needs pypdfium2", "pip install pypdfium2")
    try:
        selected = pdfread.parse_pages(pages, len(doc))[:20]
        folder = _cache_dir("render")
        out = []
        for index in selected:
            png = pdfread.render_png(doc, index, dpi)
            suffix = ""
            if region:
                png = _crop_png(png, region)
                suffix = "-box" + "-".join(f"{int(v * 100):02d}" for v in region)
            target = folder / f"{_safe_name(stem)}-p{index + 1}{suffix}.png"
            target.write_bytes(png)
            row = {"page": index + 1, "png": str(target)}
            if region:
                row["box"] = list(region)
            out.append(row)
        return {"page_count": len(doc), "dpi": dpi, "images": out}
    finally:
        doc.close()


# ------------------------------------------------------------------- e-mail ---

def _eml_message(data):
    return BytesParser(policy=policy.default).parsebytes(data)


def _eml_attachments(message):
    return list(message.iter_attachments())


def _eml_summary(message):
    def addresses(field):
        value = message.get(field)
        if not value:
            return []
        try:
            return [str(a) for a in value.addresses]
        except Exception:
            return [str(value)]

    body, source = "", ""
    part = message.get_body(preferencelist=("plain", "html"))
    if part is not None:
        try:
            content = part.get_content()
        except Exception:
            content = (part.get_payload(decode=True) or b"").decode("utf-8", "replace")
        source = "html" if part.get_content_type() == "text/html" else "plain"
        body = content if isinstance(content, str) else str(content)
    when = ""
    try:
        date = message["date"].datetime if message["date"] is not None else None
        if date is not None:
            when = date.astimezone().isoformat(timespec="minutes")
    except Exception:
        when = str(message.get("date") or "")
    attachments = []
    for index, part in enumerate(_eml_attachments(message), 1):
        is_message = part.get_content_type() == "message/rfc822"
        name = part.get_filename()
        if not name and is_message:
            inner = part.get_payload(0) if part.is_multipart() else None
            name = (inner.get("subject") if inner is not None else None) or "message"
            name = f"{name}.eml"
        payload = None if is_message else part.get_payload(decode=True)
        attachments.append({
            "index": index,
            "name": name or f"attachment-{index}",
            "mime": part.get_content_type(),
            "size": len(payload) if payload is not None else None,
            "embedded_email": is_message,
            "inline": part.get_content_maintype() == "image" and (
                part.get_content_disposition() == "inline" or bool(part.get("Content-ID"))),
        })
    return {
        "subject": str(message.get("subject") or ""),
        "from": ", ".join(addresses("from")),
        "to": addresses("to"),
        "cc": addresses("cc"),
        "bcc": addresses("bcc"),
        "date": when,
        "body": body,
        "body_source": source,
        "attachments": attachments,
    }


def _eml_attachment(message, index):
    parts = _eml_attachments(message)
    if not 1 <= index <= len(parts):
        raise FileError(f"there is no attachment {index}; this message has {len(parts)}")
    part = parts[index - 1]
    if part.get_content_type() == "message/rfc822":
        inner = part.get_payload(0) if part.is_multipart() else _eml_message(part.get_payload(decode=True) or b"")
        name = part.get_filename() or f"{inner.get('subject') or 'message'}.eml"
        return name, inner.as_bytes(policy=policy.default)
    return part.get_filename() or f"attachment-{index}", part.get_payload(decode=True) or b""


def _email_result(summary, fmt, max_chars):
    body = summary.pop("body")
    if summary.get("body_source") == "html":
        body = html_to_text(body)
    text, truncated = _clip(body.strip(), max_chars)
    attachments = summary["attachments"]
    result = {"kind": "email", "format": fmt, **summary, "text": text, "truncated": truncated}
    item_class = summary.get("message_class") or ""
    if item_class and not item_class.upper().startswith(("IPM.NOTE", "REPORT.", "IPM.SCHEDULE")):
        result["notes"] = [f"This is an Outlook {item_class} item, not a message; only its "
                           "subject, body and attachments are read here. Use office-live for "
                           "the full item."]
    if attachments:
        hidden = sum(1 for a in attachments if a.get("inline"))
        result["next"] = (
            "Read an attachment in place with attachment='<index>'; save them to disk with "
            "file_attachments (save_attachments in Python)."
            + (f" {hidden} inline item(s) are usually signature images." if hidden else ""))
    return result


def _msg_text_dump(msg):
    s = msg.summary()
    body = s["body"] if s["body_source"] != "html" else html_to_text(s["body"])
    lines = [f"Subject: {s['subject']}", f"From: {s['from']}", f"To: {'; '.join(s['to'])}"]
    if s["cc"]:
        lines.append(f"Cc: {'; '.join(s['cc'])}")
    lines += [f"Date: {s['date']}", "", body]
    return "\n".join(lines)


def _attachment_of(data, kind, index):
    """(name, bytes, kind) of one attachment; an attached .msg e-mail comes back
    as its Msg object under kind 'msg-object'."""
    if kind == "msg":
        msg = msgfile.open_msg(data)
        name, payload = msg.attachment(index)
        if isinstance(payload, msgfile.Msg):
            return name, payload, "msg-object"
        return name, payload, detect_kind(payload, name)
    name, payload = _eml_attachment(_eml_message(data), index)
    return name, payload, detect_kind(payload, name)


def _pick_attachment(data, kind, which):
    """Resolve '2', '2/1' or a file name to (name, payload, kind)."""
    steps = [s for s in str(which).split("/") if s.strip()] if re.fullmatch(r"[\d/ ]+", str(which)) else [which]
    name = ""
    for step in steps:
        if kind == "zip":                       # a member of an archive, read in place
            name, data = archive.read_member(data, step)
            kind = detect_kind(data, name)
            continue
        if kind == "msg-object":
            msg = data
        elif kind not in ("msg", "eml"):
            raise FileError(f"{name or 'this file'} is not an e-mail or a zip, so it has no attachments")
        else:
            msg = None
        listing = (msg.summary() if msg is not None else
                   (msgfile.open_msg(data).summary() if kind == "msg" else _eml_summary(_eml_message(data))))["attachments"]
        if str(step).strip().isdigit():
            index = int(step)
        else:
            matches = [a["index"] for a in listing if a["name"].lower() == str(step).lower()]
            if not matches:
                raise FileError(f"no attachment named {step!r}",
                                "Available: " + ", ".join(f"{a['index']}={a['name']}" for a in listing))
            index = matches[0]
        if msg is not None:
            name, payload = msg.attachment(index)
            data, kind = (payload, "msg-object") if isinstance(payload, msgfile.Msg) else (payload, detect_kind(payload, name))
        else:
            name, data, kind = _attachment_of(data, kind, index)
    return name, data, kind


# ------------------------------------------------------------ office / text ---

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _docx_paragraph(p):
    out = []
    for node in p.iter():
        if node.tag == _W + "t":
            out.append(node.text or "")
        elif node.tag == _W + "tab":
            out.append("\t")
        elif node.tag in (_W + "br", _W + "cr"):
            out.append("\n")
    return "".join(out)


def _docx_block(element, lines):
    for child in element:
        if child.tag == _W + "p":
            lines.append(_docx_paragraph(child))
        elif child.tag == _W + "tbl":
            for row in child.iter(_W + "tr"):
                cells = []
                for cell in row.findall(_W + "tc"):
                    cells.append(" ".join(_docx_paragraph(p) for p in cell.iter(_W + "p")).strip())
                lines.append("\t".join(cells))
            lines.append("")
        elif child.tag in (_W + "sdt", _W + "sdtContent", _W + "customXml", _W + "smartTag"):
            _docx_block(child, lines)


def read_docx(data, max_chars=TEXT_LIMIT):
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(_W + "body")
    lines = []
    _docx_block(body if body is not None else root, lines)
    text = re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()
    text, truncated = _clip(text, max_chars)
    result = {"kind": "docx", "text": text, "truncated": truncated}
    try:
        details = docmeta.docx_details(data)
    except Exception:
        details = {}
    for key in ("comments", "revisions", "checkboxes"):
        if details.get(key):
            result[key] = details[key]
    result["notes"] = ["text is the document as it reads now: tracked deletions left out, insertions "
                       "in. revisions lists both with author and date. Headers and footers are not "
                       "read. For a document open on screen use office-live."]
    return result


def read_pptx(data, max_chars=TEXT_LIMIT):
    rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    p_ns = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    r_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        names = set(zf.namelist())
        order = []
        try:
            rels = ET.fromstring(zf.read("ppt/_rels/presentation.xml.rels"))
            target = {r.get("Id"): r.get("Target") for r in rels.iter(rel_ns + "Relationship")}
            pres = ET.fromstring(zf.read("ppt/presentation.xml"))
            for sld in pres.iter(p_ns + "sldId"):
                t = target.get(sld.get(r_ns + "id"), "")
                order.append("ppt/" + t.lstrip("/").removeprefix("ppt/"))
        except KeyError:
            pass
        if not order:
            order = sorted((n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                           key=lambda n: int(re.findall(r"\d+", n)[-1]))
        parts = []
        for number, name in enumerate(order, 1):
            if name not in names:
                continue
            root = ET.fromstring(zf.read(name))
            paragraphs = []
            for para in root.iter(_A + "p"):
                line = "".join(t.text or "" for t in para.iter(_A + "t")).strip()
                if line:
                    paragraphs.append(line)
            parts.append(f"--- slide {number} ---\n" + "\n".join(paragraphs))
    text, truncated = _clip("\n\n".join(parts), max_chars)
    return {"kind": "pptx", "slide_count": len(order), "text": text, "truncated": truncated}


def _ole_has_workbook(data):
    """True when an OLE file holds an Excel 97-2003 workbook stream."""
    try:
        import olefile
    except ImportError:
        return False
    try:
        ole = olefile.OleFileIO(data)
    except Exception:
        return False
    try:
        return ole.exists("Workbook") or ole.exists("Book")
    finally:
        ole.close()


def _xls_book(data):
    try:
        import xlrd
    except ImportError as exc:
        raise FileError("reading .xls needs xlrd", "pip install xlrd") from exc
    try:
        return xlrd, xlrd.open_workbook(file_contents=data, on_demand=True)
    except Exception as exc:     # xlrd raises XLRDError, CompDocError and plain struct errors
        if "encrypt" in str(exc).lower():
            raise FileError("the .xls is password-protected", "Ask the user for an unprotected copy.") from exc
        raise FileError(f"the .xls could not be read ({type(exc).__name__}: {exc})",
                        "Open it with office-live: doc = open_doc(path), then read it there.") from exc


def _xls_cell(xlrd, book, cell):
    if cell.ctype == xlrd.XL_CELL_EMPTY or cell.ctype == xlrd.XL_CELL_BLANK:
        return ""
    if cell.ctype == xlrd.XL_CELL_DATE:
        try:
            value = xlrd.xldate_as_datetime(cell.value, book.datemode)
        except Exception:
            return str(cell.value)
        return value.date().isoformat() if value.time() == value.min.time() else value.isoformat(sep=" ")
    if cell.ctype == xlrd.XL_CELL_NUMBER and float(cell.value).is_integer():
        return str(int(cell.value))
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return "TRUE" if cell.value else "FALSE"
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return xlrd.error_text_from_code.get(cell.value, "#ERROR")
    return str(cell.value).replace("\n", " ").strip()


def _xls_rows(data, max_rows=XLSX_MAX_ROWS, max_cols=XLSX_MAX_COLS, max_sheets=XLSX_MAX_SHEETS):
    """(sheet names, [(sheet, row number, cells)], clipped) of an Excel 97-2003 workbook."""
    xlrd, book = _xls_book(data)
    rows, clipped = [], False
    try:
        names = list(book.sheet_names())
        for index, name in enumerate(names[:max_sheets]):
            sheet = book.sheet_by_index(index)
            if sheet.nrows > max_rows or sheet.ncols > max_cols:
                clipped = True
            for r in range(min(sheet.nrows, max_rows)):
                cells = [_xls_cell(xlrd, book, sheet.cell(r, c)) for c in range(min(sheet.ncols, max_cols))]
                if any(cells):
                    rows.append((name, r + 1, cells))
            book.unload_sheet(index)
    finally:
        book.release_resources()
    return names, rows, clipped


def read_xls(data, max_chars=TEXT_LIMIT):
    """An Excel 97-2003 .xls workbook, read with xlrd (BSD), in the same shape
    as read_xlsx. Values only: formulas show their last computed result."""
    names, rows, clipped = _xls_rows(data)
    parts = []
    for name in names[:XLSX_MAX_SHEETS]:
        lines = ["\t".join(cells).rstrip() for sheet, _, cells in rows if sheet == name]
        parts.append(f"--- sheet {name} ---\n" + "\n".join(lines))
    notes = []
    if len(names) > XLSX_MAX_SHEETS:
        notes.append(f"{len(names) - XLSX_MAX_SHEETS} further sheet(s) not shown.")
    if clipped:
        notes.append(f"Each sheet is clipped to {XLSX_MAX_ROWS} rows and {XLSX_MAX_COLS} columns; "
                     "use office-live for the rest.")
    text, truncated = _clip("\n\n".join(parts), max_chars)
    result = {"kind": "xls", "sheets": names, "text": text, "truncated": truncated}
    if notes:
        result["notes"] = notes
    return result


def read_xlsx(data, max_chars=TEXT_LIMIT):
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise FileError("reading .xlsx needs openpyxl", "pip install openpyxl") from exc
    notes = []
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        names = list(wb.sheetnames)
        parts = []
        clipped = False
        for name in names[:XLSX_MAX_SHEETS]:
            lines = []
            for index, row in enumerate(wb[name].iter_rows(values_only=True)):
                if index >= XLSX_MAX_ROWS:
                    clipped = True
                    break
                if len(row) > XLSX_MAX_COLS:
                    clipped = True
                cells = ["" if c is None else str(c).replace("\n", " ").strip() for c in row[:XLSX_MAX_COLS]]
                if any(cells):
                    lines.append("\t".join(cells).rstrip())
            parts.append(f"--- sheet {name} ---\n" + "\n".join(lines))
    finally:
        wb.close()
    if len(names) > XLSX_MAX_SHEETS:
        notes.append(f"{len(names) - XLSX_MAX_SHEETS} further sheet(s) not shown.")
    if clipped:
        notes.append(f"Each sheet is clipped to {XLSX_MAX_ROWS} rows and {XLSX_MAX_COLS} columns; "
                     "use openpyxl or office-live for the rest.")
    text, truncated = _clip("\n\n".join(parts), max_chars)
    result = {"kind": "xlsx", "sheets": names, "text": text, "truncated": truncated}
    if notes:
        result["notes"] = notes
    return result


# ---------------------------------------------------------------- dispatch ---

def _read_bytes(data, name, kind, pages, ocr, max_chars):
    if kind == "pdf":
        return read_pdf(data, pages=pages, ocr=ocr, max_chars=max_chars)
    if kind == "image":
        return read_image(data, pages=pages, ocr=ocr, max_chars=max_chars)
    if kind == "msg":
        return _email_result(msgfile.open_msg(data).summary(), "msg", max_chars)
    if kind == "msg-object":
        return _email_result(data.summary(), "msg", max_chars)
    if kind == "eml":
        return _email_result(_eml_summary(_eml_message(data)), "eml", max_chars)
    if kind == "docx":
        return read_docx(data, max_chars)
    if kind == "xlsx":
        return read_xlsx(data, max_chars)
    if kind == "xls":
        return read_xls(data, max_chars)
    if kind == "pptx":
        return read_pptx(data, max_chars)
    if kind == "html":
        text, truncated = _clip(html_to_text(_decode_text(data)), max_chars)
        return {"kind": "html", "text": text, "truncated": truncated}
    if kind == "text":
        raw = _decode_text(data)
        if Path(name).suffix.lower() == ".rtf" and raw.lstrip().startswith("{\\rtf"):
            raw = msgfile.rtf_to_text(data)
        text, truncated = _clip(raw, max_chars)
        return {"kind": "text", "text": text, "truncated": truncated}
    if kind == "legacy-office":
        raise FileError(
            f"{name} is a legacy Office file (.doc or .ppt)",
            "Open it with office-live: doc = open_doc(path), then read it there.")
    if kind == "zip":
        members, total = archive.list_members(data)
        return {"kind": "zip", "member_count": total, "members": members[:200],
                "next": "Read a member in place with attachment='<index or name>', search inside with "
                        "file_find, or extract with file_unpack."}
    raise FileError(f"{name} is not a format this reader knows ({kind})",
                    "PDF, images, .msg, .eml, .docx, .xlsx, .xls, .pptx, .zip, HTML and text are supported.")


def extract_labels(text, labels):
    """For each label, the value written after it: 'Tulajdonos: Kiss Anna' gives
    {'tulajdonos': 'Kiss Anna'}. Case and accents are ignored; the value runs to
    the end of the line. A label not found is left out."""
    found = {}
    folded, index = _fold(text or "")
    for label in labels or []:
        key = _fold(str(label).strip())[0]
        if not key:
            continue
        pattern = re.compile(r"(?<!\w)" + r"\s+".join(map(re.escape, key.split())) + r"\s*[:=\-–]?[ \t]*([^\n]{1,200})")
        match = pattern.search(folded)
        if match and match.group(1).strip():
            start, end = index[match.start(1)], index[match.end(1) - 1] + 1
            found[str(label)] = text[start:end].strip()
    return found


def read_file(path, pages=None, ocr="auto", attachment=None, max_chars=TEXT_LIMIT, labels=None):
    """Read one file, or one attachment or archive member inside it, as text.
    labels: names whose values to pick out of the text (see extract_labels)."""
    p = _resolve(path)
    data = p.read_bytes()
    kind = detect_kind(data, p.name)
    name = p.name
    if attachment not in (None, ""):
        name, data, kind = _pick_attachment(data, kind, attachment)
    result = {"path": str(p)}
    if attachment not in (None, ""):
        result["attachment"] = {"selector": str(attachment), "name": name}
    result.update(_read_bytes(data, name, kind, pages, ocr, int(max_chars or TEXT_LIMIT)))
    if labels:
        result["labels"] = extract_labels(result.get("text", ""), _as_list(labels))
    if kind not in ("msg-object",) and attachment in (None, ""):
        result["size"] = len(data)
    return result


def _as_list(value):
    if value in (None, ""):
        return []
    if isinstance(value, (list, tuple)):
        return [str(v) for v in value if str(v).strip()]
    return [part.strip() for part in str(value).split(",") if part.strip()]


# ------------------------------------------------------------------- search ---

SEARCHABLE_EXTENSIONS = {".pdf", ".msg", ".eml", ".docx", ".xlsx", ".xls", ".pptx", ".zip"} | HTML_EXTENSIONS | TEXT_EXTENSIONS
FIND_MAX_FILES = 300
FIND_MAX_HITS = 50
FIND_CONTEXT = 80
FIND_ATTACHMENT_DEPTH = 2
FIND_FOLDER_MAX_BYTES = 20 * 1024 * 1024   # bigger files are searched one at a time
FIND_TIME_BUDGET = 45                      # seconds a folder search may take
FIND_XLSX_MAX_ROWS = 200_000


def _fold(text):
    """Lower-case and strip accents, keeping a map back to the original
    positions: 'Összeférhetetlenség' and 'osszeferhetetlenseg' match."""
    import unicodedata
    chars, index = [], []
    for i, ch in enumerate(text):
        for c in unicodedata.normalize("NFKD", ch):
            if unicodedata.combining(c):
                continue
            for lc in c.lower():
                chars.append(lc)
                index.append(i)
    return "".join(chars), index


def _strip_accents(text):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))


def _matcher(query, regex=False, exact=False):
    """A function text -> [(start, end)] over the ORIGINAL text."""
    query = (query or "").strip()
    if not query:
        raise FileError("the search query is empty")
    if regex:
        try:
            pattern = re.compile(query if exact else _strip_accents(query), 0 if exact else re.I)
        except re.error as exc:
            raise FileError(f"invalid regular expression: {exc}",
                            "Pass regex=false to search for the words as written.") from exc
    else:
        # Words in order, any whitespace between them: PDF text breaks lines
        # in the middle of a phrase.
        words = (query if exact else _fold(query)[0]).split()
        pattern = re.compile(r"\s+".join(re.escape(w) for w in words))

    def find(text):
        if exact:
            return [(m.start(), m.end()) for m in pattern.finditer(text) if m.end() > m.start()]
        folded, index = _fold(text)
        out = []
        for m in pattern.finditer(folded):
            if m.end() > m.start():
                out.append((index[m.start()], index[m.end() - 1] + 1))
        return out
    return find


def _snippet(text, start, end, context=FIND_CONTEXT):
    a, b = max(0, start - context), min(len(text), end + context)
    before = re.sub(r"\s+", " ", text[a:start]).lstrip()
    match = re.sub(r"\s+", " ", text[start:end])
    after = re.sub(r"\s+", " ", text[end:b]).rstrip()
    return ("…" if a > 0 else "") + before + "«" + match + "»" + after + ("…" if b < len(text) else "")


def _search_units(data, name, kind, pages, ocr, meta, prefix="", depth=0):
    """Yield (where, text, line_numbers) pieces of one file to search."""
    if kind == "image":
        if ocr == "never":
            meta["unsearched"].append(f"{prefix or name + ': '}image (no OCR asked)")
            return
        data, kind = image_as_pdf(data), "pdf"
    if kind == "pdf":
        r = pdfread.read_pages(data, pages=pages, ocr=ocr, tables=False)
        meta["notes"] += [prefix + n for n in r["notes"]]
        empty = [i for i in r["selected"] if not r["sources"][i]]
        if empty:
            meta["unsearched"].append(f"{prefix}page(s) {pdfread.ranges(empty)}")
        if r["cloud_offer"] and not prefix:
            meta["cloud_offer"] = r["cloud_offer"]
        if r["ocr_upgrade"] and not prefix:
            meta["ocr_upgrade"] = r["ocr_upgrade"]
        for i in r["selected"]:
            if r["sources"][i]:
                yield f"{prefix}page {i + 1}", r["texts"][i], False
        return
    if kind in ("msg", "msg-object", "eml"):
        if kind == "msg":
            summary = msgfile.open_msg(data).summary()
        elif kind == "msg-object":
            summary = data.summary()
        else:
            summary = _eml_summary(_eml_message(data))
        body = summary["body"] if summary["body_source"] != "html" else html_to_text(summary["body"])
        yield f"{prefix}subject", summary["subject"], False
        yield f"{prefix}body", body, True
        if depth >= FIND_ATTACHMENT_DEPTH:
            return
        for att in summary["attachments"]:
            if att.get("inline") or att.get("ole_object"):
                continue
            label = f"{prefix}attachment {att['index']} ({att['name']}) › "
            try:
                if kind == "msg-object":
                    sub_name, payload = data.attachment(att["index"])
                    sub_kind = "msg-object" if isinstance(payload, msgfile.Msg) else detect_kind(payload, sub_name)
                else:
                    sub_name, payload, sub_kind = _attachment_of(data, "msg" if kind == "msg" else "eml", att["index"])
                yield from _search_units(payload, sub_name, sub_kind, None, ocr, meta, label, depth + 1)
            except (FileError, msgfile.MsgError, pdfread.PdfError, archive.ArchiveError) as exc:
                meta["notes"].append(f"{label}not searched: {exc}")
        return
    if kind == "xlsx":
        try:
            from openpyxl import load_workbook
        except ImportError as exc:
            raise FileError("searching .xlsx needs openpyxl", "pip install openpyxl") from exc
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")  # openpyxl warns about styles it drops
            wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        rows = 0
        try:
            for sheet in wb.sheetnames:
                for number, row in enumerate(wb[sheet].iter_rows(values_only=True), 1):
                    rows += 1
                    if rows > FIND_XLSX_MAX_ROWS:
                        meta["notes"].append(f"{prefix}searched the first {FIND_XLSX_MAX_ROWS} rows only.")
                        return
                    cells = ["" if c is None else str(c) for c in row]
                    if any(cells):
                        yield f"{prefix}sheet {sheet}, row {number}", "\t".join(cells).rstrip(), False
        finally:
            wb.close()
        return
    if kind == "xls":
        _, rows, clipped = _xls_rows(data, max_rows=FIND_XLSX_MAX_ROWS, max_cols=10 ** 4, max_sheets=10 ** 3)
        for sheet, number, cells in rows:
            yield f"{prefix}sheet {sheet}, row {number}", "\t".join(cells).rstrip(), False
        if clipped:
            meta["notes"].append(f"{prefix}searched the first {FIND_XLSX_MAX_ROWS} rows of each sheet only.")
        return
    if kind == "pptx":
        text = read_pptx(data, max_chars=10 ** 9)["text"]
        for part in re.split(r"(?m)^--- slide (?=\d+ ---$)", text):
            head, _, body = part.partition(" ---\n")
            if head.isdigit():
                yield f"{prefix}slide {head}", body, False
        return
    if kind in ("docx", "html", "text"):
        result = _read_bytes(data, name, kind, None, "never", 10 ** 9)
        yield f"{prefix}{'document' if kind == 'docx' else 'text'}", result["text"], True
        return
    if kind == "zip":
        if depth >= FIND_ATTACHMENT_DEPTH:
            return
        members, _ = archive.list_members(data)
        for member in members:
            if member.get("unreadable") or member["suffix"] not in SEARCHABLE_EXTENSIONS | {".zip"}:
                continue
            label = f"{prefix}{member['name']} › "
            try:
                sub_name, payload = archive.read_member(data, str(member["index"]))
                yield from _search_units(payload, sub_name, detect_kind(payload, sub_name), None,
                                         "never" if ocr == "auto" else ocr, meta, label, depth + 1)
            except (FileError, msgfile.MsgError, pdfread.PdfError, archive.ArchiveError) as exc:
                meta["notes"].append(f"{label}not searched: {exc}")
        return
    meta["notes"].append(f"{prefix or name}: {kind} files are not searched.")


def _search_file(data, name, kind, find, pages, ocr, max_hits, meta):
    hits, total = [], 0
    for where, text, with_lines in _search_units(data, name, kind, pages, ocr, meta):
        for start, end in find(text or ""):
            total += 1
            if len(hits) < max_hits:
                location = where
                if with_lines:
                    location += f", line {text.count(chr(10), 0, start) + 1}"
                hits.append({"where": location, "snippet": _snippet(text, start, end)})
    return hits, total


def find_in_files(path, query, regex=False, exact=False, pages=None, ocr="auto",
                  attachment=None, recursive=False, max_hits=FIND_MAX_HITS):
    """Search inside one file, or every readable file in a folder.

    One file: every match with where it is (page, slide, sheet and row, line,
    or attachment) and a snippet with the match in «». A scanned PDF page is
    read through your OCR server like file_read does (ocr="never" skips it).
    A folder: which files match and how often, text layers only (no OCR).
    """
    max_hits = max(1, min(int(max_hits or FIND_MAX_HITS), 500))
    find = _matcher(query, regex=regex, exact=exact)
    text = os.path.expandvars(os.path.expanduser(str(path or "").strip().strip('"').strip("'")))
    target = Path(text) if text else None
    if target is not None and not target.is_absolute():
        target = Path.cwd() / target

    if target is not None and target.is_dir():
        import time
        pattern = "**/*" if recursive else "*"
        # Newest first: when the time budget runs out, what is left is old.
        candidates = sorted((p for p in target.glob(pattern)
                             if p.is_file() and p.suffix.lower() in SEARCHABLE_EXTENSIONS
                             and not p.name.startswith("~$")),
                            key=lambda p: p.stat().st_mtime, reverse=True)
        skipped = max(0, len(candidates) - FIND_MAX_FILES)
        files_with_hits, scanned_only, failed, too_big, total = [], [], [], [], 0
        started, searched, out_of_time = time.monotonic(), 0, 0
        for p in candidates[:FIND_MAX_FILES]:
            if time.monotonic() - started > FIND_TIME_BUDGET:
                out_of_time += 1
                continue
            if p.stat().st_size > FIND_FOLDER_MAX_BYTES:
                too_big.append(str(p))
                continue
            searched += 1
            meta = {"notes": [], "unsearched": [], "cloud_offer": None, "ocr_upgrade": None}
            try:
                data = p.read_bytes()
                hits, count = _search_file(data, p.name, detect_kind(data, p.name), find,
                                           None, "never", 3, meta)
            except (FileError, msgfile.MsgError, pdfread.PdfError, archive.ArchiveError) as exc:
                failed.append({"path": str(p), "error": str(exc)})
                continue
            except Exception as exc:
                failed.append({"path": str(p), "error": f"{type(exc).__name__}: {exc}"})
                continue
            if count:
                total += count
                files_with_hits.append({"path": str(p), "hits": count, "first": hits})
            elif meta["unsearched"] and p.suffix.lower() == ".pdf":
                scanned_only.append(str(p))
        files_with_hits.sort(key=lambda f: -f["hits"])
        result = {"folder": str(target), "query": query, "files_searched": searched,
                  "files_with_hits": files_with_hits[:max_hits], "hit_count": total}
        notes = []
        if out_of_time:
            notes.append(f"Stopped after {FIND_TIME_BUDGET} s; the {out_of_time} oldest file(s) were not "
                         "searched. Narrow the folder, or search those files one at a time.")
        if too_big:
            result["too_big_for_folder_search"] = too_big[:50]
            notes.append(f"{len(too_big)} file(s) above {FIND_FOLDER_MAX_BYTES // 2**20} MB were skipped; "
                         "search each on its own with file_find on that file.")
        if len(files_with_hits) > max_hits:
            notes.append(f"{len(files_with_hits) - max_hits} more file(s) matched; raise max_hits to see them.")
        if skipped:
            notes.append(f"Only the first {FIND_MAX_FILES} files were searched; {skipped} more were not.")
        if scanned_only:
            result["scanned_pdfs_not_searched"] = scanned_only[:50]
            notes.append("Folder search reads text layers only. Search a scanned PDF one at a time "
                         "with file_find on that file: its pages then go through the OCR server.")
        if failed:
            result["not_searched"] = failed[:50]
        if notes:
            result["notes"] = notes
        return result

    p = _resolve(path)
    data = p.read_bytes()
    kind = detect_kind(data, p.name)
    name = p.name
    if attachment not in (None, ""):
        name, data, kind = _pick_attachment(data, kind, attachment)
    meta = {"notes": [], "unsearched": [], "cloud_offer": None, "ocr_upgrade": None}
    hits, total = _search_file(data, name, kind, find, pages, ocr, max_hits, meta)
    result = {"path": str(p), "query": query, "kind": kind, "hit_count": total, "hits": hits}
    if attachment not in (None, ""):
        result["attachment"] = {"selector": str(attachment), "name": name}
    if total > len(hits):
        result["next"] = f"{total - len(hits)} more match(es); raise max_hits or narrow with pages=."
    if meta["unsearched"]:
        result["not_searched"] = meta["unsearched"]
    if meta["ocr_upgrade"]:
        result["ocr_upgrade"] = meta["ocr_upgrade"]
    if meta["cloud_offer"]:
        result["cloud_ocr_needs_approval"] = meta["cloud_offer"]
    if meta["notes"]:
        result["notes"] = meta["notes"]
    return result


# ----------------------------------------------------------------- collect ---

COLLECT_PREVIEW_ROWS = 100
COLLECT_EXCERPT = 300


def _collect_items(target: Path, recursive: bool, notes: list):
    """(relative name, bytes) for every readable document in a folder or a zip."""
    import time
    started = time.monotonic()
    if target.is_dir():
        pattern = "**/*" if recursive else "*"
        paths = sorted((p for p in target.glob(pattern) if p.is_file() and p.suffix.lower() in SEARCHABLE_EXTENSIONS
                        and p.suffix.lower() != ".zip" and not p.name.startswith("~$")),
                       key=lambda p: str(p).lower())
        for p in paths[:FIND_MAX_FILES]:
            if time.monotonic() - started > FIND_TIME_BUDGET:
                notes.append(f"Stopped after {FIND_TIME_BUDGET} s; narrow the folder for the rest.")
                return
            if p.stat().st_size > FIND_FOLDER_MAX_BYTES:
                notes.append(f"{p.name}: above {FIND_FOLDER_MAX_BYTES // 2**20} MB, read it on its own.")
                continue
            yield str(p.relative_to(target)), p.read_bytes()
        if len(paths) > FIND_MAX_FILES:
            notes.append(f"Only the first {FIND_MAX_FILES} files were collected.")
        return
    if target.stat().st_size > archive.MAX_ZIP_BYTES:
        raise archive.ArchiveError(f"the zip is larger than {archive.MAX_ZIP_BYTES // 2**20} MB",
                                   "Unpack it with file_unpack, then collect over the folder.")
    data = target.read_bytes()
    members, total = archive.list_members(data)
    if total > len(members):
        notes.append(f"Only the first {len(members)} of {total} archive members were collected.")
    for member in members:
        if member.get("unreadable"):
            notes.append(f"{member['name']}: {member['unreadable']}")
            continue
        if member["suffix"] not in SEARCHABLE_EXTENSIONS or member["suffix"] == ".zip":
            continue
        yield member["name"], archive.read_member(data, str(member["index"]))[1]


def _collect_row(name, data, labels, find, excerpt):
    kind = detect_kind(data, name)
    row = {"file": name, "kind": kind, "size": len(data)}
    text = ""
    if kind in ("msg", "eml"):
        s = msgfile.open_msg(data).summary() if kind == "msg" else _eml_summary(_eml_message(data))
        text = s["body"] if s["body_source"] != "html" else html_to_text(s["body"])
        row.update({"subject": s["subject"], "from": s["from"], "to": "; ".join(s["to"]),
                    "cc": "; ".join(s["cc"]), "date": s["date"],
                    "attachments": "; ".join(a["name"] for a in s["attachments"] if not a.get("inline"))})
    elif kind == "pdf":
        r = pdfread.read_pages(data, ocr="never", tables=False)
        text = "\n\n".join(r["texts"][i] for i in r["selected"] if r["sources"][i])
        empty = [i for i in r["selected"] if not r["sources"][i]]
        row["pages"] = r["count"]
        if empty:
            row["scanned_pages"] = pdfread.ranges(empty)
        details = docmeta.pdf_details(data)
        if details.get("signers"):
            row["signers"] = "; ".join(details["signers"])
        if details.get("form_fields"):
            row["form_fields"] = "; ".join(f"{k}={v}" for k, v in details["form_fields"].items())
    elif kind == "docx":
        text = read_docx(data, 10 ** 9)["text"]
        details = docmeta.docx_details(data)
        row["comments"] = len(details["comments"])
        row["revisions"] = len(details["revisions"])
        if details["checkboxes"]:
            row["checked"] = "; ".join(b["label"][:80] for b in details["checkboxes"] if b["checked"])
    elif kind in ("xlsx", "xls", "pptx", "html", "text"):
        text = _read_bytes(data, name, kind, None, "never", 10 ** 9).get("text", "")
    else:
        row["error"] = f"{kind} files are not read"
        return row
    first = next((line.strip() for line in text.splitlines()
                  if line.strip() and not line.startswith("--- ")), "")
    if kind not in ("msg", "eml"):
        row["title"] = first[:150]          # the first line of text
    if labels:
        for label, value in extract_labels(text, labels).items():
            row[f"label:{label}"] = value
    if find is not None:
        hits = find(text)
        row["hits"] = len(hits)
        if hits:
            row["first_hit"] = _snippet(text, *hits[0], context=60)
    if excerpt:
        row["excerpt"] = re.sub(r"\s+", " ", text).strip()[:excerpt]
    return row


def collect_files(path, labels=None, query=None, regex=False, recursive=False,
                  excerpt=COLLECT_EXCERPT, out=None):
    """One row per document in a folder or a .zip: what it is (sender, date,
    subject or first line, pages, signers, form fields, comments...), the
    values after the given labels, and how often the query occurs.

    Reads only; text layers only (a scanned page is marked, not OCR'd). With
    more rows than fit an answer, or with out=, the full table is also
    written as UTF-8 CSV, so a hundred letters do not flood the conversation.
    """
    text = os.path.expandvars(os.path.expanduser(str(path or "").strip().strip('"').strip("'")))
    target = Path(text) if text else None
    if target is not None and not target.is_absolute():
        target = Path.cwd() / target
    if target is None or not target.exists():
        raise FileError(f"no such folder or file: {target}", "Pass a folder or a .zip.")
    if target.is_file() and _head_kind(target) != "zip":
        raise FileError(f"{target.name} is one file", "Use file_read for one file; collect takes a folder or a .zip.")
    labels = _as_list(labels)
    find = _matcher(query, regex=regex) if query else None
    notes, rows = [], []
    for name, data in _collect_items(target, recursive, notes):
        try:
            rows.append(_collect_row(name, data, labels, find, int(excerpt or 0)))
        except Exception as exc:
            rows.append({"file": name, "error": f"{type(exc).__name__}: {exc}"})
    columns = []
    for row in rows:
        columns += [key for key in row if key not in columns]
    result = {"source": str(target), "count": len(rows), "columns": columns,
              "rows": rows[:COLLECT_PREVIEW_ROWS]}
    if out or len(rows) > COLLECT_PREVIEW_ROWS:
        import csv
        table = Path(os.path.expandvars(os.path.expanduser(out))) if out else \
            _cache_dir("collect") / f"{_safe_name(target.stem or 'collect')}.csv"
        table = _unique(table) if out and table.exists() else table
        table.parent.mkdir(parents=True, exist_ok=True)
        with table.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=columns, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
        result["csv"] = str(table)
        if len(rows) > COLLECT_PREVIEW_ROWS:
            notes.append(f"{len(rows) - COLLECT_PREVIEW_ROWS} more row(s) are in the CSV only.")
    if query:
        result["files_with_hits"] = sum(1 for r in rows if r.get("hits"))
    result["next"] = ("To give the user this table, write it into an Excel workbook open in front of "
                      "them with office-live (or open the CSV there), not into a closed file.")
    if notes:
        result["notes"] = notes
    return result


def unpack_file(path, dest=None):
    """Extract a .zip next to itself (or into dest), safely; never overwrites."""
    text = os.path.expandvars(os.path.expanduser(str(path or "").strip().strip('"').strip("'")))
    p = Path(text)
    if not p.is_absolute():
        p = Path.cwd() / p
    if not p.is_file():
        raise FileError(f"no such file: {p}")
    if _head_kind(p) != "zip":
        raise FileError(f"{p.name} is not a .zip archive")
    return archive.unpack(p, dest)


def _head_kind(p: Path):
    """detect_kind without reading the whole file. A truncated zip head would
    always look like a plain zip, so Office packages are told apart by their
    member names, read from the file's central directory."""
    with p.open("rb") as handle:
        head = handle.read(4096)
    if not head.startswith(b"PK\x03\x04"):
        return detect_kind(head, p.name)
    try:
        with zipfile.ZipFile(p) as zf:
            names = zf.namelist()
    except zipfile.BadZipFile:
        return "zip"
    if "word/document.xml" in names:
        return "docx"
    if any(n.startswith("xl/") for n in names):
        return "xlsx"
    if any(n.startswith("ppt/slides/") for n in names):
        return "pptx"
    return "zip"


def save_attachments(path, dest=None, which=None, include_inline=False):
    """Write an e-mail's attachments to a folder. Never overwrites a file."""
    p = _resolve(path)
    data = p.read_bytes()
    kind = detect_kind(data, p.name)
    if kind not in ("msg", "eml"):
        raise FileError(f"{p.name} is not an .msg or .eml e-mail")
    summary = msgfile.open_msg(data).summary() if kind == "msg" else _eml_summary(_eml_message(data))
    listing = summary["attachments"]
    if which not in (None, "", []):
        wanted = which if isinstance(which, (list, tuple)) else str(which).split(",")
        chosen = []
        for item in wanted:
            item = str(item).strip()
            match = [a for a in listing if str(a["index"]) == item or a["name"].lower() == item.lower()]
            if not match:
                raise FileError(f"no attachment {item!r}",
                                "Available: " + ", ".join(f"{a['index']}={a['name']}" for a in listing))
            chosen += match
    else:
        chosen = [a for a in listing if include_inline or not a["inline"]]
    folder = Path(os.path.expandvars(os.path.expanduser(dest))) if dest else _cache_dir("attachments", _safe_name(p.stem))
    folder.mkdir(parents=True, exist_ok=True)
    saved = []
    unsaved = []
    for item in chosen:
        if item.get("ole_object"):
            unsaved.append(f"{item['index']} ({item['name']}): embedded OLE object, not a file")
            continue
        name, payload, sub_kind = _attachment_of(data, kind, item["index"])
        if sub_kind == "msg-object":
            target = _unique(folder / (_safe_name(Path(name).stem or "message") + ".txt"))
            target.write_text(_msg_text_dump(payload), encoding="utf-8")
            note = "attached Outlook item, saved as text"
        else:
            target = _unique(folder / _safe_name(name, f"attachment-{item['index']}"))
            target.write_bytes(payload)
            note = ""
        row = {"index": item["index"], "name": name, "saved_to": str(target),
               "size": target.stat().st_size}
        if note:
            row["note"] = note
        saved.append(row)
    skipped = len(listing) - len(chosen)
    result = {"path": str(p), "folder": str(folder), "saved": saved}
    notes = []
    if skipped and which in (None, "", []):
        notes.append(f"{skipped} inline item(s) skipped (signature images); "
                     "pass include_inline=True to save them too.")
    if unsaved:
        notes.append("Not saved: " + "; ".join(unsaved))
    if notes:
        result["notes"] = notes
    return result


def render_pages(path, pages="1", dpi=None, attachment=None, box=None):
    """PDF pages or a picture as PNG files, to look at with a vision tool.
    box crops each page to a region (fractions from the top-left corner)."""
    p = _resolve(path)
    data = p.read_bytes()
    kind = detect_kind(data, p.name)
    name = p.name
    if attachment not in (None, ""):
        name, data, kind = _pick_attachment(data, kind, attachment)
    if kind == "image":
        data, kind = image_as_pdf(data), "pdf"
    if kind != "pdf":
        raise FileError(f"{name} is not a PDF or a picture", "Only PDF pages and pictures can be rendered.")
    result = render_pdf_pages(data, pages=pages, dpi=dpi, stem=Path(name).stem, box=box)
    result["path"] = str(p)
    return result


def status():
    def has(module):
        import importlib.util
        return importlib.util.find_spec(module) is not None

    return {
        "readers": {
            "pdf_text_and_render": "pypdfium2" if has("pypdfium2") else None,
            "pdf_text_fallback": "pdfminer.six" if has("pdfminer") else None,
            "pdf_tables": "pdfplumber" if has("pdfplumber") else None,
            "pdf_form_fields": "pypdf" if has("pypdf") else None,
            "msg": "olefile" if has("olefile") else None,
            "xlsx": "openpyxl" if has("openpyxl") else None,
            "xls": "xlrd" if has("xlrd") else None,
            "docx_pptx_eml_html_text": "standard library",
        },
        **pdfread.describe_settings(),
        "cache": str(_cache_dir()),
    }


# --------------------------------------------------------------------- CLI ---

def _utf8_stdout():
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def emit(data):
    print(json.dumps(data, ensure_ascii=False, indent=1, default=str))


def run(fn, *args, **kwargs):
    """Call one of the functions above and return a JSON string, errors included."""
    try:
        return json.dumps({"ok": True, **fn(*args, **kwargs)}, ensure_ascii=False, default=str)
    except FileError as exc:
        return json.dumps({"ok": False, "error": str(exc), "hint": exc.hint}, ensure_ascii=False)
    except pdfread.PdfError as exc:
        return json.dumps({"ok": False, "error": str(exc), "hint": exc.hint}, ensure_ascii=False)
    except archive.ArchiveError as exc:
        return json.dumps({"ok": False, "error": str(exc), "hint": exc.hint}, ensure_ascii=False)
    except msgfile.MsgError as exc:
        return json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False)
    except Exception as exc:
        return json.dumps({"ok": False, "error": f"{type(exc).__name__}: {exc}"}, ensure_ascii=False)


def main(argv=None):
    _utf8_stdout()
    ap = argparse.ArgumentParser(description="Read PDF, e-mail and Office files as text.")
    sub = ap.add_subparsers(dest="command", required=True)
    r = sub.add_parser("read")
    r.add_argument("path")
    r.add_argument("--pages")
    r.add_argument("--ocr", default="auto", choices=["auto", "server", "cloud", "force", "never"])
    r.add_argument("--attachment")
    r.add_argument("--max-chars", type=int, default=TEXT_LIMIT)
    r.add_argument("--labels", help="comma-separated labels whose values to pick out")
    a = sub.add_parser("attachments")
    a.add_argument("path")
    a.add_argument("--dest")
    a.add_argument("--which")
    a.add_argument("--include-inline", action="store_true")
    g = sub.add_parser("render")
    g.add_argument("path")
    g.add_argument("--pages", default="1")
    g.add_argument("--dpi", type=int, default=None, help="default 150, or 300 with --box")
    g.add_argument("--box", help="left,top,right,bottom as fractions of the page, e.g. 0,0.5,1,1")
    f = sub.add_parser("find", help="search inside a file or a folder")
    f.add_argument("path")
    f.add_argument("query")
    f.add_argument("--regex", action="store_true")
    f.add_argument("--exact", action="store_true", help="case and accents must match")
    f.add_argument("--pages")
    f.add_argument("--ocr", default="auto", choices=["auto", "server", "cloud", "never"])
    f.add_argument("--attachment")
    f.add_argument("--recursive", action="store_true")
    f.add_argument("--max-hits", type=int, default=FIND_MAX_HITS)
    c = sub.add_parser("collect", help="one row per document in a folder or a .zip")
    c.add_argument("path")
    c.add_argument("--labels", help="comma-separated labels whose values to pick out")
    c.add_argument("--query")
    c.add_argument("--regex", action="store_true")
    c.add_argument("--recursive", action="store_true")
    c.add_argument("--excerpt", type=int, default=COLLECT_EXCERPT)
    c.add_argument("--out", help="also write the table as CSV here")
    u = sub.add_parser("unpack", help="extract a .zip safely")
    u.add_argument("path")
    u.add_argument("--dest")
    sub.add_parser("status")
    args = ap.parse_args(argv)
    if args.command == "collect":
        out = run(collect_files, args.path, labels=args.labels, query=args.query, regex=args.regex,
                  recursive=args.recursive, excerpt=args.excerpt, out=args.out)
    elif args.command == "unpack":
        out = run(unpack_file, args.path, dest=args.dest)
    elif args.command == "find":
        out = run(find_in_files, args.path, args.query, regex=args.regex, exact=args.exact,
                  pages=args.pages, ocr=args.ocr, attachment=args.attachment,
                  recursive=args.recursive, max_hits=args.max_hits)
    elif args.command == "read":
        out = run(read_file, args.path, pages=args.pages, ocr=args.ocr,
                  attachment=args.attachment, max_chars=args.max_chars, labels=args.labels)
    elif args.command == "attachments":
        out = run(save_attachments, args.path, dest=args.dest, which=args.which,
                  include_inline=args.include_inline)
    elif args.command == "render":
        out = run(render_pages, args.path, pages=args.pages, dpi=args.dpi, box=args.box)
    else:
        out = run(status)
    emit(json.loads(out))
    return 0 if json.loads(out).get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
