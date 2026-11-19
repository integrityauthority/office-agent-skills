"""What a document says about itself, beyond its running text.

Word (.docx): comments with the text they anchor to, tracked changes
(insertions and deletions, with author and date), and checkboxes -- content
control checkboxes and the Wingdings / Unicode box marks forms use instead.
PDF: whether it carries a signature and the signer names written in it (the
signature itself is not validated), and the values of its form fields.

Read with the standard library; PDF form fields with pypdf (BSD) when it is
installed. Nothing here changes a file.

    from docmeta import docx_details, pdf_details
"""
from __future__ import annotations

import io
import re
import zipfile
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
MAX_ITEMS = 100
TEXT_CHARS = 2000

_CHECKED_VALUES = {"1", "true", "on"}
_BOX_MARKS = ("☐", "☑", "☒", "✓", "✔")
_CHECKED_MARKS = {"☑", "☒", "✓", "✔"}
_WINGDINGS = {"00fe": "☑", "f0fe": "☑", "fe": "☑", "00a8": "☐", "f0a8": "☐", "a8": "☐",
              "006f": "☐", "f06f": "☐", "6f": "☐"}
_WINGDINGS_2 = {"0052": "☑", "52": "☑", "0082": "☑", "82": "☑", "0054": "☑", "54": "☑",
                "00a3": "☐", "a3": "☐"}


# ------------------------------------------------------------------ Word ---

def _text(node):
    """Visible and deleted run text under node, in document order."""
    return "".join(child.text or "" for child in node.iter() if child.tag in (W + "t", W + "delText"))


def _anchors(root):
    """comment id -> the text between its commentRangeStart and End."""
    active, found = [], {}
    for node in root.iter():
        if node.tag == W + "commentRangeStart":
            active.append(node.get(W + "id") or "")
        elif node.tag == W + "commentRangeEnd":
            cid = node.get(W + "id") or ""
            if cid in active:
                active.remove(cid)
        elif node.tag == W + "t" and node.text and active:
            for cid in active:
                found.setdefault(cid, []).append(node.text)
    return found


def _box_label(box, parents):
    node = box
    for _ in range(8):
        node = parents.get(node)
        if node is None:
            return ""
        if node.tag == W + "sdt":
            content = node.find(W + "sdtContent")
            if content is not None:
                return _text(content).strip()[:300]
        if node.tag == W + "p":
            return _text(node).strip()[:300]
    return ""


def _sym_marks(paragraph):
    marks = []
    for node in paragraph.iter(W + "sym"):
        font = (node.get(W + "font") or "").casefold()
        if "wingdings" not in font:
            continue
        table = _WINGDINGS_2 if ("wingdings 2" in font or "wingdings2" in font) else _WINGDINGS
        mark = table.get((node.get(W + "char") or "").strip().lower())
        if mark:
            marks.append(mark)
    return marks


def docx_details(data):
    """{"comments": [...], "revisions": [...], "checkboxes": [...]} of a .docx."""
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        root = ET.fromstring(package.read("word/document.xml"))
        comments_xml = package.read("word/comments.xml") if "word/comments.xml" in package.namelist() else None

    comments = []
    if comments_xml:
        anchors = _anchors(root)
        for node in ET.fromstring(comments_xml).iter(W + "comment"):
            cid = node.get(W + "id") or ""
            comments.append({"id": cid, "author": node.get(W + "author") or "",
                             "date": node.get(W + "date") or "",
                             "text": _text(node).strip()[:TEXT_CHARS],
                             "anchor": "".join(anchors.get(cid, []))[:TEXT_CHARS]})

    revisions = []
    for kind, label in (("ins", "insert"), ("del", "delete")):
        for node in root.iter(W + kind):
            text = _text(node)
            if text.strip():
                revisions.append({"kind": label, "author": node.get(W + "author") or "",
                                  "date": node.get(W + "date") or "", "text": text[:TEXT_CHARS]})

    parents = {child: parent for parent in root.iter() for child in parent}
    checkboxes = []
    for box in root.iter(W14 + "checkbox"):
        checked = box.find(W14 + "checked")
        value = ""
        if checked is not None:
            value = (checked.get(W14 + "val") or checked.get("val") or "").strip().lower()
        checkboxes.append({"checked": value in _CHECKED_VALUES, "label": _box_label(box, parents),
                           "kind": "content control"})
    for paragraph in root.iter(W + "p"):
        text = _text(paragraph)
        marks = [m for m in _BOX_MARKS if m in text] + _sym_marks(paragraph)
        if marks:
            checkboxes.append({"checked": any(m in _CHECKED_MARKS for m in marks),
                               "label": text.strip()[:300], "kind": "symbol"})

    return {"comments": comments[:MAX_ITEMS], "revisions": revisions[:MAX_ITEMS],
            "checkboxes": checkboxes[:MAX_ITEMS]}


# ------------------------------------------------------------------- PDF ---

def _pdf_string(raw):
    if raw.startswith(b"\xfe\xff"):
        return raw[2:].decode("utf-16-be", errors="replace")
    text = raw.replace(b"\\(", b"(").replace(b"\\)", b")").replace(b"\\\\", b"\\")
    for encoding in ("utf-8", "cp1250", "latin-1"):
        try:
            return text.decode(encoding)
        except UnicodeDecodeError:
            continue
    return text.decode("latin-1", errors="replace")


def _signed(raw):
    return any(marker in raw for marker in (b"/Type /Sig", b"/Type/Sig", b"/FT /Sig", b"/FT/Sig"))


def _signers(raw):
    if not _signed(raw):
        return []
    names = [_pdf_string(m.group(1)) for m in re.finditer(rb"/Name\s*\(([^)\r\n]{1,200})\)", raw)]
    for match in re.finditer(rb"/Name\s*<([0-9A-Fa-f\s]{4,400})>", raw):
        try:
            names.append(_pdf_string(bytes.fromhex(re.sub(rb"\s+", b"", match.group(1)).decode("ascii"))))
        except ValueError:
            continue
    unique = []
    for name in (n.strip() for n in names):
        if name and name not in unique:
            unique.append(name)
    return unique[:10]


def _form_fields(data):
    try:
        from pypdf import PdfReader
    except ImportError:
        return None
    try:
        fields = PdfReader(io.BytesIO(data)).get_fields() or {}
    except Exception:
        return {}
    out = {}
    for name, field in list(fields.items())[:MAX_ITEMS]:
        if not hasattr(field, "get") or field.get("/FT") == "/Sig":
            continue                    # a signature is reported as signers, not as a value
        value = field.get("/V")
        if isinstance(value, (list, tuple)):
            value = ", ".join(str(v).lstrip("/") for v in value)
        if value is None or hasattr(value, "keys") or str(value).strip() in ("", "/Off"):
            continue
        out[str(name)] = str(value).lstrip("/")
    return out


def pdf_details(data):
    """{"signed", "signers", "form_fields"} of a PDF. form_fields is None when
    pypdf is not installed, {} when there are none."""
    return {"signed": _signed(data), "signers": _signers(data), "form_fields": _form_fields(data)}
