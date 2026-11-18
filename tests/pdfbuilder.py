"""Build small PDFs for the tests: a page of text, or a "scanned" page that
has only a drawing and no text layer, which is what sends a page to OCR.

Text is set in Helvetica with WinAnsiEncoding, so Western European accents
(á, é, ö, ü) work; ő and ű are outside that encoding and are not used here.
"""
from __future__ import annotations


def _escape(line: str) -> bytes:
    raw = line.encode("cp1252")
    return raw.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)")


def _content(page) -> bytes:
    if page is None:  # a "scan": a filled rectangle, no text operators at all
        return b"0.2 g 72 300 450 300 re f"
    out = []
    y = 760
    for line in page.split("\n"):
        out.append(b"BT /F1 12 Tf 72 %d Td (" % y + _escape(line) + b") Tj ET")
        y -= 16
    return b"\n".join(out)


def build_pdf(pages) -> bytes:
    """pages: a list of str (text, '\\n' between lines) or None (no text)."""
    objects = []

    def add(body: bytes) -> int:
        objects.append(body)
        return len(objects)

    catalog = add(b"")  # filled in below
    page_tree = add(b"")
    font = add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    kids = []
    for page in pages:
        stream = _content(page)
        content = add(b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream")
        kids.append(add(b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 612 792] "
                        b"/Resources << /Font << /F1 %d 0 R >> >> /Contents %d 0 R >>"
                        % (page_tree, font, content)))
    objects[catalog - 1] = b"<< /Type /Catalog /Pages %d 0 R >>" % page_tree
    objects[page_tree - 1] = (b"<< /Type /Pages /Kids [" + b" ".join(b"%d 0 R" % k for k in kids)
                              + b"] /Count %d >>" % len(kids))

    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = []
    for number, body in enumerate(objects, 1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % number + body + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    out += b"".join(b"%010d 00000 n \n" % o for o in offsets)
    out += b"trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objects) + 1, catalog, xref)
    return bytes(out)
