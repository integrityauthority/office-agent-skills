"""files.py: reading, searching, attachments and rendering, end to end on
files generated here. No real documents, no network beyond the fake OCR.

The accented words (résumé, café) are deliberate: searches must ignore
accents and case, and attachment names must survive non-ASCII."""
import io
import json
import zipfile
from email.message import EmailMessage

import pytest

import files
from msgbuilder import build_msg
from pdfbuilder import build_pdf

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
P = ('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def make_docx():
    body = ('<w:p><w:r><w:t>First paragraph</w:t></w:r><w:r><w:tab/><w:t>with a tab</w:t></w:r></w:p>'
            '<w:p><w:ins><w:r><w:t>inserted text</w:t></w:r></w:ins>'
            '<w:del><w:r><w:delText>deleted text</w:delText></w:r></w:del></w:p>'
            '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>c1</w:t></w:r></w:p></w:tc>'
            '<w:tc><w:p><w:r><w:t>Résumé</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("word/document.xml", f"<w:document {W}><w:body>{body}</w:body></w:document>")
    return buf.getvalue()


def make_pptx():
    """Two slides whose file names run against the presentation order."""
    def slide(text):
        return f'<p:sld {P} {A}><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>'
    rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Target="slides/slide2.xml"/><Relationship Id="rId2" Target="slides/slide1.xml"/>'
            '</Relationships>')
    pres = f'<p:presentation {P}><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst></p:presentation>'
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("ppt/presentation.xml", pres)
        z.writestr("ppt/_rels/presentation.xml.rels", rels)
        z.writestr("ppt/slides/slide1.xml", slide("Second slide: café budget"))
        z.writestr("ppt/slides/slide2.xml", slide("First slide: title"))
    return buf.getvalue()


def make_xlsx():
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Name", "Amount"])
    ws.append(["Smith", 100])
    ws.append(["Résumé checklist", 1])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def make_eml():
    inner = EmailMessage()
    inner["Subject"] = "Internal note"
    inner["From"] = "a@example.com"
    inner["To"] = "b@example.com"
    inner.set_content("The internal note says: the deadline is Friday.")
    m = EmailMessage()
    m["Subject"] = "Outer subject"
    m["From"] = "Zoë Smith <zs@example.com>"
    m["To"] = "x@example.com, y@example.com"
    m["Date"] = "Tue, 17 Nov 2026 10:00:00 +0100"
    m.set_content("Hi,\nplease find the report attached.\n")
    m.add_alternative("<html><body><p>Hi,</p><p>please find the report attached.</p></body></html>", subtype="html")
    m.add_attachment(build_pdf(["Report, page 1", "The review found a gap in the résumé."]),
                     maintype="application", subtype="pdf", filename="café report.pdf")
    m.add_attachment(inner)
    m.add_attachment(make_docx(), maintype="application",
                     subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename="résumé.docx")
    return bytes(m)


@pytest.fixture
def folder(tmp_path):
    (tmp_path / "report.pdf").write_bytes(build_pdf(["Introduction", "The résumé review\nand its findings"]))
    (tmp_path / "scan.pdf").write_bytes(build_pdf([None]))
    (tmp_path / "letter.eml").write_bytes(make_eml())
    (tmp_path / "letter.msg").write_bytes(build_msg(
        subject="Meeting", body="Agenda: the résumé.", sender=("Anna", "anna@example.com"),
        to=[("Bea", "bea@example.com")],
        attachments=[{"name": "agenda.pdf", "data": build_pdf(["Agenda", "Item 3: the résumé"])},
                     {"name": "logo.png", "mime": "image/png", "content_id": "l@x", "data": b"\x89PNG0000"}]))
    (tmp_path / "doc.docx").write_bytes(make_docx())
    (tmp_path / "deck.pptx").write_bytes(make_pptx())
    (tmp_path / "table.xlsx").write_bytes(make_xlsx())
    (tmp_path / "notes.txt").write_text("first line\nsecond line résumé\n", encoding="utf-8")
    (tmp_path / "page.html").write_text("<html><head><style>x{}</style></head><body><h1>Title</h1><p>Paragraph</p></body></html>",
                                        encoding="utf-8")
    return tmp_path


def ok(result):
    data = json.loads(result) if isinstance(result, str) else result
    assert data.get("ok", True), data
    return data


# ------------------------------------------------------------------ read ---

def test_detect_kind(folder):
    kinds = {p.name: files.detect_kind(p.read_bytes(), p.name) for p in folder.iterdir()}
    assert kinds == {"report.pdf": "pdf", "scan.pdf": "pdf", "letter.eml": "eml", "letter.msg": "msg",
                     "doc.docx": "docx", "deck.pptx": "pptx", "table.xlsx": "xlsx",
                     "notes.txt": "text", "page.html": "html"}


def test_read_pdf_with_page_markers(folder):
    r = files.read_file(folder / "report.pdf")
    assert r["page_count"] == 2 and r["page_sources"] == {"text-layer": "1-2"}
    assert "--- page 2 ---\nThe résumé review" in r["text"]


def test_read_pdf_truncation_points_to_the_next_pages(folder):
    r = files.read_file(folder / "report.pdf", max_chars=20)
    assert r["truncated"] and "pages='" in r["next"]


def test_read_eml(folder):
    r = files.read_file(folder / "letter.eml")
    assert r["kind"] == "email" and r["format"] == "eml"
    assert r["subject"] == "Outer subject" and r["from"] == "Zoë Smith <zs@example.com>"
    assert r["to"] == ["x@example.com", "y@example.com"]
    assert r["text"].startswith("Hi,") and r["body_source"] == "plain"
    assert [a["name"] for a in r["attachments"]] == ["café report.pdf", "Internal note.eml", "résumé.docx"]


def test_read_attachment_in_place(folder):
    r = files.read_file(folder / "letter.eml", attachment="1")
    assert r["kind"] == "pdf" and r["attachment"]["name"] == "café report.pdf"
    r = files.read_file(folder / "letter.eml", attachment="RÉSUMÉ.DOCX")   # by name, any case
    assert r["kind"] == "docx"
    r = files.read_file(folder / "letter.eml", attachment="2")
    assert r["kind"] == "email" and r["subject"] == "Internal note"


def test_read_msg_and_its_attachment(folder):
    r = files.read_file(folder / "letter.msg")
    assert r["format"] == "msg" and r["subject"] == "Meeting"
    assert "inline item" in r["next"]
    assert files.read_file(folder / "letter.msg", attachment="1")["kind"] == "pdf"


def test_read_office_and_text(folder):
    docx = files.read_file(folder / "doc.docx")["text"]
    assert "inserted text" in docx and "deleted text" not in docx      # tracked deletions left out
    assert "c1\tRésumé" in docx
    pptx = files.read_file(folder / "deck.pptx")
    assert pptx["slide_count"] == 2
    assert pptx["text"].index("First slide") < pptx["text"].index("Second slide")   # presentation order
    xlsx = files.read_file(folder / "table.xlsx")
    assert xlsx["sheets"] == ["Data"] and "Smith\t100" in xlsx["text"]
    html = files.read_file(folder / "page.html")["text"]
    assert html == "Title\n\nParagraph"
    assert "second line" in files.read_file(folder / "notes.txt")["text"]


def test_read_errors(folder):
    err = ok_false = json.loads(files.run(files.read_file, folder / "missing.pdf"))
    assert ok_false["ok"] is False and "no such file" in err["error"] and err["hint"]
    assert "folder" in json.loads(files.run(files.read_file, folder))["error"]
    bad = json.loads(files.run(files.read_file, folder / "letter.eml", attachment="9"))
    assert bad["ok"] is False and "no attachment 9" in bad["error"]
    assert json.loads(files.run(files.read_file, folder / "doc.docx", attachment="1"))["ok"] is False


# ---------------------------------------------------------------- search ---

def test_find_ignores_case_and_accents(folder):
    r = files.find_in_files(folder / "report.pdf", "RESUME")
    assert r["hit_count"] == 1
    assert r["hits"][0]["where"] == "page 2"
    assert "«résumé»" in r["hits"][0]["snippet"]


def test_find_across_a_line_break(folder):
    r = files.find_in_files(folder / "report.pdf", "review and")
    assert r["hit_count"] == 1


def test_find_exact_and_regex(folder):
    assert files.find_in_files(folder / "report.pdf", "Résumé", exact=True)["hit_count"] == 0
    assert files.find_in_files(folder / "report.pdf", "résumé", exact=True)["hit_count"] == 1
    r = files.find_in_files(folder / "report.pdf", r"introduc\w+|r[ée]sum[ée]", regex=True)
    assert r["hit_count"] == 2


def test_find_locations_in_each_format(folder):
    where = lambda name, q: [h["where"] for h in files.find_in_files(folder / name, q)["hits"]]
    assert where("deck.pptx", "cafe") == ["slide 2"]
    assert where("table.xlsx", "checklist") == ["sheet Data, row 3"]
    assert where("notes.txt", "resume") == ["text, line 2"]
    assert where("doc.docx", "resume")[0].startswith("document, line")


def test_find_reaches_into_attachments(folder):
    eml = [h["where"] for h in files.find_in_files(folder / "letter.eml", "resume")["hits"]]
    assert "attachment 1 (café report.pdf) › page 2" in eml
    assert any(w.startswith("attachment 3 (résumé.docx) › document") for w in eml)
    inner = files.find_in_files(folder / "letter.eml", "deadline")["hits"]
    assert inner[0]["where"].startswith("attachment 2 (Internal note.eml) › body")
    msg = [h["where"] for h in files.find_in_files(folder / "letter.msg", "resume")["hits"]]
    assert msg == ["body, line 1", "attachment 1 (agenda.pdf) › page 2"]


def test_find_in_a_scanned_page_uses_the_ocr_server(folder, fake_ocr, monkeypatch):
    assert files.find_in_files(folder / "scan.pdf", "resume")["not_searched"] == ["page(s) 1"]
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    r = files.find_in_files(folder / "scan.pdf", "resume")
    assert r["hit_count"] == 1 and "not_searched" not in r


def test_find_in_a_folder(folder, fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    r = files.find_in_files(folder, "resume")
    found = {p["path"].split("\\")[-1].split("/")[-1]: p["hits"] for p in r["files_with_hits"]}
    assert set(found) == {"report.pdf", "letter.eml", "letter.msg", "doc.docx", "table.xlsx", "notes.txt"}
    assert r["files_searched"] == 9
    assert [p.split("\\")[-1].split("/")[-1] for p in r["scanned_pdfs_not_searched"]] == ["scan.pdf"]
    assert fake_ocr.requests == []                 # a folder search never OCRs
    hits = [f["hits"] for f in r["files_with_hits"]]
    assert hits == sorted(hits, reverse=True)


def test_folder_search_skips_big_files(folder, monkeypatch):
    monkeypatch.setattr(files, "FIND_FOLDER_MAX_BYTES", 1000)
    r = files.find_in_files(folder, "line")
    assert r["too_big_for_folder_search"] and any("were skipped" in n for n in r["notes"])


def test_find_errors(folder):
    assert json.loads(files.run(files.find_in_files, folder / "notes.txt", "  "))["error"] == "the search query is empty"
    bad = json.loads(files.run(files.find_in_files, folder / "notes.txt", "(", regex=True))
    assert bad["ok"] is False and "regular expression" in bad["error"]


def test_find_caps_hits(folder):
    r = files.find_in_files(folder / "notes.txt", "line", max_hits=1)
    assert r["hit_count"] == 2 and len(r["hits"]) == 1 and "1 more match" in r["next"]


# ----------------------------------------------------- attachments, render ---

def test_save_attachments_never_overwrites(folder, tmp_path):
    dest = tmp_path / "out"
    first = files.save_attachments(folder / "letter.eml", dest=str(dest))
    second = files.save_attachments(folder / "letter.eml", dest=str(dest))
    names = sorted(p.name for p in dest.iterdir())
    assert len(first["saved"]) == 3 and len(names) == 6
    assert "café report (2).pdf" in names
    assert second["saved"][0]["saved_to"].endswith("café report (2).pdf")


def test_save_attachments_skips_inline_pictures(folder, tmp_path):
    r = files.save_attachments(folder / "letter.msg", dest=str(tmp_path / "m"))
    assert [s["name"] for s in r["saved"]] == ["agenda.pdf"]
    assert "inline" in r["notes"][0]
    r = files.save_attachments(folder / "letter.msg", dest=str(tmp_path / "m2"), include_inline=True)
    assert len(r["saved"]) == 2


def test_render_pages(folder):
    r = files.render_pages(folder / "report.pdf", pages="1-2", dpi=50)
    assert [i["page"] for i in r["images"]] == [1, 2]
    from pathlib import Path
    assert Path(r["images"][0]["png"]).read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"
    assert json.loads(files.run(files.render_pages, folder / "doc.docx"))["ok"] is False


def test_status_reports_ocr_configuration(fake_ocr, monkeypatch):
    s = files.status()
    assert s["ocr_server"]["configured"] is False
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    assert files.status()["ocr_server"] == {"configured": True, "model": "deepseek-ai/DeepSeek-OCR-2"}
