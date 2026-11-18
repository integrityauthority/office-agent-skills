"""files.py: reading, searching, attachments and rendering, end to end on
files generated here. No real documents, no network beyond the fake OCR."""
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
    body = ('<w:p><w:r><w:t>Első bekezdés</w:t></w:r><w:r><w:tab/><w:t>tabbal</w:t></w:r></w:p>'
            '<w:p><w:ins><w:r><w:t>beszúrt szöveg</w:t></w:r></w:ins>'
            '<w:del><w:r><w:delText>törölt szöveg</w:delText></w:r></w:del></w:p>'
            '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>c1</w:t></w:r></w:p></w:tc>'
            '<w:tc><w:p><w:r><w:t>Összeférhetetlenség</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
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
        z.writestr("ppt/slides/slide1.xml", slide("Második dia: költségvetés"))
        z.writestr("ppt/slides/slide2.xml", slide("Első dia: cím"))
    return buf.getvalue()


def make_xlsx():
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Adatok"
    ws.append(["Név", "Összeg"])
    ws.append(["Kovács", 100])
    ws.append(["Összeférhetetlenségi nyilatkozat", 1])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def make_eml():
    inner = EmailMessage()
    inner["Subject"] = "Belső levél"
    inner["From"] = "a@example.com"
    inner["To"] = "b@example.com"
    inner.set_content("A belső levél törzse: határidő péntek.")
    m = EmailMessage()
    m["Subject"] = "Külső tárgy"
    m["From"] = "Kovács János <kj@example.com>"
    m["To"] = "x@example.com, y@example.com"
    m["Date"] = "Tue, 17 Nov 2026 10:00:00 +0100"
    m.set_content("Szia,\nmellékelem a jelentést.\n")
    m.add_alternative("<html><body><p>Szia,</p><p>mellékelem a jelentést.</p></body></html>", subtype="html")
    m.add_attachment(build_pdf(["Jelentés, 1. oldal", "A vizsgálat összeférhetetlenséget talált."]),
                     maintype="application", subtype="pdf", filename="jelentés.pdf")
    m.add_attachment(inner)
    m.add_attachment(make_docx(), maintype="application",
                     subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename="melléklet.docx")
    return bytes(m)


@pytest.fixture
def folder(tmp_path):
    (tmp_path / "jelentes.pdf").write_bytes(build_pdf(["Bevezetés", "Az összeférhetetlenség vizsgálata\nés eredménye"]))
    (tmp_path / "scan.pdf").write_bytes(build_pdf([None]))
    (tmp_path / "level.eml").write_bytes(make_eml())
    (tmp_path / "level.msg").write_bytes(build_msg(
        subject="Ülés", body="A napirend: összeférhetetlenség.", sender=("Anna", "anna@example.com"),
        to=[("Bea", "bea@example.com")],
        attachments=[{"name": "napirend.pdf", "data": build_pdf(["Napirend", "3. pont: összeférhetetlenség"])},
                     {"name": "logo.png", "mime": "image/png", "content_id": "l@x", "data": b"\x89PNG0000"}]))
    (tmp_path / "doc.docx").write_bytes(make_docx())
    (tmp_path / "deck.pptx").write_bytes(make_pptx())
    (tmp_path / "tabla.xlsx").write_bytes(make_xlsx())
    (tmp_path / "jegyzet.txt").write_text("első sor\nmásodik sor összeférhetetlenség\n", encoding="utf-8")
    (tmp_path / "oldal.html").write_text("<html><head><style>x{}</style></head><body><h1>Cím</h1><p>Bekezdés</p></body></html>",
                                         encoding="utf-8")
    return tmp_path


def ok(result):
    data = json.loads(result) if isinstance(result, str) else result
    assert data.get("ok", True), data
    return data


# ------------------------------------------------------------------ read ---

def test_detect_kind(folder):
    kinds = {p.name: files.detect_kind(p.read_bytes(), p.name) for p in folder.iterdir()}
    assert kinds == {"jelentes.pdf": "pdf", "scan.pdf": "pdf", "level.eml": "eml", "level.msg": "msg",
                     "doc.docx": "docx", "deck.pptx": "pptx", "tabla.xlsx": "xlsx",
                     "jegyzet.txt": "text", "oldal.html": "html"}


def test_read_pdf_with_page_markers(folder):
    r = files.read_file(folder / "jelentes.pdf")
    assert r["page_count"] == 2 and r["page_sources"] == {"text-layer": "1-2"}
    assert "--- page 2 ---\nAz összeférhetetlenség vizsgálata" in r["text"]


def test_read_pdf_truncation_points_to_the_next_pages(folder):
    r = files.read_file(folder / "jelentes.pdf", max_chars=20)
    assert r["truncated"] and "pages='" in r["next"]


def test_read_eml(folder):
    r = files.read_file(folder / "level.eml")
    assert r["kind"] == "email" and r["format"] == "eml"
    assert r["subject"] == "Külső tárgy" and r["from"] == "Kovács János <kj@example.com>"
    assert r["to"] == ["x@example.com", "y@example.com"]
    assert r["text"].startswith("Szia,") and r["body_source"] == "plain"
    assert [a["name"] for a in r["attachments"]] == ["jelentés.pdf", "Belső levél.eml", "melléklet.docx"]


def test_read_attachment_in_place(folder):
    r = files.read_file(folder / "level.eml", attachment="1")
    assert r["kind"] == "pdf" and r["attachment"]["name"] == "jelentés.pdf"
    r = files.read_file(folder / "level.eml", attachment="MELLÉKLET.DOCX")   # by name, any case
    assert r["kind"] == "docx"
    r = files.read_file(folder / "level.eml", attachment="2")
    assert r["kind"] == "email" and r["subject"] == "Belső levél"


def test_read_msg_and_its_attachment(folder):
    r = files.read_file(folder / "level.msg")
    assert r["format"] == "msg" and r["subject"] == "Ülés"
    assert "inline item" in r["next"]
    assert files.read_file(folder / "level.msg", attachment="1")["kind"] == "pdf"


def test_read_office_and_text(folder):
    docx = files.read_file(folder / "doc.docx")["text"]
    assert "beszúrt szöveg" in docx and "törölt szöveg" not in docx      # tracked deletions left out
    assert "c1\tÖsszeférhetetlenség" in docx
    pptx = files.read_file(folder / "deck.pptx")
    assert pptx["slide_count"] == 2
    assert pptx["text"].index("Első dia") < pptx["text"].index("Második dia")   # presentation order
    xlsx = files.read_file(folder / "tabla.xlsx")
    assert xlsx["sheets"] == ["Adatok"] and "Kovács\t100" in xlsx["text"]
    html = files.read_file(folder / "oldal.html")["text"]
    assert html == "Cím\n\nBekezdés"
    assert "második sor" in files.read_file(folder / "jegyzet.txt")["text"]


def test_read_errors(folder):
    err = ok_false = json.loads(files.run(files.read_file, folder / "nincs.pdf"))
    assert ok_false["ok"] is False and "no such file" in err["error"] and err["hint"]
    assert "folder" in json.loads(files.run(files.read_file, folder))["error"]
    bad = json.loads(files.run(files.read_file, folder / "level.eml", attachment="9"))
    assert bad["ok"] is False and "no attachment 9" in bad["error"]
    assert json.loads(files.run(files.read_file, folder / "doc.docx", attachment="1"))["ok"] is False


# ---------------------------------------------------------------- search ---

def test_find_ignores_case_and_accents(folder):
    r = files.find_in_files(folder / "jelentes.pdf", "OSSZEFERHETETLENSEG")
    assert r["hit_count"] == 1
    assert r["hits"][0]["where"] == "page 2"
    assert "«összeférhetetlenség»" in r["hits"][0]["snippet"]


def test_find_across_a_line_break(folder):
    r = files.find_in_files(folder / "jelentes.pdf", "vizsgálata és")
    assert r["hit_count"] == 1


def test_find_exact_and_regex(folder):
    assert files.find_in_files(folder / "jelentes.pdf", "Összeférhetetlenség", exact=True)["hit_count"] == 0
    assert files.find_in_files(folder / "jelentes.pdf", "összeférhetetlenség", exact=True)["hit_count"] == 1
    r = files.find_in_files(folder / "jelentes.pdf", r"bevezet\w+|eredm[ée]ny", regex=True)
    assert r["hit_count"] == 2


def test_find_locations_in_each_format(folder):
    where = lambda name, q: [h["where"] for h in files.find_in_files(folder / name, q)["hits"]]
    assert where("deck.pptx", "koltsegvetes") == ["slide 2"]
    assert where("tabla.xlsx", "nyilatkozat") == ["sheet Adatok, row 3"]
    assert where("jegyzet.txt", "osszeferhetetlenseg") == ["text, line 2"]
    assert where("doc.docx", "osszeferhetetlenseg")[0].startswith("document, line")


def test_find_reaches_into_attachments(folder):
    eml = [h["where"] for h in files.find_in_files(folder / "level.eml", "osszeferhetetlenseg")["hits"]]
    assert "attachment 1 (jelentés.pdf) › page 2" in eml
    assert any(w.startswith("attachment 3 (melléklet.docx) › document") for w in eml)
    inner = files.find_in_files(folder / "level.eml", "hatarido")["hits"]
    assert inner[0]["where"].startswith("attachment 2 (Belső levél.eml) › body")
    msg = [h["where"] for h in files.find_in_files(folder / "level.msg", "osszeferhetetlenseg")["hits"]]
    assert msg == ["body, line 1", "attachment 1 (napirend.pdf) › page 2"]


def test_find_in_a_scanned_page_uses_the_ocr_server(folder, fake_ocr, monkeypatch):
    assert files.find_in_files(folder / "scan.pdf", "szoveg")["not_searched"] == ["page(s) 1"]
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    r = files.find_in_files(folder / "scan.pdf", "szoveg")
    assert r["hit_count"] == 1 and "not_searched" not in r


def test_find_in_a_folder(folder, fake_ocr, monkeypatch):
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    r = files.find_in_files(folder, "osszeferhetetlenseg")
    found = {p["path"].split("\\")[-1].split("/")[-1]: p["hits"] for p in r["files_with_hits"]}
    assert set(found) == {"jelentes.pdf", "level.eml", "level.msg", "doc.docx", "tabla.xlsx", "jegyzet.txt"}
    assert r["files_searched"] == 9
    assert [p.split("\\")[-1].split("/")[-1] for p in r["scanned_pdfs_not_searched"]] == ["scan.pdf"]
    assert fake_ocr.requests == []                 # a folder search never OCRs
    hits = [f["hits"] for f in r["files_with_hits"]]
    assert hits == sorted(hits, reverse=True)


def test_folder_search_skips_big_files(folder, monkeypatch):
    monkeypatch.setattr(files, "FIND_FOLDER_MAX_BYTES", 1000)
    r = files.find_in_files(folder, "sor")
    assert r["too_big_for_folder_search"] and any("were skipped" in n for n in r["notes"])


def test_find_errors(folder):
    assert json.loads(files.run(files.find_in_files, folder / "jegyzet.txt", "  "))["error"] == "the search query is empty"
    bad = json.loads(files.run(files.find_in_files, folder / "jegyzet.txt", "(", regex=True))
    assert bad["ok"] is False and "regular expression" in bad["error"]


def test_find_caps_hits(folder):
    r = files.find_in_files(folder / "jegyzet.txt", "sor", max_hits=1)
    assert r["hit_count"] == 2 and len(r["hits"]) == 1 and "1 more match" in r["next"]


# ----------------------------------------------------- attachments, render ---

def test_save_attachments_never_overwrites(folder, tmp_path):
    dest = tmp_path / "out"
    first = files.save_attachments(folder / "level.eml", dest=str(dest))
    second = files.save_attachments(folder / "level.eml", dest=str(dest))
    names = sorted(p.name for p in dest.iterdir())
    assert len(first["saved"]) == 3 and len(names) == 6
    assert "jelentés (2).pdf" in names
    assert second["saved"][0]["saved_to"].endswith("jelentés (2).pdf")


def test_save_attachments_skips_inline_pictures(folder, tmp_path):
    r = files.save_attachments(folder / "level.msg", dest=str(tmp_path / "m"))
    assert [s["name"] for s in r["saved"]] == ["napirend.pdf"]
    assert "inline" in r["notes"][0]
    r = files.save_attachments(folder / "level.msg", dest=str(tmp_path / "m2"), include_inline=True)
    assert len(r["saved"]) == 2


def test_render_pages(folder):
    r = files.render_pages(folder / "jelentes.pdf", pages="1-2", dpi=50)
    assert [i["page"] for i in r["images"]] == [1, 2]
    from pathlib import Path
    assert Path(r["images"][0]["png"]).read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"
    assert json.loads(files.run(files.render_pages, folder / "doc.docx"))["ok"] is False


def test_status_reports_ocr_configuration(fake_ocr, monkeypatch):
    s = files.status()
    assert s["ocr_server"]["configured"] is False
    monkeypatch.setenv("OCR_URL", fake_ocr.url)
    assert files.status()["ocr_server"] == {"configured": True, "model": "deepseek-ai/DeepSeek-OCR-2"}
