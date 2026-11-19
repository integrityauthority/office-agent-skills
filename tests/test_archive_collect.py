"""Zip archives, labelled values, document details and folder tables:
archive.py, docmeta.py and the files.py functions built on them. Every file
is generated here."""
import csv
import io
import json
import zipfile

import pytest

import archive
import docmeta
import files
from msgbuilder import build_msg
from pdfbuilder import build_pdf

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W = f'xmlns:w="{W_NS}" xmlns:w14="{W14_NS}"'


def ok(result):
    data = json.loads(result) if isinstance(result, str) else result
    assert data.get("ok", True), data
    return data


def make_zip(members):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in members.items():
            z.writestr(name, data)
    return buf.getvalue()


def make_docx(body, comments=None):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("word/document.xml", f"<w:document {W}><w:body>{body}</w:body></w:document>")
        if comments:
            z.writestr("word/comments.xml", f"<w:comments {W}>{comments}</w:comments>")
    return buf.getvalue()


def rich_docx():
    body = (
        '<w:p><w:r><w:t>Iktatószám: 123/2026</w:t></w:r></w:p>'
        '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>A határidő péntek.</w:t></w:r>'
        '<w:commentRangeEnd w:id="0"/></w:p>'
        '<w:p><w:ins w:id="1" w:author="Kovács" w:date="2026-11-01T10:00:00Z"><w:r><w:t>új mondat</w:t></w:r></w:ins>'
        '<w:del w:id="2" w:author="Kovács"><w:r><w:delText>régi mondat</w:delText></w:r></w:del></w:p>'
        '<w:sdt><w:sdtPr><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr>'
        '<w:sdtContent><w:r><w:t>☒</w:t></w:r></w:sdtContent></w:sdt>'
        '<w:p><w:r><w:t>☐ Nem kérek értesítést</w:t></w:r></w:p>'
        '<w:p><w:r><w:sym w:font="Wingdings" w:char="F0FE"/><w:t> Hozzájárulok</w:t></w:r></w:p>'
    )
    comments = ('<w:comment w:id="0" w:author="Nagy Éva" w:date="2026-11-02T09:00:00Z">'
                '<w:p><w:r><w:t>Ellenőrizni kell.</w:t></w:r></w:p></w:comment>')
    return make_docx(body, comments)


def raw_pdf(objects):
    """A PDF from object bodies (1-based, object 1 the catalog), with a correct xref."""
    out = bytearray(b"%PDF-1.7\n")
    offsets = []
    for number, body in enumerate(objects, 1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % number + body + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    for offset in offsets:
        out += b"%010d 00000 n \n" % offset
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objects) + 1, xref)
    return bytes(out)


def signed_form_pdf():
    stream = b"BT /F1 12 Tf 72 760 Td (Hatarozat) Tj ET"
    return raw_pdf([
        b"<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R 6 0 R] /SigFlags 3 >> >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R "
        b"/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> "
        b"/Annots [5 0 R 6 0 R] >>",
        b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
        b"<< /FT /Tx /T (Ugyfel) /V (Kiss Anna) /Type /Annot /Subtype /Widget /Rect [0 0 0 0] /P 3 0 R >>",
        b"<< /FT /Sig /T (Alairas) /V 7 0 R /Type /Annot /Subtype /Widget /Rect [0 0 0 0] /P 3 0 R >>",
        b"<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached "
        b"/Name (Dr. Szabo Peter) /M (D:20261103120000+01'00') /Contents <00> /ByteRange [0 0 0 0] >>",
    ])


# --------------------------------------------------------------- archive ---

def test_list_and_read_members():
    data = make_zip({"a/egy.txt": "első", "ketto.pdf": build_pdf(["Második"]), "ures/": ""})
    members, total = archive.list_members(data)
    assert total == 2 and [m["name"] for m in members] == ["a/egy.txt", "ketto.pdf"]
    assert archive.read_member(data, "1") == ("egy.txt", "első".encode())
    assert archive.read_member(data, "EGY.TXT")[0] == "egy.txt"        # bare name, any case
    with pytest.raises(archive.ArchiveError):
        archive.read_member(data, "3")
    with pytest.raises(archive.ArchiveError) as err:
        archive.read_member(data, "nincs.doc")
    assert "1=a/egy.txt" in err.value.hint


def test_a_zip_bomb_member_is_refused():
    data = make_zip({"bomb.txt": b"0" * 3_000_000})
    members, _ = archive.list_members(data)
    assert members[0]["unreadable"] == "compression ratio too high"
    with pytest.raises(archive.ArchiveError):
        archive.read_member(data, "1")


@pytest.mark.parametrize("name", ["../kint.txt", "/abszolut.txt", "C:/meghajto.txt", "a/../../kint.txt",
                                  "pont./x.txt", "ket:pont.txt"])
def test_unpack_skips_paths_that_escape(tmp_path, name):
    z = tmp_path / "csomag.zip"
    z.write_bytes(make_zip({name: "x", "jo/fajl.txt": "rendben"}))
    result = archive.unpack(z)
    assert result["count"] == 1 and result["files"][0]["path"].replace("\\", "/") == "jo/fajl.txt"
    assert result["skipped"] == [{"name": name, "reason": "path escapes the destination"}]
    assert not (tmp_path / "kint.txt").exists()


def test_unpack_never_overwrites(tmp_path):
    z = tmp_path / "csomag.zip"
    z.write_bytes(make_zip({"level.txt": "új"}))
    (tmp_path / "csomag").mkdir()
    (tmp_path / "csomag" / "level.txt").write_text("régi", encoding="utf-8")
    result = ok(files.run(files.unpack_file, str(z)))
    assert (tmp_path / "csomag" / "level.txt").read_text(encoding="utf-8") == "régi"
    assert (tmp_path / "csomag" / "level (2).txt").read_text(encoding="utf-8") == "új"
    assert result["by_suffix"] == {".txt": 1}


def test_unpack_refuses_what_is_not_a_zip(tmp_path):
    docx = tmp_path / "irat.docx"
    docx.write_bytes(rich_docx())
    result = json.loads(files.run(files.unpack_file, str(docx)))
    assert result["ok"] is False and "not a .zip" in result["error"]


# ------------------------------------------------------ zip through files ---

def test_read_find_and_read_member_of_a_zip(tmp_path):
    inner = make_zip({"belso.txt": "mélyen"})
    z = tmp_path / "iratok.zip"
    z.write_bytes(make_zip({"jelentes.pdf": build_pdf(["Bevezetés", "Az összeférhetetlenség vizsgálata"]),
                            "irat.docx": rich_docx(), "belso.zip": inner}))
    listing = ok(files.run(files.read_file, str(z)))
    assert listing["kind"] == "zip" and listing["member_count"] == 3
    member = ok(files.run(files.read_file, str(z), attachment="jelentes.pdf"))
    assert member["attachment"]["name"] == "jelentes.pdf" and "--- page 2 ---" in member["text"]
    found = ok(files.run(files.find_in_files, str(z), "osszeferhetetlenseg"))
    assert found["hit_count"] >= 1 and "jelentes.pdf" in found["hits"][0]["where"]
    assert "page 2" in found["hits"][0]["where"]


# ---------------------------------------------------------- doc details ---

def test_docx_comments_revisions_and_checkboxes():
    d = docmeta.docx_details(rich_docx())
    assert d["comments"] == [{"id": "0", "author": "Nagy Éva", "date": "2026-11-02T09:00:00Z",
                              "text": "Ellenőrizni kell.", "anchor": "A határidő péntek."}]
    assert [(r["kind"], r["text"]) for r in d["revisions"]] == [("insert", "új mondat"), ("delete", "régi mondat")]
    assert d["revisions"][0]["author"] == "Kovács"
    boxes = [(b["kind"], b["checked"]) for b in d["checkboxes"]]
    assert ("content control", True) in boxes
    assert ("symbol", False) in boxes                     # ☐ Nem kérek értesítést
    wingdings = [b for b in d["checkboxes"] if "Hozzájárulok" in b["label"]]
    assert wingdings and wingdings[0]["checked"] is True


def test_read_docx_reports_the_details(tmp_path):
    p = tmp_path / "irat.docx"
    p.write_bytes(rich_docx())
    r = ok(files.run(files.read_file, str(p)))
    assert r["comments"][0]["author"] == "Nagy Éva" and len(r["revisions"]) == 2 and r["checkboxes"]


def test_pdf_signers_and_form_fields():
    pytest.importorskip("pypdf")
    d = docmeta.pdf_details(signed_form_pdf())
    assert d["signed"] is True and d["signers"] == ["Dr. Szabo Peter"]
    assert d["form_fields"] == {"Ugyfel": "Kiss Anna"}
    plain = docmeta.pdf_details(build_pdf(["Nincs aláírás"]))
    assert plain["signed"] is False and plain["signers"] == [] and plain["form_fields"] == {}


# ---------------------------------------------------------------- labels ---

def test_extract_labels_ignores_case_and_accents():
    text = "IKTATÓSZÁM: 123/2026\nHatáridő - 2026. december 1.\nÜgyintéző:Kiss Anna\n"
    found = files.extract_labels(text, ["iktatoszam", "Határidő", "ügyintéző", "Nincs ilyen"])
    assert found == {"iktatoszam": "123/2026", "Határidő": "2026. december 1.", "ügyintéző": "Kiss Anna"}


def test_read_with_labels(tmp_path):
    p = tmp_path / "irat.docx"
    p.write_bytes(rich_docx())
    r = ok(files.run(files.read_file, str(p), labels="Iktatószám"))
    assert r["labels"] == {"Iktatószám": "123/2026"}


# --------------------------------------------------------------- collect ---

@pytest.fixture
def letters(tmp_path):
    d = tmp_path / "levelek"
    d.mkdir()
    (d / "a.msg").write_bytes(build_msg(subject="Kérelem", body="Iktatószám: 11/2026\nKérem az összeférhetetlenség vizsgálatát.",
                                        sender=("Anna", "anna@example.com"), to=[("Bea", "bea@example.com")]))
    (d / "b.docx").write_bytes(rich_docx())
    (d / "c.pdf").write_bytes(signed_form_pdf())
    (d / "d.pdf").write_bytes(build_pdf([None]))
    (d / "e.zip").write_bytes(make_zip({"x.txt": "kimarad"}))      # a zip in a folder is not opened
    return d


def test_collect_a_folder(letters):
    r = ok(files.run(files.collect_files, str(letters), labels="Iktatószám", query="osszeferhetetlenseg"))
    rows = {row["file"]: row for row in r["rows"]}
    assert set(rows) == {"a.msg", "b.docx", "c.pdf", "d.pdf"} and r["count"] == 4
    assert rows["a.msg"]["subject"] == "Kérelem" and rows["a.msg"]["label:Iktatószám"] == "11/2026"
    assert rows["a.msg"]["hits"] == 1 and r["files_with_hits"] == 1
    assert rows["b.docx"]["label:Iktatószám"] == "123/2026" and rows["b.docx"]["comments"] == 1
    assert "Hozzájárulok" in rows["b.docx"]["checked"]
    assert rows["c.pdf"]["signers"] == "Dr. Szabo Peter"
    assert rows["d.pdf"]["scanned_pages"] == "1"
    assert "csv" not in r and "office-live" in r["next"]


def test_collect_a_zip_and_write_csv(tmp_path, letters):
    z = tmp_path / "levelek.zip"
    z.write_bytes(make_zip({p.name: p.read_bytes() for p in letters.iterdir() if p.suffix != ".zip"}))
    out = tmp_path / "tabla.csv"
    r = ok(files.run(files.collect_files, str(z), labels="Iktatószám", excerpt=0, out=str(out)))
    assert r["count"] == 4 and r["csv"] == str(out)
    with out.open(encoding="utf-8-sig", newline="") as handle:
        table = list(csv.DictReader(handle))
    assert {row["file"] for row in table} == {"a.msg", "b.docx", "c.pdf", "d.pdf"}
    assert "excerpt" not in table[0]


def test_collect_refuses_a_single_document(letters):
    r = json.loads(files.run(files.collect_files, str(letters / "b.docx")))
    assert r["ok"] is False and "file_read" in r["hint"]


def test_collect_writes_csv_when_the_table_is_long(tmp_path, monkeypatch):
    monkeypatch.setattr(files, "COLLECT_PREVIEW_ROWS", 2)
    monkeypatch.setattr(files, "_cache_dir", lambda *parts: tmp_path / "cache")
    d = tmp_path / "sok"
    d.mkdir()
    for n in range(3):
        (d / f"{n}.txt").write_text(f"Sor {n}", encoding="utf-8")
    r = ok(files.run(files.collect_files, str(d)))
    assert r["count"] == 3 and len(r["rows"]) == 2 and r["csv"].endswith("sok.csv")
    assert any("CSV only" in note for note in r["notes"])


# ---------------------------------------------------------------- images ---

def make_image(fmt="PNG", frames=1):
    Image = pytest.importorskip("PIL.Image")
    pictures = [Image.new("RGB", (400, 200), (255, 255, 255)) for _ in range(frames)]
    buf = io.BytesIO()
    pictures[0].save(buf, format=fmt, save_all=frames > 1, append_images=pictures[1:], dpi=(300, 300))
    return buf.getvalue()


def test_an_image_is_read_with_ocr(tmp_path, fake_local_ocr):
    p = tmp_path / "kepernyokep.png"
    p.write_bytes(make_image())
    assert files.detect_kind(p.read_bytes(), p.name) == "image"
    r = ok(files.run(files.read_file, str(p)))
    assert r["kind"] == "image" and r["page_sources"] == {"windows-ocr": "1"}
    assert fake_local_ocr.reply in r["text"]


def test_a_multipage_tiff_reads_page_by_page(tmp_path, fake_local_ocr):
    p = tmp_path / "fax.tif"
    p.write_bytes(make_image("TIFF", frames=3))
    r = ok(files.run(files.read_file, str(p), pages="2-3"))
    assert r["page_count"] == 3 and r["pages_read"] == "2-3" and len(fake_local_ocr.calls) == 2


def test_an_image_without_ocr_says_so(tmp_path):
    p = tmp_path / "foto.jpg"
    p.write_bytes(make_image("JPEG"))
    r = ok(files.run(files.read_file, str(p), ocr="never"))
    assert r["kind"] == "image" and r["text"] == "" and r["notes"]


def test_find_in_an_image(tmp_path, fake_local_ocr):
    fake_local_ocr.reply = "Iktatószám: 77/2026, határidő december 1."
    p = tmp_path / "scan.png"
    p.write_bytes(make_image())
    r = ok(files.run(files.find_in_files, str(p), "hatarido"))
    assert r["hit_count"] == 1 and r["hits"][0]["where"] == "page 1"
