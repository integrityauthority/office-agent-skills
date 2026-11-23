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
        '<w:p><w:r><w:t>Reference no: 123/2026</w:t></w:r></w:p>'
        '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>The deadline is Friday.</w:t></w:r>'
        '<w:commentRangeEnd w:id="0"/></w:p>'
        '<w:p><w:ins w:id="1" w:author="Smith" w:date="2026-11-01T10:00:00Z"><w:r><w:t>new sentence</w:t></w:r></w:ins>'
        '<w:del w:id="2" w:author="Smith"><w:r><w:delText>old sentence</w:delText></w:r></w:del></w:p>'
        # The ballot-box characters are what Word writes for checkboxes.
        '<w:sdt><w:sdtPr><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr>'
        '<w:sdtContent><w:r><w:t>☒</w:t></w:r></w:sdtContent></w:sdt>'
        '<w:p><w:r><w:t>☐ No notifications, please</w:t></w:r></w:p>'
        '<w:p><w:r><w:sym w:font="Wingdings" w:char="F0FE"/><w:t> I consent</w:t></w:r></w:p>'
    )
    comments = ('<w:comment w:id="0" w:author="Jane Doe" w:date="2026-11-02T09:00:00Z">'
                '<w:p><w:r><w:t>Needs checking.</w:t></w:r></w:p></w:comment>')
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
    stream = b"BT /F1 12 Tf 72 760 Td (Decision) Tj ET"
    return raw_pdf([
        b"<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R 6 0 R] /SigFlags 3 >> >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R "
        b"/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> "
        b"/Annots [5 0 R 6 0 R] >>",
        b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
        b"<< /FT /Tx /T (Client) /V (Anna Smith) /Type /Annot /Subtype /Widget /Rect [0 0 0 0] /P 3 0 R >>",
        b"<< /FT /Sig /T (Signature) /V 7 0 R /Type /Annot /Subtype /Widget /Rect [0 0 0 0] /P 3 0 R >>",
        b"<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached "
        b"/Name (Dr. John Smith) /M (D:20261103120000+01'00') /Contents <00> /ByteRange [0 0 0 0] >>",
    ])


# --------------------------------------------------------------- archive ---

def test_list_and_read_members():
    data = make_zip({"a/one.txt": "first café", "two.pdf": build_pdf(["Second"]), "empty/": ""})
    members, total = archive.list_members(data)
    assert total == 2 and [m["name"] for m in members] == ["a/one.txt", "two.pdf"]
    assert archive.read_member(data, "1") == ("one.txt", "first café".encode())
    assert archive.read_member(data, "ONE.TXT")[0] == "one.txt"        # bare name, any case
    with pytest.raises(archive.ArchiveError):
        archive.read_member(data, "3")
    with pytest.raises(archive.ArchiveError) as err:
        archive.read_member(data, "missing.doc")
    assert "1=a/one.txt" in err.value.hint


def test_a_zip_bomb_member_is_refused():
    data = make_zip({"bomb.txt": b"0" * 3_000_000})
    members, _ = archive.list_members(data)
    assert members[0]["unreadable"] == "compression ratio too high"
    with pytest.raises(archive.ArchiveError):
        archive.read_member(data, "1")


@pytest.mark.parametrize("name", ["../outside.txt", "/absolute.txt", "C:/drive.txt", "a/../../outside.txt",
                                  "dot./x.txt", "colon:name.txt"])
def test_unpack_skips_paths_that_escape(tmp_path, name):
    z = tmp_path / "bundle.zip"
    z.write_bytes(make_zip({name: "x", "good/file.txt": "fine"}))
    result = archive.unpack(z)
    assert result["count"] == 1 and result["files"][0]["path"].replace("\\", "/") == "good/file.txt"
    assert result["skipped"] == [{"name": name, "reason": "path escapes the destination"}]
    assert not (tmp_path / "outside.txt").exists()


def test_unpack_never_overwrites(tmp_path):
    z = tmp_path / "bundle.zip"
    z.write_bytes(make_zip({"letter.txt": "new"}))
    (tmp_path / "bundle").mkdir()
    (tmp_path / "bundle" / "letter.txt").write_text("old", encoding="utf-8")
    result = ok(files.run(files.unpack_file, str(z)))
    assert (tmp_path / "bundle" / "letter.txt").read_text(encoding="utf-8") == "old"
    assert (tmp_path / "bundle" / "letter (2).txt").read_text(encoding="utf-8") == "new"
    assert result["by_suffix"] == {".txt": 1}


def test_unpack_refuses_what_is_not_a_zip(tmp_path):
    docx = tmp_path / "document.docx"
    docx.write_bytes(rich_docx())
    result = json.loads(files.run(files.unpack_file, str(docx)))
    assert result["ok"] is False and "not a .zip" in result["error"]


# ------------------------------------------------------ zip through files ---

def test_read_find_and_read_member_of_a_zip(tmp_path):
    inner = make_zip({"inner.txt": "deep down"})
    z = tmp_path / "documents.zip"
    z.write_bytes(make_zip({"report.pdf": build_pdf(["Introduction", "The résumé review"]),
                            "document.docx": rich_docx(), "inner.zip": inner}))
    listing = ok(files.run(files.read_file, str(z)))
    assert listing["kind"] == "zip" and listing["member_count"] == 3
    member = ok(files.run(files.read_file, str(z), attachment="report.pdf"))
    assert member["attachment"]["name"] == "report.pdf" and "--- page 2 ---" in member["text"]
    found = ok(files.run(files.find_in_files, str(z), "resume"))
    assert found["hit_count"] >= 1 and "report.pdf" in found["hits"][0]["where"]
    assert "page 2" in found["hits"][0]["where"]


# ---------------------------------------------------------- doc details ---

def test_docx_comments_revisions_and_checkboxes():
    d = docmeta.docx_details(rich_docx())
    assert d["comments"] == [{"id": "0", "author": "Jane Doe", "date": "2026-11-02T09:00:00Z",
                              "text": "Needs checking.", "anchor": "The deadline is Friday."}]
    assert [(r["kind"], r["text"]) for r in d["revisions"]] == [("insert", "new sentence"), ("delete", "old sentence")]
    assert d["revisions"][0]["author"] == "Smith"
    boxes = [(b["kind"], b["checked"]) for b in d["checkboxes"]]
    assert ("content control", True) in boxes
    assert ("symbol", False) in boxes                     # ☐ No notifications, please
    wingdings = [b for b in d["checkboxes"] if "I consent" in b["label"]]
    assert wingdings and wingdings[0]["checked"] is True


def test_read_docx_reports_the_details(tmp_path):
    p = tmp_path / "document.docx"
    p.write_bytes(rich_docx())
    r = ok(files.run(files.read_file, str(p)))
    assert r["comments"][0]["author"] == "Jane Doe" and len(r["revisions"]) == 2 and r["checkboxes"]


def test_pdf_signers_and_form_fields():
    pytest.importorskip("pypdf")
    d = docmeta.pdf_details(signed_form_pdf())
    assert d["signed"] is True and d["signers"] == ["Dr. John Smith"]
    assert d["form_fields"] == {"Client": "Anna Smith"}
    plain = docmeta.pdf_details(build_pdf(["No signature"]))
    assert plain["signed"] is False and plain["signers"] == [] and plain["form_fields"] == {}


# ---------------------------------------------------------------- labels ---

def test_extract_labels_ignores_case_and_accents():
    text = "REFERENCE NO: 123/2026\nDeadline - 1 December 2026\nRÉSUMÉ RECEIVED:2026-11-01\n"
    found = files.extract_labels(text, ["reference no", "Deadline", "resume received", "No such label"])
    assert found == {"reference no": "123/2026", "Deadline": "1 December 2026", "resume received": "2026-11-01"}


def test_read_with_labels(tmp_path):
    p = tmp_path / "document.docx"
    p.write_bytes(rich_docx())
    r = ok(files.run(files.read_file, str(p), labels="Reference no"))
    assert r["labels"] == {"Reference no": "123/2026"}


# --------------------------------------------------------------- collect ---

@pytest.fixture
def letters(tmp_path):
    d = tmp_path / "letters"
    d.mkdir()
    (d / "a.msg").write_bytes(build_msg(subject="Request", body="Reference no: 11/2026\nPlease review the attached résumé.",
                                        sender=("Anna", "anna@example.com"), to=[("Bea", "bea@example.com")]))
    (d / "b.docx").write_bytes(rich_docx())
    (d / "c.pdf").write_bytes(signed_form_pdf())
    (d / "d.pdf").write_bytes(build_pdf([None]))
    (d / "e.zip").write_bytes(make_zip({"x.txt": "left out"}))      # a zip in a folder is not opened
    return d


def test_collect_a_folder(letters):
    r = ok(files.run(files.collect_files, str(letters), labels="Reference no", query="resume"))
    rows = {row["file"]: row for row in r["rows"]}
    assert set(rows) == {"a.msg", "b.docx", "c.pdf", "d.pdf"} and r["count"] == 4
    assert rows["a.msg"]["subject"] == "Request" and rows["a.msg"]["label:Reference no"] == "11/2026"
    assert rows["a.msg"]["hits"] == 1 and r["files_with_hits"] == 1
    assert rows["b.docx"]["label:Reference no"] == "123/2026" and rows["b.docx"]["comments"] == 1
    assert "I consent" in rows["b.docx"]["checked"]
    assert rows["c.pdf"]["signers"] == "Dr. John Smith"
    assert rows["d.pdf"]["scanned_pages"] == "1"
    assert "csv" not in r and "office-live" in r["next"]


def test_collect_a_zip_and_write_csv(tmp_path, letters):
    z = tmp_path / "letters.zip"
    z.write_bytes(make_zip({p.name: p.read_bytes() for p in letters.iterdir() if p.suffix != ".zip"}))
    out = tmp_path / "table.csv"
    r = ok(files.run(files.collect_files, str(z), labels="Reference no", excerpt=0, out=str(out)))
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
    d = tmp_path / "many"
    d.mkdir()
    for n in range(3):
        (d / f"{n}.txt").write_text(f"Line {n}", encoding="utf-8")
    r = ok(files.run(files.collect_files, str(d)))
    assert r["count"] == 3 and len(r["rows"]) == 2 and r["csv"].endswith("many.csv")
    assert any("CSV only" in note for note in r["notes"])


# ---------------------------------------------------------------- images ---

def make_image(fmt="PNG", frames=1):
    Image = pytest.importorskip("PIL.Image")
    pictures = [Image.new("RGB", (400, 200), (255, 255, 255)) for _ in range(frames)]
    buf = io.BytesIO()
    pictures[0].save(buf, format=fmt, save_all=frames > 1, append_images=pictures[1:], dpi=(300, 300))
    return buf.getvalue()


def test_an_image_is_read_with_ocr(tmp_path, fake_local_ocr):
    p = tmp_path / "screenshot.png"
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
    p = tmp_path / "photo.jpg"
    p.write_bytes(make_image("JPEG"))
    r = ok(files.run(files.read_file, str(p), ocr="never"))
    assert r["kind"] == "image" and r["text"] == "" and r["notes"]


def test_find_in_an_image(tmp_path, fake_local_ocr):
    fake_local_ocr.reply = "Reference no: 77/2026, résumé due 1 December."
    p = tmp_path / "scan.png"
    p.write_bytes(make_image())
    r = ok(files.run(files.find_in_files, str(p), "resume"))
    assert r["hit_count"] == 1 and r["hits"][0]["where"] == "page 1"
