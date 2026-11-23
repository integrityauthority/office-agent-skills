"""Zip member names written in a DOS code page, Excel 97-2003 .xls reading,
and rendering one region of a page or a picture. Every file is generated here."""
import datetime
import io
import json
import zipfile
from pathlib import Path

import pytest

import archive
import files
from pdfbuilder import build_pdf


def ok(result):
    data = json.loads(result) if isinstance(result, str) else result
    assert data.get("ok", True), data
    return data


def legacy_zip(names, codepage):
    """A zip whose member names are stored in codepage, without the UTF-8 flag,
    as Windows Explorer and older archivers write them."""
    placeholders = {}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for n, name in enumerate(names):
            raw = name.encode(codepage)
            placeholder = f"{n:03d}".encode() + b"#" * (len(raw) - 3)
            assert len(placeholder) == len(raw)
            placeholders[placeholder] = raw
            z.writestr(placeholder.decode("ascii"), f"tartalom {n}".encode())
    data = buf.getvalue()
    for placeholder, raw in placeholders.items():
        assert data.count(placeholder) == 2      # local header and central directory
        data = data.replace(placeholder, raw)
    return data


HUNGARIAN = ["Új_bejelentések/levél.txt", "Árvíztűrő tükörfúrógép.txt", "őrző értesítése.txt"]


@pytest.fixture(autouse=True)
def default_encoding(monkeypatch):
    monkeypatch.delenv("ZIP_NAME_ENCODING", raising=False)


@pytest.mark.parametrize("codepage", ["cp852", "cp1250"])
def test_dos_names_come_back_with_their_accents(codepage):
    members, total = archive.list_members(legacy_zip(HUNGARIAN, codepage))
    assert total == 3
    assert [m["name"] for m in members] == HUNGARIAN


def test_utf8_names_without_the_flag_are_read_as_utf8():
    data = legacy_zip(HUNGARIAN, "utf-8")
    assert [m["name"] for m in archive.list_members(data)[0]] == HUNGARIAN


def test_flagged_and_ascii_names_are_left_alone():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("Árvíztűrő.txt", b"x")          # Python sets the UTF-8 flag
        z.writestr("plain/name.txt", b"y")
    assert [m["name"] for m in archive.list_members(buf.getvalue())[0]] == ["Árvíztűrő.txt", "plain/name.txt"]


def test_one_fixed_code_page_can_be_forced(monkeypatch):
    data = legacy_zip(["bejelentés.txt"], "cp1250")
    monkeypatch.setenv("ZIP_NAME_ENCODING", "cp852")
    assert archive.list_members(data)[0][0]["name"] == "bejelentÚs.txt"
    monkeypatch.setenv("ZIP_NAME_ENCODING", "cp1250")
    assert archive.list_members(data)[0][0]["name"] == "bejelentés.txt"
    monkeypatch.setenv("ZIP_NAME_ENCODING", "no-such-codepage")
    assert archive.list_members(data)[0][0]["name"] == "bejelentés.txt"


def test_read_member_by_accented_name_and_unpack(tmp_path):
    data = legacy_zip(HUNGARIAN, "cp852")
    name, payload = archive.read_member(data, "levél.txt")
    assert name == "levél.txt" and payload == b"tartalom 0"
    source = tmp_path / "Új_bejelentések 10.02.zip"
    source.write_bytes(data)
    result = archive.unpack(source, tmp_path / "ki")
    assert result["count"] == 3
    assert (tmp_path / "ki" / "Új_bejelentések" / "levél.txt").read_bytes() == b"tartalom 0"
    assert (tmp_path / "ki" / "őrző értesítése.txt").is_file()


def test_file_read_lists_the_accented_names(tmp_path):
    path = tmp_path / "csomag.zip"
    path.write_bytes(legacy_zip(HUNGARIAN, "cp852"))
    r = ok(files.run(files.read_file, path))
    assert [m["name"] for m in r["members"]] == HUNGARIAN


# ------------------------------------------------------------------ .xls ---

def make_xls():
    xlwt = pytest.importorskip("xlwt")
    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet("Adatok")
    for col, value in enumerate(["Név", "Terület", "Kelte"]):
        sheet.write(0, col, value)
    sheet.write(1, 0, "Árvíztűrő Kft.")
    sheet.write(1, 1, 1234)
    sheet.write(1, 2, datetime.date(2026, 10, 3), xlwt.easyxf(num_format_str="YYYY-MM-DD"))
    sheet.write(2, 0, "tükörfúrógép")
    sheet.write(2, 1, 12.5)
    book.add_sheet("Üres")
    buf = io.BytesIO()
    book.save(buf)
    return buf.getvalue()


def test_xls_is_read_like_xlsx(tmp_path):
    pytest.importorskip("xlrd")
    path = tmp_path / "pénzügyi szolgáltatók.xls"
    path.write_bytes(make_xls())
    r = ok(files.run(files.read_file, path))
    assert r["kind"] == "xls"
    assert r["sheets"] == ["Adatok", "Üres"]
    assert "Árvíztűrő Kft.\t1234\t2026-10-03" in r["text"]
    assert "tükörfúrógép\t12.5" in r["text"]


def test_xls_is_detected_without_its_extension():
    pytest.importorskip("xlrd")
    assert files.detect_kind(make_xls(), "melleklet.bin") == "xls"


def test_xls_is_searched_with_sheet_and_row(tmp_path):
    pytest.importorskip("xlrd")
    path = tmp_path / "lista.xls"
    path.write_bytes(make_xls())
    r = ok(files.run(files.find_in_files, path, "tukorfurogep"))
    assert r["hit_count"] == 1
    assert r["hits"][0]["where"] == "sheet Adatok, row 3"


def test_damaged_xls_gives_a_hint(tmp_path):
    pytest.importorskip("xlrd")
    path = tmp_path / "rossz.xls"
    path.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 600)
    r = json.loads(files.run(files.read_file, path))
    assert r["ok"] is False and r["hint"]


# ---------------------------------------------------------------- render ---

def test_parse_box_accepts_fractions_and_refuses_the_rest():
    assert files.parse_box("0,0.5,1,1") == (0.0, 0.5, 1.0, 1.0)
    assert files.parse_box([0.1, 0.2, 0.3, 0.4]) == (0.1, 0.2, 0.3, 0.4)
    assert files.parse_box(None) is None
    for bad in ("1,2,3", "0,0,0,1", "0.5,0,0.2,1", "a,b,c,d", "0,0,1,1.5"):
        with pytest.raises(files.FileError):
            files.parse_box(bad)


def test_render_box_crops_the_region_at_a_higher_resolution(tmp_path):
    from PIL import Image
    path = tmp_path / "vn.pdf"
    path.write_bytes(build_pdf(["Vagyonnyilatkozat 1. oldal", "2. oldal"]))
    whole = files.render_pages(path, pages="1", dpi=72)
    part = files.render_pages(path, pages="1", box="0,0.5,0.5,1")
    assert part["dpi"] == files.RENDER_BOX_DPI
    assert part["images"][0]["box"] == [0.0, 0.5, 0.5, 1.0]
    full_w, full_h = Image.open(whole["images"][0]["png"]).size
    crop_w, crop_h = Image.open(part["images"][0]["png"]).size
    scale = files.RENDER_BOX_DPI / 72
    assert abs(crop_w - full_w * scale / 2) <= 2
    assert abs(crop_h - full_h * scale / 2) <= 2
    assert Path(part["images"][0]["png"]).name.endswith("-box00-50-50-100.png")


def test_render_caps_the_resolution(tmp_path):
    path = tmp_path / "a.pdf"
    path.write_bytes(build_pdf(["x"]))
    assert files.render_pages(path, dpi=5000)["dpi"] == files.RENDER_MAX_DPI


def test_render_reads_a_picture(tmp_path):
    from PIL import Image
    path = tmp_path / "kepernyokep.png"
    Image.new("RGB", (400, 200), "white").save(path)
    r = files.render_pages(path, box="0,0,0.5,1")
    assert len(r["images"]) == 1
    assert Path(r["images"][0]["png"]).read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"


def test_render_box_error_is_a_hint(tmp_path):
    path = tmp_path / "a.pdf"
    path.write_bytes(build_pdf(["x"]))
    r = json.loads(files.run(files.render_pages, path, box="0,0,2,1"))
    assert r["ok"] is False and "between 0 and 1" in r["hint"]
