"""ai_mark against fake Office objects: what gets marked after an office_exec
call, what is left alone, and that it stays off unless switched on. The same
paths were checked once against live Word, PowerPoint, Excel and Outlook."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / ".agents" / "skills" / "office-live" / "scripts"))
import ai_mark  # noqa: E402
import outlook  # noqa: E402

BLUE = 0x0F + 0x9E * 256 + 0xD5 * 65536


@pytest.fixture(autouse=True)
def clean_state(monkeypatch):
    monkeypatch.delenv("OFFICE_LIVE_MARK_AI", raising=False)
    monkeypatch.delenv("OFFICE_LIVE_MARK_COLOR", raising=False)
    ai_mark._SEEN.clear()
    ai_mark._reset()
    ai_mark._ACTIVE = False
    yield
    ai_mark._SEEN.clear()
    ai_mark._reset()
    ai_mark._ACTIVE = False


def switch_on(monkeypatch):
    monkeypatch.setenv("OFFICE_LIVE_MARK_AI", "on")


# ------------------------------------------------------------- settings ---

def test_off_unless_switched_on(monkeypatch):
    assert not ai_mark.enabled()
    for value in ("1", "true", "YES", "on"):
        monkeypatch.setenv("OFFICE_LIVE_MARK_AI", value)
        assert ai_mark.enabled()
    monkeypatch.setenv("OFFICE_LIVE_MARK_AI", "off")
    assert not ai_mark.enabled()


def test_colour_defaults_and_rejects_garbage(monkeypatch):
    assert ai_mark.color() == "#0F9ED5"
    assert ai_mark.color_com() == BLUE
    monkeypatch.setenv("OFFICE_LIVE_MARK_COLOR", "#ff0000")
    assert ai_mark.color() == "#FF0000" and ai_mark.color_com() == 0x0000FF
    monkeypatch.setenv("OFFICE_LIVE_MARK_COLOR", "blue")
    assert ai_mark.color() == "#0F9ED5"


@pytest.mark.parametrize("old,new,spans", [
    ("Hello", "Hello", []),
    ("Hello", "Hello world", [(5, 6)]),
    ("Hello world", "Hello brave world", [(6, 6)]),
    ("", "All new", [(0, 7)]),
    ("Old text", "New text", [(0, 3)]),
])
def test_inserted_spans(old, new, spans):
    assert ai_mark.inserted_spans(old, new) == spans


# ------------------------------------------------------------------- Word ---

class Font:
    def __init__(self):
        self.Underline, self.UnderlineColor = 0, -16777216


class Range:
    def __init__(self, text):
        self.Text, self.Font = text, Font()


class Revision:
    def __init__(self, doc, kind, text, author="User", date="2026-01-01 10:00:00"):
        self.doc, self.Type, self.Author, self.Date = doc, kind, author, date
        self.Range = Range(text)

    def Accept(self):
        self.doc.revisions.remove(self)
        self.doc.accepted.append(self.Range.Text)


class Revisions:
    def __init__(self, doc):
        self.doc = doc

    @property
    def Count(self):
        return len(self.doc.revisions)

    def Item(self, i):
        return self.doc.revisions[i - 1]


class WordDoc:
    def __init__(self, name, tracking=False):
        self.FullName, self.Name, self.TrackRevisions = name, name, tracking
        self.revisions, self.accepted = [], []
        self.Revisions = Revisions(self)

    def type(self, text):
        """What Word does with an insertion: a revision only while tracking."""
        if self.TrackRevisions:
            self.revisions.append(Revision(self, ai_mark.WD_REVISION_INSERT, text, "Hermes"))


def test_untracked_document_ends_marked_with_the_changes_accepted(monkeypatch):
    switch_on(monkeypatch)
    doc = WordDoc("report.docx")
    ai_mark.see("word", doc)
    ai_mark.begin("doc.Content.InsertAfter('x')")
    assert doc.TrackRevisions                 # on for the duration of the call
    doc.type("New paragraph.")
    added = doc.revisions[0].Range
    assert ai_mark.end() == {"word": 1}
    assert (added.Font.Underline, added.Font.UnderlineColor) == (6, BLUE)
    assert doc.accepted == ["New paragraph."] and doc.revisions == []
    assert doc.TrackRevisions is False        # back as the user had it


def test_tracked_document_keeps_its_revisions(monkeypatch):
    switch_on(monkeypatch)
    doc = WordDoc("contract.docx", tracking=True)
    doc.revisions.append(Revision(doc, ai_mark.WD_REVISION_INSERT, "The user's own edit."))
    users = doc.revisions[0].Range
    ai_mark.see("word", doc)
    ai_mark.begin()
    doc.type("Agent text.")
    assert ai_mark.end() == {"word": 1}
    assert users.Font.Underline == 0          # existing revisions are not the agent's
    assert doc.revisions[1].Range.Font.Underline == 6
    assert doc.accepted == [] and doc.TrackRevisions is True


def test_tracking_the_call_turned_on_stays_on(monkeypatch):
    switch_on(monkeypatch)
    doc = WordDoc("memo.docx")
    ai_mark.see("word", doc)
    ai_mark.begin("doc.TrackRevisions = True\ndoc.Content.InsertAfter('x')")
    doc.type("Tracked on purpose.")
    ai_mark.end()
    assert doc.TrackRevisions is True and len(doc.revisions) == 1 and doc.accepted == []


def test_a_document_opened_during_the_call_is_watched_from_then(monkeypatch):
    switch_on(monkeypatch)
    ai_mark.begin()
    doc = WordDoc("Document2")
    ai_mark.see("word", doc)                  # new_doc / open_doc / attach call this
    doc.type("Written into the new document.")
    assert ai_mark.end() == {"word": 1}


def test_nothing_happens_when_off():
    doc = WordDoc("report.docx")
    ai_mark.see("word", doc)
    ai_mark.begin()
    assert doc.TrackRevisions is False
    assert ai_mark.end() is None


def test_a_document_the_call_closed_is_skipped_quietly(monkeypatch):
    switch_on(monkeypatch)

    class Closing(WordDoc):
        closed = False

        def __getattribute__(self, name):
            if name == "Name" and object.__getattribute__(self, "closed"):
                raise RuntimeError("The object invoked has disconnected from its clients.")
            return object.__getattribute__(self, name)

    doc = Closing("Document3")
    ai_mark.see("word", doc)
    ai_mark.begin()
    doc.closed = True
    assert ai_mark.end() is None
    assert "document3" not in ai_mark._SEEN


# ------------------------------------------------------------- PowerPoint ---

class Font2:
    def __init__(self):
        self.UnderlineStyle = 0
        self.UnderlineColor = type("Color", (), {"RGB": 0})()


class TextRange2:
    def __init__(self, text):
        self.Text, self.Font, self.marked = text, Font2(), []

    def Characters(self, start, length):
        part = TextRange2(self.Text[start - 1:start - 1 + length])
        self.marked.append((start, length, part))
        return part


class Shape:
    Type, HasTable, HasTextFrame = 1, False, True

    def __init__(self, shape_id, text):
        self.Id = shape_id
        self.TextFrame2 = type("TF2", (), {})()
        self.TextFrame2.TextRange = TextRange2(text)


class Collection(list):
    @property
    def Count(self):
        return len(self)

    def Item(self, i):
        return self[i - 1]


class Presentation:
    def __init__(self, name):
        self.FullName = self.Name = name
        self.slide = type("Slide", (), {"SlideID": 256, "Shapes": Collection()})()
        self.Slides = Collection([self.slide])


def test_powerpoint_marks_only_the_inserted_characters(monkeypatch):
    switch_on(monkeypatch)
    pres = Presentation("deck.pptx")
    title = Shape(2, "Quarterly review")
    pres.slide.Shapes.append(title)
    ai_mark.see("ppt", pres)
    ai_mark.begin()
    title.TextFrame2.TextRange.Text = "Quarterly review, draft"
    pres.slide.Shapes.append(Shape(3, "New box"))
    assert ai_mark.end() == {"ppt": 2}
    (start, length, part), = title.TextFrame2.TextRange.marked
    assert (start, length, part.Text) == (17, 7, ", draft")
    assert part.Font.UnderlineStyle == 4 and part.Font.UnderlineColor.RGB == BLUE
    new_box = pres.slide.Shapes[1].TextFrame2.TextRange.marked
    assert [(s, n) for s, n, _p in new_box] == [(1, 7)]


# ------------------------------------------------------------------ Excel ---

class Cell:
    def __init__(self, value):
        self.Value, self.Count = value, 1
        self.Font = type("F", (), {"Underline": -4142})()


class Sheet:
    def __init__(self, book, cells):
        self.Parent, self.Name, self.cells = book, "Sheet1", cells

    def Range(self, address):
        return self.cells[address]


def test_excel_marks_written_cells_of_watched_books_only(monkeypatch):
    switch_on(monkeypatch)
    book = type("Book", (), {"FullName": "budget.xlsx", "Name": "budget.xlsx"})()
    other = type("Book", (), {"FullName": "the-users-own.xlsx", "Name": "x"})()
    ai_mark._SEEN["budget.xlsx"] = ("excel", book)
    cells = {"$B$2": Cell("Total"), "$C$3": Cell(None)}
    sheet, users = Sheet(book, cells), Sheet(other, {"$A$1": Cell("typed by the user")})
    ai_mark.begin()
    ai_mark._ExcelEvents().OnSheetChange(sheet, type("T", (), {"Address": "$B$2"})())
    ai_mark._ExcelEvents().OnSheetChange(sheet, type("T", (), {"Address": "$C$3"})())
    ai_mark._ExcelEvents().OnSheetChange(users, type("T", (), {"Address": "$A$1"})())
    assert ai_mark.end() == {"excel": 1}
    assert cells["$B$2"].Font.Underline == -4119
    assert cells["$C$3"].Font.Underline == -4142          # emptied, not marked
    assert users.cells["$A$1"].Font.Underline == -4142    # not a watched book


# ---------------------------------------------------------------- Outlook ---

def test_outlook_fragment_carries_the_mark_only_when_on(monkeypatch):
    assert "text-underline" not in outlook._fragment(body="Hello")
    switch_on(monkeypatch)
    fragment = outlook._fragment(body="Hello\n\nSecond paragraph")
    assert "text-underline:thick" in fragment and "#0F9ED5" in fragment
    assert fragment.count("<div") == 1                     # one div around it all


# ----------------------------------------------------------------- by hand ---

def test_mark_by_hand_picks_the_application():
    word_range = Range("text")
    assert ai_mark.mark(word_range) == "word" and word_range.Font.Underline == 6
    ppt_range = TextRange2("text")
    assert ai_mark.mark(ppt_range) == "ppt" and ppt_range.Font.UnderlineStyle == 4
    cell = Cell("text")
    cell.Address = "$A$1"
    assert ai_mark.mark(cell) == "excel" and cell.Font.Underline == -4119
