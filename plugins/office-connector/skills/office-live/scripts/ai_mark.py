"""Mark the text the agent writes into Office, so the user sees what to review.

Off unless the deployment turns it on:

    OFFICE_LIVE_MARK_AI=on            (1, true, yes or on)
    OFFICE_LIVE_MARK_COLOR=#0F9ED5    (optional, #RRGGBB)

Word, PowerPoint and Outlook drafts get a thick underline in that colour and
nothing else changes. Excel has neither a thick nor a coloured underline, so a
cell the agent wrote gets a double underline in the text's own colour.

The model does not apply the mark; office_exec does, around every call.
begin() notes what each document the agent works on holds, end() marks what
was added in between. The user removes a mark by taking the underline off once
they have reviewed the text.

    Word        the call runs with Track Changes on, and every insertion it
                made is marked. A document that was not tracking goes back to
                that state with those changes accepted, unless the call itself
                turned tracking on.
    PowerPoint  the text of every shape (groups and table cells included) is
                compared before and after; the inserted characters are marked.
    Excel       SheetChange events name the cells written during the call.
    Outlook     new_mail, reply and forward put the mark on the text they insert.

Only documents the agent reached through attach, new_doc or open_doc are
watched, so a document the user is typing in elsewhere is left alone.
"""
from __future__ import annotations

import difflib
import os
import re
from collections import Counter

DEFAULT_COLOR = "#0F9ED5"

WD_UNDERLINE_THICK = 6
WD_REVISION_INSERT = 1
MSO_UNDERLINE_HEAVY = 4
MSO_GROUP = 6
XL_UNDERLINE_DOUBLE = -4119
XL_CELL_CONSTANTS = 2
XL_CELL_FORMULAS = -4123

_TRUE = {"1", "true", "yes", "on"}
_TRACKING_ON = re.compile(r"TrackRevisions\s*=\s*(True|1|-1)\b")


# --------------------------------------------------------------- settings ---

def enabled() -> bool:
    return os.environ.get("OFFICE_LIVE_MARK_AI", "").strip().lower() in _TRUE


def color() -> str:
    value = os.environ.get("OFFICE_LIVE_MARK_COLOR", "").strip()
    return value.upper() if re.fullmatch(r"#[0-9A-Fa-f]{6}", value) else DEFAULT_COLOR


def color_com(hex_color: str | None = None) -> int:
    """#RRGGBB as the BGR integer that COM colour properties take."""
    h = (hex_color or color()).lstrip("#")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


def css() -> str:
    """The same mark in HTML. Outlook's editor reads the mso properties."""
    c = color()
    return (f"text-decoration:underline;text-underline:thick;"
            f"text-decoration-color:{c};text-underline-color:{c}")


def describe() -> dict:
    return {"on": enabled(), "color": color(),
            "word_ppt_outlook": "thick underline in that colour",
            "excel": "double underline (Excel has no thick or coloured one)"}


def inserted_spans(old: str, new: str) -> list[tuple[int, int]]:
    """(start, length) runs of `new` that are not in `old`, counted from 0."""
    matcher = difflib.SequenceMatcher(None, old, new, autojunk=False)
    return [(j1, j2 - j1) for tag, _i1, _i2, j1, j2 in matcher.get_opcodes()
            if tag in ("insert", "replace") and j2 > j1]


# ------------------------------------------------------------------ state ---

_SEEN: dict[str, tuple[str, object]] = {}   # key -> (host, document)
_WORD: dict[str, dict] = {}
_PPT: dict[str, dict] = {}
_XL_EVENTS: list = []                        # (sheet, address) during a call
_XL_SINKS: dict = {}                         # Excel window handle -> event sink
_NOTES: list[str] = []
_ACTIVE = False
_CODE = ""


def _key(doc) -> str:
    try:
        name = doc.FullName
    except Exception:
        name = None
    if not name:
        name = getattr(doc, "Name", "")
    return str(name).lower()


def see(host, doc):
    """A document the agent reached. Watched from now on; snapshotted at once
    when a call is running, so text written later in the same call counts."""
    if doc is None or host not in ("word", "ppt", "excel"):
        return
    try:
        key = _key(doc)
    except Exception:
        return
    _SEEN[key] = (host, doc)
    if host == "excel":
        try:
            watch_excel(doc.Application)
        except Exception:
            pass
    if _ACTIVE and enabled():
        try:
            _snapshot(host, key, doc)
        except Exception as error:
            _NOTES.append(f"{host} {key}: not watched ({error})")


def watch_excel(app):
    """Listen to SheetChange on this Excel. One listener per running Excel."""
    if not enabled():
        return
    try:
        handle = int(app.Hwnd)
    except Exception:
        return
    if handle in _XL_SINKS:
        return
    import win32com.client
    _XL_SINKS[handle] = win32com.client.DispatchWithEvents(app, _ExcelEvents)


class _ExcelEvents:
    def OnSheetChange(self, sheet, target):
        if _ACTIVE:
            try:
                _XL_EVENTS.append((sheet, target.Address))
            except Exception:
                pass


def _snapshot(host, key, doc):
    if host == "word" and key not in _WORD:
        tracking = bool(doc.TrackRevisions)
        before = Counter(_rev_key(r) for r in _revisions(doc))
        if not tracking:
            doc.TrackRevisions = True
        _WORD[key] = {"doc": doc, "tracking": tracking, "before": before}
    elif host == "ppt" and key not in _PPT:
        _PPT[key] = {"doc": doc,
                     "before": {k: text for k, (_tr, text) in _ppt_ranges(doc).items()}}


def begin(code: str = ""):
    """A call is starting: note what every watched document holds."""
    global _ACTIVE, _CODE
    _reset()
    if not enabled():
        return
    _ACTIVE, _CODE = True, code or ""
    for key, (host, doc) in list(_SEEN.items()):
        try:
            _snapshot(host, key, doc)
        except Exception:
            _SEEN.pop(key, None)                  # closed since, or gone


def end():
    """A call ended (well or not): mark what it added. Never raises.

    Returns {"word": n, "ppt": n, "excel": n, "notes": [...]} with only the
    non-empty entries, or None when nothing was marked."""
    global _ACTIVE
    if not _ACTIVE:
        return None
    counts = Counter()
    try:
        try:
            import pythoncom
            pythoncom.PumpWaitingMessages()       # deliver queued Excel events
        except Exception:
            pass
        for host, states, finish in (("word", _WORD, _end_word),
                                     ("ppt", _PPT, _end_ppt)):
            for key, state in states.items():
                if not _alive(state["doc"]):       # the call closed it
                    _SEEN.pop(key, None)
                    continue
                try:
                    counts[host] += finish(state)
                except Exception as error:
                    _NOTES.append(f"{host} {key}: {error}")
        try:
            counts["excel"] += _end_excel()
        except Exception as error:
            _NOTES.append(f"excel: {error}")
    finally:
        _ACTIVE = False
        notes = list(_NOTES)
        _reset()
    summary = {k: v for k, v in counts.items() if v}
    if notes:
        summary["notes"] = notes
    return summary or None


def _alive(doc) -> bool:
    try:
        doc.Name
        # A closed presentation still answers Name; its slides do not.
        if hasattr(doc, "Slides"):
            doc.Slides.Count
        return True
    except Exception:
        return False


def _reset():
    _WORD.clear()
    _PPT.clear()
    _XL_EVENTS.clear()
    _NOTES.clear()


# ------------------------------------------------------------------- Word ---

def _revisions(doc):
    revs = doc.Revisions
    return [revs.Item(i) for i in range(1, revs.Count + 1)]


def _rev_key(rev):
    # The revision time has whole minutes only, so it cannot tell this call's
    # changes from the user's of a minute ago: compare the sets instead.
    return (int(rev.Type), str(rev.Author), str(rev.Date), rev.Range.Text or "")


def _end_word(state) -> int:
    doc = state["doc"]
    before = Counter(state["before"])
    new = []
    for rev in _revisions(doc):
        key = _rev_key(rev)
        if before[key] > 0:
            before[key] -= 1
        else:
            new.append((key, rev))
    inserts = [rev for key, rev in new if key[0] == WD_REVISION_INSERT]
    keep_tracking = state["tracking"] or bool(_TRACKING_ON.search(_CODE))

    # Formatting applied with tracking off is not a revision of its own; the
    # insertion stays one the user can still accept or reject.
    doc.TrackRevisions = False
    for rev in inserts:
        mark_word(rev.Range)

    if keep_tracking:
        doc.TrackRevisions = True
    else:
        pending = Counter(key for key, _rev in new)
        for rev in reversed(_revisions(doc)):
            key = _rev_key(rev)
            if pending[key] > 0:
                pending[key] -= 1
                rev.Accept()
    return len(inserts)


def mark_word(rng):
    rng.Font.Underline = WD_UNDERLINE_THICK
    rng.Font.UnderlineColor = color_com()


# ------------------------------------------------------------- PowerPoint ---

def _ppt_ranges(pres) -> dict:
    out = {}
    slides = pres.Slides
    for s in range(1, slides.Count + 1):
        slide = slides.Item(s)
        _shape_ranges(slide.Shapes, (int(slide.SlideID),), out)
    return out


def _shape_ranges(shapes, prefix, out):
    for i in range(1, shapes.Count + 1):
        shape = shapes.Item(i)
        key = prefix + (int(shape.Id),)
        try:
            if int(shape.Type) == MSO_GROUP:
                _shape_ranges(shape.GroupItems, key, out)
                continue
        except Exception:
            pass
        try:
            if shape.HasTable:
                table = shape.Table
                for r in range(1, table.Rows.Count + 1):
                    for c in range(1, table.Columns.Count + 1):
                        text_range = table.Cell(r, c).Shape.TextFrame2.TextRange
                        out[key + (r, c)] = (text_range, text_range.Text or "")
                continue
        except Exception:
            pass
        try:
            if shape.HasTextFrame:
                text_range = shape.TextFrame2.TextRange
                out[key] = (text_range, text_range.Text or "")
        except Exception:
            pass


def _characters(text_range, start, length):
    """TextRange2.Characters(start, length). Late-bound COM cannot call a
    property that takes arguments, so fall back to a raw property get."""
    try:
        return text_range.Characters(start, length)
    except Exception:
        import pythoncom
        import win32com.client
        ole = text_range._oleobj_
        dispid = ole.GetIDsOfNames("Characters")
        return win32com.client.Dispatch(
            ole.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True, start, length))


def _end_ppt(state) -> int:
    before = state["before"]
    marked = 0
    for key, (text_range, text) in _ppt_ranges(state["doc"]).items():
        old = before.get(key, "")
        if text == old:
            continue
        for start, length in inserted_spans(old, text):
            mark_ppt(_characters(text_range, start + 1, length))
            marked += 1
    return marked


def mark_ppt(text_range):
    text_range.Font.UnderlineStyle = MSO_UNDERLINE_HEAVY
    text_range.Font.UnderlineColor.RGB = color_com()


# ------------------------------------------------------------------ Excel ---

def _filled(rng):
    """The non-empty cells of `rng`. SpecialCells on a single cell would search
    the whole sheet, so one cell is checked directly."""
    try:
        single = int(rng.Count) == 1
    except Exception:
        single = False
    if single:
        return [rng] if rng.Value not in (None, "") else []
    parts = []
    for kind in (XL_CELL_CONSTANTS, XL_CELL_FORMULAS):
        try:
            parts.append(rng.SpecialCells(kind))
        except Exception:                          # "No cells were found."
            pass
    return parts


def _end_excel() -> int:
    watched = {key for key, (host, _doc) in _SEEN.items() if host == "excel"}
    done, marked = set(), 0
    for sheet, address in _XL_EVENTS:
        try:
            book = _key(sheet.Parent)
            ident = (book, sheet.Name, address)
        except Exception:
            continue
        if book not in watched or ident in done:
            continue
        done.add(ident)
        for part in _filled(sheet.Range(address)):
            part.Font.Underline = XL_UNDERLINE_DOUBLE
            marked += 1
    return marked


# -------------------------------------------------------------- by hand ---

def _has(obj, name) -> bool:
    try:
        getattr(obj, name)
        return True
    except Exception:
        return False


def mark(target) -> str:
    """Mark one range by hand: a Word Range (also an Outlook draft's
    WordEditor range), a PowerPoint TextRange2 or an Excel Range. For text
    the automatic marking cannot see. Returns which kind it marked."""
    if _has(target, "Address"):
        target.Font.Underline = XL_UNDERLINE_DOUBLE
        return "excel"
    if _has(target.Font, "UnderlineStyle"):
        mark_ppt(target)
        return "ppt"
    mark_word(target)
    return "word"
