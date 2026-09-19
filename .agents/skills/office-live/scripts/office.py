"""Drive the live Microsoft Office applications on Windows over COM.

Every command prints exactly one JSON object and exits 0 on success, 1 on
failure:

    {"ok": true,  "data": ...}
    {"ok": false, "error": "...", "hint": "..."}

Design notes for whoever maintains this:

* We only ever ATTACH to a running application (GetActiveObject). Starting
  Office behind the user's back is how an agent ends up editing a document
  nobody is looking at, so --launch has to be asked for explicitly.
* `exec` is deliberately the escape hatch. Five typed commands cover the
  common cases; everything else is COM code written by the caller, with
  `app` and `doc` already in scope. That is far more useful than fifty
  half-typed wrappers.
* Word revision tracking is never turned off here. `track off` exists
  because a human may ask for it, but no write path touches it.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import shutil
import sys

HOSTS = {
    "word": ("Word.Application", "Documents"),
    "excel": ("Excel.Application", "Workbooks"),
    "ppt": ("PowerPoint.Application", "Presentations"),
}

# COM HRESULTs worth translating into something a reader can act on.
_MK_E_UNAVAILABLE = -2147221021
_CALL_REJECTED = -2147418111


class OfficeError(RuntimeError):
    """Raised by the library API; the CLI turns it into a JSON error object."""

    def __init__(self, message, hint=""):
        super().__init__(message)
        self.hint = hint


def out(data):
    print(json.dumps({"ok": True, "data": data}, ensure_ascii=False))
    sys.exit(0)


def fail(error, hint=""):
    print(json.dumps({"ok": False, "error": str(error), "hint": hint},
                     ensure_ascii=False))
    sys.exit(1)


def attach(host, name=None, launch=False):
    """Library entry point. Returns (app, doc) for 'word'/'excel'/'ppt'.

    Import this from inside execute_code when you want to drive COM directly:

        import sys; sys.path.insert(0, r"<skill>/scripts")
        from office import attach
        app, doc = attach("word")

    Raises OfficeError with an actionable .hint on every failure.
    """
    if host not in HOSTS:
        raise OfficeError(f"unknown host {host!r}",
                          f"expected one of {list(HOSTS)}")
    app = get_app(host, launch=launch)
    return app, pick(app, host, name)


def _win32():
    try:
        import win32com.client  # noqa
        return win32com.client
    except ImportError:
        raise OfficeError(
            "pywin32 is not installed",
            "COM needs native Windows Python in the user's interactive "
            "session. If this is WSL or a container, the COM route is not "
            "available at all — say so instead of retrying.")


EARLY_BINDING = None          # None = not attempted yet, True/False = result
_BINDING_NOTE = ""
_CACHE_PURGED = False


def _purge_gen_py(w):
    """Delete the makepy cache so the next EnsureDispatch regenerates it.

    The cache corrupts in practice — a half-written module then raises
    "module ... has no attribute 'CLSIDToClassMap'" forever after. It lives
    under %TEMP% and is pure cache, so deleting it is safe and self-healing.
    """
    import importlib
    import shutil
    shutil.rmtree(w.gencache.GetGeneratePath(), ignore_errors=True)
    importlib.reload(w.gencache)


def _early_bind(w, obj):
    """Re-wrap a dynamic-dispatch COM object with makepy early binding.

    Dynamic dispatch silently MISBINDS named arguments: Find.Execute(Replace=2)
    runs a plain find, returns True, and replaces NOTHING. Early binding maps
    parameter names through the type library, which fixes it.

    Falling back to the dynamic object QUIETLY is how that bug came back after
    the cache corrupted, so we now purge-and-retry once, then record the result
    in EARLY_BINDING for `status` to report loudly.
    """
    global EARLY_BINDING, _BINDING_NOTE, _CACHE_PURGED
    try:
        bound = w.gencache.EnsureDispatch(obj)
        EARLY_BINDING = True
        return bound
    except Exception as first:
        if not _CACHE_PURGED:
            _CACHE_PURGED = True
            try:
                _purge_gen_py(w)
                bound = w.gencache.EnsureDispatch(obj)
                EARLY_BINDING = True
                _BINDING_NOTE = "gen_py cache was corrupt; rebuilt automatically"
                return bound
            except Exception as second:
                first = second
        EARLY_BINDING = False
        _BINDING_NOTE = (
            "early binding UNAVAILABLE (%s: %s). Find.Execute(Replace=2) "
            "silently replaces NOTHING in this mode - verify every write, or "
            "fix pywin32 (its bitness must match Office) before editing."
            % (type(first).__name__, first))
        return obj


def binding_status():
    """What binding mode COM is actually running in, for status/preflight."""
    return {"early_binding": EARLY_BINDING, "note": _BINDING_NOTE}


def get_app(host, launch=False):
    w = _win32()
    prog, _ = HOSTS[host]
    try:
        return _early_bind(w, w.GetActiveObject(prog))
    except OfficeError:
        raise
    except Exception as e:
        code = getattr(e, "hresult", None) or (e.args[0] if e.args else None)
        if launch:
            try:
                app = _early_bind(w, w.Dispatch(prog))
                app.Visible = True
                return app
            except Exception as e2:
                raise OfficeError(f"could not start {prog}: {e2}")
        if code == _MK_E_UNAVAILABLE:
            raise OfficeError(
                f"{prog} is not reachable over COM",
                "Either the app is closed, or it is running with NO document "
                "open — Office only registers itself once a document is "
                "loaded. Ask the user to open the file. Do not pass --launch "
                "on your own initiative.")
        raise OfficeError(f"could not attach to {prog}: {e}")


def collection(app, host):
    return getattr(app, HOSTS[host][1])


def pick(app, host, name=None):
    """Return the named document, or the active one when name is None."""
    coll = collection(app, host)
    if coll.Count == 0:
        raise OfficeError(
            f"{host}: no open document",
            "The application is running but empty. Ask the user which file "
            "to open.")
    if name:
        wanted = name.lower()
        hits = [d for d in coll if wanted in d.Name.lower()]
        if not hits:
            names = [d.Name for d in coll]
            raise OfficeError(f"no open {host} document matches {name!r}",
                              f"open documents: {names}")
        if len(hits) > 1:
            raise OfficeError(f"{name!r} is ambiguous",
                              f"matches: {[d.Name for d in hits]}")
        return hits[0]
    if host == "word":
        return app.ActiveDocument
    if host == "excel":
        return app.ActiveWorkbook
    return app.ActivePresentation


# ------------------------------------------------- discovery & verification ---
#
# These two exist because the agent WRITES THE SCRIPT. A recipe list can only
# cover what we thought of; `explore` lets the model discover the rest of the
# COM surface instead of guessing method names, and `probe` turns "did anything
# actually change?" from a judgement call into a comparison.

def explore(obj, contains=None, limit=60):
    """List the live COM object's methods and property values.

    Early binding (see _early_bind) gives real names from the type library, so
    dir() actually works on these objects. Use it when no recipe covers what
    you need:

        explore(doc, "footnote")     -> everything with 'footnote' in the name
        explore(doc.Paragraphs(1))   -> what a paragraph offers

    Returns {"methods": [...], "properties": {name: value}, "truncated": bool}.
    """
    methods, props = [], {}
    # makepy hands back a COCLASS wrapper (e.g. Document) whose dir() shows only
    # plumbing — CLSID, coclass_interfaces, default_interface. The real member
    # list lives on default_interface (_Document); attribute reads still go
    # through the wrapper, so we take names from one and values from the other.
    surface = getattr(obj, "default_interface", None) or obj
    skip = {"CLSID", "coclass_interfaces", "coclass_sources",
            "default_interface", "default_source"}
    names = [n for n in dir(surface) if not n.startswith("_") and n not in skip]
    # makepy keeps PROPERTIES in _prop_map_get_/_prop_map_put_ rather than as
    # class attributes, so dir() alone shows methods only — Document.Footnotes
    # would be invisible without this.
    for attr in ("_prop_map_get_", "_prop_map_put_"):
        names.extend(getattr(surface, attr, {}) or {})
    names = sorted(set(names))
    if contains:
        needle = contains.lower()
        names = [n for n in names if needle in n.lower()]
    hit_limit = len(names) > limit
    for name in names[:limit]:
        try:
            value = getattr(obj, name)
        except Exception as e:
            props[name] = f"<unreadable: {type(e).__name__}>"
            continue
        if callable(value):
            methods.append(name)
        elif isinstance(value, (str, int, float, bool, type(None))):
            props[name] = value if not isinstance(value, str) else value[:120]
        else:
            props[name] = f"<{type(value).__name__}>"
    return {"methods": sorted(methods), "properties": props,
            "truncated": hit_limit}


def stories(doc):
    """Every text container in a Word document, not just the body.

    Word keeps headers, footers, footnotes, endnotes, comments and text boxes
    in separate "stories". doc.Content covers the body (tables included) and
    NOTHING else, which is how a find-and-replace reports success while leaving
    the header untouched.
    """
    seen = []
    for story in doc.StoryRanges:
        current = story
        while current is not None:
            seen.append(current)
            try:
                current = current.NextStoryRange
            except Exception:
                current = None
    return seen


# Word's Find.Text is capped at 255 characters; longer patterns raise
# "The string ... is too long".
FIND_TEXT_LIMIT = 255

_WD_REVISION_DELETE = 2


def final_text(rng):
    """The story's text as it will read once all revisions are ACCEPTED.

    This is the single most important function here. With Track Changes on, a
    replacement does NOT remove the old text — Word keeps those characters and
    marks them as a deletion revision, so they are still present in
    `Range.Text`. Counting raw text therefore reports before=1, after=1,
    replaced=0 for an edit that plainly worked. Comparing against the accepted
    view is what makes verification honest under tracking.
    """
    text = rng.Text or ""
    base = rng.Start
    cuts = []
    try:
        for rev in rng.Revisions:
            if rev.Type != _WD_REVISION_DELETE:
                continue
            r = rev.Range
            cuts.append((r.Start - base, r.End - base))
    except Exception:
        return text
    if not cuts:
        return text
    kept, pos = [], 0
    for a, b in sorted(cuts):
        a = max(0, min(len(text), a))
        b = max(0, min(len(text), b))
        if a > pos:
            kept.append(text[pos:a])
        pos = max(pos, b)
    kept.append(text[pos:])
    return "".join(kept)


def doc_kind(doc):
    """'word' | 'excel' | 'ppt' for any open document object."""
    if hasattr(doc, "StoryRanges"):
        return "word"
    if hasattr(doc, "Worksheets"):
        return "excel"
    if hasattr(doc, "Slides"):
        return "ppt"
    return "unknown"


def shape_textrange(shape):
    """A shape's text object, whichever application owns it.

    PowerPoint shapes expose TextFrame.TextRange; Excel's TextFrame has
    .Characters instead and only TextFrame2 carries a TextRange. TextFrame2
    exists in all three apps, so try it first and fall back.
    """
    for getter in (lambda sh: sh.TextFrame2.TextRange,
                   lambda sh: sh.TextFrame.TextRange):
        try:
            text_range = getter(shape)
            _ = text_range.Text          # force a real COM read
            return text_range
        except Exception:
            continue
    return None


def _shape_texts(shapes, out):
    """Collect text from a Shapes collection, descending into groups/tables."""
    for i in range(1, shapes.Count + 1):
        try:
            shape = shapes(i)
        except Exception:
            continue
        try:
            if shape.Type == 6:                      # msoGroup
                _shape_texts(shape.GroupItems, out)
                continue
        except Exception:
            pass
        try:
            if shape.HasTable:
                table = shape.Table
                for r in range(1, table.Rows.Count + 1):
                    for c in range(1, table.Columns.Count + 1):
                        cell = shape_textrange(table.Cell(r, c).Shape)
                        if cell is not None:
                            out.append(cell.Text or "")
                continue
        except Exception:
            pass
        text_range = shape_textrange(shape)
        if text_range is not None:
            try:
                out.append(text_range.Text or "")
            except Exception:
                continue


def text_chunks(doc, accepted=True):
    """Every piece of text in the document, whatever the application.

    Word    - one chunk per story (body, headers, footers, footnotes, ...).
              `accepted` controls whether text marked deleted by Track Changes
              is excluded; see final_text().
    Excel   - used-range cell values plus shape text, per worksheet.
    PPT     - shape text, table cells and speaker notes, per slide.

    Only Word has revision tracking, so `accepted` is a no-op elsewhere.
    """
    kind = doc_kind(doc)
    chunks = []
    if kind == "word":
        for story in stories(doc):
            try:
                chunks.append(final_text(story) if accepted else (story.Text or ""))
            except Exception:
                continue
        return chunks
    if kind == "excel":
        for ws in doc.Worksheets:
            try:
                values = ws.UsedRange.Value
            except Exception:
                values = None
            if isinstance(values, tuple):
                for row in values:
                    cells = row if isinstance(row, tuple) else (row,)
                    for cell in cells:
                        if cell is not None:
                            chunks.append(str(cell))
            elif values is not None:
                chunks.append(str(values))
            try:
                _shape_texts(ws.Shapes, chunks)
            except Exception:
                pass
        return chunks
    if kind == "ppt":
        for slide in doc.Slides:
            try:
                _shape_texts(slide.Shapes, chunks)
            except Exception:
                pass
            try:
                _shape_texts(slide.NotesPage.Shapes, chunks)
            except Exception:
                pass
        return chunks
    return chunks


def count_text(doc, needle, match_case=True, accepted=True):
    """Occurrences of `needle` anywhere in the document, in any application.

    accepted=True  -> count the text as it will read once Word revisions are
                      accepted (what the reader actually sees).
    accepted=False -> count raw characters, including Word's tracked deletions.
    """
    total = 0
    for text in text_chunks(doc, accepted=accepted):
        total += (text.count(needle) if match_case
                  else text.lower().count(needle.lower()))
    return total


def _replace_word(doc, old, new, match_case):
    long_pattern = len(old) > FIND_TEXT_LIMIT
    for story in stories(doc):
        try:
            if long_pattern:
                _replace_long(doc, story, old, new, match_case)
                continue
            find = story.Find
            find.ClearFormatting()
            find.Replacement.ClearFormatting()
            find.Text = old
            find.Replacement.Text = new
            find.Forward = True
            find.Wrap = 0                # wdFindStop: stay inside this story
            find.MatchCase = match_case
            find.MatchWildcards = False
            find.Execute(Replace=2)      # 2 = wdReplaceAll; NAME the argument
        except Exception:
            continue                     # protected or empty story
    return long_pattern


def _replace_excel(doc, old, new, match_case):
    """Cells via Excel's own Replace, plus shape text by hand."""
    for ws in doc.Worksheets:
        try:
            ws.Cells.Replace(What=old, Replacement=new, LookAt=2,
                             MatchCase=match_case)      # 2 = xlPart
        except Exception:
            pass
        try:
            _replace_shapes(ws.Shapes, old, new, match_case)
        except Exception:
            pass
    return False


def _replace_shapes(shapes, old, new, match_case):
    for i in range(1, shapes.Count + 1):
        try:
            shape = shapes(i)
        except Exception:
            continue
        try:
            if shape.Type == 6:                        # msoGroup
                _replace_shapes(shape.GroupItems, old, new, match_case)
                continue
        except Exception:
            pass
        try:
            if shape.HasTable:
                table = shape.Table
                for r in range(1, table.Rows.Count + 1):
                    for c in range(1, table.Columns.Count + 1):
                        cell = shape_textrange(table.Cell(r, c).Shape)
                        if cell is not None:
                            _replace_textrange(cell, old, new, match_case)
                continue
        except Exception:
            pass
        text_range = shape_textrange(shape)
        if text_range is not None:
            _replace_textrange(text_range, old, new, match_case)


def _replace_textrange(text_range, old, new, match_case):
    """Replace inside one shape's text.

    PowerPoint's Replace() handles a SINGLE hit and returns it, so it has to be
    looped (capped, in case the replacement contains the search text). Excel's
    TextRange2 has no usable Replace, so fall back to rewriting the whole
    string — shape text is short, and this keeps one code path for every app.
    """
    try:
        for _ in range(500):
            found = text_range.Replace(FindWhat=old, ReplaceWhat=new,
                                       MatchCase=match_case, WholeWords=False)
            if found is None:
                return
        return
    except Exception:
        pass
    try:
        text = text_range.Text or ""
    except Exception:
        return
    if match_case:
        if old in text:
            text_range.Text = text.replace(old, new)
        return
    lowered, needle = text.lower(), old.lower()
    if needle not in lowered:
        return
    rebuilt, at = [], 0
    while True:
        hit = lowered.find(needle, at)
        if hit == -1:
            rebuilt.append(text[at:])
            break
        rebuilt.append(text[at:hit])
        rebuilt.append(new)
        at = hit + len(old)
    text_range.Text = "".join(rebuilt)


def _replace_ppt(doc, old, new, match_case):
    for slide in doc.Slides:
        try:
            _replace_shapes(slide.Shapes, old, new, match_case)
        except Exception:
            pass
        try:
            _replace_shapes(slide.NotesPage.Shapes, old, new, match_case)
        except Exception:
            pass
    return False


def replace_all(doc, pairs, match_case=True):
    """Replace every (old, new) pair everywhere in the document, then prove it.

    Works on Word, Excel and PowerPoint:

        Word   - every story: body, tables, headers, footers, footnotes,
                 endnotes, comments, text boxes. With Track Changes on, each
                 hit lands as its own reviewable revision.
        Excel  - every worksheet's cells, plus shape and table text.
        PPT    - every slide's shapes, tables and speaker notes.

    Each row reports, separately:

        visible_before / visible_after  occurrences in the ACCEPTED text
        replaced                        how many actually changed
        in_revisions                    old copies kept as Word tracked
                                        deletions (always 0 outside Word)
        inserted                        occurrences of the new text
        complete                        nothing left visible

    `complete` is never True while the old text is still visible, which is the
    guarantee raw before/after counting could not give under Track Changes.
    """
    kind = doc_kind(doc)
    if kind == "unknown":
        raise OfficeError("replace_all does not know this document type",
                          "expected a Word document, Excel workbook or "
                          "PowerPoint presentation")
    rows = []
    for old, new in pairs:
        visible_before = count_text(doc, old, match_case, accepted=True)
        if kind == "word":
            long_pattern = _replace_word(doc, old, new, match_case)
        elif kind == "excel":
            long_pattern = _replace_excel(doc, old, new, match_case)
        else:
            long_pattern = _replace_ppt(doc, old, new, match_case)
        visible_after = count_text(doc, old, match_case, accepted=True)
        raw_after = count_text(doc, old, match_case, accepted=False)
        rows.append({
            "host": kind,
            "old": old,
            "new": new,
            "visible_before": visible_before,
            "visible_after": visible_after,
            "replaced": visible_before - visible_after,
            "in_revisions": max(0, raw_after - visible_after),
            "inserted": count_text(doc, new, match_case, accepted=True),
            "found": visible_before > 0,
            "long_pattern": long_pattern,
            "complete": visible_after == 0,
        })
    return rows


def probe(doc):
    """Fingerprint a document so you can prove an edit landed.

        before = probe(doc)
        ...make the edit...
        after = probe(doc)
        print(before, after)      # if they are equal, NOTHING happened

    Catches the silent no-op class of bug (a COM call returning True while
    changing nothing) without you having to eyeball the whole document.
    Auto-detects Word / Excel / PowerPoint.
    """
    import hashlib

    def digest(text):
        return hashlib.sha1(text.encode("utf-8", "replace")).hexdigest()[:12]

    if hasattr(doc, "Paragraphs") and hasattr(doc, "Revisions"):
        body = doc.Content.Text
        # Hash EVERY story: a header or footnote edit leaves the body hash
        # unchanged, which made probe() blind to exactly the edits that are
        # hardest to eyeball.
        parts = []
        for story in stories(doc):
            try:
                parts.append(final_text(story))
            except Exception:
                pass
        joined = chr(31).join(parts)
        return {"host": "word", "name": doc.Name,
                "paragraphs": doc.Paragraphs.Count,
                "revisions": doc.Revisions.Count, "chars": len(body),
                "stories": len(parts), "sha": digest(joined),
                "body_sha": digest(body), "tail": body[-80:]}
    if hasattr(doc, "Worksheets"):
        parts, cells = [], 0
        for ws in doc.Worksheets:
            try:
                used = ws.UsedRange
                parts.append(f"{ws.Name}:{used.Address}")
                cells += used.Count
            except Exception:
                parts.append(f"{ws.Name}:?")
        values = []
        try:
            for ws in doc.Worksheets:
                values.append(str(ws.UsedRange.Value))
        except Exception:
            pass
        return {"host": "excel", "name": doc.Name, "sheets": len(parts),
                "used": parts, "cells": cells, "sha": digest("".join(values))}
    if hasattr(doc, "Slides"):
        titles = []
        for sl in doc.Slides:
            try:
                titles.append(sl.Shapes.Title.TextFrame.TextRange.Text
                              if sl.Shapes.HasTitle else "")
            except Exception:
                titles.append("?")
        return {"host": "ppt", "name": doc.Name, "slides": doc.Slides.Count,
                "titles": titles, "sha": digest("|".join(titles))}
    return {"host": "unknown", "repr": repr(doc)[:120]}


# ----------------------------------------------------------------- status ---

def recovery_warning(name, path):
    """Detect an AutoRecovery / autosaved copy masquerading as the real file.

    After a crash, sleep, or a lost network share, Office reopens documents as
    recovery copies: a workbook becomes "<name> (version 1).xlsb" under the
    per-user Microsoft/Excel folder, and PowerPoint's FullName stops being a
    path at all. They look normal in the UI, so an agent will happily edit the
    copy and report success while the user's real file stays untouched.
    Observed on this machine after an overnight sleep.
    """
    low_name, low_path = (name or "").lower(), (path or "").lower()
    markers = ("autorecovery", "autorecovered", "automatikusan",
               "(version 1)", "[recovered]", "helyreall")
    if any(m in low_name or m in low_path for m in markers):
        return "this is an Office AutoRecovery copy, NOT the user's file"
    for folder in ("/microsoft/excel", "/microsoft/word", "/microsoft/powerpoint",
                   "/temp/", "/appdata/"):
        if folder in low_path.replace("\\", "/"):
            return ("the file sits in an Office recovery/temp folder, "
                    "not its real location")
    if path and not low_path.startswith("http") and ":" not in path[:3]:
        return "FullName is not a real path - the document is unsaved or recovered"
    return None


def cmd_status(a):
    w = _win32()
    report = {}
    # Touch COM once so EARLY_BINDING is populated before we report it.
    for host in HOSTS:
        try:
            get_app(host)
            break
        except Exception:
            continue
    report["com"] = binding_status()
    for host, (prog, coll_name) in HOSTS.items():
        entry = {"running": False, "documents": []}
        try:
            app = w.GetActiveObject(prog)
        except Exception:
            report[host] = entry
            continue
        entry["running"] = True
        try:
            for d in getattr(app, coll_name):
                item = {
                    "name": d.Name,
                    "path": getattr(d, "FullName", ""),
                    "saved": bool(d.Saved),
                }
                if host == "word":
                    item["track_revisions"] = bool(d.TrackRevisions)
                    item["revisions"] = d.Revisions.Count
                # Say up front whether the safety protocol's backup step is
                # even possible: a SharePoint/OneDrive document has no local
                # file to copy, so version history is the only net.
                if item["path"].lower().startswith("http"):
                    item["storage"] = "sharepoint"
                    item["backup_possible"] = False
                    item["backup_note"] = ("no local copy exists; SharePoint "
                                           "version history is the only undo")
                else:
                    item["storage"] = "local"
                    item["backup_possible"] = bool(item["path"])
                warning = recovery_warning(item["name"], item["path"])
                if warning:
                    item["WARNING"] = warning
                    item["do_not_edit_until_confirmed"] = True
                entry["documents"].append(item)
        except Exception as e:
            entry["error"] = str(e)
        report[host] = entry
    out(report)


# ------------------------------------------------------------------ word ---

def cmd_word_structure(a):
    doc = pick(get_app("word"), "word", a.name)
    headings = []
    for i, p in enumerate(doc.Paragraphs, 1):
        style = ""
        try:
            style = str(p.Style.NameLocal)
        except Exception:
            pass
        if "eading" in style or "ímsor" in style or "Címsor" in style:
            headings.append({
                "paragraph": i,
                "style": style,
                "text": p.Range.Text.strip(),
            })
    out({
        "name": doc.Name,
        "paragraphs": doc.Paragraphs.Count,
        "words": doc.Words.Count,
        "tables": doc.Tables.Count,
        "track_revisions": bool(doc.TrackRevisions),
        "revisions": doc.Revisions.Count,
        "saved": bool(doc.Saved),
        "headings": headings,
    })


def cmd_word_text(a):
    doc = pick(get_app("word"), "word", a.name)
    text = doc.Content.Text
    chunk = text[a.offset:a.offset + a.max_chars]
    out({
        "name": doc.Name,
        "offset": a.offset,
        "returned_chars": len(chunk),
        "total_chars": len(text),
        "truncated": a.offset + len(chunk) < len(text),
        "text": chunk,
    })


def cmd_word_selection(a):
    app = get_app("word")
    sel = app.Selection
    out({
        "text": sel.Text,
        "start": sel.Start,
        "end": sel.End,
        "paragraph_index": sel.Range.Paragraphs(1).Range.Start,
        "document": app.ActiveDocument.Name,
    })


def cmd_word_track(a):
    doc = pick(get_app("word"), "word", a.name)
    if a.state in ("on", "off"):
        doc.TrackRevisions = (a.state == "on")
    out({
        "name": doc.Name,
        "track_revisions": bool(doc.TrackRevisions),
        "revisions": doc.Revisions.Count,
    })


def cmd_word_revisions(a):
    doc = pick(get_app("word"), "word", a.name)
    kinds = {1: "insert", 2: "delete", 3: "property"}
    revs = []
    for i, r in enumerate(doc.Revisions, 1):
        try:
            text = r.Range.Text
        except Exception:
            text = None
        revs.append({
            "index": i,
            "type": kinds.get(r.Type, r.Type),
            "author": r.Author,
            "text": text,
        })
    out({"name": doc.Name, "count": len(revs), "revisions": revs})


# ----------------------------------------------------------------- excel ---

def cmd_excel_sheets(a):
    wb = pick(get_app("excel"), "excel", a.name)
    sheets = []
    for ws in wb.Worksheets:
        try:
            used = str(ws.UsedRange.Address)
        except Exception:
            used = ""
        sheets.append({
            "name": ws.Name,
            "used_range": used,
            "visible": ws.Visible == -1,
        })
    out({"name": wb.Name, "path": wb.FullName, "saved": bool(wb.Saved),
         "sheets": sheets})


def _resolve_range(wb, reference):
    """'Sheet1!A1:D5' or 'Sheet1!used' or plain 'A1:D5' on the active sheet."""
    if "!" in reference:
        sheet_name, addr = reference.split("!", 1)
        ws = wb.Worksheets(sheet_name)
    else:
        ws, addr = wb.ActiveSheet, reference
    if addr.lower() == "used":
        return ws, ws.UsedRange
    return ws, ws.Range(addr)


def cmd_excel_read(a):
    wb = pick(get_app("excel"), "excel", a.name)
    ws, rng = _resolve_range(wb, a.range)
    values = rng.Formula if a.formulas else rng.Value
    if not isinstance(values, tuple):
        values = ((values,),)
    grid = [[c for c in row] if isinstance(row, tuple) else [row]
            for row in values]
    out({
        "workbook": wb.Name,
        "sheet": ws.Name,
        "address": str(rng.Address),
        "formulas": bool(a.formulas),
        "values": grid,
    })


def cmd_excel_write(a):
    wb = pick(get_app("excel"), "excel", a.name)
    ws, rng = _resolve_range(wb, a.range)
    try:
        values = json.loads(a.values)
    except json.JSONDecodeError as e:
        fail(f"--values is not valid JSON: {e}",
             'It must be a 2D array, even for one cell: [["x"]]')
    if not isinstance(values, list) or not values or \
            not all(isinstance(r, list) for r in values):
        fail("--values must be a 2D JSON array",
             'One cell is [["x"]], one row is [["a","b"]].')

    previous = rng.Value
    if not isinstance(previous, tuple):
        previous = ((previous,),)
    prev_grid = [[c for c in row] if isinstance(row, tuple) else [row]
                 for row in previous]

    target = ws.Range(
        rng.Cells(1, 1),
        rng.Cells(len(values), max(len(r) for r in values)),
    )
    if a.formulas:
        target.Formula = values
    else:
        target.Value = values

    out({
        "workbook": wb.Name,
        "sheet": ws.Name,
        "address": str(target.Address),
        "written_rows": len(values),
        "previous_values": prev_grid,
    })


# ------------------------------------------------------------ powerpoint ---

def cmd_ppt_outline(a):
    pres = pick(get_app("ppt"), "ppt", a.name)
    slides = []
    for s in pres.Slides:
        shapes = []
        for sh in s.Shapes:
            try:
                if sh.HasTextFrame and sh.TextFrame.HasText:
                    shapes.append({
                        "name": sh.Name,
                        "text": sh.TextFrame.TextRange.Text,
                    })
            except Exception:
                pass
        notes = ""
        try:
            notes = s.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        except Exception:
            pass
        slides.append({"index": s.SlideIndex, "shapes": shapes,
                       "notes": notes})
    out({"name": pres.Name, "path": pres.FullName, "saved": bool(pres.Saved),
         "slide_count": pres.Slides.Count, "slides": slides})


# ---------------------------------------------------------------- backup ---

def cmd_backup(a):
    doc = pick(get_app(a.host), a.host, a.name)
    path = doc.FullName
    # Order matters: a SharePoint URL is neither absolute nor local, so it has
    # to be recognised before the "never saved" check claims it.
    if path.lower().startswith("http"):
        fail("the document lives on SharePoint/OneDrive, not a local path",
             f"FullName is {path!r}. A file copy is not possible; rely on "
             "version history there, or ask the user to save a local copy.")
    if not path or not os.path.splitext(path)[1] or not os.path.isabs(path):
        fail("this document has never been saved to disk",
             "Ask the user to save it once first. Do not edit an "
             "unbacked-up document.")
    if not doc.Saved:
        doc.Save()
    stamp = _dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    root, ext = os.path.splitext(path)
    dest = f"{root}.bak-{stamp}{ext}"
    shutil.copy2(path, dest)
    out({"document": doc.Name, "source": path, "backup": dest})


# ------------------------------------------------------------------ exec ---

def cmd_exec(a):
    if bool(a.code) == bool(a.code_file):
        fail("give exactly one of --code or --code-file")
    if a.code_file:
        with open(a.code_file, "r", encoding="utf-8") as fh:
            source = fh.read()
    else:
        source = a.code

    if a.warm:
        # Warm path: the daemon owns the namespace, so `attach` is preloaded
        # and any variable defined here survives into the next call.
        resp = daemon_request({"op": "exec", "code": source})
        print(json.dumps(resp, ensure_ascii=False))
        sys.exit(0 if resp.get("ok") else 1)

    refusal = refuse(source)
    if refusal:
        print(json.dumps(refusal, ensure_ascii=False))
        sys.exit(1)

    app = get_app(a.host, launch=a.launch)
    doc = None
    try:
        doc = pick(app, a.host, a.name) if collection(app, a.host).Count else None
    except SystemExit:
        raise
    except Exception:
        doc = None

    scope = {"app": app, "doc": doc, "json": json, "os": os}
    try:
        try:
            value = eval(compile(source, "<code>", "eval"), scope)
        except SyntaxError:
            exec(compile(source, "<code>", "exec"), scope)
            value = scope.get("result")
    except Exception as e:
        code = getattr(e, "hresult", None) or (e.args[0] if e.args else None)
        if code == _CALL_REJECTED:
            fail(f"Office rejected the call: {e}",
                 "Usually a modal dialog is open in the application. Ask the "
                 "user to dismiss it. Do not retry blindly.")
        fail(f"{type(e).__name__}: {e}")

    try:
        json.dumps(value)
    except TypeError:
        value = str(value)
    out({"host": a.host, "document": doc.Name if doc else None,
         "result": value})


# ------------------------------------------------- warm daemon (optional) ---
#
# Cold `exec` pays ~250 ms of interpreter + COM attach per call and keeps no
# state. For long multi-step sessions — or for agents that can only shell out —
# `exec --warm` talks to a background daemon that holds the COM objects and a
# persistent namespace, so variables survive between calls like a notebook.
# The daemon exits by itself after 45 idle minutes.
#
# Transport: 127.0.0.1 TCP on a random port + a per-user token, both written to
# %LOCALAPPDATA%\office-live\daemon.json. Newline-delimited JSON frames.

RUNTIME_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                           "office-live")
INFO_PATH = os.path.join(RUNTIME_DIR, "daemon.json")
IDLE_EXIT_SECONDS = 45 * 60

# Calls that destroy the human review trail or the user's work. Enforced in the
# daemon so every client inherits them, not just well-behaved ones.
FORBIDDEN = [
    (r"\bAcceptAll(Revisions)?\b", "accepting revisions erases the review trail"),
    (r"\bRejectAll(Revisions)?\b", "rejecting revisions erases the review trail"),
    (r"\.Accept\(\)", "accepting a revision is the human reviewer's call"),
    (r"\.Reject\(\)", "rejecting a revision is the human reviewer's call"),
    (r"\bSaveAs2?\b", "SaveAs can silently overwrite another file"),
    (r"\.Close\(", "closing documents is the user's call"),
    (r"\bQuit\(", "quitting the application is the user's call"),
    (r"TrackRevisions\s*=\s*(False|0)", "never turn revision tracking off"),
    (r"DisplayAlerts\s*=", "suppressing Office alerts hides destructive prompts"),
]


def refuse(code):
    """Return a refusal dict if `code` contains a forbidden call, else None."""
    import re
    for pattern, why in FORBIDDEN:
        if re.search(pattern, code):
            return {"ok": False, "error": f"refused: {why}",
                    "hint": "this tool only ADDS tracked edits; leave review "
                            "decisions and file lifecycle to the user"}
    return None


def run_snippet(code, ns):
    """Exec `code` in `ns`, notebook-style: statements run, and if the last
    node is an expression its value becomes the result."""
    import ast
    import contextlib
    import io

    stdout = io.StringIO()
    tree = ast.parse(code, mode="exec")
    last_expr = None
    if tree.body and isinstance(tree.body[-1], ast.Expr):
        last_expr = ast.Expression(tree.body.pop(-1).value)
    with contextlib.redirect_stdout(stdout):
        exec(compile(tree, "<office>", "exec"), ns)
        value = (eval(compile(last_expr, "<office>", "eval"), ns)
                 if last_expr is not None else ns.get("result"))
    try:
        json.dumps(value)
    except TypeError:
        value = repr(value)
    return {"ok": True, "result": value, "stdout": stdout.getvalue()[-4000:]}


def _seed(ns):
    ns.clear()
    ns.update({"attach": attach, "OfficeError": OfficeError, "json": json,
               "explore": explore, "probe": probe, "replace_all": replace_all,
               "count_text": count_text, "stories": stories,
               "final_text": final_text, "text_chunks": text_chunks,
               "doc_kind": doc_kind})


def _describe_exception(e):
    hint = e.hint if isinstance(e, OfficeError) else ""
    if getattr(e, "hresult", None) == _CALL_REJECTED:
        hint = ("Office rejected the call — usually a modal dialog is open; "
                "ask the user to dismiss it, do not retry blindly")
    return {"ok": False, "error": f"{type(e).__name__}: {e}", "hint": hint}


_SOURCE_MTIME = None


def _source_changed():
    """True when office.py on disk is newer than the copy this daemon loaded.

    A long-lived daemon otherwise keeps serving the code it started with, so a
    skill update appears to do nothing — the exact trap that hid a COM fix for
    a month. We step aside instead; the client respawns us automatically.
    """
    if _SOURCE_MTIME is None:
        return False
    try:
        return os.path.getmtime(os.path.abspath(__file__)) != _SOURCE_MTIME
    except OSError:
        return False


def _handle(req, ns):
    op = req.get("op")
    if op != "ping" and _source_changed():
        return {"ok": False, "_stale": True,
                "error": "office.py changed on disk; restarting the daemon",
                "_stop": True}
    if op == "ping":
        return {"ok": True, "result": "pong"}
    if op == "reset":
        # Restart rather than just clearing variables, so the fresh daemon
        # also picks up any edits to office.py.
        return {"ok": True, "_stop": True,
                "result": "daemon restarting; the next call loads the current "
                          "office.py and a clean namespace"}
    if op == "stop":
        return {"ok": True, "result": "stopping", "_stop": True}
    if op != "exec":
        return {"ok": False, "error": f"unknown op {op!r}"}
    code = req.get("code", "")
    refusal = refuse(code)
    if refusal:
        return refusal
    try:
        return run_snippet(code, ns)
    except SyntaxError as e:
        return {"ok": False, "error": f"SyntaxError: {e}",
                "hint": "send plain Python; the last expression's value is "
                        "returned automatically"}
    except Exception as e:
        return _describe_exception(e)


def daemon_main():
    import contextlib
    import secrets
    import socket
    import time

    global _SOURCE_MTIME
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import pythoncom
    pythoncom.CoInitialize()

    try:
        _SOURCE_MTIME = os.path.getmtime(os.path.abspath(__file__))
    except OSError:
        _SOURCE_MTIME = None

    ns = {}
    _seed(ns)

    srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    srv.bind(("127.0.0.1", 0))
    srv.listen(1)
    srv.settimeout(60)
    port = srv.getsockname()[1]
    token = secrets.token_hex(16)

    os.makedirs(RUNTIME_DIR, exist_ok=True)
    with open(INFO_PATH, "w", encoding="utf-8") as fh:
        json.dump({"port": port, "token": token, "pid": os.getpid()}, fh)

    last = time.time()
    try:
        while True:
            try:
                conn, _ = srv.accept()
            except socket.timeout:
                if time.time() - last > IDLE_EXIT_SECONDS:
                    return
                continue
            last = time.time()
            with conn:
                try:
                    buf = b""
                    while not buf.endswith(b"\n"):
                        chunk = conn.recv(65536)
                        if not chunk:
                            break
                        buf += chunk
                    req = json.loads(buf.decode("utf-8"))
                    resp = ({"ok": False, "error": "bad token"}
                            if req.get("token") != token else _handle(req, ns))
                    conn.sendall((json.dumps(resp, ensure_ascii=False) + "\n")
                                 .encode("utf-8"))
                    if resp.get("_stop"):
                        return
                except Exception as e:
                    with contextlib.suppress(Exception):
                        conn.sendall((json.dumps({"ok": False, "error": str(e)})
                                      + "\n").encode("utf-8"))
    finally:
        with contextlib.suppress(OSError):
            os.remove(INFO_PATH)


def _try_request(payload, timeout=120.0):
    import socket
    with open(INFO_PATH, encoding="utf-8") as fh:
        info = json.load(fh)
    payload = dict(payload, token=info["token"])
    with socket.create_connection(("127.0.0.1", info["port"]), timeout=5) as c:
        c.settimeout(timeout)
        c.sendall((json.dumps(payload) + "\n").encode("utf-8"))
        buf = b""
        while not buf.endswith(b"\n"):
            chunk = c.recv(65536)
            if not chunk:
                break
            buf += chunk
    return json.loads(buf.decode("utf-8"))


def daemon_request(payload, _retried=False):
    """Send one request, auto-spawning the daemon if it is not running.

    If the daemon reports its code is stale it shuts down; we respawn and
    resend once, so a skill update takes effect without the caller noticing.
    """
    import subprocess
    import time

    try:
        resp = _try_request(payload)
        if resp.get("_stale") and not _retried:
            time.sleep(0.3)                    # let it close the socket
            return daemon_request(payload, _retried=True)
        return resp
    except Exception:
        pass                                   # not running, or stale info file
    flags = 0x08000000 | 0x00000200            # CREATE_NO_WINDOW | NEW_PROCESS_GROUP
    subprocess.Popen([sys.executable, os.path.abspath(__file__), "daemon"],
                     creationflags=flags, close_fds=True,
                     stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                     stderr=subprocess.DEVNULL)
    for _ in range(50):
        time.sleep(0.1)
        try:
            if _try_request({"op": "ping"}).get("ok"):
                break
        except Exception:
            continue
    else:
        return {"ok": False, "error": "could not start the office daemon",
                "hint": "run 'office.py daemon' in a terminal to see why"}
    return _try_request(payload)


def cmd_daemon(a):
    daemon_main()


def cmd_daemon_op(a):
    resp = daemon_request({"op": a.cmd})
    if a.cmd == "reset" and resp.get("ok"):
        # Bring the replacement up now so the next real call is warm.
        daemon_request({"op": "ping"})
    print(json.dumps(resp, ensure_ascii=False))
    sys.exit(0 if resp.get("ok") else 1)


# ------------------------------------------------------------------- cli ---

def build_parser():
    p = argparse.ArgumentParser(
        prog="office.py",
        description="Read and edit the Office documents open on this machine.")
    sub = p.add_subparsers(dest="cmd", required=True)

    def with_name(sp):
        sp.add_argument("--name", help="document name (substring); "
                                       "default = the active document")
        return sp

    with_name(sub.add_parser("status", help="what is running and open"))
    sub.choices["status"].set_defaults(func=cmd_status)

    word = sub.add_parser("word").add_subparsers(dest="sub", required=True)
    with_name(word.add_parser("structure")).set_defaults(func=cmd_word_structure)
    t = with_name(word.add_parser("text"))
    t.add_argument("--max-chars", type=int, default=20000)
    t.add_argument("--offset", type=int, default=0)
    t.set_defaults(func=cmd_word_text)
    with_name(word.add_parser("selection")).set_defaults(func=cmd_word_selection)
    tr = with_name(word.add_parser("track"))
    tr.add_argument("state", choices=["on", "off", "status"], nargs="?",
                    default="status")
    tr.set_defaults(func=cmd_word_track)
    with_name(word.add_parser("revisions")).set_defaults(func=cmd_word_revisions)

    excel = sub.add_parser("excel").add_subparsers(dest="sub", required=True)
    with_name(excel.add_parser("sheets")).set_defaults(func=cmd_excel_sheets)
    r = with_name(excel.add_parser("read"))
    r.add_argument("--range", required=True,
                   help="Sheet1!A1:D50, Sheet1!used, or A1:D50")
    r.add_argument("--formulas", action="store_true")
    r.set_defaults(func=cmd_excel_read)
    wr = with_name(excel.add_parser("write"))
    wr.add_argument("--range", required=True)
    wr.add_argument("--values", required=True, help='2D JSON array: [[1,2]]')
    wr.add_argument("--formulas", action="store_true")
    wr.set_defaults(func=cmd_excel_write)

    ppt = sub.add_parser("ppt").add_subparsers(dest="sub", required=True)
    with_name(ppt.add_parser("outline")).set_defaults(func=cmd_ppt_outline)

    b = with_name(sub.add_parser("backup"))
    b.add_argument("--host", required=True, choices=list(HOSTS))
    b.set_defaults(func=cmd_backup)

    e = with_name(sub.add_parser("exec"))
    e.add_argument("--host", default="word", choices=list(HOSTS))
    e.add_argument("--code")
    e.add_argument("--code-file")
    e.add_argument("--warm", action="store_true",
                   help="run in the persistent daemon: variables survive "
                        "between calls, ~250ms instead of a cold attach")
    e.add_argument("--launch", action="store_true",
                   help="start the app if it is not running (ask first)")
    e.set_defaults(func=cmd_exec)

    sub.add_parser("daemon", help="run the warm daemon in the foreground "
                                  "(normally auto-spawned)").set_defaults(
        func=cmd_daemon)
    for op, helptext in (("reset", "drop the daemon's variables / stale COM refs"),
                         ("stop", "shut the daemon down")):
        sub.add_parser(op, help=helptext).set_defaults(func=cmd_daemon_op)

    return p


def main():
    args = build_parser().parse_args()
    try:
        args.func(args)
    except SystemExit:
        raise
    except OfficeError as e:
        fail(e, e.hint)
    except Exception as e:
        fail(f"{type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
