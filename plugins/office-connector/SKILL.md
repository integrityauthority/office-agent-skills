---
name: office-live
description: Work live in the user's running Microsoft Office on Windows over COM — read and edit the OPEN Word, Excel or PowerPoint document, and in classic Outlook search mail, read messages, compose new mail, replies and forwards with attachments and formatting, look up contacts and the calendar, and prepare appointments. Use when the user refers to "this document", "the open file", "my selection", "my inbox", "reply to that email", "my calendar", or asks for work in place rather than as a new file. Outlook items are composed and opened for the user to review — this skill never sends or saves mail. Do NOT use for generating a document from scratch (write a .docx/.xlsx/.pptx file instead), and do NOT use on non-Windows machines.
version: 2.1.0
author: Integritás Hatóság
platforms: [windows]
---

# office-live

Drives Word, Excel and PowerPoint **while the user has them open**, through one
file: `${HERMES_SKILL_DIR}/scripts/office.py`. Edits appear in their window
immediately. Word edits land as tracked revisions they can accept or reject.

Everything runs on the interpreter Hermes already uses — pywin32 is installed
there, so there is nothing to set up.

## How to call it

**If the `office_exec` tool exists, use it — do not write your own
`win32com` code.** It runs the snippet against one COM connection kept warm
for the whole session, and the Python namespace **persists between calls**, so
you attach once and keep the variables:

```
office_status()                      # what is open; call this first
office_exec("app, doc = attach('word')")
office_exec("doc.TrackRevisions = True; doc.Paragraphs.Count")
office_exec("doc.Content.Text[-300:]")      # verify by reading back
office_reset()                       # only after Office was closed and reopened
```

`office_exec` returns the value of the last expression, plus anything you
`print()`. Everything in the recipes below works inside it unchanged.

**Only if those tools are absent** — this skill also ships standalone, without
the plugin — write Python in `execute_code` and import the module instead.
Same functions, but the namespace does not survive between calls, so do the
whole read → modify → verify in one snippet:

```python
import sys
sys.path.insert(0, r"${HERMES_SKILL_DIR}/scripts")
from office import attach, OfficeError

app, doc = attach("word")          # "word" | "excel" | "ppt"
                                   # attach("excel", name="jelentes") picks by name
doc.TrackRevisions = True
f = doc.Content.Find
f.Text, f.Replacement.Text = "régi szöveg", "új szöveg"
f.Execute(Replace=2)               # 2 = wdReplaceAll
print(doc.Content.Text[-300:])     # verify in the SAME call
```

## Files you open yourself

`attach()` works on what the user already has open. To work on a file that is
closed, open it through the skill so it can be closed again afterwards:

```python
app, wb = open_doc("excel", r"C:\...\jelentes.xlsx")   # tracked as ours
...read or edit...
close_doc(wb)                     # allowed: we opened it

app, wb = new_doc("excel")        # create from scratch
save_as(wb, r"C:\...\uj.xlsx")   # allowed: we created it
close_doc(wb)
```

`close_doc` and `save_as` **refuse the user's own documents** — closing their
window or redirecting where their file saves is their call, not yours. Raw
`doc.Close(...)` and `doc.SaveAs(...)` stay blocked for everything, because a
regex cannot tell whose document it is; the refusal names the helper to use.

`visible=False` hides the *document*, never the application — the user's other windows stay put. `save_as` refuses an existing target (Excel would answer with a modal dialog, which over COM is a hang).

`openpyxl` is available here too, for reading or writing a closed workbook
without opening Excel at all. Prefer it when the file does not need to be
visible: it is roughly seventy times faster than COM.

**Never drive Office with hand-written `win32com.client.GetActiveObject(...)`.**
That skips the early-binding repair, the tracked-changes verification and the
AutoRecovery guard documented below — every one of which exists because the
naive version silently did the wrong thing.

`attach(host, name=None)` returns `(app, doc)` and raises `OfficeError` with an
actionable `.hint`. From there you have the whole COM object model — write COM
calls directly, there is no wrapper layer to learn.

Two more helpers come with it, and they are what make writing your own script
better than any fixed tool list:

```python
from office import attach, explore, probe

explore(doc, "footnote")      # -> {'methods': ['Footnotes'], ...}
explore(doc.Paragraphs(1))    # what a paragraph actually offers
```

**`explore(obj, contains=None)`** lists the live object's real methods and
properties. When no recipe below covers what you need, explore instead of
guessing method names — the answer comes from Word's own type library, so it
cannot be a hallucination.

```python
before = probe(doc)
...make the edit...
after = probe(doc)
print(before["sha"] != after["sha"], before["revisions"], after["revisions"])
```

**`probe(doc)`** fingerprints the document (paragraph/revision counts, a content
hash, the tail). Two probes that compare equal prove **nothing changed** — this
is how you catch a COM call that returns `True` and does nothing. Works on all
three apps.

### Bulk edits: one call, all containers, all three apps

`replace_all` works on Word, Excel and PowerPoint, and reaches the places a
naive loop misses:

| app | what it covers | tracking |
| --- | --- | --- |
| Word | every story: body, tables, headers, footers, footnotes, endnotes, comments, text boxes | each hit is a reviewable revision |
| Excel | every worksheet's cells, plus shape and table text | none — Excel has no revision tracking, so back up first |
| PowerPoint | every slide's shapes, tables and speaker notes | none |

In Word, `doc.Content` is the body **only**. Headers, footers, footnotes,
endnotes, comments and text boxes are separate **stories** — that is how a
find-and-replace reports success while the header still names the person you
just anonymised. Verified: 5 occurrences across body, header, footer and a
footnote all replaced in one call.

```python
from office import attach, replace_all, probe

app, doc = attach("word")
doc.TrackRevisions = True
rows = replace_all(doc, [("Kovács János", "[SZEMÉLY-1]"),
                         ("1234567890",  "[ADÓAZONOSÍTÓ]")])
print(rows)
# [{'old': 'Kovács János', 'visible_before': 67, 'visible_after': 0,
#   'replaced': 67, 'in_revisions': 67, 'inserted': 67,
#   'found': True, 'long_pattern': False, 'complete': True}, ...]
```

Read the row like this:

| field | meaning |
| --- | --- |
| `visible_before` / `visible_after` | occurrences in the text **as it will read once revisions are accepted** |
| `replaced` | how many actually changed |
| `in_revisions` | old copies still in the file, but only as tracked deletions — normal, that is the review trail |
| `inserted` | occurrences of the new text |
| `found` | whether the old text was there at all (`replaced: 0, found: False` means "nothing to do", not "failed") |
| `complete` | **never `True` while the old text is still visible** |

### Track Changes changes what "verify" means

With tracking on, a replacement does **not** remove the old text: Word keeps
those characters and marks them as a deletion revision. They are still in
`Range.Text`. Counting raw text therefore reports `before=1, after=1,
replaced=0` for an edit that plainly worked, and the document reads
`régi szövegúj szöveg` because the struck-out old text sits next to the new.

That is expected, not a bug. Verify against the **accepted** view instead:

```python
from office import final_text
"Kovács János" in doc.Content.Text          # True — the deletion mark
"Kovács János" in final_text(doc.Content)   # False — what the reader will see
```

`count_text(doc, needle, accepted=False)` counts raw characters if you ever
need the other view. `probe()` already hashes the accepted text.

### Long patterns: Find.Text stops at 255 characters

Word rejects a longer search string outright ("the string is too long").
Passing a whole paragraph as one `Find.Text` therefore fails. `replace_all`
detects this and switches to a range-based path automatically
(`long_pattern: True` in the row). If you are writing Find code yourself,
match on a **short, unique fragment** instead of the entire paragraph.

`count_text(doc, needle)` counts across whichever app the document belongs to.
`text_chunks(doc)` returns every piece of text, `doc_kind(doc)` says which app
it is, and `final_text(rng)` / `stories(doc)` are Word-specific. Measured: 200
replacements across 7 Word stories in ~2 s.

Everything else in this skill — `attach`, `probe`, `explore`, `status` — works
on all three apps. Only revision tracking and `final_text` are Word-only,
because only Word has revisions.

### Prefer one script over many calls

You are writing the script, so do the whole job in it: read, branch, loop, edit
every match, verify, and `print()` only the summary. Intermediate values never
enter your context, so a 200-row sweep costs the same as a one-liner. Inside
`execute_code` you can also `from hermes_tools import read_file, web_search,
terminal, ...` and combine them with COM in the same script.

```python
# every heading in one pass, with proof — one tool call
app, doc = attach("word")
before = probe(doc)
doc.TrackRevisions = True
touched = []
for i, para in enumerate(doc.Paragraphs, 1):
    if "Címsor" in para.Style.NameLocal and not para.Range.Text.strip().endswith(":"):
        para.Range.InsertAfter(":")
        touched.append(i)
after = probe(doc)
print({"touched": touched, "changed": before["sha"] != after["sha"],
       "revisions": (before["revisions"], after["revisions"])})
```

**When you need state across calls** (long sessions, or you are an agent that
can only run shell commands), use the warm daemon. Variables survive between
calls like a notebook; it auto-starts and exits after 45 idle minutes:

```
office.py exec --warm --code "app, doc = attach('word'); doc.Name"
office.py exec --warm --code "doc.Revisions.Count"        # doc is still there
office.py reset      # drop variables / stale COM refs after Office restarts
```

**Typed commands, for quick reads** (cold, ~0.5 s each):

```
office.py status                                   # all three apps at once
office.py word structure | text | selection | revisions
office.py excel sheets | read --range "Adatok!A1:D50" [--formulas]
office.py excel write --range "Adatok!B2" --values "[[1,2],[3,4]]"
office.py ppt outline
office.py backup --host word
```

Shorthand above: `office.py` = `%LOCALAPPDATA%\hermes\hermes-agent\venv\Scripts\python.exe ${HERMES_SKILL_DIR}/scripts/office.py`.
**Never wrap it in `subprocess.run` inside `execute_code`** — that is ten lines
of boilerplate to do what `attach()` does in two.

## Preflight — always first

```
office.py status
```

Reports for each app: running or not, which documents are open, saved state,
and for Word whether tracking is on and how many revisions exist.

- **"pywin32 is not installed"** → not native Windows Python. The COM route
  does not exist here. Say so and stop.
- **"not reachable over COM"** → the app is closed, *or* running with no
  document open (Office only registers itself once a document is loaded). Ask
  the user to open the file. Never launch Office yourself.

`status` also reports `com.early_binding`:

- `true` → normal. Named COM arguments work.
- `false` → **do not edit yet.** COM fell back to dynamic dispatch, where
  `Find.Execute(Replace=2)` returns `True` and replaces nothing. `attach()`
  purges and rebuilds the makepy cache automatically on first use, so `false`
  means that failed too — usually pywin32's bitness does not match Office.
  Report it instead of writing blind; every write in this mode is unverifiable.

`status` further reports, per document, `storage` and `backup_possible`.
Neither is a permission check — they tell you what to say afterwards, not
whether to proceed:

- `local` — a file exists, so back it up first.
- `sharepoint` — no local copy to back up. Edit, and say that the site's
  version history is the only undo.
- `unsaved` — never saved, so there is nothing to copy. **Edit it anyway.**
  The live document is the one on screen; there is no other copy to protect.
  Mention afterwards that it is still unsaved.

If several documents are open, name the one you mean and confirm before
writing.

**Stop if `status` flags a document with `WARNING`.** That flag means one
specific thing: the open document is an AutoRecovery copy, not the user's
file. After a crash, sleep, or a dropped network share, Office silently
reopens files that way — the workbook becomes `name (version 1).xlsb` in a
hidden per-user folder, or the file sits in a temp folder. Editing one
succeeds, reads back correctly, and changes **nothing in the user's real
file**. Observed exactly that here. Do not edit a flagged document: tell the
user what is actually open, and let them decide whether to recover it or
reopen the original.

A never-saved document is **not** this case and carries no flag. Edit it.

## Safety protocol

1. `status` — confirm you are on the right document.
2. `office.py backup --host <word|excel|ppt>` — saves and copies the file as
   `name.bak-<timestamp>.ext`. **Report the path.** When there is no file to
   copy — SharePoint/OneDrive, or a document that was never saved — this step
   simply does not apply. Say so and carry on; it is never a reason to refuse
   the edit or to ask the user to save first.
3. **Word: turn tracking on** (`doc.TrackRevisions = True`) before editing.
4. Edit **additively** — never touch existing revisions.
5. **Verify by reading back** (next section).
6. Report the backup path and the revision count so the user knows what to review.

Bulk or structural changes (delete sections, reformat everything, rewrite a
sheet) — do not do them in place. Say so and offer a new file instead.

## Verification — the rule that matters

**A call reporting success is not evidence that anything changed.** Two real
failures found on this exact stack:

- A widely used Excel MCP server returns `{"success": true,
  "cells_formatted": 4}` while applying nothing, because it silently drops
  keys it does not recognise.
- `Find.Execute(Replace=2)` under pywin32's *dynamic* dispatch runs the search,
  returns `True`, and **replaces nothing**. `attach()` works around this with
  early binding — but the class of bug is why the rule exists.
- `Footnotes.Add(rng, "text")` positionally puts the text in the `Type` slot,
  creates an EMPTY footnote, and leaks a truncated copy into the body. Named
  arguments work. When a COM call takes optional arguments, name them.
- With Track Changes on, replaced text is **kept as a deletion revision**, so
  raw text counts do not drop. Verify with `final_text` / `count_text(...,
  accepted=True)`, never with `Range.Text` alone.
- `Find.Text` is capped at 255 characters; longer patterns are rejected.
- **Writing a date over COM shifts it by your UTC offset.** pywin32 sends every
  datetime as UTC; Office stores what it gets as local time. Measured: an
  Outlook meeting set for 10:00 landed at 09:00, and an Excel deadline at
  midnight became **23:00 the day before**. Always write dates through
  `com_date(...)` — `ws.Range("B2").Value = com_date("2026-11-13")`. Reading
  needs nothing: drop the `tzinfo` and the value is the local wall clock.
- Assigning `Paragraphs(n).Range.Text = ...` **swallows the paragraph mark and
  merges the paragraph into the next one** — the paragraph count drifts and the
  edit is not atomic. For text substitution use `replace_all`; to change one
  paragraph's wording, replace its text through Find rather than by assigning
  to its Range.

So after every write, re-read the affected region and compare with what you
intended. Report what you observed, not what you asked for. Word's
`Paragraphs(n).Range` includes the paragraph mark, so `InsertAfter` lands at
the *start of paragraph n+1* — verify where the text actually went, not where
you expected it.

## Recipes

### Word

```python
app, doc = attach("word")
doc.TrackRevisions = True

doc.Content.Text                              # full text
[p.Range.Text for p in doc.Paragraphs][:20]   # by paragraph
[(r.Type, r.Author, r.Range.Text) for r in doc.Revisions]   # 1=insert 2=delete

f = doc.Content.Find                          # targeted replace
f.Text, f.Replacement.Text = "régi", "új"
f.Execute(Replace=2)

doc.Paragraphs(12).Range.InsertParagraphAfter()      # insert (see gotcha above)
doc.Paragraphs(13).Range.Text = "Új bekezdés."

doc.Comments.Add(doc.Paragraphs(4).Range, "Ez pontosításra szorul.")

# FOOTNOTE — always name the arguments. Positional silently corrupts:
# Footnotes.Add(rng, "Ket. 44. par.") puts the string in the Type slot and
# leaks a truncated "Ket. 44. p" into the BODY, leaving an empty footnote.
# Verified on this stack; cost an agent 20 turns before it worked around it.
doc.Footnotes.Add(Range=doc.Paragraphs(4).Range, Text="Ket. 44. par.")
doc.Tables(1).Cell(2, 3).Range.Text = "42"
```

### Excel

```python
app, wb = attach("excel", name="jelentes")
ws = wb.Worksheets("Adatok")

ws.UsedRange.Address                          # what is actually filled
ws.Range("A1:D50").Value                      # values (2D tuple)
ws.Range("A1:D50").Formula                    # formulas instead

ws.Range("B2").Value = 42                     # single cell
ws.Range("A5:C5").Value = [["Összesen", 3, 9]]   # a row: 2D list
ws.Range("C10").Formula = "=SUM(C2:C9)"

hdr = ws.Range("A1:D1")                       # formatting
hdr.Font.Bold = True
hdr.Interior.Color = 0xF2E1D9                 # BGR, not RGB
ws.Columns("A:D").AutoFit()
```

Excel has no revision tracking — that is exactly why the backup step is not
optional here, and why you keep the previous values before overwriting.

### PowerPoint

```python
app, pres = attach("ppt")

[(s.SlideIndex, s.Shapes.Title.TextFrame.TextRange.Text)
 for s in pres.Slides if s.Shapes.HasTitle]

s = pres.Slides(4)
s.Shapes(1).TextFrame.TextRange.Text = "Új cím"
s.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Előadói jegyzet."
pres.Slides.AddSlide(pres.Slides.Count + 1, pres.SlideMaster.CustomLayouts(2))
```

### Outlook

Classic Outlook only (`OUTLOOK.EXE`). The new Outlook (`olk.exe`) has no COM
interface at all — if `status()` says Outlook is not running and the user has
it open, that is why; say so instead of retrying.

**The approval rule is built in: you compose, the user sends.** Every Outlook
object you get is wrapped so that `.Send()`, `.Save()`, `.Delete()`,
`.Move()`, `.Copy()` and `.Close()` are refused, and setting `UnRead` or flags
on an existing message is refused too. Build the item, call `.Display()` (the
helpers do it for you), and tell the user it is open for them to review — they
press Send or Save. Never claim a message was sent.

```python
status()["outlook"]                    # running? accounts, unread, drafts

# search — summaries only, newest first, no bodies (they are large and private)
find_mail(since_days=7)
find_mail(query="hiánypótlás", since_days=30)
find_mail(sender="kovacs", unread=True)          # name OR e-mail address
find_mail(folder_name="sent", since="2026-10-01", until="2026-10-15")
get_mail(entry_id, max_chars=4000)               # one message, with body
current_mail()                                   # what the user has selected/open

# compose — opened for the user, NOT sent
new_mail(to=["kovacs.janos@pelda.hu"], subject="Hiánypótlás",
         html="<p>Tisztelt Kovács Úr!</p><p>A <b>15 napos</b> határidő ...</p>"
              "<ul><li>első</li><li>második</li></ul>",
         attachments=[r"C:\...\felhivas.pdf"])
reply(entry_id, body="Köszönöm, megnézem.")      # text goes ABOVE the quote
reply(entry_id, html="...", reply_all=True)
forward(entry_id, to="x@pelda.hu", body="Továbbítom.")
add_attachment(item, r"C:\...\melleklet.docx")
readback(item)                                   # what it really contains

# calendar — read freely, compose meetings for the user to send
calendar(days=7)                                 # recurring meetings expanded
new_appointment("Egyeztetés", "2026-11-13 10:00", minutes=45,
                location="Tárgyaló", attendees=["kovacs.janos@pelda.hu"])

# contacts
resolve("Kovács János")          # address book (GAL) -> {"name", "email"}
find_contact("Kovács")           # Contacts folder + address book
new_contact("Teszt Elek", email="...", company="...")
```

Formatting: pass `html=` for anything beyond plain paragraphs — bold, lists,
tables, links. Outlook's default font is applied around it, and when the item
is displayed the user's own signature is kept, with your text above it
(verified live). `body=` takes plain text; blank lines become paragraphs.

**Verification is built in — do not write it yourself.** End the snippet with
the item and it prints its readback: every recipient with the SMTP address it
resolved to, `unresolved` names, attachments, the start of the text, and
`sent: false`. `readback(item)` returns the same as a dict. One call composes
AND proves:

```python
m = new_mail(to="Lovas Dóra", subject="...", html="...", attachments=[...])
m        # -> {"recipients": [{"name": "Lovas Dóra dr.", "email": "dora.lovas@...",
         #     "resolved": true}], "unresolved": [], "attachments": [...], ...}
```

Do not read Outlook properties by hand to check a draft — an agent did, guessed
`Recipients.Account` (no such thing; Recipients is a collection) and spent a
turn on the error.

Things that will bite:

- **Names are fine in `to=` — check what they resolved to.** The readback shows
  the address each name landed on. If `unresolved` is not empty, or an address
  is not the person the user meant (two Kovácses), say so and ask; do not
  guess. `resolve(name)` is there when you need the address before composing.
- **Outlook collections are not callable**: `item.Recipients.Item(1)`, not
  `item.Recipients(1)` (TypeError under early binding).
- **Searching by e-mail address**: for an Exchange sender the plain
  `fromemail` field holds an X500 path, not the SMTP address. `find_mail(
  sender=...)` also checks the SMTP property; hand-written DASL must too
  (`0x5D01001F`), or it finds a fraction of the mail.
- **Dates in filters**: Jet syntax (`[ReceivedTime] >= '2026-10-01 00:00'`)
  matches local time exactly. DASL date filters read the value as UTC and miss
  messages at the edges; with a trailing `Z` they silently matched almost the
  whole inbox. The helpers chain Jet for dates and DASL for text.
- **Outlook's own AutoSave**: an unsent item left open is saved to Drafts by
  Outlook itself after a few minutes. That is Outlook, not you — nothing is
  sent — but mention it if the user leaves drafts lying open.
- `find_mail` returns every item kind (mail, meeting requests and responses,
  delivery reports) labelled with `kind`; filter on it rather than assuming
  everything is a message.

## Performance — and when NOT to use this skill

Measured on this stack, 200 replacements across a 60-paragraph document with a
table, header, footer and footnote:

| route | time | in place? | tracked changes? |
| --- | --- | --- | --- |
| `replace_all` over COM | ~2 s | yes, file stays open | yes |
| python-docx on a closed copy | ~0.03 s | no — must close, edit, reopen | no |

COM is ~70x slower and always will be: it is cross-process RPC into a running
GUI application. That is the wrong thing to optimise, because a single model
turn costs far more than 2 s. What actually wastes time is turns — so do the
whole job in one script, and never loop a COM call per paragraph.

**Use python-docx instead when** the job is bulk rewriting, the file is local
and closed (or can be), and nobody needs tracked revisions. **Use this skill
when** the document is open in front of the user, lives on SharePoint, or the
edits must land as reviewable revisions. For a long interactive session, use
`exec --warm` so the attach cost is paid once.

## Failure modes

| Symptom | Cause | What to do |
| --- | --- | --- |
| `pywin32 is not installed` | not native Windows Python | Stop; COM is unavailable here. |
| `not reachable over COM` | app closed, or open with no document | Ask the user to open the file. Never `--launch` on your own. |
| `no open document` | app running, nothing loaded | Ask which file to open. |
| Call returns `True` but nothing changed | dynamic-dispatch misbind | Use `attach()` (early binding). Always read back. |
| Command hangs | modal dialog open in Office | Ask the user to dismiss it. Do not retry blindly. |
| `Office rejected the call` | app busy mid-operation | Wait, retry once, then tell the user. |
| Stale COM errors after reopening Office | daemon holds dead refs | `office.py reset`. |
| `never been saved to disk` | the `backup` command on an untitled document | Nothing to back up. Make the edit and say it is still unsaved. |

## Boundaries

- **Never accept or reject revisions.** Not `AcceptAll()`, not `RejectAll()`,
  not `Revisions(n).Accept()`, and not "just to get a clean state first" —
  that reasoning is exactly what destroys an audit trail. Observed failure: an
  agent asked to make one edit called `AcceptAll()` first, erasing four
  revisions that were waiting for human review. Review decisions belong to the
  user; you only add.
- **Never present a file built another way as live work.** If you produced the
  result with openpyxl or python-docx instead of in the open document, say so
  plainly. Hermes recorded this lesson itself after a session that generated a
  new workbook while the user had asked for the open one — it tried to write
  it here, could not (plugin skills are read-only to the curator), and put it
  in the `xlsx` skill instead.
- Never `Close()`, `SaveAs`, or `Quit`. Saving in place after a confirmed edit
  is fine; closing and deleting are not.
- Never disable `TrackRevisions` or `DisplayAlerts` to make something succeed.

The shell paths (`office.py exec`, `--warm`) refuse these calls outright. **The
library path cannot** — once you hold the COM object, Python can call anything,
so on that path these rules are enforced only by you following them.
