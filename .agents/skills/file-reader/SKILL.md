---
name: file-reader
description: Read files on disk as text in one call — PDF and images (scanned pages, photos and screenshots go through OCR), Outlook .msg and .eml e-mails with their attachments, Word .docx, Excel .xlsx and 97-2003 .xls, PowerPoint .pptx, HTML and plain text. Search inside a document, an e-mail with its attachments, or a whole folder, with page-level locations. Read an attachment or a .zip member without saving it, save attachments, unpack a .zip safely, tabulate a folder or .zip of letters and documents (sender, date, signers, labelled values), or render PDF pages and pictures as images to look at, whole or one region enlarged. Use when the user points at a file path or an e-mail file ("olvasd el ezt a PDF-et", "mi van ebben a levélben", "nézd meg a mellékletet", "szkennelt PDF"). Do NOT use for a document the user has OPEN in Word, Excel or PowerPoint, or for mail in the running Outlook — that is office-live. Do NOT use for OneDrive or SharePoint files — that is m365-graph.
license: MIT
compatibility: Any OS with Python 3.10+. pypdfium2, pdfplumber, pdfminer.six, olefile, openpyxl, xlrd, pillow and pypdf give full coverage; OCR settings are in the repository README.
metadata:
  version: "1.6.0"
  author: Integritás Hatóság
---

# file-reader

Call `file_read` with the full path. Do not write a script to parse a PDF,
an .msg or an .eml: this tool already does it, and it is faster.

Reading happens here, at file level, in the background: nothing opens on the
user's screen. When the result is a document for the user — a reply, a
table, a summary — write it with office-live into a document open in front
of them, not into a closed file.

| Want | Call |
|---|---|
| a file's text | `file_read(path)` |
| some PDF pages | `file_read(path, pages="3-5,9")` |
| a scanned image, photo or screenshot | `file_read(path)` on the .png / .jpg / .tif: OCR, as for a scanned page |
| where something is in a long document | `file_find(path, "words")` |
| which files in a folder mention it | `file_find(folder, "words")` (`recursive=true` for subfolders) |
| an e-mail: headers, body, attachment list | `file_read(path)` on the .msg / .eml |
| one attachment, without saving it | `file_read(path, attachment="2")` |
| an attachment of an attached e-mail | `file_read(path, attachment="2/1")` |
| the attachments as files | `file_attachments(path)` (or `dest=`, `which="1,3"`) |
| to look at a PDF page or a picture | `file_render(path, pages="1")`, then open the PNG with the vision tool |
| a close look at one spot (a digit, an area, handwriting) | `file_render(path, pages="2", box="0.5,0.6,1,0.8")`: that region at 300 dpi |
| what is in a .zip | `file_read(path)` on the .zip: a member list |
| one file inside a .zip, without unpacking | `file_read(path, attachment="3")` or `attachment="name.pdf"` |
| the .zip as a folder | `file_unpack(path)` (or `dest=`) |
| values written after labels | `file_read(path, labels="Iktatószám, Határidő")` |
| a table of a folder or .zip of letters / documents | `file_collect(folder, labels=..., query=...)` |

## What comes back

Every answer has `ok`. On `false` there is an `error` and usually a `hint`;
follow the hint instead of retrying the same call.

- **PDF**: `page_count`, `pages_read`, `page_sources` and `text` with a
  `--- page N ---` line before each page. Quote page numbers from those lines.
  `page_sources` says how each page was read: `text-layer`, `windows-ocr`,
  `ocr-server`, `cloud-vision` or `no-text`.
- **Image**: like a PDF, `kind: "image"`; each frame of a multi-page TIFF
  is a page. The text is an OCR reading: check digits and dates.
- **E-mail**: `subject`, `from`, `to`, `cc`, `date`, `text` (the body) and
  `attachments`, each with `index`, `name`, `mime`, `size`, `inline`.
  `inline` items are pictures in the body, usually signature logos; leave them
  out of a summary. `embedded_email: true` is an attached e-mail, readable with
  `attachment="<index>"`.
- **Word / Excel / PowerPoint**: `text`; for Excel (.xlsx and .xls) one block per sheet, for
  PowerPoint one per slide. A .docx also gives `comments` (author, date, text
  and the text they are anchored to), `revisions` (tracked insertions and
  deletions) and `checkboxes` (`checked`, `label`) when it has any.
- **PDF extras**: `signed` and `signers` (the names written in the signature;
  the signature itself is not validated) and `form_fields` with their values.
- **.zip**: `member_count` and `members` (`index`, `name`, `size`). Read one
  with `attachment=`; `file_find` searches inside the members. A member that
  is encrypted, too large or compressed suspiciously is marked `unreadable`.
  Names stored by Windows Explorer or an older archiver in a DOS code page
  come back with their accents (see `ZIP_NAME_ENCODING` in the README).
- **Render** (`file_render`): `dpi` and `images`, each with `page` and `png`
  (and `box` when cropped). `box` is `left,top,right,bottom` as fractions of
  the page from its top-left corner. Use it to check a value an OCR reading
  is unsure of; say what you saw on the image, not what the OCR said.
- **labels**: `{"Iktatószám": "123/2026"}` — the rest of the line after each
  label. Case and accents are ignored; a label not found is left out. Check
  a value that matters against the `text`.
- **Collect** (`file_collect`): `count`, `columns` and one row per file —
  e-mails with `from`, `to`, `date`, `subject`, `attachments`; other files
  with `title` (first line), `pages`, `scanned_pages`, `signers`,
  `form_fields`, `comments`, `revisions`, `checked`; then `label:<name>`,
  `hits` and `first_hit` when asked, and an `excerpt`. Text layers only: a
  scanned PDF shows `scanned_pages`; read it with `file_read` for its text.
  Over 100 rows, or with `out=`, the full table is in the `csv` file. To hand
  the table to the user, put it in an Excel workbook open in front of them
  with office-live.
- `truncated: true` means the text was cut. `next` says which pages to ask for.
- **Search** (`file_find`): `hit_count` and `hits`, each with `where` (`page 12`,
  `slide 3`, `sheet Adatok, row 40`, `body, line 7`, `attachment 2 (x.pdf) ›
  page 3`) and a `snippet` with the match between « ». Cite `where`. A folder
  answer has `files_with_hits` (path, count, first matches), most hits first.
  `not_searched` lists pages without text (search the file on its own, or
  agree OCR); a folder search reads text layers only, newest files first,
  within 45 seconds.

## Scanned PDFs and where the pages go

The same order as m365-graph, cheapest and most private first:

1. A page with a text layer is read on this machine. It never leaves it.
2. A scanned page is read by **Windows OCR on this machine** (`ocr="auto"`,
   the default): well under a second a page, nothing leaves the computer.
   `page_sources` says `windows-ocr`.
3. **Check that text.** Windows OCR returns tables as loose lines and can
   slip a digit. If the text looks wrong, incomplete, or the user doubts
   it, call again with `ocr="server"` and those pages: the organisation's
   OCR server reads them with table structure. The answer's `ocr_upgrade`
   lists the pages Windows read. No approval is needed; the server is
   inside the organisation. Ten pages per call; `notes` names the rest.
4. Nothing goes to a cloud model on its own. When pages stay empty and a
   cloud model is configured, the answer has `cloud_ocr_needs_approval` with
   the page numbers. Ask the user. Only if they agree, call again with
   `ocr="cloud"`.

Where there is no Windows OCR, `auto` goes to the OCR server instead.
`ocr="never"` reads only the text layer. `ocr="force"` sends every requested
page to the OCR server, for a PDF whose text layer is garbage.

## Boundaries

- This reads. It never changes, moves or deletes the source file.
  `file_attachments` and `file_unpack` write new files only, and never
  overwrite one: a name that exists gets ` (2)`.
- `file_unpack` skips a member whose path would land outside the target
  folder, an encrypted member and a likely zip bomb, and caps the sizes;
  `skipped` says what was left out and why. Nothing from an archive is run.
- Legacy `.doc` and `.ppt` are not read here; open them with
  office-live (`open_doc(path)`). An `.xls` is read with xlrd: values only,
  formulas show their last computed result.
- A password-protected PDF is refused; ask the user for an open copy.
- Do not paste a whole document back to the user. Summarise, and quote the
  page or the attachment you are citing.

## Without the Hermes plugin

The same functions, from Python in `scripts/files.py`:

```python
import sys; sys.path.insert(0, r"<skill dir>\scripts")
from files import read_file, find_in_files, collect_files, unpack_file, save_attachments, render_pages, status
read_file(r"C:\path\file.pdf", pages="1-3")
find_in_files(r"C:\path\file.pdf", "words")
collect_files(r"C:\path\letters", labels="Iktatószám")
```

or from a terminal: `py scripts/files.py read <path> [--pages 1-3] [--attachment 2] [--labels "A, B"]`,
`find <path> "words" [--regex] [--recursive]`, `collect <folder|zip> [--labels ..] [--query ..] [--out x.csv]`,
`unpack <zip> [--dest ..]`, `attachments <path>`, `render <path> --pages 1 [--box 0,0.5,1,1]`, `status`.

`scripts/archive.py` (zip listing, reading and safe unpacking) and
`scripts/docmeta.py` (Word comments, revisions and checkboxes; PDF signers
and form fields) can be used on their own as well.

`scripts/pdfread.py` is the PDF and OCR reader this skill shares with
m365-graph; it is generated from `shared/pdfread.py` in the repository.

Libraries: pypdfium2 (PDF text and rendering), pdfplumber (tables, optional),
pdfminer.six (text fallback), pywinrt (Windows OCR, Windows only),
olefile (.msg), openpyxl (.xlsx), xlrd (.xls), pypdf (PDF form fields, optional),
pillow (pictures, cropping). Word,
PowerPoint, .eml, .zip, HTML and text need nothing beyond Python. `status` lists what is installed and whether OCR is
configured.
