---
name: file-reader
description: Read files on disk as text in one call — PDF (scanned pages go through OCR), Outlook .msg and .eml e-mails with their attachments, Word .docx, Excel .xlsx, PowerPoint .pptx, HTML and plain text. Read an attachment inside an e-mail without saving it, save attachments to a folder, or render PDF pages as images to look at. Use when the user points at a file path or an e-mail file ("olvasd el ezt a PDF-et", "mi van ebben a levélben", "nézd meg a mellékletet", "szkennelt PDF"). Do NOT use for a document the user has OPEN in Word, Excel or PowerPoint, or for mail in the running Outlook — that is office-live. Do NOT use for OneDrive or SharePoint files — that is m365-graph.
version: 1.1.0
author: Integritás Hatóság
---

# file-reader

Call `file_read` with the full path. Do not write a script to parse a PDF,
an .msg or an .eml: this tool already does it, and it is faster.

| Want | Call |
|---|---|
| a file's text | `file_read(path)` |
| some PDF pages | `file_read(path, pages="3-5,9")` |
| where something is in a long document | `file_find(path, "words")` |
| which files in a folder mention it | `file_find(folder, "words")` (`recursive=true` for subfolders) |
| an e-mail: headers, body, attachment list | `file_read(path)` on the .msg / .eml |
| one attachment, without saving it | `file_read(path, attachment="2")` |
| an attachment of an attached e-mail | `file_read(path, attachment="2/1")` |
| the attachments as files | `file_attachments(path)` (or `dest=`, `which="1,3"`) |
| to look at a PDF page | `file_render(path, pages="1")`, then open the PNG with the vision tool |

## What comes back

Every answer has `ok`. On `false` there is an `error` and usually a `hint`;
follow the hint instead of retrying the same call.

- **PDF**: `page_count`, `pages_read`, `page_sources` and `text` with a
  `--- page N ---` line before each page. Quote page numbers from those lines.
  `page_sources` says how each page was read: `text-layer`, `ocr-server`,
  `cloud-vision` or `no-text`.
- **E-mail**: `subject`, `from`, `to`, `cc`, `date`, `text` (the body) and
  `attachments`, each with `index`, `name`, `mime`, `size`, `inline`.
  `inline` items are pictures in the body, usually signature logos; leave them
  out of a summary. `embedded_email: true` is an attached e-mail, readable with
  `attachment="<index>"`.
- **Word / Excel / PowerPoint**: `text`; for Excel one block per sheet, for
  PowerPoint one per slide.
- `truncated: true` means the text was cut. `next` says which pages to ask for.
- **Search** (`file_find`): `hit_count` and `hits`, each with `where` (`page 12`,
  `slide 3`, `sheet Adatok, row 40`, `body, line 7`, `attachment 2 (x.pdf) ›
  page 3`) and a `snippet` with the match between « ». Cite `where`. A folder
  answer has `files_with_hits` (path, count, first matches), most hits first.
  `not_searched` lists pages without text (search the file on its own, or
  agree OCR); a folder search reads text layers only, newest files first,
  within 45 seconds.

## Scanned PDFs and where the pages go

The same order as m365-graph:

1. A page with a text layer is read on this machine. It never leaves it.
2. A page without one always goes to the organisation's own OCR server
   (`OCR_URL`) first. Ten pages per call; `notes` names the rest.
3. Nothing goes to a cloud model on its own. When pages stay empty and a
   cloud model is configured, the answer has `cloud_ocr_needs_approval` with
   the page numbers. Ask the user. Only if they agree, call again with
   `ocr="cloud"`: the OCR server is still tried first, and the pages it
   leaves empty go as images to the cloud vision model (`OCR_CLOUD_URL`).

`ocr="never"` reads only the text layer. `ocr="force"` OCRs every requested
page, for a PDF whose text layer is garbage.

## Boundaries

- This reads. It never changes, moves or deletes the source file.
  `file_attachments` writes new files only, and never overwrites one: a name
  that exists gets ` (2)`.
- Legacy `.doc`, `.xls` and `.ppt` are not read here; open them with
  office-live (`open_doc(path)`).
- A password-protected PDF is refused; ask the user for an open copy.
- Do not paste a whole document back to the user. Summarise, and quote the
  page or the attachment you are citing.

## Without the Hermes plugin

The same functions, from Python in `scripts/files.py`:

```python
import sys; sys.path.insert(0, r"<skill dir>\scripts")
from files import read_file, find_in_files, save_attachments, render_pages, status
read_file(r"C:\path\file.pdf", pages="1-3")
find_in_files(r"C:\path\file.pdf", "words")
```

or from a terminal: `py scripts/files.py read <path> [--pages 1-3] [--attachment 2]`,
`find <path> "words" [--regex] [--recursive]`, `attachments <path>`,
`render <path> --pages 1`, `status`.

`scripts/pdfread.py` is the PDF and OCR reader this skill shares with
m365-graph; it is generated from `shared/pdfread.py` in the repository.

Libraries: pypdfium2 (PDF text and rendering), pdfplumber (tables, optional),
pdfminer.six (text fallback),
olefile (.msg), openpyxl (.xlsx). Word, PowerPoint, .eml, HTML and text need
nothing beyond Python. `status` lists what is installed and whether OCR is
configured.
