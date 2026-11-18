---
name: file-reader
description: Read files on disk as text in one call — PDF (scanned pages go through OCR), Outlook .msg and .eml e-mails with their attachments, Word .docx, Excel .xlsx, PowerPoint .pptx, HTML and plain text. Read an attachment inside an e-mail without saving it, save attachments to a folder, or render PDF pages as images to look at. Use when the user points at a file path or an e-mail file ("olvasd el ezt a PDF-et", "mi van ebben a levélben", "nézd meg a mellékletet", "szkennelt PDF"). Do NOT use for a document the user has OPEN in Word, Excel or PowerPoint, or for mail in the running Outlook — that is office-live. Do NOT use for OneDrive or SharePoint files — that is m365-graph.
version: 1.0.0
author: Integritás Hatóság
---

# file-reader

Call `file_read` with the full path. Do not write a script to parse a PDF,
an .msg or an .eml: this tool already does it, and it is faster.

| Want | Call |
|---|---|
| a file's text | `file_read(path)` |
| some PDF pages | `file_read(path, pages="3-5,9")` |
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
  `page_sources` says how each page was read: `text-layer`, `local-ocr`,
  `azure-vision` or `no-text`.
- **E-mail**: `subject`, `from`, `to`, `cc`, `date`, `text` (the body) and
  `attachments`, each with `index`, `name`, `mime`, `size`, `inline`.
  `inline` items are pictures in the body, usually signature logos; leave them
  out of a summary. `embedded_email: true` is an attached e-mail, readable with
  `attachment="<index>"`.
- **Word / Excel / PowerPoint**: `text`; for Excel one block per sheet, for
  PowerPoint one per slide.
- `truncated: true` means the text was cut. `next` says which pages to ask for.

## Scanned PDFs and where the pages go

1. A page with a text layer is read on this machine. It never leaves it.
2. A page without one is rendered and sent to the organisation's internal OCR
   model (`M365_OCR_URL`). Ten pages per call; `notes` names the rest.
3. Nothing goes to a cloud model unless you pass `ocr="cloud"`. When pages
   stay empty, `notes` says so. Ask the user before using `ocr="cloud"`: the
   page images then go to the Azure vision deployment.

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
from files import read_file, save_attachments, render_pages, status
read_file(r"C:\path\file.pdf", pages="1-3")
```

or from a terminal: `py scripts/files.py read <path> [--pages 1-3] [--attachment 2]`,
`attachments <path>`, `render <path> --pages 1`, `status`.

Libraries: pypdfium2 (PDF text and rendering), pdfminer.six (text fallback),
olefile (.msg), openpyxl (.xlsx). Word, PowerPoint, .eml, HTML and text need
nothing beyond Python. `status` lists what is installed and whether OCR is
configured.
