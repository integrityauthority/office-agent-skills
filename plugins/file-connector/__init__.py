"""File Connector — let Hermes read PDF, e-mail and Office files from disk
without writing a parser each time.

A thin adapter. The code lives in ``skills/file-reader/scripts/files.py`` and
``.../msgfile.py``, the same files the portable `file-reader` Agent Skill
ships; `tools/sync_plugins.py` copies them here and CI fails if they drift.
Edit the skill, not this copy.
"""
from pathlib import Path
import sys

# The skill lives in skills/file-reader/ (the Agent Skills layout every platform reads);
# tools/sync_plugins.py copies it there from .agents/skills/file-reader/.
_SKILL_DIR = Path(__file__).parent / "skills" / "file-reader"
_SCRIPTS = _SKILL_DIR / "scripts"

READ_SCHEMA = {
    "type": "object",
    "properties": {
        "path": {"type": "string", "description": "Full path of the file (PDF, image, .msg, .eml, Office incl. .xls, .zip, text)."},
        "pages": {"type": "string",
                  "description": "PDF or multi-page TIFF pages, 1-based, e.g. '1-3,7'. Default: all."},
        "ocr": {"type": "string", "enum": ["auto", "server", "cloud", "force", "never"],
                "description": "Scanned pages. auto: Windows OCR on this machine (fast). server: the "
                               "organisation's OCR server, better for tables and numbers; use it when "
                               "the auto text looks wrong. cloud: the cloud model, only after the user "
                               "agreed. force: server even for pages with text. never: text layer only."},
        "attachment": {"type": "string",
                       "description": "For an .msg/.eml: read this attachment instead, by index "
                                      "('2'), nested index ('2/1') or file name."},
        "max_chars": {"type": "integer", "description": "Cap on returned text. Default 40000."},
        "labels": {"type": "string",
                   "description": "Comma-separated labels whose values to pick out, e.g. "
                                  "'Iktatószám, Tulajdonos': 'Iktatószám: 123/2026' gives "
                                  "{'Iktatószám': '123/2026'}. Case and accents are ignored."},
    },
    "required": ["path"],
}

COLLECT_SCHEMA = {
    "type": "object",
    "properties": {
        "path": {"type": "string", "description": "Full path of a folder, or of a .zip archive."},
        "labels": {"type": "string", "description": "Comma-separated labels whose values become "
                                                    "columns, e.g. 'Iktatószám, Határidő'."},
        "query": {"type": "string", "description": "Words to count in each file (case and accents "
                                                   "ignored); adds 'hits' and 'first_hit' columns."},
        "regex": {"type": "boolean", "description": "Treat query as a regular expression."},
        "recursive": {"type": "boolean", "description": "Folder: include subfolders."},
        "excerpt": {"type": "integer", "description": "Characters of text per row. Default 300; "
                                                      "0 for none."},
        "out": {"type": "string", "description": "Also write the table as a UTF-8 CSV here. "
                                                 "Written anyway when there are over 100 rows."},
    },
    "required": ["path"],
}

UNPACK_SCHEMA = {
    "type": "object",
    "properties": {
        "path": {"type": "string", "description": "Full path of the .zip archive."},
        "dest": {"type": "string", "description": "Folder to extract into. Default: a folder named "
                                                  "after the zip, next to it."},
    },
    "required": ["path"],
}

ATTACH_SCHEMA = {
    "type": "object",
    "properties": {
        "path": {"type": "string", "description": "Full path of the .msg or .eml file."},
        "dest": {"type": "string", "description": "Folder to write to. Default: a cache folder."},
        "which": {"type": "string", "description": "Indexes or names, comma-separated. Default: all "
                                                   "real attachments."},
        "include_inline": {"type": "boolean",
                           "description": "Also save pictures embedded in the body (logos)."},
    },
    "required": ["path"],
}

RENDER_SCHEMA = {
    "type": "object",
    "properties": {
        "path": {"type": "string", "description": "Full path of the PDF or picture, or of an e-mail "
                                                  "with `attachment` set."},
        "pages": {"type": "string", "description": "Pages, 1-based, e.g. '1' or '2-3'. Max 20."},
        "dpi": {"type": "integer", "description": "Resolution, 36-600. Default 150, or 300 with box."},
        "box": {"type": "string",
                "description": "Only this region of each page: 'left,top,right,bottom' as fractions "
                               "of the page from its top-left corner, e.g. '0.5,0.6,1,0.8'. For a "
                               "close look at a digit, an area or a handwritten entry."},
        "attachment": {"type": "string", "description": "A PDF or picture attachment inside an .msg/.eml."},
    },
    "required": ["path"],
}

FIND_SCHEMA = {
    "type": "object",
    "properties": {
        "path": {"type": "string", "description": "Full path of a file, or of a folder to search "
                                                  "every readable file in it."},
        "query": {"type": "string", "description": "Words to find. Case and accents are ignored and "
                                                   "any whitespace may separate the words."},
        "regex": {"type": "boolean", "description": "Treat query as a regular expression, e.g. "
                                                    "'kifejezés1|kifejezés2'."},
        "exact": {"type": "boolean", "description": "Case and accents must match."},
        "pages": {"type": "string", "description": "PDF pages to search, e.g. '1-20'. Default: all."},
        "ocr": {"type": "string", "enum": ["auto", "server", "cloud", "never"],
                "description": "One file only, scanned pages. auto: Windows OCR here. server: the "
                               "organisation's OCR server. cloud: only after the user agreed."},
        "attachment": {"type": "string", "description": "Search only this attachment of an .msg/.eml."},
        "recursive": {"type": "boolean", "description": "Folder: include subfolders."},
        "max_hits": {"type": "integer", "description": "Matches (or files) to return. Default 50."},
    },
    "required": ["path", "query"],
}


def register(ctx):
    if str(_SCRIPTS) not in sys.path:
        sys.path.insert(0, str(_SCRIPTS))
    import files

    # Hermes calls handler(args_dict, **context): the dict comes positionally.
    def _read(args=None, **_):
        a = args or {}
        return files.run(files.read_file, a.get("path"), pages=a.get("pages"),
                         ocr=a.get("ocr") or "auto", attachment=a.get("attachment"),
                         max_chars=a.get("max_chars") or files.TEXT_LIMIT, labels=a.get("labels"))

    def _collect(args=None, **_):
        a = args or {}
        excerpt = a.get("excerpt")
        return files.run(files.collect_files, a.get("path"), labels=a.get("labels"),
                         query=a.get("query"), regex=bool(a.get("regex")),
                         recursive=bool(a.get("recursive")),
                         excerpt=files.COLLECT_EXCERPT if excerpt is None else excerpt,
                         out=a.get("out"))

    def _unpack(args=None, **_):
        a = args or {}
        return files.run(files.unpack_file, a.get("path"), dest=a.get("dest"))

    def _attachments(args=None, **_):
        a = args or {}
        return files.run(files.save_attachments, a.get("path"), dest=a.get("dest"),
                         which=a.get("which"), include_inline=bool(a.get("include_inline")))

    def _render(args=None, **_):
        a = args or {}
        return files.run(files.render_pages, a.get("path"), pages=a.get("pages") or "1",
                         dpi=a.get("dpi") or None, attachment=a.get("attachment"), box=a.get("box"))

    def _find(args=None, **_):
        a = args or {}
        return files.run(files.find_in_files, a.get("path"), a.get("query"),
                         regex=bool(a.get("regex")), exact=bool(a.get("exact")),
                         pages=a.get("pages"), ocr=a.get("ocr") or "auto",
                         attachment=a.get("attachment"), recursive=bool(a.get("recursive")),
                         max_hits=a.get("max_hits") or files.FIND_MAX_HITS)

    ctx.register_tool(
        "file_find", "files", FIND_SCHEMA, _find,
        description=(
            "Search inside a document, e-mail or folder of files. "
            "Returns each match with where it is -- page, slide, sheet and row, line, "
            "or attachment -- and a snippet. Ignores case and accents. Use it instead "
            "of reading a long PDF end to end; a folder answer lists which files match."
        ),
        emoji="\N{RIGHT-POINTING MAGNIFYING GLASS}",
    )

    ctx.register_tool(
        "file_read", "files", READ_SCHEMA, _read,
        description=(
            # The tool-search catalogue shows the first sentence, cut at 60 characters.
            "Read a PDF, image, e-mail or Office file as text. "
            "Scanned pages and pictures (PNG, JPEG, TIFF) go through OCR; an e-mail comes back as headers, body and an "
            "attachment list, and attachment='2' reads one attachment in place -- or one "
            "member of a .zip. Word comments, tracked changes and ticked boxes, PDF signers "
            "and form fields come along; labels= picks out values. Read-only."
        ),
        emoji="\N{OPEN BOOK}",
    )
    ctx.register_tool(
        "file_collect", "files", COLLECT_SCHEMA, _collect,
        description=(
            "Tabulate a folder or .zip of letters and documents. "
            "One row per file: sender, date and subject of an e-mail, first line, pages, "
            "signers, form fields, comment and revision counts, the values after the "
            "given labels and the query hits. Large tables also go to a CSV. Read-only."
        ),
        emoji="\N{CARD INDEX DIVIDERS}",
    )
    ctx.register_tool(
        "file_unpack", "files", UNPACK_SCHEMA, _unpack,
        description=(
            "Extract a .zip archive safely to a folder. "
            "Never overwrites; skips paths that would escape the folder, encrypted members "
            "and zip bombs. To read one member, file_read with attachment= is enough."
        ),
        emoji="\N{PACKAGE}",
    )
    ctx.register_tool(
        "file_attachments", "files", ATTACH_SCHEMA, _attachments,
        description=(
            "Save the attachments of an .msg or .eml file to a folder. "
            "Never overwrites; signature images are skipped unless asked for."
        ),
        emoji="\N{PAPERCLIP}",
    )
    ctx.register_tool(
        "file_render", "files", RENDER_SCHEMA, _render,
        description=(
            "Render PDF pages or pictures to PNG to look at them. "
            "Returns the image paths; open them with the vision tool. box= crops to one "
            "region at a higher resolution, to check an uncertain digit or area."
        ),
        emoji="\N{FRAME WITH PICTURE}",
    )

    skill = _SKILL_DIR / "SKILL.md"
    if skill.exists():
        ctx.register_skill(
            "file-reader", skill,
            description="How to read PDF, e-mail and Office files, and where scanned pages go.",
        )
