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
        "path": {"type": "string", "description": "Full path of the file."},
        "pages": {"type": "string",
                  "description": "PDF pages, 1-based, e.g. '1-3,7'. Default: all."},
        "ocr": {"type": "string", "enum": ["auto", "never", "force", "cloud"],
                "description": "auto: your OCR server (OCR_URL) for pages without text. cloud also "
                               "sends pages still empty to the cloud model; only after the user agreed."},
        "attachment": {"type": "string",
                       "description": "For an .msg/.eml: read this attachment instead, by index "
                                      "('2'), nested index ('2/1') or file name."},
        "max_chars": {"type": "integer", "description": "Cap on returned text. Default 40000."},
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
        "path": {"type": "string", "description": "Full path of the PDF, or of an e-mail with "
                                                  "`attachment` set."},
        "pages": {"type": "string", "description": "Pages, 1-based, e.g. '1' or '2-3'. Max 20."},
        "dpi": {"type": "integer", "description": "Resolution. Default 150."},
        "attachment": {"type": "string", "description": "A PDF attachment inside an .msg/.eml."},
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
        "ocr": {"type": "string", "enum": ["auto", "never", "cloud"],
                "description": "One file only. auto: scanned pages go through your OCR server first. "
                               "cloud only after the user agreed."},
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
                         max_chars=a.get("max_chars") or files.TEXT_LIMIT)

    def _attachments(args=None, **_):
        a = args or {}
        return files.run(files.save_attachments, a.get("path"), dest=a.get("dest"),
                         which=a.get("which"), include_inline=bool(a.get("include_inline")))

    def _render(args=None, **_):
        a = args or {}
        return files.run(files.render_pages, a.get("path"), pages=a.get("pages") or "1",
                         dpi=a.get("dpi") or 150, attachment=a.get("attachment"))

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
            "Read a PDF, e-mail (.msg/.eml) or Office file as text. "
            "Scanned PDF pages go to your own OCR server; an e-mail comes back as "
            "headers, body and an attachment list, and attachment='2' reads one "
            "attachment in place. Also .docx, .xlsx, .pptx, HTML and text. Read-only."
        ),
        emoji="\N{OPEN BOOK}",
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
            "Render PDF pages to PNG images to look at them. "
            "Returns the image paths; open them with the vision tool."
        ),
        emoji="\N{FRAME WITH PICTURE}",
    )

    skill = _SKILL_DIR / "SKILL.md"
    if skill.exists():
        ctx.register_skill(
            "file-reader", skill,
            description="How to read PDF, e-mail and Office files, and where scanned pages go.",
        )
