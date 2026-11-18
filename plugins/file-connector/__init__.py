"""File Connector — let Hermes read PDF, e-mail and Office files from disk
without writing a parser each time.

A thin adapter. The code lives in ``scripts/files.py`` and
``scripts/msgfile.py``, the same files the portable `file-reader` Agent Skill
ships; `tools/sync_plugins.py` copies them here and CI fails if they drift.
Edit the skill, not this copy.
"""
from pathlib import Path
import sys

_SCRIPTS = Path(__file__).parent / "scripts"

READ_SCHEMA = {
    "type": "object",
    "properties": {
        "path": {"type": "string", "description": "Full path of the file."},
        "pages": {"type": "string",
                  "description": "PDF pages, 1-based, e.g. '1-3,7'. Default: all."},
        "ocr": {"type": "string", "enum": ["auto", "never", "force", "cloud"],
                "description": "auto: internal OCR for pages without text. cloud sends page "
                               "images to Azure; ask the user first."},
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

    ctx.register_tool(
        "file_read", "files", READ_SCHEMA, _read,
        description=(
            # The tool-search catalogue shows the first sentence, cut at 60 characters.
            "Read a PDF, e-mail (.msg/.eml) or Office file as text. "
            "Scanned PDF pages go through the internal OCR; an e-mail comes back as "
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

    skill = Path(__file__).parent / "SKILL.md"
    if skill.exists():
        ctx.register_skill(
            "file-reader", skill,
            description="How to read PDF, e-mail and Office files, and where scanned pages go.",
        )
