"""Office Connector — let Hermes work in the Word, Excel and PowerPoint
documents the user already has open.

This plugin is a thin adapter. Everything it exposes lives in
``scripts/office_kernel.py`` and ``scripts/office.py``, which are the same
files the portable `office-live` Agent Skill ships; `tools/sync_plugins.py`
copies them here and CI fails if the two ever drift. Edit the skill, not this
copy.

Why a plugin as well as a skill: a skill is instructions, and the agent still
has to run them through a code tool. Registering the three functions as Hermes
tools puts them in the tool catalogue, keeps one warm COM connection for the
whole session, and means the user installs one thing instead of wiring an MCP
server into `config.yaml` by hand.

Windows only. COM is the mechanism, so there is nothing to fall back to
elsewhere; the manifest says so and Hermes will not load this on another
platform.
"""
from pathlib import Path
import sys

_SCRIPTS = Path(__file__).parent / "scripts"

EXEC_SCHEMA = {
    "type": "object",
    "properties": {
        "code": {
            "type": "string",
            "description": (
                "Python to run against the live Office objects. The namespace "
                "persists between calls, so attach once and reuse the "
                "variables: app, doc = attach('word'|'excel'|'ppt'). "
                "print() what you want to see back."
            ),
        },
    },
    "required": ["code"],
}

NO_ARGS_SCHEMA = {"type": "object", "properties": {}}


def register(ctx):
    # The scripts directory has to be importable before the module that lives
    # in it, and `office_kernel` itself imports its sibling `office`.
    if str(_SCRIPTS) not in sys.path:
        sys.path.insert(0, str(_SCRIPTS))
    import office_kernel

    # One STA thread owns every COM object for the life of the session. Started
    # here rather than lazily so the first tool call is as fast as the rest.
    office_kernel.start_com_worker()

    # Hermes dispatches plugin tools as ``handler(args_dict, **context)``
    # (tools/registry.py: ``entry.handler(args, **kwargs)``), so a handler must
    # take the arguments dict positionally. Registering the typed functions
    # directly made every call fail with "takes 0 positional arguments but 1
    # was given" — and office_exec received a dict where it expected a string.
    def _status(args=None, **_):
        return office_kernel.office_status()

    def _exec(args=None, **_):
        return office_kernel.office_exec((args or {}).get("code", ""))

    def _reset(args=None, **_):
        return office_kernel.office_reset()

    ctx.register_tool(
        "office_status", "office", NO_ARGS_SCHEMA, _status,
        description=(
            "List which Office applications are running and which documents are "
            "open: name, path, saved state, whether a backup is possible, any "
            "crash-recovery warning, and for Word the revision state. Also "
            "classic Outlook: running, accounts, unread and draft counts. Call "
            "this first, and again whenever it is unclear what is active."
        ),
        emoji="\N{MEMO}",
    )

    ctx.register_tool(
        "office_exec", "office", EXEC_SCHEMA, _exec,
        description=(
            # The tool-search catalogue shows only the first sentence, cut at 60
            # characters, so it has to stand on its own.
            "Work live in open Word, Excel, PowerPoint and Outlook. "
            "Runs Python against the running Office over COM: the open "
            "document, and classic Outlook — search "
            "mail, read messages, compose new mail, replies and forwards with "
            "attachments and formatting, calendar and contacts. Outlook items "
            "are composed and opened for the user; sending and saving are "
            "refused, the user presses Send. Word edits are tracked "
            "revisions. Helpers are preloaded (find_mail, new_mail, reply, "
            "calendar, open_doc, replace_all, com_date, ...) and the "
            "namespace persists between calls."
        ),
        emoji="\N{PENCIL}",
    )

    ctx.register_tool(
        "office_reset", "office", NO_ARGS_SCHEMA, _reset,
        description=(
            "Drop every variable and COM reference held between calls. Use when "
            "Office was closed and reopened and calls start failing with stale "
            "object errors."
        ),
        emoji="\N{ANTICLOCKWISE OPEN CIRCLE ARROW}",
    )

    # The skill document as reference material. Hermes keeps a plugin-registered
    # skill out of `<available_skills>`, so this is not how the agent discovers
    # the capability -- the three tools above are. It is here for when the agent
    # needs the recipes, the COM pitfalls and the worked examples.
    skill = Path(__file__).parent / "SKILL.md"
    if skill.exists():
        ctx.register_skill(
            "office-live", skill,
            description="Recipes and pitfalls for driving the open Office document over COM.",
        )
