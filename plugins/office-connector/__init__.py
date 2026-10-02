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

    ctx.register_tool(
        "office_status", "office", NO_ARGS_SCHEMA, office_kernel.office_status,
        description=(
            "List which Office applications are running and which documents are "
            "open: name, path, saved state, and for Word whether revision "
            "tracking is on and how many tracked revisions exist. Call this "
            "first, and again whenever it is unclear which document is active."
        ),
        emoji="\N{MEMO}",
    )

    ctx.register_tool(
        "office_exec", "office", EXEC_SCHEMA, office_kernel.office_exec,
        description=(
            "Run Python against the Word, Excel or PowerPoint document the user "
            "has open, over COM. Use for reading or changing a document in "
            "place rather than producing a new file. Word edits are made as "
            "tracked revisions. The namespace persists between calls."
        ),
        emoji="\N{PENCIL}",
    )

    ctx.register_tool(
        "office_reset", "office", NO_ARGS_SCHEMA, office_kernel.office_reset,
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
