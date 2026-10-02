"""Microsoft Graph Connector — give Hermes read access to the signed-in
user's own Microsoft 365 content.

A thin adapter, like its sibling. Everything it exposes lives in
``scripts/microsoft_graph.py``, the same file the portable `m365-graph` Agent
Skill ships; `tools/sync_plugins.py` copies it here and CI fails if the two
drift. Edit the skill, not this copy.

The read-only posture is enforced in that module, not here: `graph()` refuses
any non-GET call unless it names a write area from a fixed list and the
request path belongs to that area. This plugin exposes no write tool at all,
so nothing it registers can change anything in the tenant.
"""
from pathlib import Path
import json
import sys

_SCRIPTS = Path(__file__).parent / "scripts"

SEARCH_SCHEMA = {
    "type": "object",
    "properties": {
        "query": {
            "type": "string",
            "description": (
                "Keyword query (KQL): bare words, or from:, subject:, "
                "filetype:, \"exact phrase\", AND/OR/NOT. Keyword matching, "
                "not semantic -- use words that literally appear."
            ),
        },
        "types": {
            "type": "string",
            "description": (
                "Comma-separated entity types to search. message (mail), "
                "driveItem and listItem (OneDrive and SharePoint), chatMessage "
                "(Teams chats), event, person. "
                "Default: message,driveItem,chatMessage"
            ),
        },
        "top": {"type": "integer", "description": "How many results per type. Default 15."},
    },
    "required": ["query"],
}

GET_SCHEMA = {
    "type": "object",
    "properties": {
        "path": {
            "type": "string",
            "description": "Graph path beginning with a slash, for example /me/drive/recent.",
        },
        "params": {
            "type": "object",
            "description": "Optional query string parameters, for example {\"$top\": 5}.",
        },
    },
    "required": ["path"],
}

NO_ARGS_SCHEMA = {"type": "object", "properties": {}}


def _load():
    if str(_SCRIPTS) not in sys.path:
        sys.path.insert(0, str(_SCRIPTS))
    import microsoft_graph
    return microsoft_graph


def register(ctx):
    mg = _load()

    def m365_status() -> str:
        """Signed in? Which scopes were granted, which are missing?"""
        try:
            granted = sorted(mg.granted_scopes())
        except Exception as error:  # not signed in is the common case, not a crash
            return json.dumps({"signed_in": False, "detail": str(error)}, ensure_ascii=False)
        missing = [s for s in mg.READ_SCOPES if s not in granted]
        return json.dumps(
            {"signed_in": True, "granted": granted, "missing": missing},
            ensure_ascii=False,
        )

    def m365_login() -> str:
        """Sign in. Opens a browser window, so the user has to be present."""
        mg.login()
        return m365_status()

    def m365_search(query: str, types: str = "", top: int = 15) -> str:
        """Search the user's own mail, files, SharePoint and Teams chats."""
        kinds = [t.strip() for t in types.split(",") if t.strip()] or None
        return json.dumps(mg.search(query, types=kinds, top=top), ensure_ascii=False, default=str)

    def m365_get(path: str, params: dict | None = None) -> str:
        """Any Microsoft Graph GET. Writes are refused inside the module."""
        return json.dumps(mg.graph("GET", path, params=params), ensure_ascii=False, default=str)

    ctx.register_tool(
        "m365_status", "m365", NO_ARGS_SCHEMA, m365_status,
        description=(
            "Check whether the user is signed in to Microsoft 365 and which "
            "permissions were granted. Call this before the first search, so a "
            "missing consent is reported rather than looking like an empty result."
        ),
        emoji="\N{IDENTIFICATION CARD}",
    )

    ctx.register_tool(
        "m365_login", "m365", NO_ARGS_SCHEMA, m365_login,
        description=(
            "Sign in to Microsoft 365. Opens a browser window and needs the user "
            "present, so only call it when m365_status says they are not signed in."
        ),
        emoji="\N{KEY}",
    )

    ctx.register_tool(
        "m365_search", "m365", SEARCH_SCHEMA, m365_search,
        description=(
            "Search the signed-in user's own Outlook mail, OneDrive and "
            "SharePoint documents, Teams chats and calendar. Returns only what "
            "that user can already see. Read-only."
        ),
        emoji="\N{LEFT-POINTING MAGNIFYING GLASS}",
    )

    ctx.register_tool(
        "m365_get", "m365", GET_SCHEMA, m365_get,
        description=(
            "Fetch any Microsoft Graph resource by path when the search tool is "
            "too blunt, for example /me/drive/recent or a specific message. "
            "Read-only: non-GET calls are refused."
        ),
        emoji="\N{GLOBE WITH MERIDIANS}",
    )

    skill = Path(__file__).parent / "SKILL.md"
    if skill.exists():
        ctx.register_skill(
            "m365-graph", skill,
            description="Query syntax, entity types and failure modes for Microsoft 365 search.",
        )
