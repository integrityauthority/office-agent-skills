"""Warm Office COM kernel, exposed as a 3-tool MCP server.

Why this exists
---------------
Cold-starting Python + COM for every edit costs 250-350 ms and, much worse,
makes the model re-write the same attach boilerplate every call — small local
models get that wrong often. An MCP stdio server is a process the host keeps
alive for the whole session, so it doubles as a warm kernel:

  * COM objects are attached once and reused,
  * the exec namespace PERSISTS between calls (define a variable in one call,
    use it in the next — like a Jupyter kernel),
  * only 3 tools are exported, so the tool catalog stays tiny.

Runs on the Hermes venv's own interpreter: `mcp` and `pywin32` are already
there, so deployment is this one file plus a config block.

COM threading
-------------
COM objects must be used from the thread that created them. FastMCP handlers
run on asyncio worker threads, so all COM work is funneled through ONE
dedicated STA thread via a queue. This also serializes access, which is what
Office wants anyway.
"""

from __future__ import annotations

import ast
import contextlib
import io
import json
import queue
import re
import sys
import threading
import traceback

from mcp.server.fastmcp import FastMCP

sys.stderr.reconfigure(errors="replace")

# --------------------------------------------------------------- COM thread ---

_requests: "queue.Queue[tuple]" = queue.Queue()

# The persistent namespace all office_exec calls share. Lives in the COM
# thread; nothing outside that thread may touch COM objects stored here.
_NS: dict = {}

# Edits that destroy the human review trail or the user's work. The kernel is
# for a colleague-style agent working under revision tracking; these calls
# have no legitimate place in that loop, so they are refused outright.
_FORBIDDEN = [
    (r"\bAcceptAll(Revisions)?\b", "accepting revisions erases the review trail"),
    (r"\bRejectAll(Revisions)?\b", "rejecting revisions erases the review trail"),
    (r"\.Accept\(\)", "accepting a revision is the human reviewer's call"),
    (r"\.Reject\(\)", "rejecting a revision is the human reviewer's call"),
    (r"\bSaveAs2?\b", "SaveAs can silently overwrite another file"),
    (r"\.Close\(", "closing documents is the user's call"),
    (r"\bQuit\(", "quitting the application is the user's call"),
    (r"TrackRevisions\s*=\s*(False|0)", "never turn revision tracking off"),
    (r"DisplayAlerts\s*=", "suppressing Office alerts hides destructive prompts"),
]


def _seed_namespace():
    """(Re)build the persistent namespace with the helpers preloaded."""
    import office  # noqa: the sibling module; sys.path is set in main()

    _NS.clear()
    _NS.update({
        "attach": office.attach,
        "OfficeError": office.OfficeError,
        "json": json,
    })


def _run_snippet(code: str) -> dict:
    """Exec `code` in the persistent namespace, Jupyter-style: statements run,
    and if the last node is an expression its value becomes the result."""
    stdout = io.StringIO()
    tree = ast.parse(code, mode="exec")
    last_expr = None
    if tree.body and isinstance(tree.body[-1], ast.Expr):
        last_expr = ast.Expression(tree.body.pop(-1).value)

    with contextlib.redirect_stdout(stdout):
        exec(compile(tree, "<office_exec>", "exec"), _NS)
        value = (eval(compile(last_expr, "<office_exec>", "eval"), _NS)
                 if last_expr is not None else _NS.get("result"))

    try:
        json.dumps(value)
    except TypeError:
        value = repr(value)
    return {"ok": True, "result": value, "stdout": stdout.getvalue()[-4000:]}


def _com_worker():
    import pythoncom
    pythoncom.CoInitialize()
    _seed_namespace()
    while True:
        code, reply = _requests.get()
        try:
            if code is None:                      # reset request
                _seed_namespace()
                reply.put({"ok": True, "result": "namespace reset"})
                continue
            reply.put(_run_snippet(code))
        except SyntaxError as e:
            reply.put({"ok": False, "error": f"SyntaxError: {e}",
                       "hint": "send plain Python; the last expression's "
                               "value is returned automatically"})
        except Exception as e:
            hint = ""
            import office
            if isinstance(e, office.OfficeError):
                hint = e.hint
            code_attr = getattr(e, "hresult", None)
            if code_attr == -2147418111:
                hint = ("Office rejected the call — usually a modal dialog is "
                        "open. Ask the user to dismiss it; do not retry blindly.")
            reply.put({"ok": False,
                       "error": f"{type(e).__name__}: {e}",
                       "hint": hint,
                       "traceback": traceback.format_exc(limit=3)})


def _submit(code):
    reply: "queue.Queue[dict]" = queue.Queue()
    _requests.put((code, reply))
    try:
        return reply.get(timeout=120)
    except queue.Empty:
        return {"ok": False,
                "error": "the COM call did not return within 120s",
                "hint": "a modal dialog is probably open in Office — ask the "
                        "user to dismiss it"}


# --------------------------------------------------------------------- MCP ---

mcp = FastMCP("office-kernel")


@mcp.tool()
def office_exec(code: str) -> str:
    """Run Python against the LIVE Office apps (Word/Excel/PowerPoint) open on
    this machine. The namespace PERSISTS between calls, so attach once and
    keep using the variables:

        app, doc = attach("word")        # or "excel" / "ppt"; name="..." picks one
        doc.TrackRevisions = True
        doc.Revisions.Count              # last expression = return value

    `attach(host, name=None)` and `json` are preloaded. Full COM object model
    is available — write COM calls directly, there is no wrapper layer.
    Edits appear in the user's open window immediately. ALWAYS verify a write
    by reading the affected region back in a follow-up call.
    Refused for safety: accepting/rejecting revisions, SaveAs, Close, Quit,
    disabling TrackRevisions or DisplayAlerts.
    """
    for pattern, why in _FORBIDDEN:
        if re.search(pattern, code):
            return json.dumps({
                "ok": False,
                "error": f"refused: {why}",
                "hint": "this kernel only ADDS tracked edits; leave review "
                        "decisions and file lifecycle to the user",
            }, ensure_ascii=False)
    return json.dumps(_submit(code), ensure_ascii=False)


@mcp.tool()
def office_status() -> str:
    """List which Office apps are running and which documents are open (name,
    path, saved state; for Word also whether revision tracking is on and how
    many tracked revisions exist). Call this FIRST, and re-check when unsure
    which document is active."""
    snippet = """
import office as _o
_report = {}
for _host in ("word", "excel", "ppt"):
    try:
        _app, _ = None, None
        _app = _o.get_app(_host)
        _docs = []
        for _d in _o.collection(_app, _host):
            _item = {"name": _d.Name, "path": getattr(_d, "FullName", ""),
                     "saved": bool(_d.Saved)}
            if _host == "word":
                _item["track_revisions"] = bool(_d.TrackRevisions)
                _item["revisions"] = _d.Revisions.Count
            _docs.append(_item)
        _report[_host] = {"running": True, "documents": _docs}
    except Exception as _e:
        _report[_host] = {"running": False, "note": str(_e)}
result = _report
"""
    return json.dumps(_submit(snippet), ensure_ascii=False)


@mcp.tool()
def office_reset() -> str:
    """Clear the persistent namespace (drops all variables and stale COM
    references). Use when Office was closed/reopened and calls fail with
    stale-object errors."""
    return json.dumps(_submit(None), ensure_ascii=False)


def main():
    import os
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    threading.Thread(target=_com_worker, daemon=True, name="com-sta").start()
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
