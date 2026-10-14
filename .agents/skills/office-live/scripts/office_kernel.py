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
import sys
import threading
import traceback

sys.stderr.reconfigure(errors="replace")

# --------------------------------------------------------------- COM thread ---

_requests: "queue.Queue[tuple]" = queue.Queue()

# The persistent namespace all office_exec calls share. Lives in the COM
# thread; nothing outside that thread may touch COM objects stored here.
_NS: dict = {}



def _seed_namespace():
    """(Re)build the persistent namespace with the helpers preloaded."""
    import office  # noqa: the sibling module; sys.path is set in main()

    # office._seed is the single list of what is preloaded. This module used
    # to keep its own three-name copy (attach, OfficeError, json), so every
    # helper added since — open_doc, close_doc, replace_all, probe, the
    # Outlook functions — was missing from the plugin, and the agent imported
    # them by hand in every call.
    office._seed(_NS)


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


# ------------------------------------------------------------------- tools ---
#
# Plain functions, deliberately. They were decorated with `@mcp.tool()` here,
# which made importing this module require an MCP server library -- and it
# made the module unimportable outright once `mcp.server.fastmcp` disappeared
# in mcp 2.0. Nothing about driving Office needs MCP. `main()` below wraps
# these three for the clients that speak it; a Hermes plugin registers the
# same functions directly.


def office_exec(code: str) -> str:
    """Run Python against the LIVE Office apps (Word/Excel/PowerPoint) open on
    this machine. The namespace PERSISTS between calls, so attach once and
    keep using the variables:

        app, doc = attach("word")        # or "excel" / "ppt"; name="..." picks one
        doc.TrackRevisions = True
        doc.Revisions.Count              # last expression = return value

    Preloaded: attach, status, open_doc/new_doc/save_as/close_doc,
    replace_all, probe, explore, and for Outlook find_mail, get_mail,
    current_mail, new_mail, reply, forward, calendar, new_appointment,
    resolve, find_contact. Full COM object model otherwise — write COM calls
    directly.
    Edits appear in the user's open window immediately. ALWAYS verify a write
    by reading the affected region back in a follow-up call.
    Refused for safety: accepting/rejecting revisions, raw SaveAs/Close (use
    save_as/close_doc), Quit, disabling TrackRevisions or DisplayAlerts, and
    anything that sends, saves or deletes in Outlook — compose and .Display()
    instead, the user presses Send.
    """
    # The guard lives in office.refuse. This module used to carry its own copy
    # of the pattern list, which drifted: the plugin kept refusing Close with
    # the old wording and no pointer to close_doc after office.py had moved on.
    import office
    refusal = office.refuse(code)
    if refusal:
        return json.dumps(refusal, ensure_ascii=False)
    return json.dumps(_submit(code), ensure_ascii=False)


def office_status() -> str:
    """Which Office applications are running and what is open — Word, Excel,
    PowerPoint documents with their storage, backup possibility and any
    AutoRecovery warning, plus Outlook (counts only). Call this FIRST, and
    again whenever it is unclear which document is active."""
    # office.status() is the one implementation. This used to be a private
    # snippet that predated the storage/backup fields, the AutoRecovery
    # warning, the COM binding report and Outlook.
    return json.dumps(_submit("import office as _o\nresult = _o.status()"),
                      ensure_ascii=False, default=str)


def office_reset() -> str:
    """Clear the persistent namespace (drops all variables and stale COM
    references). Use when Office was closed/reopened and calls fail with
    stale-object errors."""
    return json.dumps(_submit(None), ensure_ascii=False)


def start_com_worker():
    """Start the single STA thread every COM call is funnelled through.

    Idempotent: a second call is a no-op, so a host that loads this module more
    than once does not end up with two threads fighting over the same objects.
    """
    for thread in threading.enumerate():
        if thread.name == "com-sta":
            return
    threading.Thread(target=_com_worker, daemon=True, name="com-sta").start()


def _build_server():
    """The MCP server, from whichever library this installation has.

    `mcp.server.fastmcp` was the home of FastMCP up to mcp 1.x; mcp 2.0 renamed
    it to `mcp.server.mcpserver.MCPServer` with the same `tool()`/`run()`
    surface. Importing it here rather than at module scope is what keeps this
    file usable without any MCP library at all.
    """
    try:
        from mcp.server.fastmcp import FastMCP as Server
    except ModuleNotFoundError:
        from mcp.server.mcpserver import MCPServer as Server
    server = Server("office-kernel")
    for fn in (office_exec, office_status, office_reset):
        server.tool()(fn)
    return server


def main():
    import os
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    server = _build_server()
    start_com_worker()
    server.run(transport="stdio")


if __name__ == "__main__":
    main()
