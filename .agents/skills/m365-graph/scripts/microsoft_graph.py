"""Search and read Microsoft 365 content with the signed-in user's own rights.

Mail, calendar, OneDrive, SharePoint and Teams chats over Microsoft Graph,
using DELEGATED permissions: the user signs in through a normal browser
window and the agent sees exactly what that person can already see, nothing
more.

Every command prints exactly one JSON object and exits 0 on success, 1 on
failure — same contract as office.py:

    {"ok": true,  "data": ...}
    {"ok": false, "error": "...", "hint": "..."}

Design notes for whoever maintains this:

* Reads are open, writes are CLOSED AT THE TRANSPORT. `graph()` refuses any
  non-GET call unless the caller names a write area AND the path matches that
  area's pattern AND the token actually holds the scope. A skill that only
  searches cannot send mail even if the model decides to try, because the
  refusal happens below the model, not in prose. `POST /search/query` is
  whitelisted — it is the search API, it changes nothing.
* Scopes are granted INCREMENTALLY and remembered. Requesting the full read
  set on every silent call would fail for anyone who consented to a subset,
  so `login` records what consent actually returned and later calls ask for
  exactly that. Missing scope => a message naming the `login --add-scope`
  command to fix it, never a bare 403.
* The Search API's entityTypes cannot be freely mixed (mail, Teams messages,
  events and SharePoint items each want their own request). `search` splits
  the requested types into compatible groups, fans out, and merges. Callers
  should not have to know this.
* Search responses are enormous — one hit carries the whole resource. We
  compact per type down to the fields a human would read, and keep --raw for
  when that is not enough. This is the difference between a search that fits
  in a model's context and one that does not.
* The token cache is encrypted with Windows DPAPI, so the file is useless to
  any other account on the machine. It is not a secret store; it is a cache.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import re
import sys

GRAPH = "https://graph.microsoft.com/v1.0"
AUTHORITY = "https://login.microsoftonline.com/organizations"

# Microsoft's own "Graph Command Line Tools" public client — the app id the
# official Connect-MgGraph uses. It is registered in practically every tenant
# with an http://localhost redirect, which is why interactive sign-in works
# here with no app registration of your own.
#
# The trade-off is the audit trail: Entra sign-in logs will name that app, not
# Hermes. For anything beyond trying this out, register your own app (two
# minutes, no admin rights in a default tenant) and set MSGRAPH_CLIENT_ID.
DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

RUNTIME_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                           "m365-graph")
CACHE_PATH = os.path.join(RUNTIME_DIR, "token-cache.bin")
STATE_PATH = os.path.join(RUNTIME_DIR, "state.json")

# Read scopes a normal user can consent to themselves in a default tenant.
# Deliberately excludes anything that always needs an administrator — notably
# ChannelMessage.Read.All (Teams *channel* posts). Chat.Read covers the 1:1
# and group chats the user is in, which is what "search Teams" usually means.
READ_SCOPES = [
    "User.Read",
    "Mail.Read",
    "Calendars.Read",
    "Files.Read.All",
    "Sites.Read.All",
    "Chat.Read",
    "People.Read",
]

# The smallest useful set — try this when the full set is refused, to find out
# whether consent is blocked outright or only for the broader scopes.
MINIMAL_SCOPES = ["User.Read", "Mail.Read"]

# Write areas. Each entry: path pattern the write may target, and the scope
# the token must hold. Adding an area here is a deliberate act; nothing is
# writable by default.
WRITE_AREAS = {
    "calendar": (re.compile(r"^/me/(events|calendars?/[^/]+/events)"),
                 "Calendars.ReadWrite"),
    "mail-send": (re.compile(r"^/me/(sendMail$|messages/[^/]+/(reply|replyAll|forward|send)$)"),
                  "Mail.Send"),
    "files": (re.compile(r"^/me/drive/"), "Files.ReadWrite.All"),
}

# Non-GET calls that change nothing. The search API is a POST because the
# query is a document, not because it writes.
SAFE_POSTS = (re.compile(r"^/search/query$"),)

# Search entityTypes that may travel in one request. Graph rejects most
# cross-group combinations, so we never build one.
SEARCH_GROUPS = [
    ["driveItem", "drive", "listItem", "list", "site"],   # OneDrive + SharePoint
    ["message"],                                          # Outlook mail
    ["chatMessage"],                                      # Teams chats
    ["event"],                                            # calendar
    ["person"],
]

# AADSTS codes worth translating. Everything else falls through with its own
# text, which Microsoft writes reasonably well.
AAD_HINTS = {
    "AADSTS65001": ("consent was not granted for these permissions",
                    "Run 'login --minimal' to see whether a smaller set is "
                    "allowed. If even that fails, user consent is disabled in "
                    "the tenant and an administrator has to approve the app."),
    "AADSTS90094": ("this permission needs administrator approval",
                    "Drop the scope that triggered it and retry — "
                    "'login --minimal' is the safe baseline. Teams CHANNEL "
                    "search (ChannelMessage.Read.All) always needs an admin."),
    "AADSTS700016": ("the application is not present in this tenant",
                     "Set MSGRAPH_CLIENT_ID to an app registered in your own "
                     "tenant, or ask for the default client to be enabled."),
    "AADSTS50076": ("multi-factor authentication is required",
                    "Complete the MFA prompt in the browser window, then "
                    "retry. Do not retry without the user present."),
    "AADSTS530003": ("device compliance policy blocked the sign-in",
                     "Conditional Access requires a managed device. This "
                     "cannot be worked around from here — tell the user."),
    "AADSTS50105": ("the user is not assigned to this application",
                    "An administrator restricted who may use the app."),
}


class GraphError(RuntimeError):
    """Raised by the library API; the CLI turns it into a JSON error object."""

    def __init__(self, message, hint=""):
        super().__init__(message)
        self.hint = hint


def _utf8_stdout():
    """Force UTF-8 on stdout.

    Python on Windows writes stdout in the console codepage, which is cp852
    here: an em-dash in a hint comes out as '?', and a Hungarian accent in a
    document title raises UnicodeEncodeError mid-print, so the caller gets
    half a JSON object and a traceback.
    """
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def out(data):
    _utf8_stdout()
    print(json.dumps({"ok": True, "data": data}, ensure_ascii=False, default=str))
    sys.exit(0)


def fail(error, hint=""):
    _utf8_stdout()
    print(json.dumps({"ok": False, "error": str(error), "hint": hint},
                     ensure_ascii=False))
    sys.exit(1)


# -------------------------------------------------------------------- auth ---

def _msal():
    try:
        import msal
        return msal
    except ImportError:
        raise GraphError(
            "the msal package is not installed",
            "Install it into the interpreter Hermes uses: "
            "'<hermes venv>\\Scripts\\python.exe -m pip install msal', or "
            "re-run Setup-Hermes.ps1 which does it for you.")


def client_id():
    return os.environ.get("MSGRAPH_CLIENT_ID", "").strip() or DEFAULT_CLIENT_ID


def _dpapi(data, unprotect=False):
    """Encrypt/decrypt with the current Windows user's DPAPI key."""
    try:
        import win32crypt
    except ImportError:
        return data                     # no pywin32: cache stays plaintext
    if unprotect:
        return win32crypt.CryptUnprotectData(data, None, None, None, 0)[1]
    return win32crypt.CryptProtectData(data, "hermes m365 token cache",
                                       None, None, None, 0)


def _load_cache():
    msal = _msal()
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_PATH):
        try:
            with open(CACHE_PATH, "rb") as fh:
                cache.deserialize(_dpapi(fh.read(), unprotect=True).decode("utf-8"))
        except Exception:
            # A cache from another user, another machine, or a rotated DPAPI
            # key is unreadable, not corrupt-and-fatal: sign in again.
            pass
    return cache


def _save_cache(cache):
    if not cache.has_state_changed:
        return
    os.makedirs(RUNTIME_DIR, exist_ok=True)
    blob = _dpapi(cache.serialize().encode("utf-8"))
    with open(CACHE_PATH, "wb") as fh:
        fh.write(blob)


def _state():
    try:
        with open(STATE_PATH, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _save_state(**kw):
    st = _state()
    st.update(kw)
    os.makedirs(RUNTIME_DIR, exist_ok=True)
    with open(STATE_PATH, "w", encoding="utf-8") as fh:
        json.dump(st, fh, ensure_ascii=False, indent=2)


def granted_scopes():
    """Scopes consent actually returned at the last sign-in."""
    return _state().get("granted_scopes", [])


def _app(cache):
    msal = _msal()
    return msal.PublicClientApplication(client_id(), authority=AUTHORITY,
                                        token_cache=cache)


def _aad_error(result):
    desc = result.get("error_description") or result.get("error") or "sign-in failed"
    for code, (msg, hint) in AAD_HINTS.items():
        if code in desc:
            return GraphError(f"{msg} ({code})", hint)
    first = desc.strip().splitlines()[0][:400]
    return GraphError(f"sign-in failed: {first}",
                      "Read the message above — Microsoft usually names the "
                      "exact policy or permission at fault.")


def login(scopes=None, interactive=True):
    """Acquire a token, opening a browser window only when necessary."""
    scopes = list(scopes or granted_scopes() or READ_SCOPES)
    cache = _load_cache()
    app = _app(cache)
    result = None

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        if not interactive:
            raise GraphError(
                "no usable token for the requested scopes",
                "Someone has to sign in interactively first. Ask the user to "
                "run the login step; a browser window will open.")
        result = app.acquire_token_interactive(
            scopes=scopes,
            prompt="select_account" if not accounts else None,
        )

    if "access_token" not in result:
        _save_cache(cache)
        raise _aad_error(result)

    _save_cache(cache)
    got = (result.get("scope") or "").split()
    # Keep the union: consent is incremental, so a later --add-scope must not
    # make us forget what we already hold.
    merged = sorted(set(granted_scopes()) | set(s for s in got
                                                if not s.startswith("openid")
                                                and s not in ("profile", "email",
                                                              "offline_access")))
    account = (result.get("id_token_claims") or {})
    _save_state(granted_scopes=merged,
                account=account.get("preferred_username") or
                        (accounts[0]["username"] if accounts else None),
                tenant=account.get("tid"),
                client_id=client_id())
    return result["access_token"], merged


def _token(need_scope=None):
    """Silent token for the scopes we already hold. Never opens a browser."""
    have = granted_scopes()
    if not have:
        raise GraphError(
            "not signed in to Microsoft 365",
            "Run the login command first: it opens a browser window where the "
            "user signs in with their work account. Nothing is cached until "
            "they do.")
    if need_scope and need_scope not in have:
        raise GraphError(
            f"the current token does not include {need_scope}",
            f"Ask the user to approve it: "
            f"'microsoft_graph.py login --add-scope {need_scope}'. That opens "
            f"a consent window; it is their decision, not yours to assume.")
    token, _ = login(scopes=have, interactive=False)
    return token


# -------------------------------------------------------------------- http ---

def graph(method, path, *, params=None, body=None, allow_write=None,
          headers=None, timeout=60, all_pages=False, page_cap=10):
    """Call Microsoft Graph. Reads are open; writes must be named explicitly.

    Import this from execute_code to go beyond the typed commands:

        import sys; sys.path.insert(0, r"<skill>/scripts")
        from microsoft_graph import graph
        graph("GET", "/me/messages", params={"$top": 5})

    Writes need `allow_write=<area>` from WRITE_AREAS, and the path has to
    belong to that area — otherwise the call is refused here, before any
    request goes out.
    """
    import requests

    method = method.upper()
    path = "/" + path.lstrip("/")
    need = None

    if method != "GET":
        if path == "/$batch":
            inner = ((body or {}).get("requests") or [])
            bad = [r.get("method", "?") for r in inner
                   if str(r.get("method", "")).upper() != "GET"]
            if bad:
                raise GraphError(
                    f"refused: $batch may only carry GET sub-requests (found {bad})",
                    "Batching is for fanning out reads. Send writes one by "
                    "one through the area they belong to.")
        elif method == "POST" and any(p.match(path) for p in SAFE_POSTS):
            pass
        elif allow_write is None:
            raise GraphError(
                f"refused: {method} {path} is a write and no area was allowed",
                "This helper is read-only unless the caller passes "
                f"allow_write=<area>, one of {sorted(WRITE_AREAS)}. Tell the "
                "user what you intend to change and let them confirm first.")
        else:
            if allow_write not in WRITE_AREAS:
                raise GraphError(f"unknown write area {allow_write!r}",
                                 f"known areas: {sorted(WRITE_AREAS)}")
            pattern, need = WRITE_AREAS[allow_write]
            if not pattern.match(path):
                raise GraphError(
                    f"refused: {path} is outside the {allow_write!r} area",
                    f"{allow_write!r} only covers paths matching "
                    f"{pattern.pattern}. Pick the right area rather than a "
                    "wider one.")

    token = _token(need_scope=need)
    hdrs = {"Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json"}
    hdrs.update(headers or {})

    url = path if path.startswith("http") else GRAPH + path
    collected, pages = None, 0

    while True:
        for attempt in (1, 2, 3):
            r = requests.request(method, url, headers=hdrs, params=params,
                                 json=body, timeout=timeout)
            if r.status_code == 429 or r.status_code >= 500:
                wait = int(r.headers.get("Retry-After", 2 * attempt))
                if attempt < 3:
                    import time
                    time.sleep(min(wait, 20))
                    continue
            break

        if r.status_code == 401:
            raise GraphError("Microsoft Graph rejected the token (401)",
                             "The token expired or was revoked. Sign in again "
                             "with the login command.")
        if r.status_code == 403:
            raise GraphError(
                f"Microsoft Graph denied the request (403): {r.text[:300]}",
                "Either the token lacks the scope for this path, or the user "
                "genuinely has no access to that mailbox/site. Check "
                "'scopes' before assuming it is a permission bug.")
        if r.status_code == 404:
            raise GraphError(f"not found: {path}",
                             "Check the id or path. A deleted item and a "
                             "mistyped id look the same from here.")
        if not r.ok:
            raise GraphError(f"HTTP {r.status_code}: {r.text[:400]}")

        payload = r.json() if r.content else {}
        params = None                          # nextLink carries its own query

        if not all_pages:
            return payload
        if collected is None:
            collected = payload
        else:
            collected.setdefault("value", []).extend(payload.get("value", []))
        pages += 1
        url = payload.get("@odata.nextLink")
        if not url or pages >= page_cap:
            if url:
                collected["truncated_after_pages"] = pages
            collected.pop("@odata.nextLink", None)
            return collected


# ------------------------------------------------------------------ search ---

def _text(html, limit=600):
    """Crude HTML to text. Good enough for previews, never for reproduction."""
    if not html:
        return ""
    s = re.sub(r"(?is)<(script|style).*?</\1>", " ", html)
    s = re.sub(r"(?s)<[^>]+>", " ", s)
    s = (s.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<")
          .replace("&gt;", ">").replace("&quot;", '"').replace("&#39;", "'"))
    return re.sub(r"\s+", " ", s).strip()[:limit]


def _compact(hit):
    """One search hit down to the fields a person would actually read."""
    res = hit.get("resource") or {}
    kind = (res.get("@odata.type") or "").rsplit(".", 1)[-1]
    rec = {"type": kind or "unknown", "rank": hit.get("rank")}
    summary = _text(hit.get("summary"), 300)

    if kind == "message":
        frm = ((res.get("from") or {}).get("emailAddress") or {})
        rec.update(subject=res.get("subject"),
                   sender=frm.get("address") or frm.get("name"),
                   received=res.get("receivedDateTime"),
                   has_attachments=res.get("hasAttachments"),
                   preview=summary or _text((res.get("body") or {}).get("content"), 300),
                   id=res.get("id"), link=res.get("webLink"))
    elif kind in ("driveItem", "listItem", "drive", "list", "site"):
        parent = (res.get("parentReference") or {})
        rec.update(name=res.get("name") or res.get("displayName"),
                   modified=res.get("lastModifiedDateTime"),
                   size=res.get("size"),
                   location=parent.get("path") or parent.get("siteId"),
                   preview=summary,
                   id=res.get("id"), link=res.get("webUrl"))
    elif kind == "chatMessage":
        frm = ((res.get("from") or {}).get("user") or {})
        rec.update(sender=frm.get("displayName"),
                   created=res.get("createdDateTime"),
                   preview=summary or _text((res.get("body") or {}).get("content"), 300),
                   chat_id=res.get("chatId"), id=res.get("id"),
                   link=res.get("webUrl"))
    elif kind == "event":
        rec.update(subject=res.get("subject"),
                   start=(res.get("start") or {}).get("dateTime"),
                   organizer=(((res.get("organizer") or {}).get("emailAddress")
                               or {}).get("address")),
                   preview=summary, id=res.get("id"), link=res.get("webLink"))
    else:
        rec.update(name=res.get("displayName") or res.get("name"),
                   preview=summary, id=res.get("id"))
    return {k: v for k, v in rec.items() if v not in (None, "", [])}


def search(query, types=None, top=15, start=0, raw=False):
    """Search across M365. Types that Graph refuses to mix are fanned out."""
    wanted = [t.strip() for t in (types or ["message", "driveItem", "chatMessage"])
              if t.strip()]
    known = {t for g in SEARCH_GROUPS for t in g}
    unknown = [t for t in wanted if t not in known]
    if unknown:
        raise GraphError(f"unknown entityTypes: {unknown}",
                         f"pick from {sorted(known)}")

    # Fail fast on auth. Without this the per-group handler below turns one
    # "nobody signed in" into three identical failures wrapped in a hint about
    # SharePoint scopes, which sends the reader looking in the wrong place.
    _token()

    results, errors = [], {}
    for group in SEARCH_GROUPS:
        ask = [t for t in wanted if t in group]
        if not ask:
            continue
        body = {"requests": [{"entityTypes": ask,
                              "query": {"queryString": query},
                              "from": start, "size": top}]}
        try:
            payload = graph("POST", "/search/query", body=body)
        except GraphError as e:
            # One group failing (usually a missing scope) must not sink the
            # rest of the search — report it alongside what did work.
            errors[",".join(ask)] = str(e)
            continue
        for value in payload.get("value", []):
            for container in value.get("hitsContainers", []):
                for hit in container.get("hits", []):
                    results.append(hit if raw else _compact(hit))

    if not results and errors and len(errors) == len([
            g for g in SEARCH_GROUPS if any(t in wanted for t in g)]):
        raise GraphError("every part of the search failed: "
                         + json.dumps(errors, ensure_ascii=False),
                         "Check 'scopes' — SharePoint/OneDrive needs "
                         "Sites.Read.All, Teams chats need Chat.Read.")

    if not raw:
        results.sort(key=lambda r: (r.get("rank") or 999))
    return {"query": query, "types": wanted, "count": len(results),
            "hits": results, **({"partial_failures": errors} if errors else {})}


# -------------------------------------------------------------- commands ----

def cmd_status(a):
    st = _state()
    report = {
        "msal_installed": True,
        "client_id": client_id(),
        "using_default_client": client_id() == DEFAULT_CLIENT_ID,
        "signed_in": False,
        "account": st.get("account"),
        "granted_scopes": granted_scopes(),
        "cache": CACHE_PATH if os.path.exists(CACHE_PATH) else None,
    }
    try:
        _msal()
    except GraphError as e:
        report["msal_installed"] = False
        report["hint"] = e.hint
        out(report)
    if granted_scopes():
        try:
            me = graph("GET", "/me", params={
                "$select": "displayName,userPrincipalName,mail"})
            report["signed_in"] = True
            report["me"] = me
        except GraphError as e:
            report["token_error"] = str(e)
            report["hint"] = e.hint
    else:
        report["hint"] = ("Nobody has signed in yet. Run the login command — "
                          "a browser window opens for the user.")
    out(report)


def cmd_login(a):
    scopes = MINIMAL_SCOPES if a.minimal else list(READ_SCOPES)
    scopes = sorted(set(scopes) | set(granted_scopes()) | set(a.add_scope or []))
    _, merged = login(scopes=scopes, interactive=True)
    me = graph("GET", "/me", params={"$select": "displayName,userPrincipalName"})
    _save_state(account=me.get("userPrincipalName"))
    out({"signed_in_as": me.get("userPrincipalName"),
         "display_name": me.get("displayName"),
         "granted_scopes": merged,
         "note": "Token cached and encrypted for this Windows user. "
                 "Later runs are silent until it is revoked or expires."})


def cmd_logout(a):
    for path in (CACHE_PATH, STATE_PATH):
        try:
            os.remove(path)
        except FileNotFoundError:
            pass
    out({"removed": [CACHE_PATH, STATE_PATH],
         "note": "Local token cache cleared. The Entra-side consent still "
                 "stands — revoke that in the account's app permissions."})


def cmd_scopes(a):
    have = set(granted_scopes())
    out({"granted": sorted(have),
         "read_set_missing": [s for s in READ_SCOPES if s not in have],
         "write_areas": {k: {"scope": v[1], "paths": v[0].pattern,
                             "granted": v[1] in have}
                         for k, v in WRITE_AREAS.items()},
         "note": "Add one with: login --add-scope <Scope>. Teams CHANNEL "
                 "messages (ChannelMessage.Read.All) always need an admin."})


def cmd_whoami(a):
    out(graph("GET", "/me"))


def cmd_search(a):
    out(search(a.query, types=(a.types or "").split(",") if a.types else None,
               top=a.top, start=a.start, raw=a.raw))


def cmd_mail_list(a):
    folder = f"/me/mailFolders/{a.folder}/messages" if a.folder else "/me/messages"
    payload = graph("GET", folder, params={
        "$top": a.top, "$orderby": "receivedDateTime desc",
        "$select": "subject,from,receivedDateTime,hasAttachments,isRead,webLink"})
    items = []
    for m in payload.get("value", []):
        frm = ((m.get("from") or {}).get("emailAddress") or {})
        items.append({"subject": m.get("subject"),
                      "sender": frm.get("address"),
                      "received": m.get("receivedDateTime"),
                      "unread": not m.get("isRead"),
                      "has_attachments": m.get("hasAttachments"),
                      "id": m.get("id")})
    out({"folder": a.folder or "all", "count": len(items), "messages": items})


def cmd_mail_get(a):
    m = graph("GET", f"/me/messages/{a.id}")
    out({"subject": m.get("subject"),
         "sender": ((m.get("from") or {}).get("emailAddress") or {}).get("address"),
         "to": [((r.get("emailAddress") or {}).get("address"))
                for r in (m.get("toRecipients") or [])],
         "received": m.get("receivedDateTime"),
         "body": _text((m.get("body") or {}).get("content"), a.max_chars),
         "link": m.get("webLink")})


def cmd_calendar_list(a):
    start = _dt.datetime.now().astimezone()
    end = start + _dt.timedelta(days=a.days)
    payload = graph("GET", "/me/calendarView", params={
        "startDateTime": start.isoformat(), "endDateTime": end.isoformat(),
        "$orderby": "start/dateTime", "$top": 100,
        "$select": "subject,start,end,location,organizer,isAllDay,webLink"},
        headers={"Prefer": 'outlook.timezone="Europe/Budapest"'})
    events = [{"subject": e.get("subject"),
               "start": (e.get("start") or {}).get("dateTime"),
               "end": (e.get("end") or {}).get("dateTime"),
               "all_day": e.get("isAllDay"),
               "location": (e.get("location") or {}).get("displayName"),
               "organizer": (((e.get("organizer") or {}).get("emailAddress")
                              or {}).get("address")),
               "id": e.get("id")}
              for e in payload.get("value", [])]
    out({"from": start.date().isoformat(), "days": a.days,
         "count": len(events), "events": events})


def cmd_calendar_add(a):
    """The one write this CLI performs — deliberately narrow, and gated."""
    body = {
        "subject": a.subject,
        "start": {"dateTime": a.start, "timeZone": a.timezone},
        "end": {"dateTime": a.end, "timeZone": a.timezone},
    }
    if a.body:
        body["body"] = {"contentType": "text", "content": a.body}
    if a.location:
        body["location"] = {"displayName": a.location}
    if a.attendees:
        body["attendees"] = [{"emailAddress": {"address": x.strip()},
                              "type": "required"}
                             for x in a.attendees.split(",") if x.strip()]
    created = graph("POST", "/me/events", body=body, allow_write="calendar")
    out({"created": created.get("id"), "subject": created.get("subject"),
         "start": (created.get("start") or {}).get("dateTime"),
         "attendees_invited": bool(a.attendees),
         "link": created.get("webLink"),
         "note": "An invitation was sent to every attendee listed."})


def cmd_get(a):
    params = dict(p.split("=", 1) for p in (a.param or []))
    out(graph("GET", a.path, params=params or None, all_pages=a.all))


# ------------------------------------------------------------------- cli ----

def build_parser():
    p = argparse.ArgumentParser(
        prog="microsoft_graph.py",
        description="Search and read Microsoft 365 with the signed-in user's "
                    "own delegated permissions.")
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("status", help="preflight: signed in? which scopes?"
                   ).set_defaults(func=cmd_status)
    lg = sub.add_parser("login", help="sign in (opens a browser window)")
    lg.add_argument("--minimal", action="store_true",
                    help="ask only for User.Read + Mail.Read")
    lg.add_argument("--add-scope", action="append",
                    help="request one more scope, keeping the current ones")
    lg.set_defaults(func=cmd_login)
    sub.add_parser("logout", help="delete the local token cache"
                   ).set_defaults(func=cmd_logout)
    sub.add_parser("scopes", help="what is granted, what is missing"
                   ).set_defaults(func=cmd_scopes)
    sub.add_parser("whoami", help="GET /me").set_defaults(func=cmd_whoami)

    s = sub.add_parser("search", help="search mail, files, SharePoint, chats")
    s.add_argument("--query", required=True, help="KQL: words, from:, filetype:")
    s.add_argument("--types", help="comma list; default message,driveItem,chatMessage")
    s.add_argument("--top", type=int, default=15)
    s.add_argument("--start", type=int, default=0, help="paging offset")
    s.add_argument("--raw", action="store_true", help="full Graph hits")
    s.set_defaults(func=cmd_search)

    mail = sub.add_parser("mail").add_subparsers(dest="sub", required=True)
    ml = mail.add_parser("list")
    ml.add_argument("--top", type=int, default=20)
    ml.add_argument("--folder", help="inbox, sentitems, drafts, ...")
    ml.set_defaults(func=cmd_mail_list)
    mg = mail.add_parser("get")
    mg.add_argument("--id", required=True)
    mg.add_argument("--max-chars", type=int, default=8000)
    mg.set_defaults(func=cmd_mail_get)

    cal = sub.add_parser("calendar").add_subparsers(dest="sub", required=True)
    cl = cal.add_parser("list")
    cl.add_argument("--days", type=int, default=7)
    cl.set_defaults(func=cmd_calendar_list)
    ca = cal.add_parser("add", help="create an event (needs Calendars.ReadWrite)")
    ca.add_argument("--subject", required=True)
    ca.add_argument("--start", required=True, help="2026-08-25T10:00:00")
    ca.add_argument("--end", required=True)
    ca.add_argument("--timezone", default="Europe/Budapest")
    ca.add_argument("--location")
    ca.add_argument("--body")
    ca.add_argument("--attendees", help="comma-separated addresses; they get invited")
    ca.set_defaults(func=cmd_calendar_add)

    g = sub.add_parser("get", help="any Graph GET, e.g. /me/drive/recent")
    g.add_argument("path")
    g.add_argument("--param", action="append", help='repeatable: $top=5')
    g.add_argument("--all", action="store_true", help="follow nextLink (max 10 pages)")
    g.set_defaults(func=cmd_get)

    return p


def main():
    args = build_parser().parse_args()
    try:
        args.func(args)
    except SystemExit:
        raise
    except GraphError as e:
        fail(e, e.hint)
    except KeyboardInterrupt:
        fail("interrupted")
    except Exception as e:
        fail(f"{type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
