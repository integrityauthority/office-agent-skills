"""Classic Outlook over COM: mail, calendar and contacts — compose, never send.

The approval model is the Outlook window itself. Everything here can READ the
mailbox and COMPOSE new items (mail, replies, forwards, appointments,
contacts), and then `.Display()` them in front of the user. Sending and saving
are refused outright: the human reviews the open item and presses Send or Save
themselves. Nothing the agent writes leaves the machine or lands in a folder
without that click.

Two layers enforce it:

* every Outlook object handed out is wrapped in `_Guarded`, which refuses the
  methods that send, save, delete or move (`Send`, `Save`, `Delete`, `Move`,
  `Close` — whose default argument SAVES — ...). Attribute reads, property
  sets on new drafts, and ordinary calls pass straight through;
* office.refuse() blocks reaching Outlook WITHOUT the wrapper
  (`Dispatch("Outlook.Application")`, `get_app("outlook")`) and any `.Send(`.

The wrapper is a guard against mistakes, not a sandbox: determined code can
always find a raw COM object. It exists so that the natural thing to type is
also the safe thing.

Only classic Outlook (OUTLOOK.EXE) has a COM interface. The new Outlook
(olk.exe) has none, and nothing here can work with it.

Measured on this machine (Hungarian locale), because the obvious filter syntax
is wrong here:

* Jet date filters (`[ReceivedTime] >= 'YYYY-MM-DD HH:MM'`) match local time
  exactly — 124 of 124 in a seven-day test.
* DASL date filters read the string as UTC and missed one message at the
  boundary; with a trailing `Z` they silently matched 11 913 of 11 965.
* So dates go through Jet, text through DASL (`LIKE`), chained: a Restrict on
  a restricted collection is allowed, mixing the two syntaxes in one is not.
"""

from __future__ import annotations

import datetime as _dt
import html as _html
import os
import re

import office
from office import OfficeError

# Outlook constants (OlItemType / OlDefaultFolders / OlObjectClass).
OL_MAIL, OL_APPOINTMENT, OL_CONTACT = 0, 1, 2
CLASS_MAIL, CLASS_MEETING_REQUEST = 43, 53
KINDS = {43: "mail", 45: "post", 46: "report", 49: "task_request",
         53: "meeting_request", 54: "meeting_cancellation",
         55: "meeting_declined", 56: "meeting_accepted",
         57: "meeting_tentative"}
FOLDERS = {"inbox": 6, "sent": 5, "drafts": 16, "deleted": 3, "outbox": 4,
           "calendar": 9, "contacts": 10, "tasks": 13, "junk": 23}

# Methods that send, persist, delete or move. `Close` is here because its
# argument defaults to olSave; `Copy` because it writes a copy into the folder.
BLOCKED_CALLS = {
    "Send", "Save", "SaveAs", "Delete", "PermanentlyDelete", "Move", "Copy",
    "Close", "Respond", "Execute", "Remove", "Quit", "MarkAsTask",
    "ClearTaskFlag", "AddToPFFavorites",
}
# Property sets that change an EXISTING item immediately, without a Save.
BLOCKED_SETS = {"UnRead", "FlagRequest", "FlagStatus", "Categories"}

_BLOCK_HINT = ("compose the item and call .Display() — the user reviews it in "
               "Outlook and presses Send or Save themselves. Nothing is sent, "
               "saved, moved or deleted from here.")


# ------------------------------------------------------------------ guard ---

def _is_com(value):
    return hasattr(value, "_oleobj_")


def _unwrap(value):
    if isinstance(value, _Guarded):
        return object.__getattribute__(value, "_o")
    if isinstance(value, (list, tuple)):
        return type(value)(_unwrap(v) for v in value)
    return value


def _wrap(value):
    if isinstance(value, _Guarded) or value is None:
        return value
    if _is_com(value):
        return _Guarded(value)
    if callable(value):
        def call(*args, **kwargs):
            result = value(*[_unwrap(a) for a in args],
                           **{k: _unwrap(v) for k, v in kwargs.items()})
            return _wrap(result)
        call.__name__ = getattr(value, "__name__", "com_method")
        return call
    return value


class _Guarded:
    """Pass-through wrapper over an Outlook COM object that refuses the calls
    which send, save, delete or move. Everything it returns is wrapped too, so
    an item reached through a folder is as guarded as one from a helper."""

    __slots__ = ("_o",)

    def __init__(self, obj):
        object.__setattr__(self, "_o", obj)

    def __getattr__(self, name):
        if name in BLOCKED_CALLS:
            raise OfficeError(f"refused: Outlook .{name}() is the user's call",
                              _BLOCK_HINT)
        target = object.__getattribute__(self, "_o")
        if name == "Add" and "folders" in type(target).__name__.lower():
            raise OfficeError("refused: creating a mail folder changes the "
                              "user's mailbox", _BLOCK_HINT)
        return _wrap(getattr(target, name))

    def __setattr__(self, name, value):
        if name in BLOCKED_SETS:
            raise OfficeError(f"refused: setting {name} changes the user's "
                              "mailbox immediately", _BLOCK_HINT)
        setattr(object.__getattribute__(self, "_o"), name, _unwrap(value))

    def __call__(self, *args, **kwargs):
        target = object.__getattribute__(self, "_o")
        return _wrap(target(*[_unwrap(a) for a in args],
                            **{k: _unwrap(v) for k, v in kwargs.items()}))

    def __iter__(self):
        for item in object.__getattribute__(self, "_o"):
            yield _wrap(item)

    def __len__(self):
        return object.__getattribute__(self, "_o").Count

    def __bool__(self):
        return True

    def __repr__(self):
        return f"<guarded {object.__getattribute__(self, '_o')!r}>"


# ---------------------------------------------------------------- attach ---

def _raw_app():
    w = office._win32()
    try:
        raw = w.GetActiveObject("Outlook.Application")
    except Exception:
        raise OfficeError(
            "classic Outlook is not running",
            "Ask the user to open Outlook. Only classic Outlook (OUTLOOK.EXE) "
            "can be driven; the new Outlook (olk.exe) has no COM interface. "
            "Never start Outlook on your own initiative.")
    return office._early_bind(w, raw)


def _raw_ns():
    return _raw_app().GetNamespace("MAPI")


def attach_outlook():
    """(app, namespace), both guarded. office.attach("outlook") lands here."""
    app = _raw_app()
    return _Guarded(app), _Guarded(app.GetNamespace("MAPI"))


def folder(name="inbox"):
    """A default folder by name ("inbox", "sent", "drafts", "calendar",
    "contacts", ...) or a path under the default store, e.g. "Inbox/Projektek"."""
    return _Guarded(_raw_folder(name))


def _raw_folder(name):
    ns = _raw_ns()
    key = str(name).lower()
    if key in FOLDERS:
        return ns.GetDefaultFolder(FOLDERS[key])
    parts = [p for p in re.split(r"[\\/]", str(name)) if p]
    root = ns.GetDefaultFolder(FOLDERS["inbox"]).Parent
    current = root
    for part in parts:
        match = None
        for child in current.Folders:
            if child.Name.lower() == part.lower():
                match = child
                break
        if match is None:
            names = [c.Name for c in current.Folders]
            raise OfficeError(f"no folder {part!r} under {current.Name!r}",
                              f"available: {names}")
        current = match
    return current


# --------------------------------------------------------------- helpers ---

def _iso(value):
    try:
        return value.replace(tzinfo=None).isoformat(sep=" ", timespec="minutes")
    except Exception:
        return str(value) if value is not None else None


def _jet_date(value):
    """Jet filter literal. Verified to match local time on this locale."""
    if isinstance(value, str):
        value = _dt.datetime.fromisoformat(value)
    return value.strftime("%Y-%m-%d %H:%M")


def _dasl_quote(text):
    return str(text).replace("'", "''")


def _sender_email(item):
    try:
        if item.SenderEmailType == "EX":
            user = item.Sender.GetExchangeUser()
            if user is not None:
                return user.PrimarySmtpAddress
        return item.SenderEmailAddress
    except Exception:
        return None


def summary(item):
    """A small, context-friendly description of a mail item — no body."""
    raw = _unwrap(item)
    out = {"entry_id": getattr(raw, "EntryID", None)}
    try:
        out["kind"] = KINDS.get(raw.Class, f"class {raw.Class}")
    except Exception:
        pass
    for key, attr in (("subject", "Subject"), ("sender", "SenderName"),
                      ("to", "To"), ("cc", "CC")):
        try:
            out[key] = getattr(raw, attr)
        except Exception:
            pass
    out["sender_email"] = _sender_email(raw)
    for key, attr in (("received", "ReceivedTime"), ("sent", "SentOn")):
        try:
            out[key] = _iso(getattr(raw, attr))
        except Exception:
            pass
    try:
        out["unread"] = bool(raw.UnRead)
        out["attachments"] = [a.FileName for a in raw.Attachments]
    except Exception:
        pass
    return out


# ------------------------------------------------------------------ mail ---

def find_mail(query=None, sender=None, folder_name="inbox", since_days=None,
              since=None, until=None, unread=None, limit=20):
    """Search one folder, newest first. Returns summaries, never bodies.

    query    words in the subject or body (LIKE, case-insensitive)
    sender   part of the sender's name or address
    since_days / since / until   date window, local time
    unread   True / False / None
    """
    items = _raw_folder(folder_name).Items
    items.Sort("[ReceivedTime]", True)
    if since_days is not None and since is None:
        since = _dt.datetime.now() - _dt.timedelta(days=since_days)
    if since is not None:
        items = items.Restrict(f"[ReceivedTime] >= '{_jet_date(since)}'")
    if until is not None:
        items = items.Restrict(f"[ReceivedTime] < '{_jet_date(until)}'")
    if unread is not None:
        items = items.Restrict(f"[UnRead] = {'True' if unread else 'False'}")
    if query:
        q = _dasl_quote(query)
        items = items.Restrict(
            '@SQL=("urn:schemas:httpmail:subject" LIKE \'%' + q + '%\' OR '
            '"urn:schemas:httpmail:textdescription" LIKE \'%' + q + '%\')')
    if sender:
        # For an Exchange sender, fromemail holds an X500 path
        # (/O=EXCHANGELABS/...), not the SMTP address — a search by e-mail
        # found 36 of 890 here. PR_SENDER_SMTP_ADDRESS (0x5D01001F) has it.
        s = _dasl_quote(sender)
        items = items.Restrict(
            '@SQL=("urn:schemas:httpmail:fromname" LIKE \'%' + s + '%\' OR '
            '"urn:schemas:httpmail:fromemail" LIKE \'%' + s + '%\' OR '
            '"http://schemas.microsoft.com/mapi/proptag/0x5D01001F" LIKE '
            '\'%' + s + '%\')')
    items.Sort("[ReceivedTime]", True)
    # Every item comes back, labelled with its kind. Filtering to plain mail
    # here silently dropped meeting responses and delivery reports, so the
    # count no longer matched what the user sees in the folder.
    found, item = [], items.GetFirst()
    while item is not None and len(found) < limit:
        try:
            found.append(summary(item))
        except Exception:
            pass
        item = items.GetNext()
    return found


def get_mail(entry_id, max_chars=4000):
    """One message with its plain-text body (truncated), as a dict."""
    raw = _raw_ns().GetItemFromID(entry_id)
    out = summary(raw)
    body = raw.Body or ""
    out["body"] = body[:max_chars]
    out["body_truncated"] = len(body) > max_chars
    return out


def open_item(entry_id):
    """The guarded item itself, to reply to or forward."""
    return _Guarded(_raw_ns().GetItemFromID(entry_id))


def current_mail():
    """What the user is looking at: the open item window, else the first
    selected item in the main window. Guarded. None if nothing."""
    app = _raw_app()
    try:
        inspector = app.ActiveInspector()
        if inspector is not None:
            return _Guarded(inspector.CurrentItem)
    except Exception:
        pass
    try:
        explorer = app.ActiveExplorer()
        if explorer is not None and explorer.Selection.Count:
            return _Guarded(explorer.Selection(1))
    except Exception:
        pass
    return None


def _text_to_html(text):
    paragraphs = [p for p in re.split(r"\n\s*\n", str(text).strip()) if p]
    return "".join(
        "<p>" + _html.escape(p).replace("\n", "<br>") + "</p>" for p in paragraphs)


# Outlook's own default body font, so inserted text does not look pasted in.
_FONT = 'font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt'


def _fragment(body=None, html=None):
    if html is None and body is None:
        return ""
    inner = html if html is not None else _text_to_html(body)
    return f'<div style="{_FONT}">{inner}</div>'


def _insert_top(raw_item, fragment):
    """Put our text at the top of the body, above the signature and any quoted
    message, without disturbing either."""
    if not fragment:
        return
    current = raw_item.HTMLBody or ""
    match = re.search(r"<body[^>]*>", current, re.IGNORECASE)
    if match:
        raw_item.HTMLBody = current[:match.end()] + fragment + current[match.end():]
    else:
        raw_item.HTMLBody = fragment + current


def _attach(raw_item, attachments):
    for path in attachments or ():
        path = os.path.abspath(path)
        if not os.path.isfile(path):
            raise OfficeError(f"attachment not found: {path}",
                              "give an absolute path to an existing file")
        raw_item.Attachments.Add(path)


def _recipients(raw_item, to=None, cc=None, bcc=None):
    def join(value):
        return "; ".join(value) if isinstance(value, (list, tuple)) else value
    if to:
        raw_item.To = join(to)
    if cc:
        raw_item.CC = join(cc)
    if bcc:
        raw_item.BCC = join(bcc)
    try:
        raw_item.Recipients.ResolveAll()
    except Exception:
        pass


def new_mail(to=None, subject="", body=None, html=None, cc=None, bcc=None,
             attachments=(), display=True):
    """Compose a new message and open it for the user to review and send.

    body  plain text (paragraphs on blank lines), or
    html  HTML for real formatting: <b>, <ul>, <table>, links.
    With display=True the window opens first, so Outlook inserts the user's
    signature, and the text goes above it.
    """
    raw = _raw_app().CreateItem(OL_MAIL)
    raw.BodyFormat = 2                                   # olFormatHTML
    _recipients(raw, to, cc, bcc)
    raw.Subject = subject or ""
    if display:
        raw.Display(False)
    _insert_top(raw, _fragment(body, html))
    _attach(raw, attachments)
    return _Guarded(raw)


def reply(item, body=None, html=None, reply_all=False, attachments=(),
          display=True):
    """Reply (or reply-all) to `item` — an entry_id or a guarded item. The
    quoted original and the signature are kept; the new text goes on top."""
    raw_src = (_raw_ns().GetItemFromID(item) if isinstance(item, str)
               else _unwrap(item))
    raw = raw_src.ReplyAll() if reply_all else raw_src.Reply()
    if display:
        raw.Display(False)
    _insert_top(raw, _fragment(body, html))
    _attach(raw, attachments)
    return _Guarded(raw)


def forward(item, to=None, body=None, html=None, attachments=(), display=True):
    raw_src = (_raw_ns().GetItemFromID(item) if isinstance(item, str)
               else _unwrap(item))
    raw = raw_src.Forward()
    _recipients(raw, to)
    if display:
        raw.Display(False)
    _insert_top(raw, _fragment(body, html))
    _attach(raw, attachments)
    return _Guarded(raw)


def add_attachment(item, path):
    _attach(_unwrap(item), [path])
    return [a.FileName for a in _unwrap(item).Attachments]


# -------------------------------------------------------------- calendar ---

def calendar(start=None, end=None, days=7, limit=50):
    """Appointments overlapping [start, end), recurring ones expanded."""
    start = (_dt.datetime.fromisoformat(start) if isinstance(start, str)
             else start or _dt.datetime.now().replace(hour=0, minute=0,
                                                       second=0, microsecond=0))
    end = (_dt.datetime.fromisoformat(end) if isinstance(end, str)
           else end or start + _dt.timedelta(days=days))
    items = _raw_folder("calendar").Items
    # Order matters: Sort, then IncludeRecurrences, then Restrict — otherwise
    # recurring meetings show up once, at their first occurrence.
    items.Sort("[Start]")
    items.IncludeRecurrences = True
    items = items.Restrict(f"[Start] < '{_jet_date(end)}' AND "
                           f"[End] > '{_jet_date(start)}'")
    out, item = [], items.GetFirst()
    while item is not None and len(out) < limit:
        try:
            out.append({
                "subject": item.Subject, "start": _iso(item.Start),
                "end": _iso(item.End), "location": item.Location,
                "organizer": item.Organizer, "all_day": bool(item.AllDayEvent),
                "recurring": bool(item.IsRecurring),
                "busy_status": item.BusyStatus, "entry_id": item.EntryID,
            })
        except Exception:
            pass
        item = items.GetNext()
    return out


def new_appointment(subject, start, end=None, minutes=60, location=None,
                    body=None, html=None, attendees=None, reminder_minutes=15,
                    display=True):
    """Compose an appointment — or, with attendees, a meeting request — and
    open it for the user. They press Send (meeting) or Save (appointment)."""
    start = _dt.datetime.fromisoformat(start) if isinstance(start, str) else start
    if end is None:
        end = start + _dt.timedelta(minutes=minutes)
    elif isinstance(end, str):
        end = _dt.datetime.fromisoformat(end)
    raw = _raw_app().CreateItem(OL_APPOINTMENT)
    raw.Subject = subject
    # com_date: a plain datetime arrives one UTC offset early (10:00 -> 09:00).
    raw.Start = office.com_date(start)
    raw.End = office.com_date(end)
    if location:
        raw.Location = location
    if reminder_minutes is not None:
        raw.ReminderSet = True
        raw.ReminderMinutesBeforeStart = reminder_minutes
    if attendees:
        raw.MeetingStatus = 1                              # olMeeting
        for person in attendees:
            raw.Recipients.Add(person)
        raw.Recipients.ResolveAll()
    if body is not None:
        raw.Body = body
    if display:
        raw.Display(False)
    if html is not None:
        # AppointmentItem has no HTMLBody; RTFBody is the rich route and is
        # not worth the trouble for a description. Plain text it is.
        raw.Body = re.sub(r"<[^>]+>", "", html)
    return _Guarded(raw)


# -------------------------------------------------------------- contacts ---

def resolve(name):
    """Look a person up in the address book (GAL included). Returns
    {"name", "email"} or None. For "send it to Kovács" — resolve first."""
    recipient = _raw_ns().CreateRecipient(name)
    if not recipient.Resolve():
        return None
    entry = recipient.AddressEntry
    email = None
    try:
        user = entry.GetExchangeUser()
        if user is not None:
            email = user.PrimarySmtpAddress
    except Exception:
        pass
    return {"name": entry.Name, "email": email or entry.Address}


def find_contact(query, limit=10):
    """Search the user's Contacts folder, plus the address book."""
    q = _dasl_quote(query)
    items = _raw_folder("contacts").Items.Restrict(
        '@SQL=("urn:schemas:contacts:cn" LIKE \'%' + q + '%\' OR '
        '"urn:schemas:contacts:email1" LIKE \'%' + q + '%\' OR '
        '"urn:schemas:contacts:o" LIKE \'%' + q + '%\')')
    out, item = [], items.GetFirst()
    while item is not None and len(out) < limit:
        try:
            out.append({"name": item.FullName, "email": item.Email1Address,
                        "company": item.CompanyName,
                        "phone": item.BusinessTelephoneNumber or
                                 item.MobileTelephoneNumber,
                        "source": "contacts", "entry_id": item.EntryID})
        except Exception:
            pass
        item = items.GetNext()
    if len(out) < limit:
        hit = resolve(query)
        if hit and all(hit["email"] != c.get("email") for c in out):
            out.append({**hit, "source": "address book"})
    return out


def new_contact(full_name, email=None, company=None, phone=None, mobile=None,
                job_title=None, notes=None, display=True):
    raw = _raw_app().CreateItem(OL_CONTACT)
    raw.FullName = full_name
    for attr, value in (("Email1Address", email), ("CompanyName", company),
                        ("BusinessTelephoneNumber", phone),
                        ("MobileTelephoneNumber", mobile),
                        ("JobTitle", job_title), ("Body", notes)):
        if value:
            setattr(raw, attr, value)
    if display:
        raw.Display(False)
    return _Guarded(raw)


# ---------------------------------------------------------------- status ---

def outlook_status():
    """Is classic Outlook up, and the shape of the mailbox — counts only."""
    try:
        app = _raw_app()
    except OfficeError as e:
        return {"running": False, "note": str(e)}
    ns = app.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(FOLDERS["inbox"])
    out = {"running": True, "version": app.Version,
           "accounts": [a.DisplayName for a in ns.Accounts],
           "inbox": {"items": inbox.Items.Count,
                     "unread": inbox.UnReadItemCount},
           "drafts": ns.GetDefaultFolder(FOLDERS["drafts"]).Items.Count}
    try:
        explorer = app.ActiveExplorer()
        out["selected"] = explorer.Selection.Count if explorer else 0
        out["item_window_open"] = app.ActiveInspector() is not None
    except Exception:
        pass
    return out
