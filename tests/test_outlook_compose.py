"""outlook.py compose helpers against fake Outlook objects: what new_mail,
reply and forward return, and the ways a model calls them. No Outlook needed."""
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / ".agents" / "skills" / "office-live" / "scripts"))
import office  # noqa: E402
import outlook  # noqa: E402


class FakeRecipient:
    def __init__(self, name, address):
        self.Type, self.Name, self.Address, self.Resolved = 1, name, address, True
        self.AddressEntry = None


class FakeRecipients:
    def __init__(self):
        self.items = []

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]

    def ResolveAll(self):
        return True


class FakeAttachments(list):
    def Add(self, path):
        self.append(type("A", (), {"FileName": Path(path).name})())


class FakeMail:
    _oleobj_ = object()                 # what outlook._is_com looks for
    Class = 43

    def __init__(self, subject="", body="", sender="Jane Doe"):
        self.Subject, self.Body, self.HTMLBody = subject, body, "<html><body></body></html>"
        self.Recipients, self.Attachments = FakeRecipients(), FakeAttachments()
        self.Sent, self.displayed, self.sent_calls, self.sender = False, False, 0, sender
        self.To = ""

    def Display(self, modal=False):
        self.displayed = True

    def Send(self):
        self.sent_calls += 1

    def Reply(self):
        answer = FakeMail("RE: " + self.Subject)
        answer.Recipients.items.append(FakeRecipient(self.sender, "jane.doe@example.org"))
        return answer

    ReplyAll = Reply

    def Forward(self):
        return FakeMail("FW: " + self.Subject)


class FakeNamespace:
    def __init__(self, items):
        self.items = items

    def GetItemFromID(self, entry_id):
        return self.items[entry_id]


@pytest.fixture
def outlook_app(monkeypatch):
    original = FakeMail("Annual leave", "You have 8 days left.")
    ns = FakeNamespace({"ID-1": original})
    app = type("App", (), {"CreateItem": lambda self, kind: FakeMail(),
                           "GetNamespace": lambda self, name: ns})()
    monkeypatch.setattr(outlook, "_raw_app", lambda: app)
    return original


def test_new_mail_result_is_json_and_says_not_sent(outlook_app):
    m = outlook.new_mail(to="jane.doe@example.org", subject="Leave", body="I will take it.")
    data = json.loads(json.dumps(m))            # what the model did; it used to raise
    assert data["subject"] == "Leave" and data["sent"] is False
    assert data["status"].startswith("NOT sent")
    assert "I will take it." in m.HTMLBody      # attributes still reach the item
    with pytest.raises(office.OfficeError):
        m.Send()                                # and the guard still holds


@pytest.mark.parametrize("call", [
    lambda: outlook.reply("ID-1", body="Thanks."),
    lambda: outlook.reply(entry_id="ID-1", body="Thanks."),
    lambda: outlook.reply({"entry_id": "ID-1", "subject": "Annual leave"}, body="Thanks."),
])
def test_reply_takes_an_entry_id_however_it_is_given(outlook_app, call):
    r = call()
    assert r["subject"] == "RE: Annual leave" and r["recipients"][0]["email"] == "jane.doe@example.org"
    assert r["status"].startswith("NOT sent")


def test_reply_without_a_message_says_how(outlook_app):
    with pytest.raises(office.OfficeError) as err:
        outlook.reply(body="Thanks.")
    assert "entry_id" in err.value.hint


def test_forward_and_attach_accept_the_composed_result(outlook_app, tmp_path):
    f = outlook.forward(entry_id="ID-1", to="x@example.org", body="FYI.")
    annex = tmp_path / "annex.txt"
    annex.write_text("x", encoding="utf-8")
    assert outlook.add_attachment(f, str(annex)) == ["annex.txt"]
    assert outlook.readback(f)["attachments"] == ["annex.txt"]


def test_summary_has_read_and_unread(outlook_app):
    outlook_app.UnRead = True
    row = outlook.summary(outlook_app)
    assert row["unread"] is True and row["read"] is False
