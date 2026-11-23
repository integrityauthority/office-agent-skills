"""office.py's command line, the way an agent without the Hermes plugin runs
it: through a pipe, which on Windows defaults to cp1252."""
import io
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / ".agents" / "skills" / "office-live" / "scripts"))
import office  # noqa: E402


def test_emit_writes_central_european_text_through_a_cp1252_pipe(monkeypatch):
    raw = io.BytesIO()
    pipe = io.TextIOWrapper(raw, encoding="cp1252")      # what Windows gives a pipe
    monkeypatch.setattr(sys, "stdout", pipe)
    monkeypatch.delenv("HERMES_HOME", raising=False)
    # Encoding test case: ő and ű are not in cp1252, so this only round-trips
    # if office.out switches the pipe to UTF-8.
    name = "Meeting notes – Erdős, Szűcs.docx"
    with pytest.raises(SystemExit) as ended:
        office.out({"document": name})
    pipe.flush()
    assert ended.value.code == 0
    assert json.loads(raw.getvalue().decode("utf-8")) == {"ok": True, "data": {"document": name}}
