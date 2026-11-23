"""Shared test setup: import paths, a clean OCR environment, and a fake OCR
server so the OCR order can be tested over real HTTP without a real model."""
from __future__ import annotations

import json
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
for path in (ROOT / "tests", ROOT / "tools",
             ROOT / ".agents" / "skills" / "file-reader" / "scripts",
             ROOT / ".agents" / "skills" / "m365-graph" / "scripts"):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

OCR_VARS = ("OCR_LOCAL", "OCR_LOCAL_LANGUAGE", "OCR_URL", "OCR_MODEL", "OCR_API_KEY",
            "OCR_CLOUD_URL", "OCR_CLOUD_MODEL", "OCR_CLOUD_API_KEY", "OCR_CLOUD_API_KEY_ENV")

# A port nothing listens on: connections are refused at once.
UNREACHABLE = "http://127.0.0.1:9/v1"


@pytest.fixture(autouse=True)
def clean_ocr_env(monkeypatch):
    """No test sees the OCR settings of the machine it runs on. Windows OCR is
    off unless a test turns it on (or fakes it), so results do not depend on
    whether the machine is Windows with pywinrt."""
    for name in OCR_VARS:
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setenv("OCR_LOCAL", "off")


@pytest.fixture
def fake_local_ocr(monkeypatch):
    """Windows OCR replaced by a stand-in, so the order can be tested anywhere.
    Every call is recorded; set .reply or .error to change what it does."""
    import pdfread

    class Fake:
        reply = "Windows OCR text of the page, long enough to count."
        error = None
        calls = []

    def recognize(png, language=""):
        Fake.calls.append(language)
        if Fake.error:
            raise Fake.error
        return Fake.reply

    Fake.calls = []
    monkeypatch.setenv("OCR_LOCAL", "on")
    monkeypatch.setattr(pdfread, "local_settings", lambda: (pdfread._env("OCR_LOCAL_LANGUAGE"),))
    monkeypatch.setattr(pdfread, "local_ocr_page", recognize)
    return Fake


class FakeOcr:
    """Answers /v1/chat/completions like DeepSeek-OCR in grounding mode, and
    /v1/responses like a cloud vision model. Records every request."""

    SERVER_TEXT = ("<|ref|>text<|/ref|><|det|>[[10, 20, 300, 40]]<|/det|>\n"
                   "Text of a scanned page, with the résumé\n"
                   "<|ref|>table<|/ref|><|det|>[[1, 2, 3, 4]]<|/det|>\n"
                   "<table><tr><td>Item</td><td>Amount</td></tr><tr><td>A</td><td>100</td></tr></table>")
    CLOUD_TEXT = "A page read in the cloud, a text long enough to count."

    def __init__(self):
        self.requests = []
        fake = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args):
                pass

            def do_POST(self):
                body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
                fake.requests.append({"path": self.path, "headers": dict(self.headers), "body": body})
                if self.path.endswith("/chat/completions"):
                    answer = {"choices": [{"message": {"content": fake.SERVER_TEXT}}]}
                elif self.path.endswith("/responses"):
                    answer = {"output": [{"content": [{"type": "output_text", "text": fake.CLOUD_TEXT}]}]}
                else:
                    self.send_response(404)
                    self.end_headers()
                    return
                raw = json.dumps(answer).encode()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)

        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.url = f"http://127.0.0.1:{self.httpd.server_address[1]}/v1"
        threading.Thread(target=self.httpd.serve_forever, daemon=True).start()

    def calls(self, suffix):
        return [r for r in self.requests if r["path"].endswith(suffix)]


@pytest.fixture
def fake_ocr():
    server = FakeOcr()
    yield server
    server.httpd.shutdown()
