"""Local Web Reader — a keyless ``web_extract`` backend for Hermes.

Hermes' ``web_extract`` tool needs an extract backend, and every hosted one
wants an API key. This plugin registers a web provider named ``local`` that
reads the page on this machine instead; select it with
``web.extract_backend: local``. It implements extract only, so web search
stays with whichever search backend is configured.

A thin adapter. The work is in ``scripts/web_reader.py``, which has no Hermes
dependency and can be imported and tested on its own.
"""

from __future__ import annotations

import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any

_SCRIPTS = Path(__file__).parent / "scripts"
NAME = "local"
MAX_PARALLEL = 4


def _reader():
    if str(_SCRIPTS) not in sys.path:
        sys.path.insert(0, str(_SCRIPTS))
    import web_reader

    return web_reader


def _entry(result: dict) -> dict:
    """One result in the shape web_extract expects: url, title, content,
    raw_content and metadata, or url and error for a page that failed."""
    url = result.get("url", "")
    if result.get("error"):
        return {"url": url, "title": "", "content": "", "error": result["error"]}
    content = result.get("content", "")
    title = result.get("title", "")
    return {
        "url": url,
        "title": title,
        "content": content,
        "raw_content": content,
        "metadata": {"sourceURL": result.get("final_url") or url, "title": title, "via": result.get("via", "")},
    }


def extract_urls(urls: list[str], max_chars: int | None = None) -> list[dict]:
    """Read every URL, a few at a time. Never raises: a failure is a per-URL error."""
    reader = _reader()
    limit = int(max_chars) if max_chars else reader.MAX_CHARS

    def one(u: str) -> dict:
        try:
            return _entry(reader.read(u, max_chars=limit))
        except Exception as exc:  # a reader bug must not take the tool down
            return {"url": u, "title": "", "content": "", "error": f"{type(exc).__name__}: {exc}"}

    urls = [u for u in (urls or []) if u]
    if len(urls) <= 1:
        return [one(u) for u in urls]
    with ThreadPoolExecutor(max_workers=min(MAX_PARALLEL, len(urls))) as pool:
        return list(pool.map(one, urls))


def _provider_class():
    # Imported here, not at module level: the base class exists only inside Hermes.
    from agent.web_search_provider import WebSearchProvider

    class LocalReaderProvider(WebSearchProvider):
        """Reads pages on this machine. Extract only: search stays with the configured search backend."""

        @property
        def name(self) -> str:
            return NAME

        @property
        def display_name(self) -> str:
            return "Local reader (no key)"

        def is_available(self) -> bool:
            return True

        def supports_search(self) -> bool:
            return False

        def supports_extract(self) -> bool:
            return True

        def extract(self, urls: list[str], **kwargs: Any) -> list[dict]:
            return extract_urls(urls, kwargs.get("max_chars"))

        def get_setup_schema(self) -> dict[str, Any]:
            return {
                "name": self.display_name,
                "badge": "no key",
                "tag": "Reads the page on this machine (HTTP, HTML, PDF); pages that need JavaScript go to the browser tool.",
                "env_vars": [],
            }

    return LocalReaderProvider


def register(ctx):
    ctx.register_web_search_provider(_provider_class()())
