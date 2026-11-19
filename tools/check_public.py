"""Fail if anything in this repository looks like it belongs inside one
organisation rather than in a public repository.

This repo is public. The skills here are used inside a public body whose
internal hostnames, model deployments and staff addresses must not end up in
it, and a commit cannot be taken back once someone has fetched it. The values
that make a skill work in a particular place (endpoints, model names, keys)
come from the environment at install time, never from a default in the code.

What it looks for, in every tracked file:

- internal names: server code names and internal domains, matched against
  SHA-256 digests so that the list itself does not publish them
- private IPv4 addresses
- Azure resource hostnames (``<name>.openai.azure.com`` and friends)
- e-mail addresses other than GitHub's noreply form
- strings shaped like keys: long hex runs, known token prefixes
- GUIDs, except the public Microsoft identifiers listed in ``ALLOWED_GUIDS``

With ``--commits A..B`` it also checks every commit in that range: the author
and committer e-mail, and the message. That is how CI checks a pull request.

    python tools/check_public.py
    python tools/check_public.py --commits origin/main..HEAD
    python tools/check_public.py --digest some-name     # a line for DENIED_DIGESTS

Exit status 1 and one line per finding when something matches. A finding that
is genuinely public goes into an allowlist below, with a comment saying why.
"""
from __future__ import annotations

import argparse
import hashlib
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Names that must not appear: internal server code names, internal domains.
# Stored as SHA-256 of the lower-cased name, because a list of the names in a
# public file would publish exactly what it is meant to keep out. Every word,
# every part of a dotted or hyphenated name, and every domain suffix of a host
# is hashed and looked up. Add one with: check_public.py --digest <name>
DENIED_DIGESTS = {
    "bd549a1f143594ba4b1ab4386419e803e4a469cd7f3c4a8e445be89f7e286fe3",
    "3f1ed99a79147aefd798580f2271940d6af7acf64afc9368230d629ecdb3f45d",
    "d41cf97a51369e89142dde076d84dac980bd52c4aa1f4f72f976dfe302dc83f0",
    "b64c83288462c1cfca39473614e478521a19d7191fc2895a43966f412b8cd445",
    "5370bce317c57c558cd3d9676d5ea7c65ca9ef8336d958ccb94cd69ee6c3d80c",
    "afd99d7090d47d57a4aff3592d49bf1dc6ebb29e76893d357cb665deb209664c",
    "9b4d23d368142232b9ca9758662c1bd270bfc94882e4c38b91fd7b8e395e537e",
}

ALLOWED_GUIDS = {
    # Microsoft Graph Command Line Tools: Microsoft's own public client id,
    # documented and used by every Graph PowerShell user.
    "14d82eec-204b-4c2f-b7e8-296a70dab67e",
}

# Placeholders the docs use on purpose.
ALLOWED_AZURE_HOSTS = {"example-resource", "your-resource", "contoso"}

# noreply forms, and the placeholder domains the docs use in examples
# (pelda.hu is Hungarian for example.com; a@b.hu style one-letter domains).
ALLOWED_EMAIL = re.compile(
    r"(@users\.noreply\.github\.com$|@example\.(com|org|net)$|@pelda\.hu$"
    r"|@[a-z]\.[a-z]{2,3}$|^noreply@anthropic\.com$|^noreply@github\.com$)",
    re.I,
)

CHECKS: list[tuple[str, re.Pattern[str]]] = [
    ("private IP", re.compile(
        r"\b(?:10(?:\.\d{1,3}){3}|192\.168(?:\.\d{1,3}){2}|172\.(?:1[6-9]|2\d|3[01])(?:\.\d{1,3}){2})\b")),
    ("Azure resource", re.compile(
        r"\b([a-z0-9][a-z0-9-]{1,62})\.(?:openai|cognitiveservices|services\.ai|search\.windows)"
        r"\.(?:azure\.com|net)\b", re.I)),
    ("e-mail address", re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}")),
    ("key-like hex", re.compile(r"\b[0-9a-fA-F]{32,}\b")),
    ("token prefix", re.compile(
        r"\b(?:apb_live_|sk-(?:proj-|ant-)?|ghp_|gho_|github_pat_|xox[abp]-|AKIA)[A-Za-z0-9_-]{8,}")),
    ("GUID", re.compile(r"\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b", re.I)),
]

# Files that describe the patterns above and so would match themselves.
SELF = {
    "tools/check_public.py",
    # Its tests feed it a sample of each finding on purpose.
    "tests/test_check_public.py",
}


def allowed(kind: str, match: re.Match[str]) -> bool:
    text = match.group(0)
    if kind == "GUID":
        return text.lower() in ALLOWED_GUIDS
    if kind == "Azure resource":
        return match.group(1).lower() in ALLOWED_AZURE_HOSTS
    if kind == "e-mail address":
        return bool(ALLOWED_EMAIL.search(text))
    return False


def tracked_files() -> list[str]:
    out = subprocess.run(["git", "ls-files", "-z"], cwd=ROOT, check=True,
                         capture_output=True).stdout
    return [p for p in out.decode("utf-8").split("\0") if p]


_TOKEN = re.compile(r"[A-Za-z0-9][A-Za-z0-9._-]*[A-Za-z0-9]|[A-Za-z0-9]")


def _digest(word: str) -> str:
    return hashlib.sha256(word.lower().encode("utf-8")).hexdigest()


def denied_names(line: str) -> list[str]:
    """Tokens of the line whose name, part or domain suffix is denied."""
    hits = []
    for token in _TOKEN.findall(line):
        candidates = {token}
        parts = [p for p in re.split(r"[._-]", token) if p]
        # Every run of consecutive parts, so a hyphenated denied name is found
        # inside a longer token too (name-swap, NAME_LOCAL_URL).
        for i in range(len(parts)):
            for j in range(i + 1, len(parts) + 1):
                for joiner in ("-", "_", "."):
                    candidates.add(joiner.join(parts[i:j]))
        labels = token.split(".")
        candidates.update(".".join(labels[i:]) for i in range(len(labels) - 1))
        if any(_digest(c) in DENIED_DIGESTS for c in candidates):
            hits.append(token)
    return hits


def scan_text(label: str, text: str) -> list[str]:
    found = []
    for lineno, line in enumerate(text.splitlines(), 1):
        for token in denied_names(line):
            found.append(f"{label}:{lineno}: internal name: {token}")
        for kind, pattern in CHECKS:
            for m in pattern.finditer(line):
                if not allowed(kind, m):
                    found.append(f"{label}:{lineno}: {kind}: {m.group(0)}")
    return found


def scan_files() -> list[str]:
    found = []
    for rel in tracked_files():
        if rel in SELF or not (ROOT / rel).is_file():
            continue  # this file, or one deleted but not yet staged
        data = (ROOT / rel).read_bytes()
        if b"\0" in data[:8192]:
            continue  # binary
        found += scan_text(rel, data.decode("utf-8", errors="replace"))
    return found


def scan_commits(rng: str) -> list[str]:
    out = subprocess.run(
        ["git", "log", "--format=%h%x00%an <%ae>%x00%cn <%ce>", rng],
        cwd=ROOT, check=True, capture_output=True, text=True, encoding="utf-8").stdout
    found = []
    for line in out.splitlines():
        sha, author, committer = line.split("\0")
        for role, who in (("author", author), ("committer", committer)):
            email = who[who.rfind("<") + 1:-1]
            if email and not ALLOWED_EMAIL.search(email):
                found.append(f"commit {sha}: {role} e-mail is not a noreply address: {email}")
        # A commit message is as public as the code.
        message = subprocess.run(["git", "log", "-1", "--format=%B", sha], cwd=ROOT, check=True,
                                 capture_output=True, text=True, encoding="utf-8").stdout
        found += scan_text(f"commit {sha} message", message)
    return found


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--commits", metavar="RANGE",
                    help="also check commit e-mails and messages in this range, e.g. origin/main..HEAD")
    ap.add_argument("--digest", metavar="NAME", help="print the DENIED_DIGESTS line for NAME and exit")
    args = ap.parse_args()
    if args.digest:
        print(f'    "{_digest(args.digest)}",')
        return 0

    found = scan_files()
    if args.commits:
        found += scan_commits(args.commits)

    for line in found:
        print(line)
    if found:
        print(f"\n{len(found)} finding(s). This repository is public: move the value "
              "into an environment variable set by the installer, or add a genuinely "
              "public value to the allowlist in tools/check_public.py.", file=sys.stderr)
        return 1
    print("public-safety check: clean")
    return 0


if __name__ == "__main__":
    sys.exit(main())
