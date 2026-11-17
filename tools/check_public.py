"""Fail if anything in this repository looks like it belongs inside one
organisation rather than in a public repository.

This repo is public. The skills here are used inside a public body whose
internal hostnames, model deployments and staff addresses must not end up in
it, and a commit cannot be taken back once someone has fetched it. The values
that make a skill work in a particular place (endpoints, model names, keys)
come from the environment at install time, never from a default in the code.

What it looks for, in every tracked file:

- hosts on internal domains, and private IPv4 addresses
- Azure resource hostnames (``<name>.openai.azure.com`` and friends)
- e-mail addresses other than GitHub's noreply form
- strings shaped like keys: long hex runs, known token prefixes
- GUIDs, except the public Microsoft identifiers listed in ``ALLOWED_GUIDS``

With ``--commits A..B`` it also checks the author and committer e-mail of
every commit in that range, which is how CI checks a pull request.

    python tools/check_public.py
    python tools/check_public.py --commits origin/main..HEAD

Exit status 1 and one line per finding when something matches. A finding that
is genuinely public goes into an allowlist below, with a comment saying why.
"""
from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Domains that only resolve inside the organisation, or identify its staff.
INTERNAL_DOMAINS = ("inhat.hu", "integritashatosag.hu")

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
    ("internal host", re.compile(
        r"[a-z0-9.-]*\b(?:" + "|".join(re.escape(d) for d in INTERNAL_DOMAINS) + r")\b",
        re.I)),
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
SELF = {"tools/check_public.py"}


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


def scan_text(label: str, text: str) -> list[str]:
    found = []
    for lineno, line in enumerate(text.splitlines(), 1):
        for kind, pattern in CHECKS:
            for m in pattern.finditer(line):
                if not allowed(kind, m):
                    found.append(f"{label}:{lineno}: {kind}: {m.group(0)}")
    return found


def scan_files() -> list[str]:
    found = []
    for rel in tracked_files():
        if rel in SELF:
            continue
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
    return found


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--commits", metavar="RANGE",
                    help="also check commit e-mails in this git range, e.g. origin/main..HEAD")
    args = ap.parse_args()

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
