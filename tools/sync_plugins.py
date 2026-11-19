#!/usr/bin/env python3
"""Build every plugin directory from its skill, for every platform, and fail
if anything drifts.

There is one source of truth for every line here: ``.agents/skills/<skill>/``
(and ``shared/`` for code more than one skill needs). Each directory under
``plugins/`` is the same skill packaged three ways at once:

    plugins/<plugin>/
      plugin.yaml, __init__.py      Hermes Agent (hand-written: the tool adapter)
      plugin.json                   Agent Plugins 1.0 (Codex, Cursor, Copilot, VS Code, ...)
      .claude-plugin/plugin.json    Claude Code
      skills/<skill>/               the skill itself: SKILL.md and scripts/, copied

and ``.claude-plugin/marketplace.json`` at the repository root lists them for
Claude Code. Hermes reads plugin.yaml and ignores plugin.json beside it, so
the three never compete.

The copies are not a design choice, they are what installers need: each one
takes a single directory and nothing outside it. This script makes them
derived artefacts rather than second originals:

    py tools/sync_plugins.py            refresh copies and manifests
    py tools/sync_plugins.py --check    fail if any differ (CI)

Edit the skill, shared/ or plugin.yaml. Run the sync. Commit all of it.
"""
from __future__ import annotations

import argparse
import filecmp
import json
import re
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

REPOSITORY = "https://github.com/integrityauthority/office-agent-skills"
AUTHOR = {"name": "Integritás Hatóság", "url": "https://github.com/integrityauthority"}
LICENSE = "MIT"
AGENT_PLUGINS_SCHEMA = "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json"
MARKETPLACE = {
    "name": "office-agent-skills",
    "description": "Agent skills for office work: read PDF, e-mail and Office files, work live in "
                   "Word, Excel, PowerPoint and Outlook, search your own Microsoft 365.",
}

# plugin directory -> the skill it carries, and discovery keywords.
PLUGINS = {
    "office-connector": ("office-live", ["office", "word", "excel", "powerpoint", "outlook", "windows", "com"]),
    "ms-graph-connector": ("m365-graph", ["microsoft-365", "graph", "onedrive", "sharepoint", "outlook", "teams"]),
    "file-connector": ("file-reader", ["pdf", "ocr", "msg", "eml", "email", "docx", "xlsx", "search"]),
}

# What travels from a skill directory. Anything else stays there.
COPIED = ("SKILL.md", "scripts")

# Modules more than one skill needs. A skill has to stay self-contained -- it
# is installed as one directory -- so each gets a copy in its scripts/, made
# from the one source under shared/. Step one of every run, so the plugin
# copies then pick the fresh file up like any other script.
SHARED = {
    "shared/pdfread.py": ("file-reader", "m365-graph"),
}


class Drift:
    def __init__(self, check: bool):
        self.check = check
        self.problems = 0

    def file(self, source: Path, destination: Path, label: str):
        if destination.is_file() and filecmp.cmp(source, destination, shallow=False):
            return
        if self.check:
            print(f"     {label} differs from {source.relative_to(ROOT).as_posix()}")
            self.problems += 1
            return
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
        print(f"     copied {label}")

    def text(self, content: str, destination: Path, label: str):
        if destination.is_file() and destination.read_text(encoding="utf-8") == content:
            return
        if self.check:
            print(f"     {label} is out of date")
            self.problems += 1
            return
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(content, encoding="utf-8", newline="\n")
        print(f"     wrote {label}")

    def remove(self, path: Path, label: str):
        if not path.exists():
            return
        if self.check:
            print(f"     {label} should not exist any more")
            self.problems += 1
            return
        shutil.rmtree(path) if path.is_dir() else path.unlink()
        print(f"     removed {label}")


def skill_files(skill: str) -> list[tuple[Path, Path]]:
    """(source, path relative to the skill directory) for one skill."""
    base = ROOT / ".agents" / "skills" / skill
    out = []
    for name in COPIED:
        item = base / name
        if item.is_file():
            out.append((item, Path(name)))
        elif item.is_dir():
            for child in sorted(item.rglob("*")):
                if child.is_file() and "__pycache__" not in child.parts:
                    out.append((child, child.relative_to(base)))
    return out


def read_manifest(plugin: str) -> dict:
    """name, version and description from plugin.yaml, without a YAML library:
    flat `key: value` lines, and a folded `key: >-` block."""
    lines = (ROOT / "plugins" / plugin / "plugin.yaml").read_text(encoding="utf-8").splitlines()
    out, i = {}, 0
    while i < len(lines):
        match = re.match(r"^([a-z_]+):\s*(.*)$", lines[i])
        i += 1
        if not match:
            continue
        key, value = match.groups()
        if value in (">-", ">", "|"):
            block = []
            while i < len(lines) and (lines[i].startswith("  ") or not lines[i].strip()):
                block.append(lines[i].strip())
                i += 1
            value = " ".join(part for part in block if part)
        out[key] = value.strip().strip('"').strip("'")
    return out


def manifests(plugin: str, keywords: list[str]) -> tuple[str, str]:
    meta = read_manifest(plugin)
    common = {
        "name": meta["name"],
        "version": meta["version"],
        "description": meta["description"],
        "author": AUTHOR,
        "homepage": f"{REPOSITORY}/tree/main/plugins/{plugin}",
        "repository": REPOSITORY,
        "license": LICENSE,
        "keywords": keywords,
    }
    agent_plugins = {"$schema": AGENT_PLUGINS_SCHEMA, **common}
    return (json.dumps(agent_plugins, ensure_ascii=False, indent=2) + "\n",
            json.dumps(common, ensure_ascii=False, indent=2) + "\n")


def marketplace() -> str:
    entries = []
    for plugin in PLUGINS:
        meta = read_manifest(plugin)
        entries.append({"name": meta["name"], "source": f"./plugins/{plugin}", "description": meta["description"]})
    data = {"name": MARKETPLACE["name"], "description": MARKETPLACE["description"],
            "owner": {"name": AUTHOR["name"]}, "plugins": entries}
    return json.dumps(data, ensure_ascii=False, indent=2) + "\n"


def sync_shared(drift: Drift):
    for source_rel, skills in SHARED.items():
        source = ROOT / source_rel
        for skill in skills:
            drift.file(source, ROOT / ".agents" / "skills" / skill / "scripts" / source.name,
                       f"{skill}/scripts/{source.name}")


def sync_plugin(drift: Drift, plugin: str, skill: str, keywords: list[str]):
    target = ROOT / "plugins" / plugin
    if not (target / "plugin.yaml").is_file():
        print(f"FAIL {plugin}: no plugin.yaml")
        drift.problems += 1
        return
    before = drift.problems
    skill_dir = target / "skills" / skill
    expected = skill_files(skill)
    for source, relative in expected:
        drift.file(source, skill_dir / relative, f"{plugin}/skills/{skill}/{relative.as_posix()}")

    # A file dropped from the skill has to disappear from the plugin too,
    # otherwise the plugin keeps shipping code nobody maintains any more.
    wanted = {skill_dir / relative for _, relative in expected}
    for existing in sorted((target / "skills").rglob("*")) if (target / "skills").is_dir() else []:
        if existing.is_file() and existing not in wanted and "__pycache__" not in existing.parts:
            drift.remove(existing, f"{plugin}/{existing.relative_to(target).as_posix()}")
    for other in sorted((target / "skills").iterdir()) if (target / "skills").is_dir() else []:
        if other.is_dir() and other.name != skill:
            drift.remove(other, f"{plugin}/skills/{other.name}/")
    # The layout before the per-platform packaging kept the skill at the root.
    drift.remove(target / "SKILL.md", f"{plugin}/SKILL.md (moved to skills/{skill}/)")
    drift.remove(target / "scripts", f"{plugin}/scripts/ (moved to skills/{skill}/)")

    agent_plugins, claude = manifests(plugin, keywords)
    drift.text(agent_plugins, target / "plugin.json", f"{plugin}/plugin.json")
    drift.text(claude, target / ".claude-plugin" / "plugin.json", f"{plugin}/.claude-plugin/plugin.json")

    if drift.problems == before:
        print(f"ok   {plugin} <- {skill} ({len(expected)} file(s), 3 manifests)")
    else:
        print(f"FAIL {plugin}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--check", action="store_true",
                        help="report drift and exit non-zero instead of writing")
    arguments = parser.parse_args()

    drift = Drift(arguments.check)
    sync_shared(drift)
    for plugin, (skill, keywords) in PLUGINS.items():
        sync_plugin(drift, plugin, skill, keywords)
    drift.text(marketplace(), ROOT / ".claude-plugin" / "marketplace.json", ".claude-plugin/marketplace.json")

    print()
    if drift.problems:
        print(f"{drift.problems} file(s) out of step. Run: py tools/sync_plugins.py", file=sys.stderr)
        return 1
    print("Plugins match their skills.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
