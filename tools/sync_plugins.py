#!/usr/bin/env python3
"""Copy each skill into the Hermes plugin that wraps it, and fail if they drift.

There is one source of truth for every line of Python here, and it is
``.agents/skills/<skill>/``. The plugins under ``plugins/`` are adapters: a
manifest, a ``register()`` that exposes the functions as Hermes tools, and a
*copy* of the skill's files.

The copy is not a design choice, it is what ``hermes plugins install`` needs.
It installs a single directory -- the repository's own subdirectory named in
the catalogue entry -- and nothing outside that directory comes with it. A
plugin that reached up into ``.agents/`` would validate here and be broken for
everyone who installed it.

So the copy exists, and this script makes it a derived artefact rather than a
second original:

    py tools/sync_plugins.py            refresh the copies
    py tools/sync_plugins.py --check    fail if they differ (CI)

Edit the skill (or the module under shared/). Run the sync. Commit all of it.
"""
from __future__ import annotations

import argparse
import filecmp
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# plugin directory -> the skill it wraps. A plugin may ship the skill document
# too; the agent does not discover a plugin-registered skill from
# `<available_skills>`, but it is useful reference material once the tools have
# been reached for.
PAIRS = {
    "office-connector": "office-live",
    "ms-graph-connector": "m365-graph",
    "file-connector": "file-reader",
}

# What travels. Anything else in a skill directory stays there.
COPIED = ("SKILL.md", "scripts")

# Modules more than one skill needs. A skill has to stay self-contained -- it
# is installed as one directory -- so each gets a copy in its scripts/, made
# here from the one source under shared/. Step one of every run, so the
# plugin copies below then pick the fresh file up like any other script.
SHARED = {
    "shared/pdfread.py": ("file-reader", "m365-graph"),
}


def sync_shared(check: bool) -> int:
    problems = 0
    for source_rel, skills in SHARED.items():
        source = ROOT / source_rel
        for skill in skills:
            destination = ROOT / ".agents" / "skills" / skill / "scripts" / source.name
            if destination.is_file() and filecmp.cmp(source, destination, shallow=False):
                continue
            if check:
                print(f"     {skill}/scripts/{source.name} differs from {source_rel}")
                problems += 1
                continue
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)
            print(f"     copied {source_rel} -> {skill}/scripts/{source.name}")
    if problems:
        print("FAIL shared modules")
    else:
        print(f"ok   shared modules ({sum(len(v) for v in SHARED.values())} copies)")
    return problems


def sources(skill: str) -> list[tuple[Path, Path]]:
    """(source, destination-relative-path) for one skill's shipped files."""
    base = ROOT / ".agents" / "skills" / skill
    out: list[tuple[Path, Path]] = []
    for name in COPIED:
        item = base / name
        if item.is_file():
            out.append((item, Path(name)))
        elif item.is_dir():
            for child in sorted(item.rglob("*")):
                if child.is_file() and "__pycache__" not in child.parts:
                    out.append((child, child.relative_to(base)))
    return out


def sync(plugin: str, skill: str, check: bool) -> int:
    target = ROOT / "plugins" / plugin
    if not (target / "plugin.yaml").is_file():
        print(f"FAIL {plugin}: no plugin.yaml")
        return 1

    expected = sources(skill)
    if not expected:
        print(f"FAIL {plugin}: {skill} has nothing to copy")
        return 1

    problems = 0
    wanted = {relative for _, relative in expected}

    for source, relative in expected:
        destination = target / relative
        same = destination.is_file() and filecmp.cmp(source, destination, shallow=False)
        if same:
            continue
        if check:
            print(f"     {plugin}/{relative.as_posix()} differs from {skill}")
            problems += 1
            continue
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
        print(f"     copied {relative.as_posix()}")

    # A file dropped from the skill has to disappear from the plugin too,
    # otherwise the plugin keeps shipping code nobody maintains any more.
    for existing in sorted((target / "scripts").rglob("*")) if (target / "scripts").is_dir() else []:
        if not existing.is_file():
            continue
        relative = existing.relative_to(target)
        if relative in wanted:
            continue
        if check:
            print(f"     {plugin}/{relative.as_posix()} is not in {skill} any more")
            problems += 1
        else:
            existing.unlink()
            print(f"     removed {relative.as_posix()}")

    if problems:
        print(f"FAIL {plugin}")
    else:
        print(f"ok   {plugin} <- {skill} ({len(expected)} file(s))")
    return problems


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true",
                        help="report drift and exit non-zero instead of copying")
    arguments = parser.parse_args()

    total = sync_shared(arguments.check)
    total += sum(sync(plugin, skill, arguments.check) for plugin, skill in PAIRS.items())

    print()
    if total:
        print(f"{total} file(s) out of step. Run: py tools/sync_plugins.py", file=sys.stderr)
        return 1
    print("Plugins match their skills.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
