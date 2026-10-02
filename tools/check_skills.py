#!/usr/bin/env python3
"""Check every skill in this repository against what the publishing path accepts.

A skill leaves here by three routes and each one has its own rules:

  * copied into an agent's skill directory, where `SKILL.md` must carry the
    `name` and `description` the Agent Skills specification requires;
  * pushed to an AgentPlaybooks playbook, where a bundled file has to satisfy
    the `safe_filename` and `max_file_size` constraints on `skill_attachments`;
  * served from `/.well-known/skills/<name>/`, where the file name becomes a
    URL path segment.

The three agree, and this asserts the agreement locally so a problem surfaces
in a pull request rather than halfway through an upload.

Usage:  py tools/check_skills.py [--root .agents/skills]
"""
from __future__ import annotations

import argparse
import ast
import re
import sys
from pathlib import Path

# Mirrors `safe_filename` on skill_attachments, `isSafeSkillFile` in the web
# app, and the same rule in the AgentPlaybooks CLI.
SKILL_FILE_DIRECTORIES = ("scripts", "references", "assets", "examples", "templates")
SAFE_FILE = re.compile(
    r"^(?:(?:" + "|".join(SKILL_FILE_DIRECTORIES) + r")/)?[A-Za-z0-9][A-Za-z0-9._-]*$"
)
MAX_FILE_BYTES = 262_144
MAX_FILES_PER_SKILL = 10
MAX_FILENAME = 100

# Extensions the attachment store knows a type for. A file with any other
# extension is rejected on upload, so it is rejected here.
KNOWN_EXTENSIONS = {
    ".ts", ".tsx", ".js", ".jsx", ".mjs", ".py", ".go", ".rs", ".sql",
    ".md", ".json", ".yaml", ".yml", ".txt", ".csv", ".sh", ".cursorrules",
}

SAFE_SKILL_NAME = re.compile(r"^[a-z0-9](?:[a-z0-9-]{0,62}[a-z0-9])?$")
FRONTMATTER = re.compile(r"\A---\r?\n(.*?)\r?\n---(?:\r?\n|\Z)", re.DOTALL)


def frontmatter_fields(text: str) -> dict[str, str] | None:
    """The top-level scalar keys of the YAML block, without a YAML dependency.

    Only flat `key: value` lines matter here; a nested block is skipped rather
    than misparsed, because nothing this checks for lives inside one.
    """
    match = FRONTMATTER.match(text)
    if match is None:
        return None
    fields: dict[str, str] = {}
    for line in match.group(1).splitlines():
        if not line or line.startswith((" ", "\t", "#")):
            continue
        key, separator, value = line.partition(":")
        if separator:
            fields[key.strip()] = value.strip()
    return fields


def bundled_files(directory: Path) -> list[str]:
    """Every file the skill would publish, as a posix path relative to itself."""
    names = []
    for entry in sorted(directory.iterdir()):
        if entry.is_file() and entry.name != "SKILL.md":
            names.append(entry.name)
        elif entry.is_dir() and entry.name in SKILL_FILE_DIRECTORIES:
            names.extend(f"{entry.name}/{child.name}" for child in sorted(entry.iterdir()) if child.is_file())
        elif entry.is_dir() and entry.name != "__pycache__":
            names.append(f"{entry.name}/")  # reported as unpublishable below
    return names


def module_exports(path: Path) -> set[str] | None:
    """Top-level names a module defines, or None if it will not parse."""
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    except SyntaxError:
        return None
    names: set[str] = set()
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            names.add(node.name)
        elif isinstance(node, ast.Assign):
            names.update(t.id for t in node.targets if isinstance(t, ast.Name))
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            names.add(node.target.id)
        elif isinstance(node, (ast.Import, ast.ImportFrom)):
            names.update(a.asname or a.name.split(".")[0] for a in node.names)
    return names


def check_documented_imports(directory: Path, text: str) -> list[str]:
    """Every `from <module> import a, b` in the document must name something
    the module really defines.

    This is the failure the repository exists to prevent, one level deeper
    than a missing file: the SKILL.md listed `status` beside `attach` and
    `probe` as though all three were importable. Only two were, so an agent
    that believed the document got an ImportError at the first call.
    """
    problems = []
    pattern = r"^\s*from\s+([A-Za-z_][A-Za-z0-9_]*)\s+import\s+([^\n(]+)"
    for module, imported in re.findall(pattern, text, re.M):
        source = directory / "scripts" / f"{module}.py"
        if not source.is_file():
            continue  # a third-party module, not ours to verify
        exports = module_exports(source)
        if exports is None:
            continue  # the parse failure is reported on its own
        for name in (n.strip().split(" as ")[0].strip() for n in imported.split(",")):
            if name and name != "*" and name not in exports:
                problems.append(f"SKILL.md imports '{name}' from {module}, which does not define it")
    return problems


def check_skill(directory: Path) -> list[str]:
    problems: list[str] = []
    name = directory.name

    if not SAFE_SKILL_NAME.match(name):
        problems.append(f"directory name '{name}' is not lowercase kebab-case")

    document = directory / "SKILL.md"
    if not document.is_file():
        return [f"{name}: no SKILL.md"]

    text = document.read_text(encoding="utf-8")
    fields = frontmatter_fields(text)
    if fields is None:
        problems.append("SKILL.md has no YAML frontmatter")
    else:
        for key in ("name", "description"):
            if not fields.get(key):
                problems.append(f"SKILL.md frontmatter is missing '{key}'")
        if fields.get("name") and fields["name"] != name:
            problems.append(f"SKILL.md declares name '{fields['name']}' but the directory is '{name}'")

    files = bundled_files(directory)
    if len(files) > MAX_FILES_PER_SKILL:
        problems.append(f"{len(files)} bundled files; the store accepts {MAX_FILES_PER_SKILL}")

    for relative in files:
        if relative.endswith("/"):
            problems.append(f"'{relative}' is a directory the publishing path cannot carry; "
                            f"use one of {', '.join(SKILL_FILE_DIRECTORIES)}")
            continue
        if len(relative) > MAX_FILENAME:
            problems.append(f"'{relative}' is longer than {MAX_FILENAME} characters")
        if not SAFE_FILE.match(relative):
            problems.append(f"'{relative}' is not a publishable file name")
            continue

        path = directory / relative
        size = path.stat().st_size
        if size > MAX_FILE_BYTES:
            problems.append(f"'{relative}' is {size} bytes; the limit is {MAX_FILE_BYTES}")
        if path.suffix.lower() not in KNOWN_EXTENSIONS:
            problems.append(f"'{relative}' has an extension the attachment store has no type for")

        if path.suffix == ".py":
            try:
                ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            except SyntaxError as error:
                problems.append(f"'{relative}' does not parse: line {error.lineno}: {error.msg}")

        # A SKILL.md that tells the agent to import a file the skill does not
        # ship is the failure this whole check exists to prevent, so the
        # reference is verified rather than assumed.
    problems.extend(check_documented_imports(directory, text))

    mentioned = re.findall(r"(?:scripts|references|assets|examples|templates)/[A-Za-z0-9][A-Za-z0-9._-]*", text)
    # A path at the end of a sentence picks up the full stop; a file name never
    # ends in one, so trailing punctuation is prose rather than part of it.
    for referenced in sorted({name.rstrip(".,;:") for name in mentioned}):
        if referenced not in files:
            problems.append(f"SKILL.md refers to '{referenced}', which the skill does not bundle")

    return [f"{name}: {problem}" for problem in problems]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", default=".agents/skills", help="directory holding the skills")
    arguments = parser.parse_args()

    root = Path(arguments.root)
    if not root.is_dir():
        print(f"No skills directory at {root}", file=sys.stderr)
        return 1

    directories = sorted(entry for entry in root.iterdir() if entry.is_dir())
    if not directories:
        print(f"No skills under {root}", file=sys.stderr)
        return 1

    failures: list[str] = []
    for directory in directories:
        problems = check_skill(directory)
        if problems:
            failures.extend(problems)
            print(f"FAIL {directory.name}")
            for problem in problems:
                print(f"     {problem}")
        else:
            files = bundled_files(directory)
            print(f"ok   {directory.name} ({len(files)} bundled file(s))")

    if failures:
        print(f"\n{len(failures)} problem(s).", file=sys.stderr)
        return 1
    print(f"\n{len(directories)} skill(s) ready to publish.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
