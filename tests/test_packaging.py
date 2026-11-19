"""Every plugin directory is one skill packaged for every platform at once:
Hermes (plugin.yaml + __init__.py), Agent Plugins 1.0 (plugin.json), Claude
Code (.claude-plugin/plugin.json), plus the marketplace at the root. These
tests hold the packages to the rules each platform's loader enforces, so a
store rejection is found here first."""
import json
import re
from pathlib import Path

import pytest

import sync_plugins

ROOT = Path(__file__).resolve().parent.parent
PLUGINS = sync_plugins.PLUGINS

# Agent Plugins 1.0 (the rules Hermes' portable loader applies, too).
AP_SCHEMA = "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json"
AP_FIELDS = {"$schema", "name", "version", "description", "author", "homepage", "repository",
             "license", "keywords", "extensions"}
AP_NAME = re.compile(r"^(?!.*(?:--|\.\.))[a-z0-9](?:[a-z0-9.-]*[a-z0-9])?$")
# Agent Skills.
SKILL_NAME = re.compile(r"^(?!.*--)[a-z0-9]+(?:-[a-z0-9]+)*$")
SKILL_FIELDS = {"name", "description", "license", "compatibility", "metadata", "allowed-tools"}
# Claude Code: names that pass as Anthropic's own are refused.
CLAUDE_RESERVED = re.compile(r"^(claude|anthropic|anthropics|cc-plugin)(-|$)|^claude-code$")


def frontmatter(path: Path) -> dict:
    block = path.read_text(encoding="utf-8").split("---")[1]
    out, current = {}, None
    for line in block.splitlines():
        if not line.strip():
            continue
        if line.startswith(" ") and current:
            key, _, value = line.strip().partition(":")
            out.setdefault(current, {})[key] = value.strip().strip('"')
            continue
        key, _, value = line.partition(":")
        current = key.strip()
        out[current] = value.strip() if value.strip() else {}
    return out


def test_the_sync_is_current():
    drift = sync_plugins.Drift(check=True)
    sync_plugins.sync_shared(drift)
    for plugin, (skill, keywords) in PLUGINS.items():
        sync_plugins.sync_plugin(drift, plugin, skill, keywords)
    assert drift.problems == 0, "run: py tools/sync_plugins.py"


@pytest.mark.parametrize("plugin", PLUGINS)
def test_standard_layout(plugin):
    root = ROOT / "plugins" / plugin
    skill = PLUGINS[plugin][0]
    assert (root / "plugin.yaml").is_file() and (root / "__init__.py").is_file()       # Hermes
    assert (root / "plugin.json").is_file()                                              # Agent Plugins
    assert (root / ".claude-plugin" / "plugin.json").is_file()                           # Claude Code
    assert (root / "skills" / skill / "SKILL.md").is_file()
    assert not (root / "SKILL.md").exists() and not (root / "scripts").exists()          # old layout gone
    assert [p.name for p in (root / "skills").iterdir()] == [skill]


@pytest.mark.parametrize("plugin", PLUGINS)
def test_agent_plugins_manifest(plugin):
    data = json.loads((ROOT / "plugins" / plugin / "plugin.json").read_text(encoding="utf-8"))
    assert data["$schema"] == AP_SCHEMA
    assert set(data) <= AP_FIELDS
    assert AP_NAME.fullmatch(data["name"]) and len(data["name"]) <= 64
    assert set(data["author"]) <= {"name", "email", "url"}
    assert all(isinstance(k, str) for k in data["keywords"])
    assert data["license"] == "MIT"


@pytest.mark.parametrize("plugin", PLUGINS)
def test_all_manifests_agree(plugin):
    yaml_meta = sync_plugins.read_manifest(plugin)
    agent = json.loads((ROOT / "plugins" / plugin / "plugin.json").read_text(encoding="utf-8"))
    claude = json.loads((ROOT / "plugins" / plugin / ".claude-plugin" / "plugin.json").read_text(encoding="utf-8"))
    for key in ("name", "version", "description"):
        assert agent[key] == claude[key] == yaml_meta[key], key
    assert agent["name"] == plugin
    assert not CLAUDE_RESERVED.search(claude["name"])


@pytest.mark.parametrize("plugin", PLUGINS)
def test_skill_frontmatter_follows_the_spec(plugin):
    skill = PLUGINS[plugin][0]
    fm = frontmatter(ROOT / "plugins" / plugin / "skills" / skill / "SKILL.md")
    assert set(fm) <= SKILL_FIELDS, set(fm) - SKILL_FIELDS
    assert fm["name"] == skill and SKILL_NAME.fullmatch(skill)
    assert 1 <= len(fm["description"]) <= 1024
    assert len(fm.get("compatibility", "")) <= 500
    assert all(isinstance(v, str) for v in fm.get("metadata", {}).values())


def test_marketplace():
    data = json.loads((ROOT / ".claude-plugin" / "marketplace.json").read_text(encoding="utf-8"))
    assert data["name"] and data["owner"]["name"]
    assert [p["name"] for p in data["plugins"]] == list(PLUGINS)
    for entry in data["plugins"]:
        assert entry["source"].startswith("./") and ".." not in entry["source"]
        source = ROOT / entry["source"]
        assert source.is_dir()
        manifest = json.loads((source / ".claude-plugin" / "plugin.json").read_text(encoding="utf-8"))
        assert manifest["name"] == entry["name"]           # entry name == manifest name
