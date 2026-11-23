"""The Hermes plugins: registration, the dispatcher contract, and catalogue
descriptions. office-connector needs Windows and Office and is checked
statically only."""
import ast
import importlib.util
import json
import re
from pathlib import Path

import pytest

from pdfbuilder import build_pdf

PLUGINS = Path(__file__).resolve().parent.parent / "plugins"
CATALOGUE_LIMIT = 60     # Hermes tool search shows the first sentence, cut here


class FakeContext:
    def __init__(self):
        self.tools, self.descriptions, self.skills, self.schemas = {}, {}, [], {}

    def register_tool(self, name, toolset, schema, handler, description="", emoji=""):
        # Hermes sends {**schema, "name": name} to the model as the function
        # definition (tools/registry.py, get_definitions), so the schema must be
        # a whole OpenAI function with a description and a parameters object.
        # A bare parameter schema reached the model as a tool with neither.
        assert schema.get("name") == name
        assert schema.get("description", "").strip(), f"{name}: no description in the schema"
        params = schema.get("parameters")
        assert isinstance(params, dict) and params.get("type") == "object", f"{name}: no parameters"
        self.tools[name] = handler
        self.descriptions[name] = description
        self.schemas[name] = schema

    def register_skill(self, name, path, description=""):
        assert Path(path).is_file()
        self.skills.append(name)


def load(plugin):
    spec = importlib.util.spec_from_file_location(
        f"plugin_{plugin.replace('-', '_')}", PLUGINS / plugin / "__init__.py",
        submodule_search_locations=[str(PLUGINS / plugin)])
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    ctx = FakeContext()
    module.register(ctx)
    return ctx


def first_sentence(description):
    return description.split(". ")[0].rstrip(".") + "."


def test_file_connector_registers_and_dispatches(tmp_path):
    ctx = load("file-connector")
    assert set(ctx.tools) == {"file_read", "file_find", "file_attachments", "file_render",
                              "file_collect", "file_unpack"}
    assert ctx.skills == ["file-reader"]
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(build_pdf(["The full text of the café opening rules"]))
    # Hermes calls handler(args_dict, **context); the dict comes positionally.
    found = json.loads(ctx.tools["file_find"]({"path": str(pdf), "query": "cafe"}, task_id="t"))
    assert found["ok"] and found["hit_count"] == 1
    read = json.loads(ctx.tools["file_read"]({"path": str(pdf)}, task_id="t", session_id="s"))
    assert read["ok"] and read["page_sources"] == {"text-layer": "1"}
    missing = json.loads(ctx.tools["file_read"]({}, task_id="t"))
    assert missing["ok"] is False and missing["hint"]


def test_ms_graph_connector_registers():
    ctx = load("ms-graph-connector")
    assert set(ctx.tools) == {"m365_status", "m365_login", "m365_search", "m365_get"}


@pytest.mark.parametrize("plugin", ["file-connector", "ms-graph-connector"])
def test_first_sentences_fit_the_catalogue(plugin):
    for name, description in load(plugin).descriptions.items():
        assert len(first_sentence(description)) <= CATALOGUE_LIMIT, name


def test_office_connector_descriptions_statically():
    """Registering office-connector starts a COM thread; read its source instead."""
    tree = ast.parse((PLUGINS / "office-connector" / "__init__.py").read_text(encoding="utf-8"))
    calls = [n for n in ast.walk(tree) if isinstance(n, ast.Call)
             and getattr(n.func, "attr", "") == "register_tool"
             and getattr(n.func.value, "id", "") == "ctx"]
    assert len(calls) == 3
    for call in calls:
        description = next(k.value for k in call.keywords if k.arg == "description")
        text = "".join(c.value for c in ast.walk(description) if isinstance(c, ast.Constant)
                       and isinstance(c.value, str))
        assert len(first_sentence(text)) <= CATALOGUE_LIMIT


@pytest.mark.parametrize("plugin", ["file-connector", "ms-graph-connector", "office-connector"])
def test_register_wraps_the_context_for_function_schemas(plugin):
    """office-connector cannot be registered here, so every plugin is also
    checked in its source: register() starts by wrapping ctx."""
    tree = ast.parse((PLUGINS / plugin / "__init__.py").read_text(encoding="utf-8"))
    register = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "register")
    first = register.body[0]
    assert isinstance(first, ast.Assign) and ast.unparse(first) == "ctx = _FunctionSchemas(ctx)"


def test_function_schema_wrapper():
    source = (PLUGINS / "office-connector" / "__init__.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == "_FunctionSchemas")
    ns = {}
    exec(compile(ast.Module(body=[cls], type_ignores=[]), "<adapter>", "exec"), ns)
    ctx = ns["_FunctionSchemas"](FakeContext())
    ctx.register_tool("t", "x", {"type": "object", "properties": {"code": {"type": "string"}}},
                      lambda a=None, **_: "", description="Run it. More text.")
    assert ctx.schemas["t"]["parameters"]["properties"] == {"code": {"type": "string"}}
    assert ctx.skills == []                       # attributes pass through to the real context


@pytest.mark.parametrize("plugin", ["file-connector", "ms-graph-connector", "office-connector"])
def test_manifest_lists_the_registered_tools(plugin):
    manifest = (PLUGINS / plugin / "plugin.yaml").read_text(encoding="utf-8")
    listed = set(re.findall(r"^  - (\w+)$", manifest.split("provides_tools:")[1].split("\n\n")[0], re.M))
    source = (PLUGINS / plugin / "__init__.py").read_text(encoding="utf-8")
    registered = set(re.findall(r'register_tool\(\s*"(\w+)"', source))
    assert listed == registered
