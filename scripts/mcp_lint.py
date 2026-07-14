#!/usr/bin/env python3
"""Lint MCP tool ergonomics against the audit Pass 1 taxonomy.

Usage:
    python3 scripts/mcp_lint.py [--staged] [--all] [--json] [files...]

The default mode lints every provided Python file. Use --staged from
pre-commit to lint only MCP tools whose definition overlaps staged changes.
Use --all without files to scan likely MCP server files in the repository.
"""

from __future__ import annotations

import argparse
import ast
import importlib
import json
import re
import subprocess
import sys
import tomllib
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

DESTRUCTIVE_VERB_RE = re.compile(
    r"(^|_)(delete|remove|purge|drop|send|publish|finalize|cleanup)(_|$)"
)
MUTATING_VERB_RE = re.compile(
    r"(^|_)(add|annotate|assign|create|delete|disable|drop|enable|finalize|"
    r"manage|move|publish|purge|remove|rename|restore|save|send|set|ship|"
    r"unassign|update|write)(_|$)"
)
ID_PARAM_RE = re.compile(r"(^id$|_id$|^name$|_name$|^path$|_path$|^key$|_key$)")
DISCOVERY_RE = re.compile(r"(?im)^\s*Discovery\s*:")
MCP_PATH_RE = re.compile(
    r"(^|/)mcp_server\.py$|(^|/)x_mcp_server\.py$|(^|/)[A-Za-z0-9_]+_mcp_server\.py$|"
    r"(^|/)mcp_servers/.*\.py$|(^|/)mcp-server/.*\.py$|(^|/)server\.py$"
)

SAFETY_PARAMS = {
    "confirm",
    "confirm_token",
    "dry_run",
    "force",
    "permanent",
    "preview",
}
LEGACY_BYPASS_PARAMS = {"schedule_now", "send_now"}
IGNORED_REQUIRED_PARAMS = {"self", "cls", "ctx", "context"}

EXPECTED_LOCAL_TOOLS = (
    "gsheets_list_tabs",
    "gsheets_read_range",
    "gsheets_write_range",
    "gsheets_append_rows",
    "gsheets_create_spreadsheet",
    "gsheets_copy_spreadsheet",
    "gsheets_search_spreadsheets",
    "gsheets_clear_range",
    "gsheets_recalculate_range",
)
EXPECTED_BROKER_TOOLS = tuple(
    name for name in EXPECTED_LOCAL_TOOLS if name != "gsheets_search_spreadsheets"
)
EXPECTED_ANNOTATIONS = {
    "gsheets_list_tabs": (True, False, True, True),
    "gsheets_read_range": (True, False, True, True),
    "gsheets_write_range": (False, True, False, True),
    "gsheets_append_rows": (False, False, False, True),
    "gsheets_create_spreadsheet": (False, False, False, True),
    "gsheets_copy_spreadsheet": (False, False, False, True),
    "gsheets_search_spreadsheets": (True, False, True, True),
    "gsheets_clear_range": (False, True, True, True),
    "gsheets_recalculate_range": (False, True, False, True),
}
REMOVED_TOOL_NAMES = {
    "gsheet_list_tabs",
    "gsheet_read_range",
    "gsheet_update_range",
    "gsheet_append_rows",
    "gsheet_create",
    "gsheet_copy_spreadsheet",
    "gsheet_search",
    "gsheet_clear_range",
    "gsheet_touch_range",
}


@dataclass(frozen=True)
class Issue:
    severity: str
    rule_id: str
    path: str
    line: int
    tool: str
    message: str

    def to_dict(self) -> dict[str, object]:
        return {
            "severity": self.severity,
            "rule_id": self.rule_id,
            "path": self.path,
            "line": self.line,
            "tool": self.tool,
            "message": self.message,
        }


@dataclass(frozen=True)
class ToolFunction:
    name: str
    node: ast.FunctionDef | ast.AsyncFunctionDef
    start_line: int
    end_line: int
    registration_lines: frozenset[int]
    structured_error_envelope: bool = False

    def overlaps(self, changed_lines: set[int] | None) -> bool:
        if changed_lines is None:
            return True
        own_lines = range(self.start_line, self.end_line + 1)
        return any(line in changed_lines for line in own_lines) or bool(
            self.registration_lines & changed_lines
        )


def is_mcp_tool(decorator: ast.expr) -> bool:
    """Return True for @mcp.tool, @mcp.tool(...), @server.tool, or @tool."""
    func = decorator.func if isinstance(decorator, ast.Call) else decorator
    if isinstance(func, ast.Attribute) and func.attr == "tool":
        return True
    if isinstance(func, ast.Name) and func.id == "tool":
        return True
    return False


def lint_file(path: Path, changed_lines: set[int] | None = None) -> list[Issue]:
    """Lint one file and return all MCP tool issues."""
    try:
        source = path.read_text(encoding="utf-8")
        tree = ast.parse(source, filename=str(path))
    except (OSError, SyntaxError, UnicodeDecodeError):
        return []

    tools = _find_tools(tree)
    sibling_names = {tool.name for tool in tools}
    issues: list[Issue] = []
    for tool in tools:
        if not tool.overlaps(changed_lines):
            continue
        issues.extend(_lint_tool(path, tool, sibling_names))
    return issues


def _find_tools(tree: ast.AST) -> list[ToolFunction]:
    registered = _registered_tool_lines(tree)
    structured_decorators = _structured_tool_decorator_names(tree)
    module_structured_tool_wrapper = _module_structured_tool_wrapper(tree)
    tools: list[ToolFunction] = []
    seen: set[str] = set()
    for node in ast.walk(tree):
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        decorated = any(
            is_mcp_tool(decorator)
            or _decorator_name(decorator) in structured_decorators
            for decorator in node.decorator_list
        )
        registration_lines = frozenset(registered.get(node.name, set()))
        if not decorated and not registration_lines:
            continue
        if registration_lines and node.name == "wrapped" and _has_wraps_decorator(node):
            continue

        decorator_lines = [decorator.lineno for decorator in node.decorator_list]
        start_line = min([node.lineno, *decorator_lines])
        end_line = getattr(node, "end_lineno", None) or node.lineno
        tools.append(
            ToolFunction(
                name=node.name,
                node=node,
                start_line=start_line,
                end_line=end_line,
                registration_lines=registration_lines,
                structured_error_envelope=module_structured_tool_wrapper
                or any(
                    _decorator_name(decorator) in structured_decorators
                    for decorator in node.decorator_list
                ),
            )
        )
        seen.add(node.name)

    return sorted(tools, key=lambda tool: (tool.start_line, tool.name))


def _registered_tool_lines(tree: ast.AST) -> dict[str, set[int]]:
    """Find mcp.tool()(function_name) registrations in a module."""
    out: dict[str, set[int]] = {}
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        if not isinstance(node.func, ast.Call) or not is_mcp_tool(node.func):
            continue
        if len(node.args) != 1 or not isinstance(node.args[0], ast.Name):
            continue
        out.setdefault(node.args[0].id, set()).add(node.lineno)
    return out


def _structured_tool_decorator_names(tree: ast.AST) -> set[str]:
    """Find local decorator variables created by _structured_tool(mcp)."""
    names: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            value = node.value
            targets = node.targets
        elif isinstance(node, ast.AnnAssign):
            value = node.value
            targets = [node.target]
        else:
            continue
        if (
            not isinstance(value, ast.Call)
            or _call_name(value.func) != "_structured_tool"
        ):
            continue
        for target in targets:
            if isinstance(target, ast.Name):
                names.add(target.id)
    return names


def _module_structured_tool_wrapper(tree: ast.AST) -> bool:
    """Detect modules that patch mcp.tool to wrap every tool in a structured envelope."""
    functions = {
        node.name: node
        for node in ast.walk(tree)
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
    }
    for node in ast.walk(tree):
        if not isinstance(node, ast.Assign):
            continue
        if not any(
            isinstance(target, ast.Attribute) and target.attr == "tool"
            for target in node.targets
        ):
            continue
        wrapper_name = _call_name(node.value)
        wrapper = functions.get(wrapper_name)
        if wrapper is None:
            continue
        if any(
            isinstance(call, ast.Call)
            and _call_name(call.func) == "_with_structured_error_envelope"
            for call in ast.walk(wrapper)
        ):
            return True
    return False


def _decorator_name(decorator: ast.expr) -> str:
    func = decorator.func if isinstance(decorator, ast.Call) else decorator
    return _call_name(func)


def _lint_tool(path: Path, tool: ToolFunction, sibling_names: set[str]) -> list[Issue]:
    func = tool.node
    doc = (ast.get_docstring(func) or "").strip()
    params = _param_names(func)
    required_params = _required_param_names(func)
    body_src = _safe_unparse(func)
    issues: list[Issue] = []

    def add(severity: str, rule_id: str, message: str) -> None:
        issues.append(
            Issue(
                severity=severity,
                rule_id=rule_id,
                path=str(path),
                line=tool.start_line,
                tool=tool.name,
                message=message,
            )
        )

    if len(doc) <= 20:
        add(
            "error",
            "L-001",
            f"{tool.name}: docstring is empty or <=20 chars (current: {len(doc)})",
        )

    if (
        DESTRUCTIVE_VERB_RE.search(tool.name)
        and not _is_preview_tool_name(tool.name)
        and not (params & SAFETY_PARAMS)
    ):
        add(
            "error",
            "L-002",
            f"{tool.name}: destructive verb in name but no safety param "
            f"({', '.join(sorted(SAFETY_PARAMS))})",
        )

    legacy_bypass_params = sorted(params & LEGACY_BYPASS_PARAMS)
    if legacy_bypass_params:
        add(
            "error",
            "L-008",
            f"{tool.name}: legacy immediate-execution bypass params are not allowed "
            f"({', '.join(legacy_bypass_params)}); require confirm_token or split draft/send tools",
        )

    enum_values = _literal_enum_values(func)
    if (
        enum_values
        and not _doc_surfaces_enum_values(doc, enum_values)
        and "Literal[" not in body_src
    ):
        preview = ", ".join(sorted(enum_values)[:5])
        add(
            "warn",
            "L-003",
            f"{tool.name}: validates enum values ({preview}) but docstring does not surface them",
        )

    id_like_required = sorted(
        param
        for param in required_params
        if param not in IGNORED_REQUIRED_PARAMS and ID_PARAM_RE.search(param)
    )
    if id_like_required and not DISCOVERY_RE.search(doc):
        add(
            "warn",
            "L-004",
            f"{tool.name}: required ID-like params {id_like_required} lack a Discovery: block",
        )

    near_siblings = sorted(
        name
        for name in sibling_names
        if name != tool.name and _share_two_token_prefix(name, tool.name)
    )
    if near_siblings and not any(name in doc for name in near_siblings):
        add(
            "warn",
            "L-005",
            f"{tool.name}: near-name siblings {near_siblings[:3]} exist but docstring does not cross-reference them",
        )

    if (
        _returns_bare_string(func)
        and _raises_generic_error(func)
        and not _has_try_envelope(func)
    ):
        add(
            "warn",
            "L-006",
            f"{tool.name}: returns a bare string and raises generic errors; prefer a structured error envelope",
        )

    if _looks_mutating(tool.name, func) and not (
        tool.structured_error_envelope or _has_try_envelope(func)
    ):
        add(
            "warn",
            "L-007",
            f"{tool.name}: appears to mutate state without an explicit try/except result envelope",
        )

    return issues


def _param_names(func: ast.FunctionDef | ast.AsyncFunctionDef) -> set[str]:
    return {
        *(arg.arg for arg in func.args.posonlyargs),
        *(arg.arg for arg in func.args.args),
        *(arg.arg for arg in func.args.kwonlyargs),
    }


def _has_wraps_decorator(func: ast.FunctionDef | ast.AsyncFunctionDef) -> bool:
    return any(
        _call_name(decorator.func if isinstance(decorator, ast.Call) else decorator)
        == "wraps"
        for decorator in func.decorator_list
    )


def _is_preview_tool_name(name: str) -> bool:
    return (
        name.startswith("preview_") or name.endswith("_preview") or "_preview_" in name
    )


def _required_param_names(func: ast.FunctionDef | ast.AsyncFunctionDef) -> set[str]:
    required: set[str] = set()
    positional = [*func.args.posonlyargs, *func.args.args]
    default_start = len(positional) - len(func.args.defaults)
    for index, arg in enumerate(positional):
        if index < default_start:
            required.add(arg.arg)
    for arg, default in zip(func.args.kwonlyargs, func.args.kw_defaults):
        if default is None:
            required.add(arg.arg)
    return required


def _safe_unparse(node: ast.AST) -> str:
    try:
        return ast.unparse(node)
    except Exception:
        return ""


def _literal_enum_values(func: ast.FunctionDef | ast.AsyncFunctionDef) -> set[str]:
    values: set[str] = set()
    for node in ast.walk(func):
        if not isinstance(node, ast.Compare):
            continue
        if not any(isinstance(op, (ast.In, ast.NotIn)) for op in node.ops):
            continue
        for comparator in node.comparators:
            values.update(_constant_container_values(comparator))
    return values


def _constant_container_values(node: ast.AST) -> set[str]:
    if not isinstance(node, (ast.Set, ast.Tuple, ast.List)):
        return set()
    values: set[str] = set()
    for elt in node.elts:
        if isinstance(elt, ast.Constant) and isinstance(elt.value, (str, int, float)):
            values.add(str(elt.value))
    return values


def _doc_surfaces_enum_values(doc: str, values: set[str]) -> bool:
    if not values:
        return True
    lower_doc = doc.lower()
    if "valid" in lower_doc or "allowed" in lower_doc or "one of" in lower_doc:
        return True
    return all(str(value).lower() in lower_doc for value in values)


def _share_two_token_prefix(a: str, b: str) -> bool:
    left = a.split("_")
    right = b.split("_")
    return len(left) >= 2 and len(right) >= 2 and left[:2] == right[:2]


def _returns_bare_string(func: ast.FunctionDef | ast.AsyncFunctionDef) -> bool:
    if func.returns is not None and _safe_unparse(func.returns) in {
        "str",
        "builtins.str",
    }:
        return True
    for node in ast.walk(func):
        if not isinstance(node, ast.Return) or node.value is None:
            continue
        if isinstance(node.value, (ast.Constant, ast.JoinedStr)):
            if isinstance(node.value, ast.JoinedStr) or isinstance(
                getattr(node.value, "value", None), str
            ):
                return True
        if (
            isinstance(node.value, ast.Call)
            and isinstance(node.value.func, ast.Name)
            and node.value.func.id == "str"
        ):
            return True
    return False


def _raises_generic_error(func: ast.FunctionDef | ast.AsyncFunctionDef) -> bool:
    for node in ast.walk(func):
        if not isinstance(node, ast.Raise):
            continue
        exc = node.exc
        if exc is None:
            return True
        if isinstance(exc, ast.Call):
            exc = exc.func
        if isinstance(exc, ast.Name) and exc.id in {
            "Exception",
            "RuntimeError",
            "ValueError",
        }:
            return True
    return False


def _looks_mutating(name: str, func: ast.FunctionDef | ast.AsyncFunctionDef) -> bool:
    if name.startswith(
        ("get_", "list_", "find_", "search_", "preview_", "show_", "read_")
    ):
        return False
    if MUTATING_VERB_RE.search(name):
        return True
    for node in ast.walk(func):
        if not isinstance(node, ast.Call):
            continue
        called = _call_name(node.func)
        if called in {
            "commit",
            "delete",
            "execute",
            "patch",
            "post",
            "put",
            "rename",
            "replace",
            "run",
            "unlink",
            "write",
        }:
            return True
    return False


def _call_name(func: ast.AST) -> str:
    if isinstance(func, ast.Attribute):
        return func.attr
    if isinstance(func, ast.Name):
        return func.id
    return ""


def _has_try_envelope(func: ast.FunctionDef | ast.AsyncFunctionDef) -> bool:
    if any(isinstance(node, ast.Try) for node in ast.walk(func)):
        return True
    return any(
        _call_name(node.func) in {"_run_tool", "run_tool"}
        for node in ast.walk(func)
        if isinstance(node, ast.Call)
    )


def discover_mcp_files(root: Path) -> list[Path]:
    files: list[Path] = []
    for path in root.rglob("*.py"):
        if _is_ignored_path(path):
            continue
        rel = path.relative_to(root).as_posix()
        if MCP_PATH_RE.search(rel):
            files.append(path)
    return sorted(files)


def _is_ignored_path(path: Path) -> bool:
    ignored_parts = {
        ".git",
        ".mypy_cache",
        ".pytest_cache",
        ".ruff_cache",
        ".venv",
        "__pycache__",
        "node_modules",
        "dist",
    }
    return any(part in ignored_parts for part in path.parts)


def staged_changed_lines(
    files: Sequence[Path], cwd: Path | None = None
) -> dict[Path, set[int]]:
    """Return staged changed destination lines per file."""
    if not files:
        return {}
    cwd = cwd or Path.cwd()
    cmd = [
        "git",
        "diff",
        "--cached",
        "--unified=0",
        "--",
        *[str(path) for path in files],
    ]
    try:
        result = subprocess.run(
            cmd, cwd=cwd, capture_output=True, text=True, check=False
        )
    except OSError:
        return {}
    if result.returncode != 0:
        return {}
    return _parse_unified_diff_changed_lines(result.stdout, cwd)


def _parse_unified_diff_changed_lines(
    diff_text: str, cwd: Path
) -> dict[Path, set[int]]:
    out: dict[Path, set[int]] = {}
    current: Path | None = None
    new_line: int | None = None
    for raw_line in diff_text.splitlines():
        if raw_line.startswith("+++ "):
            current = _diff_path_to_path(raw_line[4:].strip(), cwd)
            if current is not None:
                out.setdefault(current, set())
            new_line = None
            continue
        if raw_line.startswith("@@ "):
            match = re.search(r"\+(\d+)(?:,(\d+))?", raw_line)
            new_line = int(match.group(1)) if match else None
            continue
        if current is None or new_line is None:
            continue
        if raw_line.startswith("+") and not raw_line.startswith("+++"):
            out.setdefault(current, set()).add(new_line)
            new_line += 1
        elif raw_line.startswith("-") and not raw_line.startswith("---"):
            continue
        else:
            new_line += 1
    return out


def _diff_path_to_path(raw_path: str, cwd: Path) -> Path | None:
    if raw_path == "/dev/null":
        return None
    if raw_path.startswith("b/"):
        raw_path = raw_path[2:]
    return (cwd / raw_path).resolve()


def _normalize_files(values: Iterable[str], root: Path, all_files: bool) -> list[Path]:
    explicit = [Path(value) for value in values]
    if not explicit and all_files:
        return discover_mcp_files(root)
    out: list[Path] = []
    for path in explicit:
        resolved = path if path.is_absolute() else root / path
        if resolved.exists() and resolved.suffix == ".py":
            out.append(resolved.resolve())
    return out


def run_lint(
    files: Sequence[Path], staged_only: bool = False, cwd: Path | None = None
) -> list[Issue]:
    cwd = cwd or Path.cwd()
    changes = staged_changed_lines(files, cwd=cwd) if staged_only else {}
    issues: list[Issue] = []
    for path in files:
        changed_lines = changes.get(path.resolve(), set()) if staged_only else None
        if staged_only and not changed_lines:
            continue
        issues.extend(lint_file(path, changed_lines=changed_lines))
    return issues


def _registry_issue(root: Path, rule_id: str, tool: str, message: str) -> Issue:
    return Issue(
        severity="error",
        rule_id=rule_id,
        path=str(root / "src/gsheets_mcp/server.py"),
        line=1,
        tool=tool,
        message=message,
    )


def _object_schemas_without_closed_world(schema: object, path: str = "$") -> list[str]:
    failures: list[str] = []
    if isinstance(schema, dict):
        if (
            schema.get("type") == "object"
            and schema.get("additionalProperties") is not False
        ):
            failures.append(path)
        for key, value in schema.items():
            failures.extend(
                _object_schemas_without_closed_world(value, f"{path}.{key}")
            )
    elif isinstance(schema, list):
        for index, value in enumerate(schema):
            failures.extend(
                _object_schemas_without_closed_world(value, f"{path}[{index}]")
            )
    return failures


def _schema_has_key(schema: object, target: str) -> bool:
    if isinstance(schema, dict):
        return target in schema or any(
            _schema_has_key(value, target) for value in schema.values()
        )
    if isinstance(schema, list):
        return any(_schema_has_key(value, target) for value in schema)
    return False


def lint_live_registry(root: Path) -> list[Issue]:
    """Validate the executable registry and schemas without loading credentials."""
    issues: list[Issue] = []
    src_path = str(root / "src")
    if src_path not in sys.path:
        sys.path.insert(0, src_path)
    try:
        package = importlib.import_module("gsheets_mcp")
        server_module = importlib.import_module("gsheets_mcp.server")
        from pydantic import ValidationError
    except Exception as exc:
        return [
            _registry_issue(
                root,
                "R-001",
                "registry",
                f"live registry could not be imported ({type(exc).__name__})",
            )
        ]

    if package.__version__ != "1.0.0" or package.CONTRACT_VERSION != "1.0":
        issues.append(
            _registry_issue(
                root,
                "R-002",
                "registry",
                "package version must be 1.0.0 and wire contract must be 1.0",
            )
        )

    try:
        local_specs = server_module.tool_specs("local")
        broker_specs = server_module.tool_specs("broker")
    except Exception as exc:
        issues.append(
            _registry_issue(
                root,
                "R-003",
                "registry",
                f"credential-free registry construction failed ({type(exc).__name__})",
            )
        )
        return issues

    local_names = tuple(spec.name for spec in local_specs)
    broker_names = tuple(spec.name for spec in broker_specs)
    if local_names != EXPECTED_LOCAL_TOOLS:
        issues.append(
            _registry_issue(
                root,
                "R-004",
                "registry",
                "local mode must expose the exact nine-tool contract",
            )
        )
    if broker_names != EXPECTED_BROKER_TOOLS:
        issues.append(
            _registry_issue(
                root,
                "R-005",
                "registry",
                "broker mode must expose exactly eight tools without search",
            )
        )
    if REMOVED_TOOL_NAMES & set(local_names):
        issues.append(
            _registry_issue(
                root, "R-006", "registry", "removed tool names remain executable"
            )
        )

    for spec in local_specs:
        definition = spec.definition()
        input_schema = definition.inputSchema
        output_schema = definition.outputSchema or {}
        if _object_schemas_without_closed_world(input_schema):
            issues.append(
                _registry_issue(
                    root,
                    "R-007",
                    spec.name,
                    "input object schemas must recursively forbid unknown fields",
                )
            )
        if _object_schemas_without_closed_world(output_schema):
            issues.append(
                _registry_issue(
                    root,
                    "R-008",
                    spec.name,
                    "output object schemas must recursively forbid unknown fields",
                )
            )
        if output_schema.get("discriminator", {}).get("propertyName") != "status":
            issues.append(
                _registry_issue(
                    root,
                    "R-009",
                    spec.name,
                    "output schema must be a status-discriminated success/error union",
                )
            )
        if _schema_has_key(output_schema, "result"):
            issues.append(
                _registry_issue(
                    root, "R-010", spec.name, "generic result wrappers are forbidden"
                )
            )
        expected = EXPECTED_ANNOTATIONS[spec.name]
        actual = (
            definition.annotations.readOnlyHint,
            definition.annotations.destructiveHint,
            definition.annotations.idempotentHint,
            definition.annotations.openWorldHint,
        )
        if actual != expected:
            issues.append(
                _registry_issue(
                    root,
                    "R-011",
                    spec.name,
                    f"annotations {actual!r} do not match the canonical contract",
                )
            )
        try:
            spec.input_model.model_validate(
                {**spec.example_arguments.model_dump(), "dry_run": True}
            )
        except ValidationError:
            pass
        else:
            issues.append(
                _registry_issue(
                    root,
                    "R-012",
                    spec.name,
                    "unknown safety fields must fail strict validation",
                )
            )

    legacy_paths = (
        root / "run_server.py",
        root / "src/__init__.py",
        root / "src/server.py",
    )
    for path in legacy_paths:
        if path.exists():
            issues.append(
                _registry_issue(
                    root,
                    "R-013",
                    "package",
                    f"legacy runtime path still exists: {path.relative_to(root)}",
                )
            )

    try:
        project = tomllib.loads((root / "pyproject.toml").read_text(encoding="utf-8"))[
            "project"
        ]
        dependencies = set(project["dependencies"])
        scripts = project["scripts"]
    except (OSError, KeyError, tomllib.TOMLDecodeError, TypeError) as exc:
        issues.append(
            _registry_issue(
                root,
                "R-014",
                "package",
                f"pyproject contract could not be read ({type(exc).__name__})",
            )
        )
    else:
        if project.get("version") != "1.0.0":
            issues.append(
                _registry_issue(
                    root, "R-015", "package", "pyproject version must be 1.0.0"
                )
            )
        for dependency in {"mcp>=1.26.0,<2", "pydantic>=2.12,<3"}:
            if dependency not in dependencies:
                issues.append(
                    _registry_issue(
                        root,
                        "R-016",
                        "package",
                        f"missing dependency constraint: {dependency}",
                    )
                )
        if scripts.get("gsheets-mcp") != "gsheets_mcp.cli:main":
            issues.append(
                _registry_issue(
                    root,
                    "R-017",
                    "package",
                    "console entry point must target gsheets_mcp.cli:main",
                )
            )

    return issues


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("files", nargs="*", help="Python files to lint")
    parser.add_argument(
        "--staged",
        action="store_true",
        help="lint only MCP tool definitions touched by the staged diff",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="with no files, scan likely MCP server files under the current directory",
    )
    parser.add_argument("--json", action="store_true", help="emit JSON issue records")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv if argv is not None else sys.argv[1:])
    root = Path.cwd()
    files = _normalize_files(args.files, root=root, all_files=args.all)
    issues = lint_live_registry(root)
    issues.extend(run_lint(files, staged_only=args.staged, cwd=root))
    if args.json:
        print(
            json.dumps([issue.to_dict() for issue in issues], indent=2, sort_keys=True)
        )
    else:
        for issue in issues:
            prefix = "ERROR" if issue.severity == "error" else "WARN"
            print(
                f"[{prefix} {issue.rule_id}] {issue.path}:{issue.line}: {issue.message}",
                file=sys.stderr,
            )

    if any(issue.severity == "error" for issue in issues):
        if not args.json:
            print(
                "\nMCP tool lint failed. See docs/audits/_adapter/mcp_lint_design.md.",
                file=sys.stderr,
            )
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
