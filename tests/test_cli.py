"""CLI and installed-package acceptance tests for the atomic cutover."""

from __future__ import annotations

import json
import os
import subprocess
import sys
from importlib.metadata import version
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CONSOLE = Path(sys.executable).with_name("gsheets-mcp")


def _run(*arguments: str, env: dict[str, str] | None = None, cwd: Path = ROOT):
    return subprocess.run(
        [str(CONSOLE), *arguments],
        cwd=cwd,
        env=env,
        capture_output=True,
        text=True,
        timeout=15,
        check=False,
    )


def test_bare_cli_and_help_are_explicit_and_do_not_start_stdio() -> None:
    bare = _run()
    help_result = _run("--help")

    assert bare.returncode == 0
    assert help_result.returncode == 0
    assert bare.stderr == ""
    assert help_result.stderr == ""
    assert "usage: gsheets-mcp" in bare.stdout
    assert "serve" in bare.stdout
    assert "capabilities" in bare.stdout
    assert bare.stdout == help_result.stdout


def test_version_is_the_exact_package_version_only() -> None:
    result = _run("--version")

    assert result.returncode == 0
    assert result.stdout == "1.0.0\n"
    assert result.stderr == ""
    assert version("gsheets-mcp") == "1.0.0"


def test_capabilities_json_is_credential_free_and_mode_specific() -> None:
    local_environment = dict(os.environ)
    local_environment.update(
        {
            "GSHEETS_TOKEN_MODE": "local",
            "GOOGLE_CREDENTIALS_FILE": "/definitely/not/present.json",
            "GOOGLE_TOKEN_FILE": "/definitely/not/present.pickle",
            "GSHEETS_HEADLESS": "1",
        }
    )
    broker_environment = dict(local_environment)
    broker_environment["GSHEETS_TOKEN_MODE"] = "broker"
    broker_environment.pop("GSHEETS_BROKER_URL", None)
    broker_environment.pop("GSHEETS_BROKER_SESSION_TOKEN", None)

    local_result = _run("capabilities", "--json", env=local_environment)
    broker_result = _run("capabilities", "--json", env=broker_environment)

    assert local_result.returncode == broker_result.returncode == 0
    assert local_result.stderr == broker_result.stderr == ""
    local = json.loads(local_result.stdout)
    broker = json.loads(broker_result.stdout)
    assert (local["package_version"], local["contract_version"]) == ("1.0.0", "1.0")
    assert local["mode"] == "local"
    assert local["tool_count"] == 9
    assert broker["mode"] == "broker"
    assert broker["tool_count"] == 8
    assert "gsheets_search_spreadsheets" in {tool["name"] for tool in local["tools"]}
    assert "gsheets_search_spreadsheets" not in {
        tool["name"] for tool in broker["tools"]
    }
    assert all(
        tool["_meta"]["gsheets/contractVersion"] == "1.0" for tool in local["tools"]
    )


def test_invalid_mode_fails_closed_with_stable_stderr_diagnostic() -> None:
    environment = dict(os.environ)
    environment["GSHEETS_TOKEN_MODE"] = "legacy"

    result = _run("capabilities", "--json", env=environment)

    assert result.returncode == 2
    assert result.stdout == ""
    diagnostic = json.loads(result.stderr)
    assert diagnostic == {
        "status": "error",
        "error": {
            "action": "correct_configuration",
            "code": "invalid_configuration",
            "message": "GSHEETS_TOKEN_MODE must be either 'local' or 'broker'.",
        },
    }


def test_installed_package_imports_from_neutral_directory(tmp_path: Path) -> None:
    environment = dict(os.environ)
    environment.pop("PYTHONPATH", None)
    result = subprocess.run(
        [
            sys.executable,
            "-c",
            (
                "import pathlib, gsheets_mcp; "
                "print(gsheets_mcp.__version__); "
                "print(pathlib.Path(gsheets_mcp.__file__).parent.name)"
            ),
        ],
        cwd=tmp_path,
        env=environment,
        capture_output=True,
        text=True,
        timeout=15,
        check=False,
    )

    assert result.returncode == 0, result.stderr
    assert result.stdout.splitlines() == ["1.0.0", "gsheets_mcp"]
    assert result.stderr == ""


def test_removed_runtime_paths_do_not_exist() -> None:
    assert not (ROOT / "run_server.py").exists()
    assert not (ROOT / "src" / "__init__.py").exists()
    assert not (ROOT / "src" / "server.py").exists()
    assert not (ROOT / "src" / "sheets_client.py").exists()
