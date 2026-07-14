"""Command-line entry point for gsheets-mcp."""

from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Sequence

import anyio

from . import __version__
from .server import capabilities, run_stdio
from .sheets_client import SheetsClientError


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="gsheets-mcp",
        description="Strict Google Sheets MCP server.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=__version__,
    )
    subparsers = parser.add_subparsers(dest="command", metavar="COMMAND")

    subparsers.add_parser(
        "serve",
        help="Run the MCP server over stdio.",
        description=(
            "Run the MCP server over stdio. Standard output is reserved for MCP "
            "protocol frames. GSHEETS_TOKEN_MODE selects local or broker mode."
        ),
    )
    capabilities_parser = subparsers.add_parser(
        "capabilities",
        help="Show the mode-specific tool contract without loading credentials.",
    )
    capabilities_parser.add_argument(
        "--json",
        action="store_true",
        help="Emit one machine-readable JSON object.",
    )
    return parser


def _print_capabilities(*, as_json: bool) -> None:
    payload = capabilities()
    if as_json:
        print(json.dumps(payload, sort_keys=True, separators=(",", ":")))
        return
    print(
        f"gsheets-mcp {payload['package_version']} "
        f"(contract {payload['contract_version']}, mode {payload['mode']})"
    )
    print(f"tools: {payload['tool_count']}")
    for tool in payload["tools"]:
        print(f"  {tool['name']}: {tool.get('description', '')}")


def _print_cli_error(error: SheetsClientError) -> None:
    diagnostic = {
        "status": "error",
        "error": {
            "code": error.code,
            "message": str(error),
            "action": error.details.get("retry_action", "correct_configuration"),
        },
    }
    print(json.dumps(diagnostic, sort_keys=True), file=sys.stderr)


def main(argv: Sequence[str] | None = None) -> int:
    """Run the explicit command-oriented CLI."""
    parser = _parser()
    arguments = parser.parse_args(argv)
    if arguments.command is None:
        parser.print_help()
        return 0

    try:
        if arguments.command == "capabilities":
            _print_capabilities(as_json=arguments.json)
            return 0
        if arguments.command == "serve":
            anyio.run(run_stdio)
            return 0
    except SheetsClientError as exc:
        _print_cli_error(exc)
        return 2

    parser.error(f"unsupported command: {arguments.command}")
    return 2


if __name__ == "__main__":  # pragma: no cover - exercised by console entry point
    raise SystemExit(main())
