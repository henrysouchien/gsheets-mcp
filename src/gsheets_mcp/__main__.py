"""Allow ``python -m gsheets_mcp`` to use the canonical CLI."""

from .cli import main


raise SystemExit(main())
