"""Load `OPENAI_API_KEY` from the project-root `.env` file (python-dotenv).

Shell exports still work: variables already set in the environment are not overridden
unless you change `override` below.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]


def load_project_env() -> None:
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    load_dotenv(_ROOT / ".env", override=False)


def get_openai_api_key(*, required: bool) -> str | None:
    """Return the API key after loading `.env`, or exit/return depending on *required*."""
    load_project_env()
    key = (os.environ.get("OPENAI_API_KEY") or "").strip()
    if key:
        return key
    if not required:
        return None
    env_file = _ROOT / ".env"
    print(
        "Missing OPENAI_API_KEY.\n"
        f"  Create {env_file} with OPENAI_API_KEY=... (copy from .env.example in the repo root).\n"
        "  Do not commit real API keys; .env is gitignored.",
        file=sys.stderr,
    )
    raise SystemExit(1)
