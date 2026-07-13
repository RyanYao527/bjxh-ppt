"""
config.py — template path resolution and company-specific data.

Template lookup order:
    1. BJXH_TEMPLATE environment variable
    2. config.json in the same directory as this script
    3. Hardcoded fallback (Windows-only legacy path)

Company contact info is read from config.json; the hardcoded defaults
below are placeholders that should be replaced per deployment.
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any


def _find_config() -> dict[str, Any]:
    """Load config.json from the same directory as this script, or return {}."""
    config_path = Path(__file__).parent / "config.json"
    if config_path.exists():
        try:
            return json.loads(config_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            print(f"Warning: failed to parse {config_path}", file=sys.stderr)
    return {}


def resolve_template_path() -> str:
    """Resolve the master template .pptx path.

    Returns the first of:
        - BJXH_TEMPLATE env var
        - config.json → template_path
        - legacy fallback (may not exist on non-Windows or different machines)
    """
    env_path: str | None = os.environ.get("BJXH_TEMPLATE")
    if env_path and Path(env_path).exists():
        return env_path

    cfg: dict[str, Any] = _find_config()
    cfg_path: str | None = cfg.get("template_path")
    if cfg_path and Path(cfg_path).exists():
        return cfg_path

    # Legacy fallback — unlikely to exist outside the original author's machine.
    legacy: str = (
        r"C:\工作\04-总结与报告\2026年工作\2026合伙人大会\北京兴华模板.pptx"
    )
    if Path(legacy).exists():
        return legacy

    print(
        "ERROR: Cannot find the BJXH master template.\n"
        "  Set the BJXH_TEMPLATE environment variable, or\n"
        f"  create {Path(__file__).parent / 'config.json'} with a 'template_path' key, or\n"
        "  pass the template path as the third CLI argument to from_outline.py.",
        file=sys.stderr,
    )
    sys.exit(1)


def get_company_info() -> dict[str, str]:
    """Return company contact info from config.json, or sensible defaults.

    Keys: phone, fax, address, company_name.
    """
    cfg: dict[str, Any] = _find_config()
    company: dict[str, Any] = cfg.get("company", {})
    return {
        "company_name": company.get("company_name", "北京兴华集团"),
        "phone": company.get("phone", ""),
        "fax": company.get("fax", ""),
        "address": company.get("address", ""),
    }
