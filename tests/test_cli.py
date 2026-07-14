"""Tests for cli.py — command-line argument parsing and error paths."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

CLI = str(Path(__file__).parent.parent / "scripts" / "cli.py")


def _run(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, CLI, *args],
        capture_output=True, text=True, timeout=30,
    )


def test_cli_missing_outline() -> None:
    result = _run("nonexistent.md", "out.pptx")
    assert result.returncode != 0


def test_cli_missing_template() -> None:
    result = _run(
        str(Path(__file__).parent.parent / "examples" / "audit_demo.md"),
        "/tmp/test_out.pptx",
        "/nonexistent/template.pptx",
    )
    assert result.returncode != 0


def test_cli_help_shows_usage() -> None:
    result = _run("--help")
    assert result.returncode == 0
    assert "usage" in result.stdout.lower() or "usage" in result.stderr.lower()


def test_cli_no_args_exits_nonzero() -> None:
    result = _run()
    assert result.returncode != 0


def test_cli_no_chapter_covers_flag_accepted() -> None:
    # Just verify the flag is recognized; actual E2E needs template
    result = _run(
        "--help",
    )
    assert "--no-chapter-covers" in result.stdout
