"""Tests for preview_html.py — overflow estimation and HTML generation."""

from __future__ import annotations

import io
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
from preview_html import estimate_overflow  # noqa: E402
from shared import emu_to_inches  # noqa: E402


def emu_to_in(v: int | None) -> float:
    """Alias for shared.emu_to_inches (preview_html used this name)."""
    return emu_to_inches(v)


# -- emu_to_in -------------------------------------------------------------


def test_emu_to_in_typical_value() -> None:
    assert emu_to_in(914400) == pytest.approx(1.0, abs=0.01)


def test_emu_to_in_zero() -> None:
    assert emu_to_in(0) == 0.0


def test_emu_to_in_none() -> None:
    assert emu_to_in(None) == 0.0


# -- estimate_overflow -----------------------------------------------------


def test_overflow_short_text_fits() -> None:
    chars_per_line, lines = estimate_overflow("短标题", box_w_in=3.87, font_pt=20)
    # Short CJK text should fit in 1-2 lines on 3.87" at 20pt
    assert lines <= 2


def test_overflow_long_text_overflows() -> None:
    long_text = "这是一个非常长的审计行业AI应用分析报告标题包含大量内容"
    chars_per_line, lines = estimate_overflow(long_text, box_w_in=3.87, font_pt=20)
    # Long text needs more lines
    assert lines >= 2


def test_overflow_ascii_fits_more() -> None:
    ascii_text = "Short ASCII title"
    chars_ascii, lines_ascii = estimate_overflow(ascii_text, box_w_in=3.87, font_pt=20)
    cjk_text = "短中文标题测试"
    chars_cjk, lines_cjk = estimate_overflow(cjk_text, box_w_in=3.87, font_pt=20)
    # ASCII chars are narrower, more fit per line
    assert chars_ascii >= chars_cjk


def test_overflow_empty_text() -> None:
    chars_per_line, lines = estimate_overflow("", box_w_in=3.87, font_pt=20)
    assert lines == 1
