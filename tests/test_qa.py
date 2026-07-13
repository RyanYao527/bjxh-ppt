"""Tests for qa.py check functions.

These tests exercise individual check functions with mock data, as well as
the regex patterns and constant values used by the QA script.
"""
import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
from qa import (  # noqa: E402
    BJH_PATTERN,
    CANVAS_H,
    CANVAS_W,
    EDIT_NOTE_GREY,
    FORBIDDEN_FONTS,
)


# -- constants -------------------------------------------------------------

def test_canvas_dimensions_are_16_9() -> None:
    """13.333" × 7.5" = 16:9 widescreen."""
    assert CANVAS_W == 12192000
    assert CANVAS_H == 6858000
    ratio = (CANVAS_W / 914400) / (CANVAS_H / 914400)
    assert abs(ratio - (16 / 9)) < 0.01


def test_edit_note_grey_is_445469() -> None:
    assert EDIT_NOTE_GREY == (0x44, 0x54, 0x69)


def test_forbidden_fonts_contains_common_violations() -> None:
    assert "宋体" in FORBIDDEN_FONTS
    assert "Times New Roman" in FORBIDDEN_FONTS
    assert "Calibri" in FORBIDDEN_FONTS
    # Allowed fonts should NOT be in the forbidden set
    assert "Microsoft YaHei" not in FORBIDDEN_FONTS
    assert "Arial" not in FORBIDDEN_FONTS


# -- BJH pattern -----------------------------------------------------------

@pytest.mark.parametrize("text,should_match", [
    ("B J H", True),            # classic B/J/H residue
    ("B X H", True),
    ("BJ H", True),
    ("关键词 B", True),          # trailing isolated B
    ("X 点击编辑标题", True),    # leading isolated X
    ("B JH", True),
    ("J ", True),
    (" B ", True),
    ("BJH", True),              # three chars joined — new regex catches this
    ("B.J.H", True),            # dot-separated triplet
    ("B.J", True),              # dot-separated pair
    ("B H", True),              # space-separated pair
    # Should NOT match:
    ("BEIJING XINGHUA", False),  # full words
    ("这是一个B示例", False),     # B embedded in CJK text
    ("正文B正文", False),         # B embedded
    ("ABCDEFG", False),          # normal English
    ("北京兴华", False),         # normal Chinese
    ("Key: B", True),            # colon then space then B
])
def test_bjh_pattern(text: str, should_match: bool) -> None:
    matched = BJH_PATTERN.search(text) is not None
    assert matched == should_match, (
        f"Expected {'match' if should_match else 'no match'} for {text!r}"
    )


# -- regex patterns --------------------------------------------------------

def test_h1_regex() -> None:
    """Verify H1 regex from from_outline.py matches expected lines."""
    H1_RE = re.compile(r"^#\s+(.+?)\s*$")
    assert H1_RE.match("# 年度工作汇报")
    assert H1_RE.match("# Cover")
    assert not H1_RE.match("## Not H1")
    assert not H1_RE.match("#")


def test_h2_regex() -> None:
    H2_RE = re.compile(r"^##\s+(.+?)\s*$")
    assert H2_RE.match("## 第一章：回顾")
    assert not H2_RE.match("# Not H2")
    assert not H2_RE.match("### Not H2")


def test_h3_regex() -> None:
    H3_RE = re.compile(r"^###\s+(.+?)\s*$")
    assert H3_RE.match("### 页面标题")
    assert not H3_RE.match("## Not H3")


def test_directive_regex() -> None:
    DIRECTIVE_RE = re.compile(r"^>\s*(\w+):\s*(.+?)\s*$")
    m = DIRECTIVE_RE.match("> layout: 无图分段-3项")
    assert m is not None
    assert m.group(1) == "layout"
    assert m.group(2) == "无图分段-3项"

    m = DIRECTIVE_RE.match("> note: 这是备注")
    assert m is not None
    assert m.group(1) == "note"

    # Malformed directive — missing colon
    assert DIRECTIVE_RE.match("> layout 无图分段-3项") is None


def test_bullet_regex() -> None:
    UL_RE = re.compile(r"^[-*]\s+(.+?)\s*$")
    assert UL_RE.match("- 要点一")
    assert UL_RE.match("* 要点二")
    assert not UL_RE.match("-")
    assert not UL_RE.match("-- not a bullet")
