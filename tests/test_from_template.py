"""Tests for from_template.py — parse_layout_list and error paths."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
from from_template import parse_layout_list, find_layout  # noqa: E402
from pptx import Presentation  # noqa: E402


# -- parse_layout_list -----------------------------------------------------


def test_parse_layout_list_single() -> None:
    assert parse_layout_list("主题-封面") == ["主题-封面"]


def test_parse_layout_list_multiple() -> None:
    result = parse_layout_list("主题-封面,主题-目录页,文字模板1")
    assert result == ["主题-封面", "主题-目录页", "文字模板1"]


def test_parse_layout_list_strips_whitespace() -> None:
    result = parse_layout_list(" 主题-封面 , 标题页-空白 , 文字模板1 ")
    assert result == ["主题-封面", "标题页-空白", "文字模板1"]


def test_parse_layout_list_empty() -> None:
    assert parse_layout_list("") == []


def test_parse_layout_list_skips_empty_items() -> None:
    assert parse_layout_list("主题-封面,,标题页-空白") == ["主题-封面", "标题页-空白"]


# -- find_layout -----------------------------------------------------------


def test_find_layout_returns_none_for_unknown() -> None:
    prs = Presentation()
    assert find_layout(prs, "不存在的版式名称") is None


def test_find_layout_with_default_template() -> None:
    prs = Presentation()
    found = find_layout(prs, prs.slide_masters[0].slide_layouts[0].name)
    assert found is not None
