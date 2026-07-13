"""Tests for from_outline.parse_outline()."""
import sys
from pathlib import Path

import pytest

# Allow running tests from repo root without installing the package
sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
from from_outline import PageSpec, parse_outline  # noqa: E402


# -- basic structure -------------------------------------------------------

def test_basic_cover_and_chapters() -> None:
    text = """\
# 年度工作汇报

## 第一章：回顾

### 页面1
- 要点A
- 要点B

### 页面2
- 要点C

## 第二章：展望

### 页面3
- 要点D
"""
    pages = parse_outline(text)
    kinds = [p.kind for p in pages]
    assert kinds == ["cover", "toc", "chapter", "content", "content",
                     "chapter", "content", "closing"]


def test_cover_title() -> None:
    text = "# 年度工作汇报"
    pages = parse_outline(text)
    cover = pages[0]
    assert cover.kind == "cover"
    assert cover.title == "年度工作汇报"
    assert cover.layout == "主题-封面"


def test_chapter_layout() -> None:
    text = """\
# Cover

## Chapter One

### Page 1
- bullet
"""
    pages = parse_outline(text)
    chapter = [p for p in pages if p.kind == "chapter"][0]
    assert chapter.layout == "标题页-空白"


def test_content_default_layout() -> None:
    text = """\
# Cover

### Page 1
- bullet
"""
    pages = parse_outline(text)
    content = [p for p in pages if p.kind == "content"][0]
    assert content.layout == "无图分段-3项"


# -- directives ------------------------------------------------------------

def test_layout_override() -> None:
    text = """\
# Cover

### Page 1
- bullet
> layout: 文字模板1
"""
    pages = parse_outline(text)
    content = [p for p in pages if p.kind == "content"][0]
    assert content.layout == "文字模板1"


def test_note_directive() -> None:
    text = """\
# Cover

### Page 1
- bullet
> note: speaker notes here
> layout: 图表-1
"""
    pages = parse_outline(text)
    content = [p for p in pages if p.kind == "content"][0]
    assert content.note == "speaker notes here"
    assert content.layout == "图表-1"


def test_bullets_after_directive() -> None:
    text = """\
# Cover

### Page 1
> layout: 无图分段-4项
- bullet 1
- bullet 2
- bullet 3
- bullet 4
"""
    pages = parse_outline(text)
    content = [p for p in pages if p.kind == "content"][0]
    assert len(content.bullets) == 4


# -- --no-chapter-covers ---------------------------------------------------

def test_no_chapter_covers() -> None:
    text = """\
# Cover

## Direct Content
- bullet A
- bullet B

## Another
- bullet C
"""
    pages = parse_outline(text, no_chapter_covers=True)
    kinds = [p.kind for p in pages]
    # No chapter-kind pages; H2 becomes content
    assert "chapter" not in kinds
    content_pages = [p for p in pages if p.kind == "content"]
    assert len(content_pages) == 2
    assert content_pages[0].title == "Direct Content"
    assert content_pages[0].bullets == ["bullet A", "bullet B"]


# -- options ---------------------------------------------------------------

def test_no_toc() -> None:
    text = """\
# Cover

## Chapter

### Page
- bullet
"""
    pages = parse_outline(text, add_toc=False)
    kinds = [p.kind for p in pages]
    assert "toc" not in kinds


def test_no_closing() -> None:
    text = "# Cover"
    pages = parse_outline(text, add_closing=False)
    kinds = [p.kind for p in pages]
    assert "closing" not in kinds


# -- edge cases ------------------------------------------------------------

def test_empty_input() -> None:
    pages = parse_outline("")
    assert pages == []


def test_whitespace_only() -> None:
    pages = parse_outline("   \n\n  \n")
    assert pages == []


def test_only_bullets_no_heading() -> None:
    text = """\
- stray bullet
- another bullet
"""
    pages = parse_outline(text)
    assert pages == []


def test_multiple_h1_uses_first_only() -> None:
    text = """\
# First Cover
# Second Cover
"""
    pages = parse_outline(text)
    covers = [p for p in pages if p.kind == "cover"]
    assert len(covers) == 1
    assert covers[0].title == "First Cover"


def test_directive_before_any_page_is_ignored() -> None:
    text = """\
> layout: 文字模板1
# Cover
"""
    pages = parse_outline(text)
    assert pages[0].kind == "cover"


def test_h2_without_h1_still_works() -> None:
    text = """\
## Chapter One

### Page 1
- bullet
"""
    pages = parse_outline(text)
    kinds = [p.kind for p in pages]
    assert "cover" not in kinds  # no H1 → no cover
    assert "chapter" in kinds


# -- closing_title ----------------------------------------------------------

def test_closing_title_default() -> None:
    text = "# Cover"
    pages = parse_outline(text)
    closing = [p for p in pages if p.kind == "closing"][0]
    assert closing.title == "北京兴华集团"


def test_closing_title_custom() -> None:
    text = "# Cover"
    pages = parse_outline(text, closing_title="Custom Corp")
    closing = [p for p in pages if p.kind == "closing"][0]
    assert closing.title == "Custom Corp"


# -- PageSpec --------------------------------------------------------------

def test_pagespec_repr() -> None:
    ps = PageSpec(kind="content", title="Test", layout="文字模板1",
                  bullets=["a", "b"], note="n")
    r = repr(ps)
    assert "content" in r
    assert "Test" in r
    assert "文字模板1" in r
    assert "n=2" in r
