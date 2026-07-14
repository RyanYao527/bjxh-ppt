"""
parse.py — Markdown outline parser for bjxh-ppt.

Converts a Markdown outline into a list of PageSpec objects that drive
render.py.  See examples/audit_demo.md for the expected input format.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field, replace

from layout_utils import suggest_layout


# ---- regex patterns -------------------------------------------------------

H1_RE: re.Pattern[str] = re.compile(r"^#\s+(.+?)\s*$")
H2_RE: re.Pattern[str] = re.compile(r"^##\s+(.+?)\s*$")
H3_RE: re.Pattern[str] = re.compile(r"^###\s+(.+?)\s*$")
UL_RE: re.Pattern[str] = re.compile(r"^[-*]\s+(.+?)\s*$")
DIRECTIVE_RE: re.Pattern[str] = re.compile(r"^>\s*(\w+):\s*(.+?)\s*$")

# H3 default layout. 文字模板1 has too few placeholders — qa.py would flag
# content pages as THIN-PAGE.  无图分段-3项 provides structured sections while
# passing the visual-element check.
DEFAULT_CONTENT_LAYOUT: str = "标题页-空白"


# ---- PageSpec -------------------------------------------------------------


@dataclass(frozen=True)
class PageSpec:
    """One output page.

    A page can represent a cover, table-of-contents, chapter divider, content
    slide, or closing slide.  The ``kind`` field drives layout selection in
    render_page().
    """

    kind: str   # 'cover' | 'toc' | 'chapter' | 'content' | 'closing'
    title: str
    layout: str
    bullets: list[str] = field(default_factory=list)
    note: str = ""
    toc_page_numbers: list[str] = field(default_factory=list)

    def __repr__(self) -> str:
        return (
            f"PageSpec({self.kind}, {self.title!r}, "
            f"layout={self.layout!r}, n={len(self.bullets)})"
        )


# ---- main parser ----------------------------------------------------------


def parse_outline(
    text: str,
    *,
    add_toc: bool = True,
    add_closing: bool = True,
    no_chapter_covers: bool = False,
    closing_title: str = "北京兴华集团",
) -> list[PageSpec]:
    """Parse a Markdown outline into a list of PageSpec objects.

    Parameters
    ----------
    text :
        Raw markdown text following the bjxh-ppt outline convention
        (``#`` cover, ``##`` chapters, ``###`` pages, ``-`` bullets,
        ``> layout:`` / ``> note:`` directives).
    add_toc :
        Insert a table-of-contents page after the cover.
    add_closing :
        Append a closing (封底) page at the end.
    no_chapter_covers :
        Treat ``##`` headings as content pages instead of chapter-dividers.
        Useful for compact decks (cover + TOC + N content pages + closing).
    closing_title :
        Title text for the closing slide.  Defaults to the BJXH company name.
        Set this from config or CLI when using a different organisation.
    """
    pages: list[PageSpec] = []
    current: PageSpec | None = None
    h2_titles: list[str] = []
    cover_added: bool = False
    line_number: int = 0
    # Layout rotation tracker — last 3 used layouts are excluded from suggestion
    recent_layouts: list[str] = []

    def flush() -> None:
        nonlocal current, recent_layouts
        if current is not None:
            # Auto-select layout if still on default (no explicit > layout: directive)
            if current.kind == "content" and current.layout == DEFAULT_CONTENT_LAYOUT:
                used_set = frozenset(recent_layouts[-3:])
                current = replace(
                    current,
                    layout=suggest_layout(current.bullets, used_layouts=used_set),
                )
            # Track for rotation
            recent_layouts.append(current.layout)
            if len(recent_layouts) > 5:
                recent_layouts = recent_layouts[-5:]
            pages.append(current)
            current = None

    for raw_line in text.splitlines():
        line_number += 1
        line = raw_line.rstrip()
        if not line.strip():
            continue

        m = H1_RE.match(line)
        if m:
            if not cover_added:
                flush()
                current = PageSpec(
                    kind="cover",
                    title=m.group(1).strip(),
                    layout="主题-封面",
                )
                cover_added = True
            else:
                print(
                    f"Warning: line {line_number}: duplicate H1 "
                    f"'# {m.group(1).strip()[:40]}' ignored "
                    f"(only the first H1 is used as the cover title).",
                    file=sys.stderr,
                )
            continue

        m = H2_RE.match(line)
        if m:
            flush()
            chapter_title = m.group(1).strip()
            h2_titles.append(chapter_title)
            if no_chapter_covers:
                current = PageSpec(
                    kind="content",
                    title=chapter_title,
                    layout=DEFAULT_CONTENT_LAYOUT,
                )
            else:
                current = PageSpec(
                    kind="chapter",
                    title=chapter_title,
                    layout="标题页-空白",
                )
            continue

        m = H3_RE.match(line)
        if m:
            flush()
            current = PageSpec(
                kind="content",
                title=m.group(1).strip(),
                layout=DEFAULT_CONTENT_LAYOUT,
            )
            continue

        m = UL_RE.match(line)
        if m and current is not None:
            current = replace(
                current, bullets=current.bullets + [m.group(1).strip()]
            )
            continue

        m = DIRECTIVE_RE.match(line)
        if m and current is not None:
            key, val = m.group(1).lower(), m.group(2).strip()
            if key == "layout":
                current = replace(current, layout=val)
            elif key == "note":
                current = replace(current, note=val)
            else:
                print(
                    f"Warning: line {line_number}: unknown directive "
                    f"'> {line[1:].strip()[:60]}' ignored "
                    f"(supported: layout, note).",
                    file=sys.stderr,
                )
            continue

        # Line doesn't match any known pattern — warn if we're inside a page.
        if current is not None:
            print(
                f"Warning: line {line_number}: unrecognized line ignored: "
                f"'{line[:80]}'",
                file=sys.stderr,
            )

    flush()

    # ---- post-processing: insert TOC ----------------------------------
    if add_toc and h2_titles:
        toc_page_numbers: list[str] = []
        for title in h2_titles:
            idx = next(
                (j for j, p in enumerate(pages) if p.title == title),
                len(pages),
            )
            toc_page_numbers.append(str(idx + 2))
        toc = PageSpec(
            kind="toc",
            title="目录",
            layout="主题-目录页",
            bullets=list(h2_titles),
            toc_page_numbers=toc_page_numbers,
        )
        insert_at: int = 1 if pages and pages[0].kind == "cover" else 0
        pages.insert(insert_at, toc)

    # ---- post-processing: append closing -------------------------------
    if add_closing and pages and pages[-1].kind != "closing":
        pages.append(
            PageSpec(
                kind="closing",
                title=closing_title,
                layout="自定义版式",
            )
        )

    return pages
