"""
parse.py — Markdown outline parser for bjxh-ppt.

Converts a Markdown outline into a list of PageSpec objects that drive
render.py.  See examples/audit_demo.md for the expected input format.
"""

import re
import sys

from config import get_company_info


# ---- regex patterns -------------------------------------------------------

H1_RE: re.Pattern[str] = re.compile(r"^#\s+(.+?)\s*$")
H2_RE: re.Pattern[str] = re.compile(r"^##\s+(.+?)\s*$")
H3_RE: re.Pattern[str] = re.compile(r"^###\s+(.+?)\s*$")
UL_RE: re.Pattern[str] = re.compile(r"^[-*]\s+(.+?)\s*$")
DIRECTIVE_RE: re.Pattern[str] = re.compile(r"^>\s*(\w+):\s*(.+?)\s*$")

# H3 default layout. 文字模板1 has too few placeholders — qa.py would flag
# content pages as THIN-PAGE.  无图分段-3项 provides structured sections while
# passing the visual-element check.
DEFAULT_CONTENT_LAYOUT: str = "无图分段-3项"


# ---- PageSpec -------------------------------------------------------------


class PageSpec:
    """One output page.

    A page can represent a cover, table-of-contents, chapter divider, content
    slide, or closing slide.  The ``kind`` field drives layout selection in
    render_page().
    """

    def __init__(
        self,
        kind: str,
        title: str,
        layout: str,
        bullets: list[str],
        note: str = "",
        toc_page_numbers: list[str] | None = None,
    ) -> None:
        self.kind = kind          # 'cover' | 'toc' | 'chapter' | 'content' | 'closing'
        self.title = title
        self.layout = layout
        self.bullets = bullets
        self.note = note
        self.toc_page_numbers = toc_page_numbers or []

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
    """
    pages: list[PageSpec] = []
    current: PageSpec | None = None
    h2_titles: list[str] = []
    cover_added: bool = False
    line_number: int = 0

    def flush() -> None:
        nonlocal current
        if current is not None:
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
                    bullets=[],
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
                    bullets=[],
                )
            else:
                current = PageSpec(
                    kind="chapter",
                    title=chapter_title,
                    layout="标题页-空白",
                    bullets=[],
                )
            continue

        m = H3_RE.match(line)
        if m:
            flush()
            current = PageSpec(
                kind="content",
                title=m.group(1).strip(),
                layout=DEFAULT_CONTENT_LAYOUT,
                bullets=[],
            )
            continue

        m = UL_RE.match(line)
        if m and current is not None:
            current.bullets.append(m.group(1).strip())
            continue

        m = DIRECTIVE_RE.match(line)
        if m and current is not None:
            key, val = m.group(1).lower(), m.group(2).strip()
            if key == "layout":
                current.layout = val
            elif key == "note":
                current.note = val
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
        company: dict[str, str] = get_company_info()
        pages.append(
            PageSpec(
                kind="closing",
                title=company["company_name"],
                layout="自定义版式",
                bullets=[],
            )
        )

    return pages
