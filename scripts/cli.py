"""
cli.py — Command-line entry point for bjxh-ppt.

Usage:
    python cli.py outline.md out.pptx [template.pptx]
"""

import argparse
import shutil
import sys
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn

from config import resolve_template_path
from parse import parse_outline
from render import render_page


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert a Markdown outline to a Beijing Xinghua .pptx."
    )
    parser.add_argument("outline", help="path to outline.md")
    parser.add_argument("output", help="output .pptx path")
    parser.add_argument(
        "template",
        nargs="?",
        default=None,
        help="template .pptx path "
        "(optional: uses BJXH_TEMPLATE env var or config.json if omitted)",
    )
    parser.add_argument("--no-toc", dest="add_toc", action="store_false")
    parser.add_argument("--no-closing", dest="add_closing", action="store_false")
    parser.add_argument(
        "--no-chapter-covers",
        dest="no_chapter_covers",
        action="store_true",
        help="Treat H2 headings as content pages instead of chapter title pages "
        "(useful for compact decks with cover+TOC+content+closing).",
    )
    args = parser.parse_args()

    outline_path = Path(args.outline)
    if not outline_path.exists():
        print(f"ERROR: outline not found: {outline_path}", file=sys.stderr)
        return 1

    if args.template:
        template_path = Path(args.template)
    else:
        template_path = Path(resolve_template_path())
    if not template_path.exists():
        print(
            f"ERROR: template not found: {template_path}", file=sys.stderr
        )
        print(
            "Set BJXH_TEMPLATE env var or create scripts/config.json "
            "with 'template_path'.",
            file=sys.stderr,
        )
        return 1

    text = outline_path.read_text(encoding="utf-8")
    pages = parse_outline(
        text,
        add_toc=args.add_toc,
        add_closing=args.add_closing,
        no_chapter_covers=args.no_chapter_covers,
    )

    # Copy template → temp, then strip existing slides.  python-pptx has no
    # "create from scratch with theme" mode, so we copy the master, drop its
    # slides, and add ours.
    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        suffix=".pptx", delete=False, dir=tempfile.gettempdir()
    ) as tmp:
        tmp_path = Path(tmp.name)
    shutil.copy2(str(template_path), str(tmp_path))

    prs = Presentation(str(tmp_path))

    xml_slides = prs.slides._sldIdLst
    for sld_id in list(xml_slides):
        rId = sld_id.get(qn("r:id"))
        prs.part.drop_rel(rId)
        xml_slides.remove(sld_id)

    for spec in pages:
        render_page(prs, spec, template_path)

    prs.save(str(out_path))
    tmp_path.unlink(missing_ok=True)

    print(f"OK: wrote {out_path}")
    print(f"  pages: {len(pages)}")
    for i, spec in enumerate(pages, 1):
        print(f"  {i}. [{spec.kind:8s}] {spec.title}  ({spec.layout})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
