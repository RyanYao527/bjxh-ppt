"""
from_outline.py — Convert a Markdown outline to a Beijing Xinghua .pptx.

Outline syntax (deliberately minimal — see examples/audit_demo.md for full sample):

    # 主标题 (H1)        — 整份演示的标题，仅第一行使用 → 主题-封面
    ## 章节名 (H2)        — 章节封面 → 标题页-空白
    ### 页面标题 (H3)     — 一张内容页 → 默认版式（可指定）
    - 要点                 — 页内要点
    - 要点 2
    > layout: 文字模板1   — 显式指定当前 H3 章节的版式（可选，覆盖默认）
    > note: 备注文字       — 添加演讲者备注（可选）

Usage:
    python from_outline.py outline.md out.pptx [template.pptx]
        template.pptx 默认为 BJXH_TEMPLATE 环境变量或 config.json 中的路径

版式自动选择规则（见 SKILL.md §6.1）:
    H1 (1 个)              → 主题-封面
    第一个 H2 之前插入     → 主题-目录页（可选：--no-toc 关闭）
    每个 H2                → 标题页-空白（章节封面，可用 --no-chapter-covers 改为内容页）
    每个 H3                → 无图分段-3项（默认）/ > layout: 覆盖
    末尾                    → 自定义版式（可选：--no-closing 关闭）

Implementation spread across three modules for maintainability:
    parse.py   — markdown parsing  (PageSpec, parse_outline)
    render.py  — slide generation  (PLACEHOLDER_MAP, render_page)
    cli.py     — CLI entry point   (main)
"""

from __future__ import annotations

# Re-exports for backwards compatibility — scripts that import from
# from_outline continue to work unchanged.
from cli import main
from parse import DEFAULT_CONTENT_LAYOUT, PageSpec, parse_outline  # noqa: F401
from render import (  # noqa: F401
    PLACEHOLDER_MAP,
    find_layout,
    find_placeholder,
    render_page,
    set_placeholder_text,
)

if __name__ == "__main__":
    import sys
    sys.exit(main())
