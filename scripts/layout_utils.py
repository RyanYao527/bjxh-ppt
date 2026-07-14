"""
layout_utils.py — Layout helpers for bjxh-ppt render pipeline.

Provides dynamic font sizing (prevents text overflow) and placeholder
cleanup (removes "Click here to add text" residue).
"""

from __future__ import annotations

import sys as _sys
from typing import Any

from shared import emu_to_inches


def count_cjk(text: str) -> int:
    """Return the number of CJK / wide characters in *text*."""
    return sum(1 for c in text if ord(c) > 0x2E80)


def estimated_char_width_pt(text: str) -> float:
    """Rough per-character width in points, accounting for CJK vs ASCII.

    CJK characters are roughly 1.0× font-size wide; ASCII ~0.55×.
    Returns a weighted average for the given text.
    """
    if not text:
        return 0.55
    cjk = count_cjk(text)
    ascii_count = len(text) - cjk
    return (1.0 * cjk + 0.55 * ascii_count) / len(text)


def calc_safe_font_size(
    ph: Any,
    text: str,
    *,
    max_pt: int = 20,
    min_pt: int = 8,
    line_spacing: float = 1.3,
) -> int:
    """Find the largest font size (pt) that fits *text* in *ph*.

    Steps down from *max_pt* until the text fits both width and height,
    bottoming out at *min_pt*.
    """
    if not text:
        return max_pt

    box_w = emu_to_inches(ph.width) if ph.width else 13.3
    box_h = emu_to_inches(ph.height) if ph.height else 7.5
    char_w_ratio = estimated_char_width_pt(text)

    for size in range(max_pt, min_pt - 1, -1):
        char_width_in = (size / 72) * char_w_ratio * 1.10  # +10% safety margin
        chars_per_line = max(1, int(box_w / char_width_in))
        lines_needed = -(-len(text) // chars_per_line)
        line_height_in = (size / 72) * line_spacing * 1.10
        lines_fit = max(1, int(box_h / line_height_in))
        if lines_needed <= lines_fit:
            return size
    return min_pt


# Content-aware layout categories (all names must match PLACEHOLDER_MAP).
# Only layouts in PLACEHOLDER_MAP are eligible for auto-selection.
_STRUCTURED = ["无图分段-3项", "无图分段-4项", "无图分段-5项"]
_TEXT = ["标题页-空白", "2_标题页-空白"]
_CHART = ["图表-1", "图表-2"]
_IMAGE_TEXT = ["有图分段式-16", "有图分段式-8"]

# Keywords that hint at content type
_DATA_KEYWORDS = [
    "%", "增长", "下降", "占比", "同比", "环比", "数据", "比例",
    "亿", "万", "倍", "个百分点", "统计", "趋势",
    # Audit-specific data indicators
    "累计投入", "约", "万元", "亿元", "试点", "覆盖率",
    "处理", "份", "个", "识别", "节约", "提升",
]


def _validate_layout(name: str) -> str:
    """Ensure *name* is actually in PLACEHOLDER_MAP, falling back safely."""
    from render import PLACEHOLDER_MAP  # avoid circular import at module level
    if name in PLACEHOLDER_MAP:
        return name
    fallback = "标题页-空白"
    _sys.stderr.write(
        f"Warning: layout '{name}' not in PLACEHOLDER_MAP, "
        f"falling back to '{fallback}'.\n"
    )
    return fallback


def suggest_layout(bullets: list[str], *, used_layouts: frozenset[str] | None = None) -> str:
    """Suggest an appropriate layout for a content page based on its bullets.

    Heuristics (in priority order):
    1. Data-heavy bullets → chart layout
    2. Process/step bullets → logic layout
    3. Short bullets (≤3 items) → structured card layout
    4. Everything else → rotated through structured variants

    *used_layouts* is a set of layouts already used nearby; the function
    avoids repeating the same layout.
    """
    used = used_layouts or frozenset()
    text = " ".join(bullets)

    # Data-heavy bullets → try chart layout first
    if any(kw in text for kw in _DATA_KEYWORDS) and len(bullets) <= 6:
        for c in _CHART:
            if c not in used:
                return _validate_layout(c)

    # Assemble all available variants for rotation
    all_candidates = list(_STRUCTURED)
    # Insert text and image-text layouts for variety
    if len(bullets) <= 3:
        all_candidates = _TEXT + _IMAGE_TEXT + all_candidates
    else:
        all_candidates = all_candidates + _IMAGE_TEXT

    # Exclude recently-used layouts
    candidates = [lay for lay in all_candidates if lay not in used]
    if not candidates:
        candidates = all_candidates

    # Prefer match on bullet count for structured layouts
    n = len(bullets)
    structured_preferred = [lay for lay in candidates if lay in _STRUCTURED]
    text_preferred = [lay for lay in candidates if lay in _TEXT]

    if n <= 3:
        # Short pages: alternate between text and structured
        if text_preferred and sum(1 for lay in used if lay in _STRUCTURED) >= 2:
            return text_preferred[0]
        ordered = [lay for lay in candidates if "3项" in lay or lay in _TEXT]
    elif n == 4:
        ordered = [lay for lay in candidates if "4项" in lay]
    elif n >= 5:
        ordered = [lay for lay in candidates if "5项" in lay]
    else:
        ordered = []

    fallback = ordered or structured_preferred or text_preferred or candidates
    return _validate_layout(fallback[0])


def truncate_title(title: str, max_chars: int = 14) -> tuple[str, str]:
    """Truncate a long title for the narrow 3.87\"×0.54\" title placeholder.

    Returns (display_title, overflow_note).  *overflow_note* is the full
    original title (or empty string) so it can be saved as a speaker note
    or subtitle.
    """
    if len(title) <= max_chars:
        return title, ""
    # Try to break at a natural separator
    for sep in ["：", ":", "——", "——", "，", ","]:
        idx = title.find(sep)
        if 4 <= idx <= max_chars:
            return title[:idx], title
    # Hard truncation
    return title[:max_chars], title


def extract_segment_title(bullet: str, max_chars: int = 8) -> str:
    """Extract a short title from a bullet string for segment card headers.

    Priority: separator-split prefix → first N CJK chars → raw truncation.
    Avoids producing garbled fragments like '第一阶段(0' or '成本结构变'.
    """
    # 1) Try splitting on colon / Chinese colon — use the prefix
    for sep in ["：", ":"]:
        if sep in bullet:
            prefix = bullet.split(sep, 1)[0].strip()
            if 2 <= len(prefix) <= max_chars + 2:
                return prefix
            # If prefix is too long, take first N chars of it
            if len(prefix) > max_chars:
                return prefix[:max_chars]

    # 2) Try splitting on comma / Chinese comma / period — prefix
    for sep in ["，", ",", "。", "."]:
        if sep in bullet:
            prefix = bullet.split(sep, 1)[0].strip()
            if 2 <= len(prefix) <= max_chars:
                return prefix

    # 3) Try parentheses → drop parenthetical content for cleaner title
    import re
    cleaned = re.sub(r"[（(][^)）]*[)）]", "", bullet)
    if 2 <= len(cleaned) <= max_chars:
        return cleaned.strip()
    if len(cleaned) > max_chars:
        return cleaned[:max_chars].strip()

    # 4) Fallback: first N chars
    return bullet[:max_chars]


def clear_unused_placeholders(slide: Any, used_idx_set: set[int]) -> None:
    """Clear text from any placeholder whose idx is NOT in *used_idx_set*.

    This prevents "单击此处添加文本" (Click to add text) template residue
    from appearing in the final .pptx when PowerPoint renders unfilled
    placeholders.
    """
    for ph in slide.placeholders:
        idx: int = ph.placeholder_format.idx
        if idx not in used_idx_set and ph.has_text_frame:
            # Write a single space so PowerPoint doesn't show placeholder prompt
            tf = ph.text_frame
            tf.text = " "
