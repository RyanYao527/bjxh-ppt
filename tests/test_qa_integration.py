"""Integration tests for qa.py using committed example .pptx files.

These tests exercise the QA check functions against known-good example
files, verifying that the QA pipeline doesn't regress.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
from pptx import Presentation  # noqa: E402
from qa import (  # noqa: E402
    check_bjh_residue,
    check_canvas,
    check_edit_grey_color,
    check_fonts_and_size,
    check_layout_variety,
)


@pytest.fixture(scope="module")
def audit_demo() -> Presentation:
    path = Path(__file__).parent.parent / "examples" / "audit_demo.pptx"
    if not path.exists():
        pytest.skip("audit_demo.pptx not found")
    return Presentation(str(path))


@pytest.fixture(scope="module")
def xxb_report() -> Presentation:
    path = Path(__file__).parent.parent / "examples" / "xxb_2025_2026_report.pptx"
    if not path.exists():
        pytest.skip("xxb_2025_2026_report.pptx not found")
    return Presentation(str(path))


# -- canvas check ----------------------------------------------------------

@pytest.mark.parametrize("prs_fixture", ["audit_demo", "xxb_report"])
def test_canvas_is_16_9(request: pytest.FixtureRequest, prs_fixture: str) -> None:
    prs = request.getfixturevalue(prs_fixture)
    issues: list[str] = []
    check_canvas(prs, issues)
    assert not issues, f"Canvas check failed: {issues}"


# -- fonts check -----------------------------------------------------------

@pytest.mark.parametrize("prs_fixture", ["audit_demo", "xxb_report"])
def test_no_forbidden_fonts(request: pytest.FixtureRequest, prs_fixture: str) -> None:
    prs = request.getfixturevalue(prs_fixture)
    issues: list[str] = []
    check_fonts_and_size(prs, issues)
    # Both example files should pass — no forbidden fonts, no None sizes
    font_or_size_issues = [
        i for i in issues if "[FONT]" in i or "[SIZE-NONE]" in i or "[FONT-NONE]" in i
    ]
    assert not font_or_size_issues, f"Font/size issues: {font_or_size_issues}"


# -- layout variety --------------------------------------------------------

@pytest.mark.parametrize("prs_fixture", ["audit_demo", "xxb_report"])
def test_layout_variety(request: pytest.FixtureRequest, prs_fixture: str) -> None:
    prs = request.getfixturevalue(prs_fixture)
    issues: list[str] = []
    check_layout_variety(prs, issues)
    assert not issues, f"Layout variety check failed: {issues}"


# -- BJH residue -----------------------------------------------------------

@pytest.mark.parametrize("prs_fixture", ["audit_demo", "xxb_report"])
def test_no_bjh_residue(request: pytest.FixtureRequest, prs_fixture: str) -> None:
    prs = request.getfixturevalue(prs_fixture)
    issues: list[str] = []
    check_bjh_residue(prs, issues)
    assert not issues, f"BJH residue found: {issues}"


# -- grey note color -------------------------------------------------------

@pytest.mark.parametrize("prs_fixture", ["audit_demo", "xxb_report"])
def test_no_edit_grey_color(request: pytest.FixtureRequest, prs_fixture: str) -> None:
    prs = request.getfixturevalue(prs_fixture)
    issues: list[str] = []
    check_edit_grey_color(prs, issues)
    assert not issues, f"Edit-note grey found: {issues}"
