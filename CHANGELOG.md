# Changelog

## v1.2 (unreleased)

### Added
- `scripts/shared.py`: shared `apply_minimal_formatting` utility (removes `sys.path` hack).
- `scripts/config.py`: template path resolution via `BJXH_TEMPLATE` env var → `config.json` → legacy fallback.
- `scripts/config.example.json`: template for local configuration.
- `scripts/inspect_placeholders.py`: dump placeholder idx/name/position per layout for PLACEHOLDER_MAP verification.
- `tests/`: 40 pytest tests covering `parse_outline`, QA regexes, and constants.
- `pyproject.toml` with pytest + ruff + mypy configuration.
- `requirements-dev.txt` for dev dependencies.

### Changed
- **Split `from_outline.py`** (638 lines → 45 lines): logic moved to `parse.py`, `render.py`, `cli.py`.
- `from_template.py`: `apply_minimal_formatting` now imported from `shared`.
- `qa.py`: replaced silent `except: pass` with `hasattr` guard; full type annotations.
- `parse_outline()`: duplicate `# H1`, unrecognized lines, and unknown directives now emit stderr warnings.
- **Fixed**: `render_page()` closing-slide code previously referenced `main()`'s local `template_path` variable (would `NameError` at runtime). Now an explicit parameter.
- Template path is no longer hardcoded — resolved via the lookup chain described above.
- Company contact info (phone/fax/address) read from `config.json` instead of embedded in source.

### Removed
- Dead variable `current_chapter` in `parse_outline()`.

---

## v1.1 (2026-06-18)

- Corrected theme-color references (`MSO_THEME_COLOR_INDEX` import path).
- Adjusted default content layout from `文字模板1` to `无图分段-3项`.
- Added `inspect_placeholders.py` references for PLACEHOLDER_MAP verification.
- v3 spec revisions: cover title 66pt, subtitle 18pt, closing title repositioned to y=2.50in.
- Added `preview_html.py` for visual overflow checking.
- Added `--no-chapter-covers` compact mode.

---

## v1.0 (initial)

- Initial release with `from_outline.py`, `from_template.py`, `qa.py`, `dump_template_spec.py`.
- Based on 2026 partner-meeting template (8 sample slides, 111 layouts).
- Markdown outline → `.pptx` pipeline with 9-rule QA validator.
