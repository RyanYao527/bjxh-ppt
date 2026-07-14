# Changelog

## v2.0 (2026-07-14)

### Added
- **排版引擎** (`layout_utils.py`): 动态字号、占位符清理、内容感知版式选择、标题智能截断
- **新模板适配**: 5.27定稿-北京兴华标准化模板 (59 slides, 80 layouts, 10 mapped)
- 图表版式支持 (`图表-1`, `图表-2`) + 数据关键词自动检测
- `inspect_placeholders.py`: dump 版式 placeholder 元数据
- E2E smoke test in CI
- `requirements.txt` (production deps)

### Changed
- **模块拆分**: `from_outline.py` 638→45行 (→ `parse.py`/`render.py`/`cli.py`)
- **默认版式**: `无图分段-3项` → `标题页-空白` (匹配模板实际使用)
- **默认字号基线**: 16pt → 14pt (匹配新模板统计)
- **TOC 行为**: >3章节自动禁用目录 (新模板限制)
- config.py: legacy 路径改为 5.27 模板
- SKILL.md: 更新模板规格为 5.27 定稿

### Fixed
- `render_page()` 引用 `main()` 局部变量 (改为参数传入)
- 空占位符残留 "单击此处添加文本" (自动清理)
- logo zip 提取增加错误处理
- BJH 正则增加 `BJH`/`B.J.H` 匹配
- suggest_layout 运行时校验 PLACEHOLDER_MAP

---

## v1.2

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
