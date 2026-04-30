# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

`pytools` is an interactive CLI tool for Excel/WPS report batch processing (banking report workflows). It exposes 7 menu groups: 时序工具 / 去重工具 / 打印工具 / 转换工具 / 配置工具 / 汇总工具 / 一二批处理工具 (+ 调研选项统计). All user interaction happens through the menu in `pytools/main.py`; there is no library API.

## Commands

```bash
# Run from source (cwd = project root)
python pytools/main.py

# Run tests (pytest, end-to-end fixtures only)
python -m pytest pytools/tests -v
python -m pytest pytools/tests/test_e2e.py::<test_name> -v   # single test

# Build single-file Windows exe (writes dist/pytools.exe)
build_exe.bat
```

`build_exe.bat` calls PyInstaller via `pytools.spec` (onefile, console). It auto-`cd`s to its own directory and prefers `%LocalAppData%\Programs\Python\Python310\python.exe`, falling back to `py -3`.

## Architecture

### Entry & path resolution

- `pytools_entry.py` — PyInstaller entry point; only re-exports `pytools.main:main`.
- `pytools/main.py` — interactive menu loop. `_app_base_dir()` switches between source mode (`__file__.parent`) and frozen mode (`sys.executable.parent`). **Consequence:** in the built exe, `config/config.xlsx`, `output/`, and `logs/` live next to `pytools.exe`, not inside any bundled directory.

### Two engines for the same feature (pandas vs COM)

Most heavy features ship in pairs: a pandas/openpyxl version and a `_com`-suffixed Excel-COM version (uses `pywin32`). Examples: `compare_393.py` ↔ `compare_393_com.py`, `timeline_396.py` ↔ `timeline_396_com.py`, `wide_397_398.py` ↔ `wide_397_398_com.py`, `print_310.py` ↔ `print_310_com.py`.

The COM branch uses `core/com_engine.py`, which exposes:
- `start_excel_app()` — boots Excel/WPS (`Excel.Application` → `ket.Application` fallback), silences alerts/events/calculation.
- `ArrayLike` — a thin 2D-list shim with `.shape` and `.iat[r, c]`, **deliberately mimicking the pandas DataFrame surface** that pure-computation helpers (e.g. `extract_cells`) rely on. This lets the same downstream extraction logic serve both engines.

When editing a feature, check whether a `_com` twin exists and apply the same change to both unless intentionally diverging.

### Config: one xlsx, many sheets

All runtime config lives in `pytools/config/config.xlsx`. Sheet names and required columns are declared as constants in `core/config_xlsx.py` (`SHEET_GLOBAL`, `SHEET_TIMELINE_RULE`, `SHEET_PATH_MAP`, `SHEET_DEDUP_TASK`, `SHEET_PRINT_CONFIG`, `SHEET_CONFIG_RENAME`, `SHEET_INSTITUTION_MAPPING`, `SHEET_EXTRACT_CONFIG`). `core/config_init.py` (menu `5 → 1`) creates missing sheets and repairs the header row + cell comments **without** wiping data rows.

Loaders return typed dataclasses; `validate_sheets_for_feature()` is the single gate features call before doing real work.

### IO conventions

- `core/io_excel.py` — pandas-side helpers. `read_sheet_2d()` is `lru_cache`d on `(path, mtime, size, sheet)`; **call `clear_sheet_cache()` after finishing one source workbook** to bound memory. `safe_sheet_name()` enforces Excel's 31-char + illegal-char rules with collision suffixes.
- `core/logger.py` — each feature gets its own logger named `pytools.<feature>`, writing to `<logs>/<feature>_<YYYYMMDD_HHMMSS>.log` (UTF-8) plus stderr. Handlers are reset on each call so re-entering a menu doesn't double-log.
- Output goes to `pytools/output/<feature>_<timestamp>.xlsx` (or paths from 全局配置 in production).

### Feature naming

The numeric infixes (`1_4`, `33_34_35_37`, `393`, `396`, `397_398`, `310`, `311`, `321_322`) trace back to VBA menu numbers in the sister project `通用汇总工具vbaExcel`. `feature_1_2_normalize_institution.py` corresponds to that project's menu `1.2`, etc. Don't renumber — the user navigates by these IDs.

## Editing rules specific to this repo

- **All `.py` files must remain UTF-8.** Some Windows editors silently save as GBK/ANSI, which breaks `python pytools/main.py` with `SyntaxError: Non-UTF-8 code starting with...`. If that happens: `python -c "from pathlib import Path; p=Path('pytools/main.py'); p.write_bytes(p.read_bytes().decode('gbk').encode('utf-8'))"`.
- Don't import `win32com` at module top level in non-`_com` files — it would break the pandas branch on machines without `pywin32`. The COM modules and `core/com_engine.py` are the only legitimate importers.
- When adding a new menu item, update both the menu string constant (e.g. `TIMELINE_MENU`) and the dispatch block in `main.py`, then mirror the entry in `pytools/使用指南.md`.

## Sister project

`C:\Users\AZI\Desktop\VibeCode项目\通用汇总工具vbaExcel` is the original VBA implementation. It is a separate codebase; don't cross-edit. Python tooling for VBA module encoding/patches lives there (`convert.py`, `scripts/*.py`) and stays there.
