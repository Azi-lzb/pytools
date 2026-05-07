from __future__ import annotations

from pathlib import Path
from typing import Iterable

from .com_engine import safe_close, safe_quit, start_excel_app
from .feature_1_2_normalize_institution import _norm_name
from .logger import get_logger


def _find_last_row(ws, col: int) -> int:
    used = ws.UsedRange
    last_row = int(used.Row) + int(used.Rows.Count) - 1
    for r in range(last_row, 0, -1):
        v = ws.Cells(r, col).Value
        if v is not None and str(v).strip() != "":
            return r
    return 0


def _col_has_data(ws, col: int, last_row: int) -> bool:
    for r in range(2, last_row + 1):
        v = ws.Cells(r, col).Value
        if v is not None and str(v).strip() != "":
            return True
    return False


def run_normalize_institution_com(target_wbs: Iterable[Path], mapping: dict[str, str],
                                  log_dir: Path) -> dict:
    logger = get_logger("feature_1_2_normalize_institution_com", log_dir)
    if not mapping:
        return {"files": 0, "branch_sheets": 0, "mapped_sheets": 0,
                "mapped_rows": 0, "failed_files": 0, "mapping_empty": True}

    keys_lower = [(k.lower(), k) for k in mapping.keys()]
    total = {"files": 0, "branch_sheets": 0, "mapped_sheets": 0,
             "mapped_rows": 0, "failed_files": 0, "mapping_empty": False}

    app = start_excel_app()
    try:
        for path in target_wbs:
            wb = None
            try:
                wb = app.Workbooks.Open(str(path.resolve()), ReadOnly=False, UpdateLinks=0)
                stat = {"branch_sheets": 0, "mapped_sheets": 0, "mapped_rows": 0}
                changed = False
                for ws in wb.Worksheets:
                    if "分机构" not in str(ws.Name):
                        continue
                    stat["branch_sheets"] += 1
                    last_row = max(_find_last_row(ws, 1), _find_last_row(ws, 2))
                    if last_row < 2:
                        continue
                    if _col_has_data(ws, 1, last_row):
                        target_col = 1
                    elif _col_has_data(ws, 2, last_row):
                        target_col = 2
                    else:
                        continue
                    n = 0
                    for r in range(2, last_row + 1):
                        v = ws.Cells(r, target_col).Value
                        if v is None:
                            continue
                        original = str(v).strip()
                        if not original:
                            continue
                        mapped = _norm_name(original, mapping, keys_lower)
                        if mapped != original:
                            ws.Cells(r, target_col).Value = mapped
                            n += 1
                    if n > 0:
                        stat["mapped_sheets"] += 1
                        stat["mapped_rows"] += n
                        changed = True
                if changed:
                    wb.Save()
                total["files"] += 1
                for k in ("branch_sheets", "mapped_sheets", "mapped_rows"):
                    total[k] += stat[k]
            except Exception as e:
                total["failed_files"] += 1
                logger.error(f"文件失败: {path} - {e}")
            finally:
                safe_close(wb)
    finally:
        safe_quit(app)
    return total

