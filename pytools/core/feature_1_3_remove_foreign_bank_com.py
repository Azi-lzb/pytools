from __future__ import annotations

import shutil
from pathlib import Path
from typing import Iterable

from .com_engine import safe_close, safe_quit, start_excel_app
from .feature_1_3_remove_foreign_bank import COPY_SUFFIX
from .logger import get_logger


def _build_copy_path(src: Path) -> Path:
    return src.with_name(f"{src.stem}{COPY_SUFFIX}{src.suffix}")


def _last_row_col_a(ws) -> int:
    used = ws.UsedRange
    last_row = int(used.Row) + int(used.Rows.Count) - 1
    for r in range(last_row, 0, -1):
        v = ws.Cells(r, 1).Value
        if v is not None and str(v).strip() != "":
            return r
    return 0


def run_remove_foreign_bank_com(target_wbs: Iterable[Path], foreign_set: set[str],
                                log_dir: Path) -> dict:
    logger = get_logger("feature_1_3_remove_foreign_bank_com", log_dir)
    if not foreign_set:
        return {"files": 0, "ok_files": 0, "branch_sheets": 0,
                "deleted_rows": 0, "failed_files": 0, "foreign_empty": True}

    total = {"files": 0, "ok_files": 0, "branch_sheets": 0,
             "deleted_rows": 0, "failed_files": 0, "foreign_empty": False}
    app = start_excel_app()
    try:
        for src in target_wbs:
            total["files"] += 1
            dst = _build_copy_path(src)
            wb = None
            try:
                shutil.copy2(src, dst)
                wb = app.Workbooks.Open(str(dst.resolve()), ReadOnly=False, UpdateLinks=0)
                deleted_rows = 0
                branch_sheets = 0
                for ws in wb.Worksheets:
                    if "分机构" not in str(ws.Name):
                        continue
                    branch_sheets += 1
                    last_row = _last_row_col_a(ws)
                    marked: list[bool] = [False] * (last_row + 1)
                    for r in range(1, last_row + 1):
                        v = ws.Cells(r, 1).Value
                        if v is None:
                            continue
                        name = str(v).strip()
                        if name and name in foreign_set:
                            marked[r] = True
                    r = last_row
                    while r >= 1:
                        if marked[r]:
                            end = r
                            while r >= 1 and marked[r]:
                                r -= 1
                            start = r + 1
                            ws.Rows(f"{start}:{end}").Delete()
                            deleted_rows += end - start + 1
                        else:
                            r -= 1
                wb.Save()
                total["ok_files"] += 1
                total["branch_sheets"] += branch_sheets
                total["deleted_rows"] += deleted_rows
            except Exception as e:
                total["failed_files"] += 1
                logger.error(f"文件失败: {src} - {e}")
            finally:
                safe_close(wb)
    finally:
        safe_quit(app)
    return total

