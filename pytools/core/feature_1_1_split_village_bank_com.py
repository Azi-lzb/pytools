from __future__ import annotations

from pathlib import Path
from typing import Iterable

from .com_engine import safe_close, safe_quit, start_excel_app
from .feature_1_1_split_village_bank import (
    INDICATORS,
    SPLIT_ROW_1_NAME,
    SPLIT_ROW_2_NAME,
    VILLAGE_BANK_KEY,
)
from .logger import get_logger


def _to_num(v) -> float:
    if v is None:
        return 0.0
    try:
        return float(v)
    except Exception:
        try:
            return float(str(v).strip())
        except Exception:
            return 0.0


def _yoy(curr: float, prev: float):
    if prev == 0:
        return 0 if curr == 0 else "∞"
    return round((curr / prev - 1) * 100, 2)


def _sheet_type(name: str) -> str | None:
    if "本外币" in name:
        return "本外币"
    if "人民币" in name:
        return "人民币"
    if "外汇" in name:
        return "外汇"
    return None


def _find_sheet_by_substr(wb, substr: str):
    for ws in wb.Worksheets:
        if substr in str(ws.Name):
            return ws
    return None


def _find_village_row(ws) -> int | None:
    used = ws.UsedRange
    last_row = int(used.Row) + int(used.Rows.Count) - 1
    for r in range(1, last_row + 1):
        v = ws.Cells(r, 1).Value
        if v is not None and str(v).strip() == VILLAGE_BANK_KEY:
            return r
    return None


def _compute_matrix_com(cur_wb, prev_wb) -> list[list]:
    matrix: list[list] = [[0] * 8 for _ in range(6)]
    for idx, indicator in enumerate(INDICATORS, start=1):
        ws_cur = _find_sheet_by_substr(cur_wb, indicator)
        if ws_cur is None:
            continue
        ws_prev = None
        for x in prev_wb.Worksheets:
            if str(x.Name) == str(ws_cur.Name):
                ws_prev = x
                break
        if ws_prev is None:
            continue
        for row_offset, base_row in enumerate((5, 6)):
            target_row = idx + row_offset * 3
            curr_dep = _to_num(ws_cur.Cells(base_row, 3).Value)
            prev_dep = _to_num(ws_prev.Cells(base_row, 3).Value)
            matrix[target_row - 1][0] = curr_dep
            matrix[target_row - 1][1] = ws_cur.Cells(base_row, 4).Value
            matrix[target_row - 1][2] = ws_cur.Cells(base_row, 5).Value
            matrix[target_row - 1][3] = _yoy(curr_dep, prev_dep)

            curr_loan = _to_num(ws_cur.Cells(base_row, 87).Value)  # CI
            prev_loan = _to_num(ws_prev.Cells(base_row, 87).Value)
            matrix[target_row - 1][4] = curr_loan
            matrix[target_row - 1][5] = ws_cur.Cells(base_row, 88).Value  # CJ
            matrix[target_row - 1][6] = ws_cur.Cells(base_row, 89).Value  # CK
            matrix[target_row - 1][7] = _yoy(curr_loan, prev_loan)
    return matrix


def _fill(ws, anchor_row: int, matrix: list[list], t: str) -> None:
    row_map = {"本外币": (0, 3), "人民币": (1, 4), "外汇": (2, 5)}
    top_idx, bot_idx = row_map[t]
    for i in range(8):
        ws.Cells(anchor_row, i + 2).Value = matrix[top_idx][i]
        ws.Cells(anchor_row + 1, i + 2).Value = matrix[bot_idx][i]


def run_split_village_bank_com(current_wb: Path, prev_wb: Path,
                               target_wbs: Iterable[Path], log_dir: Path) -> dict:
    logger = get_logger("feature_1_1_split_village_bank_com", log_dir)
    total = {"files": 0, "branch_sheets": 0, "filled_sheets": 0, "skipped_sheets": 0,
             "failed_files": 0, "matrix_empty": False}

    app = start_excel_app()
    cur = prev = None
    try:
        cur = app.Workbooks.Open(str(current_wb.resolve()), ReadOnly=True, UpdateLinks=0)
        prev = app.Workbooks.Open(str(prev_wb.resolve()), ReadOnly=True, UpdateLinks=0)
        matrix = _compute_matrix_com(cur, prev)
    finally:
        safe_close(prev)
        safe_close(cur)

    if all(matrix[0][i] in (0, None) for i in (0, 4)):
        total["matrix_empty"] = True
        return total

    for path in target_wbs:
        wb = None
        try:
            wb = app.Workbooks.Open(str(path.resolve()), ReadOnly=False, UpdateLinks=0)
            stat = {"branch_sheets": 0, "filled_sheets": 0, "skipped_sheets": 0}
            for ws in wb.Worksheets:
                title = str(ws.Name)
                if "分机构" not in title:
                    continue
                stat["branch_sheets"] += 1
                t = _sheet_type(title)
                if t is None:
                    stat["skipped_sheets"] += 1
                    continue
                row = _find_village_row(ws)
                if row is None:
                    stat["skipped_sheets"] += 1
                    continue
                ws.Rows(row).Delete()
                ws.Rows(row).Insert()
                ws.Rows(row).Insert()
                ws.Cells(row, 1).Value = SPLIT_ROW_1_NAME
                ws.Cells(row + 1, 1).Value = SPLIT_ROW_2_NAME
                _fill(ws, row, matrix, t)
                stat["filled_sheets"] += 1
            wb.Save()
            total["files"] += 1
            for k in ("branch_sheets", "filled_sheets", "skipped_sheets"):
                total[k] += stat[k]
        except Exception as e:
            total["failed_files"] += 1
            logger.error(f"文件失败: {path} - {e}")
        finally:
            safe_close(wb)
    safe_quit(app)
    return total

