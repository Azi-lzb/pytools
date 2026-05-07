from __future__ import annotations

from pathlib import Path

from .com_engine import safe_close, safe_quit, start_excel_app
from .feature_1_8_adjust_rural_loan import (
    AREA_SHEET_KEYS,
    CITY_NAME,
    HEADER_SCAN_ROWS,
    HEADER_TEXT,
    INST_DATA_START_ROW,
    INST_DIFF_COL,
    INST_NAME_COL,
    INST_SHEET_KEYS,
    INST_VALUE_COL,
    NAME_SCAN_COLS,
    TOLERANCE,
    _normalize,
    _to_num,
)
from .logger import get_logger


def _last_row(ws) -> int:
    used = ws.UsedRange
    return int(used.Row) + int(used.Rows.Count) - 1


def _last_col(ws) -> int:
    used = ws.UsedRange
    return int(used.Column) + int(used.Columns.Count) - 1


def _find_sheet(wb, keys: tuple[str, ...]):
    hits = [ws for ws in wb.Worksheets if all(k in str(ws.Name) for k in keys)]
    if len(hits) > 1:
        raise ValueError(f"匹配到多个目标工作表（关键词 {keys}），请保留唯一目标表。")
    return hits[0] if hits else None


def _build_inst_map(ws) -> tuple[dict[str, float], dict[str, int], list[str]]:
    values: dict[str, float] = {}
    rows: dict[str, int] = {}
    dups: list[str] = []
    last_r = _last_row(ws)
    for r in range(INST_DATA_START_ROW, last_r + 1):
        name = _normalize(ws.Cells(r, INST_NAME_COL).Value)
        if not name:
            continue
        if name in values:
            dups.append(name)
            continue
        values[name] = _to_num(ws.Cells(r, INST_VALUE_COL).Value)
        rows[name] = r
    return values, rows, dups


def _find_header_col(ws, header: str) -> int:
    max_r = min(_last_row(ws), HEADER_SCAN_ROWS)
    lc = _last_col(ws)
    for r in range(1, max_r + 1):
        for c in range(1, lc + 1):
            if _normalize(ws.Cells(r, c).Value) == header:
                return c
    return 0


def _find_city_row(ws, city: str) -> int:
    lr = _last_row(ws)
    for r in range(1, lr + 1):
        for c in range(1, NAME_SCAN_COLS + 1):
            if _normalize(ws.Cells(r, c).Value) == city:
                return r
    return 0


def run_adjust_rural_loan_com(current_path: Path, previous_path: Path, log_dir: Path) -> dict:
    logger = get_logger("feature_1_8_adjust_rural_loan_com", log_dir)
    result = {
        "matched": 0, "fill_total": 0.0, "area_total": 0.0,
        "unmatched_current": [], "previous_only": [],
        "consistent": False, "error": None,
    }

    app = start_excel_app()
    wb_cur = wb_prev = None
    try:
        wb_cur = app.Workbooks.Open(str(current_path.resolve()), ReadOnly=False, UpdateLinks=0)
        wb_prev = app.Workbooks.Open(str(previous_path.resolve()), ReadOnly=True, UpdateLinks=0)

        ws_cur_inst = _find_sheet(wb_cur, INST_SHEET_KEYS)
        if ws_cur_inst is None:
            raise ValueError("本期文件未找到包含“惠州市”且包含“涉农贷款分机构”的工作表。")
        ws_prev_inst = _find_sheet(wb_prev, INST_SHEET_KEYS)
        if ws_prev_inst is None:
            raise ValueError("上期文件未找到包含“惠州市”且包含“涉农贷款分机构”的工作表。")
        ws_cur_area = _find_sheet(wb_cur, AREA_SHEET_KEYS)
        if ws_cur_area is None:
            raise ValueError("本期文件未找到包含“本外币”且包含“涉农贷款分地区”的工作表。")

        cur_values, cur_rows, cur_dups = _build_inst_map(ws_cur_inst)
        if cur_dups:
            raise ValueError("本期分机构表存在重复机构名称：" + "、".join(cur_dups))
        prev_values, _, prev_dups = _build_inst_map(ws_prev_inst)
        if prev_dups:
            raise ValueError("上期分机构表存在重复机构名称：" + "、".join(prev_dups))

        fill_total = 0.0
        unmatched: list[str] = []
        for name, r in cur_rows.items():
            if name in prev_values:
                diff = cur_values[name] - prev_values[name]
                ws_cur_inst.Cells(r, INST_DIFF_COL).Value = diff
                fill_total += diff
                result["matched"] += 1
            else:
                ws_cur_inst.Cells(r, INST_DIFF_COL).Value = None
                unmatched.append(name)
        previous_only = [n for n in prev_values if n not in cur_values]

        header_col = _find_header_col(ws_cur_area, HEADER_TEXT)
        if header_col == 0:
            raise ValueError('未在“本外币涉农贷款分地区”表中找到表头“涉农贷款比上月”。')
        city_row = _find_city_row(ws_cur_area, CITY_NAME)
        if city_row == 0:
            raise ValueError('未在“本外币涉农贷款分地区”表中找到“惠州市”所在行。')
        area_total = _to_num(ws_cur_area.Cells(city_row, header_col).Value)

        wb_cur.Save()
        result["fill_total"] = fill_total
        result["area_total"] = area_total
        result["unmatched_current"] = unmatched
        result["previous_only"] = previous_only
        result["consistent"] = abs(fill_total - area_total) <= TOLERANCE
    except Exception as e:
        result["error"] = str(e)
        logger.error(f"失败: {e}")
    finally:
        safe_close(wb_prev)
        safe_close(wb_cur)
        safe_quit(app)
    return result

