"""1.6 提取工作表数据 — COM 版本

策略：
- full_extract=True：用 src_ws.Copy(After=...) 整表过去，再把公式就地转为值，**保留批注/格式/合并/列宽，公式按源结果落地**
- full_extract=False：用 Range.Value2 批量读值 + 行列过滤 + 紧凑后写入新 sheet
  （部分提取场景需要"剔除空行/空列"，这一步要求纯值视图，因此不再保留格式）
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

from .com_engine import (
    open_readonly,
    safe_close,
    safe_quit,
    start_excel_app,
)
from .feature_1_6_extract_sheets import (
    _build_exact_name,
    _compact,
    _extract_values,
    _parse_indexes,  # noqa: F401  - re-exported indirectly via _extract_values
    _sanitize_sheet_name,
    _sheet_matches_all,
)
from .logger import get_logger


XL_OPEN_XML_WORKBOOK = 51


def _list_sheet_names(wb) -> list[str]:
    return [str(ws.Name) for ws in wb.Worksheets]


def _find_matched_sheets_com(wb, keywords: list[str]):
    """精确匹配（用 - 拼起来）+ 模糊匹配（所有非空关键词均为子串）。"""
    matched = []
    seen: set[str] = set()
    exact = _build_exact_name(keywords)
    if exact:
        for ws in wb.Worksheets:
            if str(ws.Name) == exact:
                matched.append(ws)
                seen.add(str(ws.Name))
                break
    for ws in wb.Worksheets:
        title = str(ws.Name)
        if title in seen:
            continue
        if _sheet_matches_all(title, keywords):
            matched.append(ws)
            seen.add(title)
    return matched


def _unique_sheet_name_com(wb, base: str) -> str:
    base = _sanitize_sheet_name(base)
    existing = set(_list_sheet_names(wb))
    if base not in existing:
        return base
    root = base[:28] if len(base) > 28 else base
    for i in range(2, 101):
        cand = f"{root}_{i}"
        if cand not in existing:
            return cand
    return f"{root}_{datetime.now().strftime('%H%M%S')}"


def _add_clean_workbook(app):
    """新建只含 1 张默认 sheet 的工作簿。"""
    old = None
    try:
        old = app.SheetsInNewWorkbook
        app.SheetsInNewWorkbook = 1
    except Exception:
        pass
    wb = app.Workbooks.Add()
    if old is not None:
        try:
            app.SheetsInNewWorkbook = old
        except Exception:
            pass
    return wb


def _read_full_2d(src_ws) -> list[list]:
    """整张 sheet 的值（用于 partial 提取）。"""
    ur = src_ws.UsedRange
    last_row = int(ur.Row) + int(ur.Rows.Count) - 1
    last_col = int(ur.Column) + int(ur.Columns.Count) - 1
    if last_row < 1 or last_col < 1:
        return []
    rng = src_ws.Range(src_ws.Cells(1, 1), src_ws.Cells(last_row, last_col))
    v = rng.Value2
    if v is None:
        return []
    if not isinstance(v, tuple):
        return [[v]]
    return [list(row) for row in v]


def _write_values_com(dst_ws, values: list[list]) -> None:
    if not values:
        return
    nrow = len(values)
    ncol = max((len(r) for r in values), default=0)
    if ncol == 0:
        return
    # 对齐为等长 tuple of tuples，pywin32 需要规整结构
    payload = tuple(
        tuple(row[i] if i < len(row) else None for i in range(ncol))
        for row in values
    )
    dst_ws.Range(dst_ws.Cells(1, 1), dst_ws.Cells(nrow, ncol)).Value2 = payload


def _is_blank_spec(spec: str) -> bool:
    return str(spec or "").strip() == ""


def run_extract_sheets_com(target_wbs: Iterable[Path], tasks: list[dict],
                           output_dir: Path, log_dir: Path) -> dict:
    logger = get_logger("feature_1_6_extract_sheets_com", log_dir)
    if not tasks:
        logger.warning("'工作表提取'配置无启用任务")
        return {"files": 0, "extracted_sheets": 0, "output_files": 0,
                "failed_files": 0, "paths": []}

    ts_default = f"提取结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    output_dir.mkdir(parents=True, exist_ok=True)

    out_wbs: dict[str, object] = {}        # name -> COM Workbook (in memory)
    out_paths: dict[str, Path] = {}        # name -> save path
    out_default_sheet: dict[str, str] = {}  # name -> default sheet to drop later
    stat = {"files": 0, "extracted_sheets": 0, "output_files": 0,
            "failed_files": 0, "paths": []}

    app = start_excel_app()
    try:
        for src in target_wbs:
            src_wb = None
            try:
                src_wb = open_readonly(app, src)
            except Exception as e:
                stat["failed_files"] += 1
                logger.error(f"打开失败: {src} - {e}")
                continue

            stat["files"] += 1
            logger.info(f"[COM] 打开源文件: {src}")
            try:
                for task in tasks:
                    matched = _find_matched_sheets_com(src_wb, task["keywords"])
                    if not matched:
                        continue

                    out_name = task["output_name"] or ts_default
                    if out_name not in out_wbs:
                        wb_new = _add_clean_workbook(app)
                        out_wbs[out_name] = wb_new
                        out_paths[out_name] = output_dir / f"{out_name}.xlsx"
                        out_default_sheet[out_name] = str(wb_new.Worksheets(1).Name)
                    wb_out = out_wbs[out_name]

                    for src_sheet in matched:
                        src_title = str(src_sheet.Name)
                        sheet_title = _unique_sheet_name_com(wb_out, src_title)
                        if task["full_extract"]:
                            # 整表 Copy（保留批注/格式/合并/列宽），随后把公式就地转为值
                            try:
                                src_sheet.Copy(After=wb_out.Worksheets(wb_out.Worksheets.Count))
                                new_ws = wb_out.Worksheets(wb_out.Worksheets.Count)
                                try:
                                    new_ws.Name = sheet_title
                                except Exception:
                                    pass
                                # 公式 → 值（避免跨簿引用断裂、保持纯值视图）
                                try:
                                    ur = new_ws.UsedRange
                                    ur.Value = ur.Value
                                except Exception:
                                    pass
                                stat["extracted_sheets"] += 1
                                logger.info(
                                    f"  整表提取: {src.name} | {src_title} -> "
                                    f"{out_name}.xlsx[{sheet_title}]"
                                )
                            except Exception as e:
                                logger.warning(f"整表 Copy 失败: {e}")
                        else:
                            # COM 场景下，未指定 rows/cols 时直接复制 UsedRange，保留合并/边框等格式
                            if _is_blank_spec(task.get("rows_spec", "")) and _is_blank_spec(task.get("cols_spec", "")):
                                try:
                                    new_ws = wb_out.Worksheets.Add(
                                        After=wb_out.Worksheets(wb_out.Worksheets.Count)
                                    )
                                    try:
                                        new_ws.Name = sheet_title
                                    except Exception:
                                        pass
                                    src_sheet.UsedRange.Copy(new_ws.Range("A1"))
                                    stat["extracted_sheets"] += 1
                                    logger.info(
                                        f"  UsedRange复制: {src.name} | {src_title} -> "
                                        f"{out_name}.xlsx[{sheet_title}]"
                                    )
                                    continue
                                except Exception as e:
                                    logger.warning(f"UsedRange 复制失败，回退值提取: {e}")

                            # 部分提取（rows/cols spec），按值视图
                            data = _read_full_2d(src_sheet)
                            values = _extract_values(
                                data, task["rows_spec"],
                                task["cols_spec"], full_extract=False,
                            )
                            values = _compact(values)
                            if not values:
                                logger.info(f"  [skip] {src.name} | {src_title} 无数据")
                                continue
                            new_ws = wb_out.Worksheets.Add(
                                After=wb_out.Worksheets(wb_out.Worksheets.Count)
                            )
                            try:
                                new_ws.Name = sheet_title
                            except Exception:
                                pass
                            _write_values_com(new_ws, values)
                            stat["extracted_sheets"] += 1
                            logger.info(
                                f"  部分提取: {src.name} | {src_title} -> "
                                f"{out_name}.xlsx[{sheet_title}]"
                            )
            finally:
                safe_close(src_wb)

        # 保存所有输出工作簿，先删默认空白 sheet
        for name, wb_out in out_wbs.items():
            try:
                default = out_default_sheet.get(name, "")
                if default:
                    for ws in wb_out.Worksheets:
                        if str(ws.Name) == default and wb_out.Worksheets.Count > 1:
                            ws.Delete()
                            break
                path = out_paths[name]
                wb_out.SaveAs(str(path.resolve()), FileFormat=XL_OPEN_XML_WORKBOOK)
                stat["output_files"] += 1
                stat["paths"].append(str(path))
                logger.info(f"已保存: {path}")
            except Exception as e:
                logger.warning(f"保存输出失败 {name}: {e}")
            finally:
                safe_close(wb_out)
    finally:
        safe_quit(app)

    return stat
