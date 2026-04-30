"""1.4 外汇页眉修改 — COM 版本

直接操作 Excel PageSetup 的 LeftHeader/CenterHeader/RightHeader（含偶数页/首页变体）。
对比 openpyxl 版的优势：偶数页/首页页眉处理更稳，page setup 副属性不会丢。
"""
from __future__ import annotations

from pathlib import Path
from typing import Iterable

from .com_engine import safe_close, safe_quit, start_excel_app
from .logger import get_logger


FX_SHEET_KEY = "外汇"
OLD_UNIT = "万元"
NEW_UNIT = "万美元"


def _replace_in_default(ps) -> bool:
    """默认（奇数）页页眉。COM 里这三个属性是字符串。"""
    changed = False
    for attr in ("LeftHeader", "CenterHeader", "RightHeader"):
        try:
            txt = getattr(ps, attr) or ""
            if OLD_UNIT in txt:
                setattr(ps, attr, txt.replace(OLD_UNIT, NEW_UNIT))
                changed = True
        except Exception:
            pass
    return changed


def _replace_in_subpage(sub_obj) -> bool:
    """偶数页 / 首页页眉。COM 里 EvenPage.LeftHeader 是 HeaderFooter 对象，需要 .Text。"""
    changed = False
    if sub_obj is None:
        return False
    for attr in ("LeftHeader", "CenterHeader", "RightHeader"):
        try:
            hf = getattr(sub_obj, attr)
            if hf is None:
                continue
            txt = hf.Text or ""
            if OLD_UNIT in txt:
                hf.Text = txt.replace(OLD_UNIT, NEW_UNIT)
                changed = True
        except Exception:
            pass
    return changed


def _replace_unit_in_headers(ps) -> bool:
    changed = _replace_in_default(ps)
    # 偶数页：仅当工作表启用 OddAndEvenPagesHeaderFooter 时
    try:
        if ps.OddAndEvenPagesHeaderFooter:
            if _replace_in_subpage(ps.EvenPage):
                changed = True
    except Exception:
        pass
    # 首页：仅当工作表启用 DifferentFirstPageHeaderFooter 时
    try:
        if ps.DifferentFirstPageHeaderFooter:
            if _replace_in_subpage(ps.FirstPage):
                changed = True
    except Exception:
        pass
    return changed


def run_fx_header_fix_com(target_wbs: Iterable[Path], log_dir: Path) -> dict:
    logger = get_logger("feature_1_4_fx_header_com", log_dir)
    total = {"files": 0, "saved_files": 0, "fx_sheets": 0,
             "modified_sheets": 0, "failed_files": 0}
    valid = [p for p in target_wbs
             if p.exists() and p.suffix.lower() in (".xlsx", ".xlsm", ".xls")]
    if not valid:
        return total

    app = start_excel_app()
    try:
        for path in valid:
            wb = None
            try:
                logger.info(f"[COM] 处理文件: {path}")
                wb = app.Workbooks.Open(str(path.resolve()), ReadOnly=False, UpdateLinks=0)
            except Exception as e:
                total["failed_files"] += 1
                logger.error(f"打开失败: {path} - {e}")
                continue
            try:
                fx_count = 0
                modified_count = 0
                file_changed = False
                for ws in wb.Worksheets:
                    if FX_SHEET_KEY not in str(ws.Name):
                        continue
                    fx_count += 1
                    if _replace_unit_in_headers(ws.PageSetup):
                        modified_count += 1
                        file_changed = True
                        logger.info(f"  sheet '{ws.Name}' 页眉已修改")
                total["files"] += 1
                total["fx_sheets"] += fx_count
                total["modified_sheets"] += modified_count
                if file_changed:
                    wb.Save()
                    total["saved_files"] += 1
                logger.info(
                    f"文件完成: {path.name} 外汇sheet={fx_count} "
                    f"修改={modified_count} 保存={file_changed}"
                )
            except Exception as e:
                total["failed_files"] += 1
                logger.error(f"处理失败: {path} - {e}")
            finally:
                safe_close(wb)
    finally:
        safe_quit(app)
    return total
