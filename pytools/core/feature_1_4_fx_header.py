from __future__ import annotations

from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook

from .logger import get_logger


FX_SHEET_KEY = "外汇"
OLD_UNIT = "万元"
NEW_UNIT = "万美元"


def _replace_header_part(part, changed_flag: list[bool]) -> None:
    txt = getattr(part, "text", None)
    if txt and OLD_UNIT in txt:
        part.text = txt.replace(OLD_UNIT, NEW_UNIT)
        changed_flag[0] = True


def _process_sheet(ws) -> bool:
    changed = [False]
    for hf in (ws.oddHeader, ws.evenHeader, ws.firstHeader):
        if hf is None:
            continue
        for part in (hf.left, hf.center, hf.right):
            if part is not None:
                _replace_header_part(part, changed)
    return changed[0]


def _process_one_workbook(path: Path, logger) -> dict:
    stat = {"fx_sheets": 0, "modified_sheets": 0, "saved": False}
    wb = load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
    try:
        file_changed = False
        for ws in wb.worksheets:
            if FX_SHEET_KEY not in ws.title:
                continue
            stat["fx_sheets"] += 1
            if _process_sheet(ws):
                stat["modified_sheets"] += 1
                file_changed = True
                logger.info(f"  sheet '{ws.title}' 页眉已修改")
        if file_changed:
            wb.save(path)
            stat["saved"] = True
    finally:
        wb.close()
    return stat


def run_fx_header_fix(target_wbs: Iterable[Path], log_dir: Path) -> dict:
    logger = get_logger("feature_1_4_fx_header", log_dir)
    total = {"files": 0, "saved_files": 0, "fx_sheets": 0,
             "modified_sheets": 0, "failed_files": 0}
    for path in target_wbs:
        total["files"] += 1
        try:
            logger.info(f"处理文件: {path}")
            stat = _process_one_workbook(path, logger)
            total["fx_sheets"] += stat["fx_sheets"]
            total["modified_sheets"] += stat["modified_sheets"]
            if stat["saved"]:
                total["saved_files"] += 1
            logger.info(
                f"文件完成: {path.name} 外汇sheet={stat['fx_sheets']} "
                f"修改={stat['modified_sheets']} 保存={stat['saved']}"
            )
        except Exception as e:
            total["failed_files"] += 1
            logger.error(f"文件失败: {path} - {e}")
    return total
