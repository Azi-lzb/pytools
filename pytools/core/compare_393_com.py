"""3.9.3 表头路径比对（混合）— COM 版本

用 Excel COM 代替 pandas/openpyxl 读取源/模板的单元格矩阵，
比对与输出逻辑与 compare_393.py 完全一致，便于 A/B 对照。
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from .com_engine import (
    ArrayLike,
    find_sheet_com,
    list_sheet_names_com,
    open_readonly,
    read_sheet_values,
    safe_close,
    safe_quit,
    start_excel_app,
)
from .config_xlsx import GlobalConfig, TimelineRule
from .compare_393 import (
    RESULT_COLS,
    _path_diff_rows,
    _position_diff_rows,
    _resolve_template_sheet,
)
from .io_excel import write_workbook
from .logger import get_logger
from .rule_engine import match_all_keywords, pick_rules_for_workbook


def run_compare_com(template_path: Path, source_paths: list[Path],
                    rules: list[TimelineRule], g: GlobalConfig) -> Path:
    """与 run_compare 同功能，改用 Excel COM 读取数据。"""
    log = get_logger("3_9_3_compare_com", g.log_dir)
    log.info(f"[COM] 模板: {template_path.name}, 源 {len(source_paths)} 个, 规则 {len(rules)} 条")

    if not template_path.exists():
        raise FileNotFoundError(f"模板不存在: {template_path}")

    # 护栏：剔除与模板同路径的源文件
    tmpl_resolved = template_path.resolve()
    filtered_sources: list[Path] = []
    skipped_same = 0
    for p in source_paths:
        try:
            if p.resolve() == tmpl_resolved:
                skipped_same += 1
                log.warning(f"跳过源文件（与模板同路径）: {p}")
                continue
        except Exception:
            pass
        filtered_sources.append(p)
    if skipped_same:
        log.info(f"已跳过 {skipped_same} 个与模板同路径的源文件")

    app = start_excel_app()
    template_wb = None
    summary_rows: list[list] = []
    diff_rows: list[list] = []
    total_pos = total_path = 0

    try:
        template_wb = open_readonly(app, template_path)
        template_sheets_all = list_sheet_names_com(template_wb)
        # 模板 sheet → 2D 数组缓存，避免多条规则命中同一 sheet 时重复读取
        tmpl_data_cache: dict[str, ArrayLike] = {}

        def _tmpl_data(name: str) -> ArrayLike:
            if name not in tmpl_data_cache:
                tws = find_sheet_com(template_wb, name)
                tmpl_data_cache[name] = read_sheet_values(tws) if tws is not None else ArrayLike([])
            return tmpl_data_cache[name]

        for src in filtered_sources:
            src_wb = None
            try:
                src_wb = open_readonly(app, src)
            except Exception as e:
                log.warning(f"打开源失败 {src.name}: {e}")
                summary_rows.append([src.name, "", 0, 0, f"打开失败: {e}"])
                continue

            try:
                src_sheet_names = list_sheet_names_com(src_wb)
                src_data_cache: dict[str, ArrayLike] = {}

                def _src_data(name: str) -> ArrayLike:
                    if name not in src_data_cache:
                        sws = find_sheet_com(src_wb, name)
                        src_data_cache[name] = read_sheet_values(sws) if sws is not None else ArrayLike([])
                    return src_data_cache[name]

                per_src_pos = per_src_path = 0
                for rule in pick_rules_for_workbook(rules, src.name, g.timeline_rule_match_mode):
                    if not match_all_keywords(src.name, rule.wb_keyword):
                        continue
                    src_match_sheets = [s for s in src_sheet_names
                                        if match_all_keywords(s, rule.sheet_keyword)]
                    if not src_match_sheets:
                        continue
                    tmpl_match_sheets = [s for s in template_sheets_all
                                         if match_all_keywords(s, rule.sheet_keyword)]
                    if not tmpl_match_sheets:
                        log.warning(f"模板未匹配: 规则={rule.name}")
                        continue

                    for s_sheet in src_match_sheets:
                        t_sheet = _resolve_template_sheet(tmpl_match_sheets, s_sheet)
                        if not t_sheet:
                            log.warning(f"模板表匹配不唯一: 规则={rule.name} 源表={s_sheet}")
                            continue
                        tdf = _tmpl_data(t_sheet)
                        sdf = _src_data(s_sheet)

                        pos = _position_diff_rows(tdf, sdf, rule)
                        for target, sp, tp in pos:
                            diff_rows.append([
                                "", "", rule.name,
                                rule.wb_keyword, rule.sheet_keyword,
                                target, "精确",
                                sp, tp, "按位置",
                                src.name, s_sheet,
                            ])
                        per_src_pos += len(pos)

                        if pos:
                            path_diffs = _path_diff_rows(tdf, sdf, rule)
                            for target, sp, tp in path_diffs:
                                diff_rows.append([
                                    "", "", rule.name,
                                    rule.wb_keyword, rule.sheet_keyword,
                                    target, "精确",
                                    sp, tp, "按路径",
                                    src.name, s_sheet,
                                ])
                            per_src_path += len(path_diffs)

                total_pos += per_src_pos
                total_path += per_src_path
                status = "通过" if (per_src_pos == 0 and per_src_path == 0) else "存在差异"
                summary_rows.append([src.name, status, per_src_pos, per_src_path, ""])
            finally:
                safe_close(src_wb)
    finally:
        safe_close(template_wb)
        safe_quit(app)

    out_path = g.output_dir / f"compare_3_9_3_com_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    diff_df = pd.DataFrame(diff_rows, columns=RESULT_COLS)
    summary_df = pd.DataFrame(
        summary_rows,
        columns=["源文件", "状态", "按位置差异数", "按路径差异数", "备注"],
    )
    write_workbook(out_path, {"表头比对结果": diff_df, "汇总": summary_df})
    log.info(f"[COM] 完成: 源 {len(source_paths)}, 按位置 {total_pos}, 按路径 {total_path} → {out_path}")
    return out_path


