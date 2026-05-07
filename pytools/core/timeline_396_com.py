"""3.9.6 时序提取（快速精简）— COM 版本

数据读取走 Excel COM；合并区域走 openpyxl 缓存（metadata 解析很轻）；
下游纯计算（extract_cells_from_arrays）与 pandas 版本共用同一套代码。
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from .com_engine import (
    ArrayLike,
    get_merge_ranges_com,
    list_sheet_names_com,
    open_readonly,
    read_merged_aware_texts,
    read_sheet_values,
    safe_close,
    safe_quit,
    start_excel_app,
)
from .config_xlsx import GlobalConfig, PathMapRule, TimelineRule
from .io_excel import append_to_target, list_source_files, write_workbook
from .logger import get_logger
from .rule_engine import _kw_match, extract_cells_from_arrays, pick_rules_for_workbook
from .timeline_396 import SLIM_COLS


def _build_header_coords(rule: TimelineRule, rows_n: int, cols_n: int) -> tuple[list[tuple[int, int]], tuple]:
    r_start = rule.data_row_start - 1 if rule.data_row_start else max(r - 1 for r in rule.col_header_rows) + 1
    r_end = (rule.data_row_end - 1) if rule.data_row_end else rows_n - 1
    c_start = rule.data_col_start - 1 if rule.data_col_start else (
        (max((c - 1 for c in getattr(rule, "row_header_cols", []) if c > 0), default=-1) + 1)
        if getattr(rule, "row_header_col_specified", True) else 0
    )
    c_end = (rule.data_col_end - 1) if rule.data_col_end else cols_n - 1
    header_rows = [r - 1 for r in rule.col_header_rows if 0 < r <= rows_n]
    row_header_cols = [c - 1 for c in getattr(rule, "row_header_cols", []) if 0 < c <= cols_n]
    if not row_header_cols and getattr(rule, "row_header_col_specified", True):
        if rule.row_header_col > 0 and rule.row_header_col <= cols_n:
            row_header_cols = [rule.row_header_col - 1]

    coords_set: set[tuple[int, int]] = set()
    for r0 in header_rows:
        for c0 in range(max(c_start, 0), min(c_end, cols_n - 1) + 1):
            coords_set.add((r0, c0))
    if getattr(rule, "row_header_col_specified", True):
        for r0 in range(max(r_start, 0), min(r_end, rows_n - 1) + 1):
            for c0 in row_header_cols:
                coords_set.add((r0, c0))
    sig = (
        tuple(sorted(header_rows)),
        max(r_start, 0), min(r_end, rows_n - 1),
        max(c_start, 0), min(c_end, cols_n - 1),
        tuple(sorted(row_header_cols)),
        bool(getattr(rule, "row_header_col_specified", True)),
    )
    return list(coords_set), sig


def run_timeline_slim_com(rules: list[TimelineRule], path_maps: list[PathMapRule],
                          g: GlobalConfig,
                          source_paths: list[Path] | None = None) -> Path | None:
    log = get_logger("3_9_6_timeline_slim_com", g.log_dir)
    if source_paths is None:
        sources = list_source_files(g.input_dir, g.source_exts)
    else:
        ext_ok = {e.lower() for e in g.source_exts} | {".xls"}
        sources = [p for p in source_paths if p.exists() and p.suffix.lower() in ext_ok]
    log.info(f"[COM] 开始时序提取（精简）: 源文件 {len(sources)} 个，启用规则 {len(rules)} 条")

    rows: list[list] = []
    dedup: set[tuple] = set()
    matched_sheets = 0
    rule_stats: dict[str, dict[str, int]] = {}
    for r in rules:
        rule_stats[r.name] = {"wb_miss": 0, "sheet_miss": 0, "cells": 0}

    app = start_excel_app()
    try:
        for src in sources:
            src_wb = None
            try:
                src_wb = open_readonly(app, src)
            except Exception as e:
                log.warning(f"打开源失败 {src.name}: {e}")
                continue
            try:
                all_sheets = list_sheet_names_com(src_wb)
                sheet_cache: dict[str, tuple[ArrayLike, tuple[tuple[int, int, int, int], ...]]] = {}
                ws_cache: dict[str, object] = {}

                matched_by_rule: dict[str, list[str]] = {}
                for rule in pick_rules_for_workbook(rules, src.name, g.timeline_rule_match_mode):
                    if not _kw_match(src.name, rule.wb_keyword):
                        rule_stats[rule.name]["wb_miss"] += 1
                        continue
                    matched_by_rule[rule.name] = [sn for sn in all_sheets if _kw_match(sn, rule.sheet_keyword)]

                for rule in pick_rules_for_workbook(rules, src.name, g.timeline_rule_match_mode):
                    hit_sheet = False
                    for sn in matched_by_rule.get(rule.name, []):
                        hit_sheet = True
                        if sn not in sheet_cache:
                            ws = src_wb.Worksheets(sn)
                            ws_cache[sn] = ws
                            sheet_cache[sn] = (read_sheet_values(ws), get_merge_ranges_com(ws))
                        ws = ws_cache[sn]
                        df, merge_ranges = sheet_cache[sn]
                        rows_n, cols_n = df.shape
                        coords, _sig = _build_header_coords(rule, rows_n, cols_n)
                        cell_text_overrides = read_merged_aware_texts(ws, coords)
                        cells = extract_cells_from_arrays(
                            df, merge_ranges, src, rule, sn, path_maps,
                            row_suffix_enabled=True,
                            cell_text_overrides=cell_text_overrides,
                        )
                        if cells:
                            matched_sheets += 1
                            rule_stats[rule.name]["cells"] += len(cells)
                        for ec in cells:
                            key = (ec.source_file, ec.sheet_name, ec.data_date,
                                   ec.row_path, ec.col_path, str(ec.value))
                            if key in dedup:
                                continue
                            dedup.add(key)
                            rows.append([ec.source_file, ec.sheet_name, ec.rule_name,
                                         ec.data_date, ec.row_path, ec.col_path, ec.value])
                    if not hit_sheet:
                        rule_stats[rule.name]["sheet_miss"] += 1
            finally:
                safe_close(src_wb)
    finally:
        safe_quit(app)

    df_out = pd.DataFrame(rows, columns=SLIM_COLS)
    if df_out.empty:
        for rule_name, st in rule_stats.items():
            log.info("规则[%s] wb_miss=%s sheet_miss=%s cells=%s",
                     rule_name, st["wb_miss"], st["sheet_miss"], st["cells"])
        log.info("[COM] 完成: 无匹配结果，不生成输出文件")
        return None

    out_path = g.output_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_时序提取（快速精简）_COM.xlsx"
    write_workbook(out_path, {"时序提取结果": df_out})
    log.info(f"[COM] 完成: 匹配 sheet {matched_sheets}，输出 {len(df_out)} 行 → {out_path}")

    # 写入目标簿（与 pandas 版行为一致）
    targets: dict[tuple, list[list]] = {}
    for rule in pick_rules_for_workbook(rules, src.name, g.timeline_rule_match_mode):
        if rule.target_write_enabled and rule.target_wb_path and rule.target_sheet:
            targets[(rule.target_wb_path, rule.target_sheet)] = []
    if targets:
        for row in rows:
            for rule in pick_rules_for_workbook(rules, src.name, g.timeline_rule_match_mode):
                if (row[2] == rule.name and rule.target_write_enabled
                        and rule.target_wb_path and rule.target_sheet):
                    targets[(rule.target_wb_path, rule.target_sheet)].append(row)
        for (wb, sn), rs in targets.items():
            if rs:
                stat = append_to_target(
                    wb, sn, SLIM_COLS, rs,
                    dedup_key_idx=[0, 1, 3, 4, 5],
                    required_prefix=SLIM_COLS[:6],
                    allow_extend=False,
                )
                if stat.get("written", 0) == 0 and stat.get("skipped_reason") == "header_mismatch":
                    log.warning(f"跳过写入目标 {wb}::{sn}: 表头不匹配（保护旧数据）")
                else:
                    log.info(
                        "追加到目标 %s::%s 输入=%s 批内去重后=%s 目标去重后新增=%s",
                        wb, sn,
                        stat.get("input_rows", 0),
                        stat.get("batch_dedup_rows", 0),
                        stat.get("existing_filtered_rows", 0),
                    )

    return out_path


