"""3.9.7 / 3.9.8 宽表汇总 — COM 版本

数据读取走 Excel COM；合并区域走 openpyxl 缓存；
下游分桶/冲突检测/写入目标逻辑与 pandas 版对齐。
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import time

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
from .rule_engine import (
    _kw_match,
    extract_cells_from_arrays,
    pick_rules_for_workbook,
    compute_wide_col_paths_from_arrays,
)
from .wide_397_398 import (
    FIXED_COLS,
    FIXED_COLS_NO_ROWPATH,
    _row_no_from_addr,
    _is_wide_rule_compatible,
    _disambiguate_headers,
    _merged_value,
)


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


def run_wide_summary_com(rules: list[TimelineRule], path_maps: list[PathMapRule],
                         g: GlobalConfig, row_suffix_enabled: bool,
                         source_paths: list[Path] | None = None) -> Path | None:
    tag = "3_9_7" if row_suffix_enabled else "3_9_8"
    log = get_logger(f"{tag}_wide_com", g.log_dir)
    if source_paths is None:
        sources = list_source_files(g.input_dir, g.source_exts)
    else:
        ext_ok = {e.lower() for e in g.source_exts} | {".xls"}
        sources = [p for p in source_paths if p.exists() and p.suffix.lower() in ext_ok]
    log.info(
        f"[COM] 开始宽表汇总({tag}): 源 {len(sources)}，规则 {len(rules)}，"
        f"行头后缀={row_suffix_enabled}"
    )

    by_rule: dict[str, dict[tuple, dict[str, object]]] = {}
    set_values_by_rule: dict[str, dict[tuple, list[object]]] = {}
    rule_output_rowpath: dict[str, bool] = {}
    rule_set_headers: dict[str, list[str]] = {}
    rule_set_addrs: dict[str, list[str]] = {}
    col_order: dict[str, list[str]] = {}
    col_seen: dict[str, set[str]] = {}
    conflict = 0
    rule_stats: dict[str, dict[str, int]] = {}
    for r in rules:
        rule_stats[r.name] = {"wb_miss": 0, "sheet_miss": 0, "cells": 0}

    app = start_excel_app()
    try:
        for src in sources:
            src_wb = None
            try:
                log.info(f"[COM] 正在打开源文件: {src}")
                src_wb = open_readonly(app, src)
                log.info(f"[COM] 已打开源文件: {src.name}")
            except Exception as e:
                log.warning(f"打开源失败 {src.name}: {e}")
                continue
            try:
                all_sheets = list_sheet_names_com(src_wb)
                log.info(f"[COM] 工作表数量 {len(all_sheets)}: {src.name}")
                sheet_cache: dict[str, tuple[ArrayLike, tuple[tuple[int, int, int, int], ...]]] = {}
                ws_cache: dict[str, object] = {}

                matched_by_rule: dict[str, list[str]] = {}
                for rule in pick_rules_for_workbook(rules, src.name, g.timeline_rule_match_mode):
                    if not _kw_match(src.name, rule.wb_keyword):
                        rule_stats[rule.name]["wb_miss"] += 1
                        continue
                    if not _is_wide_rule_compatible(rule, log):
                        continue
                    matched = [sn for sn in all_sheets if _kw_match(sn, rule.sheet_keyword)]
                    matched_by_rule[rule.name] = matched

                for rule in pick_rules_for_workbook(rules, src.name, g.timeline_rule_match_mode):
                    if getattr(rule, "set_parse_error", ""):
                        log.warning("规则[%s] set区域配置非法，已跳过: %s", rule.name, rule.set_parse_error)
                        continue
                    rule_t0 = time.perf_counter()
                    cells_before = rule_stats[rule.name]["cells"]
                    matched_sheets = matched_by_rule.get(rule.name, [])
                    hit_sheet = len(matched_sheets) > 0
                    if matched_sheets:
                        log.info(
                            "[COM] 规则[%s] 命中工作表 %s 个: %s",
                            rule.name, len(matched_sheets), ",".join(matched_sheets[:8])
                        )
                    for sn in matched_sheets:
                        hit_sheet = True
                        if sn not in sheet_cache:
                            log.info(f"[COM] 读取工作表: {src.name}::{sn}")
                            ws = src_wb.Worksheets(sn)
                            ws_cache[sn] = ws
                            sheet_cache[sn] = (read_sheet_values(ws), get_merge_ranges_com(ws))
                        ws = ws_cache[sn]
                        df, merge_ranges = sheet_cache[sn]
                        if rule.name not in rule_set_headers:
                            aliases = [x[0] for x in getattr(rule, "set_items", [])]
                            headers = _disambiguate_headers(aliases) if aliases else []
                            addrs = [x[1] for x in getattr(rule, "set_items", [])]
                            rule_set_headers[rule.name] = headers
                            rule_set_addrs[rule.name] = addrs
                        set_headers = rule_set_headers.get(rule.name, [])
                        set_addrs = rule_set_addrs.get(rule.name, [])
                        set_vals: list[object] = []
                        if set_headers and set_addrs:
                            for addr in set_addrs:
                                col = 0
                                row = 0
                                i = 0
                                while i < len(addr) and addr[i].isalpha():
                                    col = col * 26 + (ord(addr[i].upper()) - 64)
                                    i += 1
                                if i < len(addr):
                                    row = int(addr[i:])
                                set_vals.append(_merged_value(df, merge_ranges, row - 1, col - 1))
                        rows_n, cols_n = df.shape
                        coords, _sig = _build_header_coords(rule, rows_n, cols_n)
                        cell_text_overrides = read_merged_aware_texts(ws, coords)
                        cells = extract_cells_from_arrays(
                            df, merge_ranges, src, rule, sn, path_maps, row_suffix_enabled,
                            cell_text_overrides=cell_text_overrides,
                        )
                        ch_rows = [r - 1 for r in rule.col_header_rows if 0 < r <= rows_n]
                        row_header_specified = getattr(rule, "row_header_col_specified", True)
                        rh_cols_cfg = getattr(rule, "row_header_cols", None)
                        if row_header_specified:
                            if rh_cols_cfg:
                                rh_cols = [c - 1 for c in rh_cols_cfg if c > 0]
                            else:
                                rh_cols = [rule.row_header_col - 1]
                        else:
                            rh_cols = []
                        if ch_rows:
                            r_start = rule.data_row_start - 1 if rule.data_row_start else max(ch_rows) + 1
                            c_start = (
                                rule.data_col_start - 1
                                if rule.data_col_start
                                else ((max(rh_cols) + 1) if row_header_specified and rh_cols else 0)
                            )
                            c_end = (rule.data_col_end - 1) if rule.data_col_end else cols_n - 1
                            c_start = max(c_start, 0)
                            c_end = min(c_end, cols_n - 1)
                            if r_start <= rows_n - 1 and c_start <= c_end:
                                _, ordered_cols = compute_wide_col_paths_from_arrays(
                                    df,
                                    merge_ranges,
                                    rule,
                                    sn,
                                    path_maps,
                                    src.name,
                                    c_start=c_start,
                                    c_end=c_end,
                                    ch_rows=ch_rows,
                                    cell_text_overrides=cell_text_overrides,
                                )
                                cs = col_seen.setdefault(rule.name, set())
                                co = col_order.setdefault(rule.name, [])
                                for cp in ordered_cols:
                                    if cp not in cs:
                                        cs.add(cp)
                                        co.append(cp)
                        if cells:
                            rule_stats[rule.name]["cells"] += len(cells)
                        for ec in cells:
                            bucket = by_rule.setdefault(ec.rule_name, {})
                            row_no = _row_no_from_addr(ec.cell_addr)
                            include_row_path = getattr(rule, "row_header_col_specified", True)
                            rule_output_rowpath.setdefault(ec.rule_name, include_row_path)
                            if row_suffix_enabled:
                                row_unique = ec.row_path
                            else:
                                row_unique = (
                                    f"{ec.row_path}#R{row_no}" if row_no else ec.row_path
                                )
                            row_path_out = ec.row_path if include_row_path else ""
                            key = (ec.source_file, ec.sheet_name, ec.data_date,
                                   row_path_out, row_unique)
                            row_dict = bucket.setdefault(key, {})
                            if set_headers and key not in set_values_by_rule.setdefault(ec.rule_name, {}):
                                set_values_by_rule.setdefault(ec.rule_name, {})[key] = list(set_vals)
                            if ec.col_path in row_dict:
                                conflict += 1
                                log.warning(
                                    f"宽表冲突: 规则={ec.rule_name} 簿={ec.source_file} "
                                    f"表={ec.sheet_name} 日期={ec.data_date} 行={ec.row_path} "
                                    f"列={ec.col_path} 当前={ec.cell_addr} 保留首值"
                                )
                                continue
                            row_dict[ec.col_path] = ec.value
                    if not hit_sheet:
                        rule_stats[rule.name]["sheet_miss"] += 1
                    log.info(
                        "[COM] 规则[%s] 用时 %.2fs, 产出cells=%s",
                        rule.name, time.perf_counter() - rule_t0, rule_stats[rule.name]["cells"] - cells_before
                    )
            finally:
                safe_close(src_wb)
    finally:
        safe_quit(app)

    sheets: dict[str, pd.DataFrame] = {}
    total_rows = 0
    for rule_name, bucket in by_rule.items():
        include_row_path = rule_output_rowpath.get(rule_name, True)
        set_headers = rule_set_headers.get(rule_name, [])
        if include_row_path:
            fixed_cols = ["工作簿名", "工作表名", "数据日期"] + set_headers + ["行头路径"]
        else:
            fixed_cols = ["工作簿名", "工作表名", "数据日期"] + set_headers
        cols = fixed_cols + col_order[rule_name]
        rows: list[list] = []
        for (wb, sn, dt, rp, _ru), m in bucket.items():
            set_vals = set_values_by_rule.get(rule_name, {}).get((wb, sn, dt, rp, _ru), [""] * len(set_headers))
            if include_row_path:
                row = [wb, sn, dt] + list(set_vals) + [rp] + [m.get(cp, "") for cp in col_order[rule_name]]
            else:
                row = [wb, sn, dt] + list(set_vals) + [m.get(cp, "") for cp in col_order[rule_name]]
            rows.append(row)
        df_out = pd.DataFrame(rows, columns=cols)
        sheets[rule_name or "未命名规则"] = df_out
        total_rows += len(df_out)

    if not sheets or total_rows == 0:
        for rule_name, st in rule_stats.items():
            log.info("规则[%s] wb_miss=%s sheet_miss=%s cells=%s",
                     rule_name, st["wb_miss"], st["sheet_miss"], st["cells"])
        log.info("[COM] 完成: 无匹配结果，不生成输出文件")
        return None

    if row_suffix_enabled:
        menu_name = "宽表规则汇总（行头加后缀）"
    else:
        menu_name = "宽表规则汇总（行头不加后缀）"
    out_path = g.output_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{menu_name}_COM.xlsx"
    write_workbook(out_path, sheets)
    log.info(f"[COM] 完成: 规则数 {len(by_rule)}，输出行 {total_rows}，冲突 {conflict} → {out_path}")

    # 写入目标簿（与 pandas 版一致）
    for rule in rules:
        if not rule.target_write_enabled:
            continue
        if not (rule.target_wb_path and rule.target_sheet):
            continue
        if rule.name not in sheets:
            continue
        df_t = sheets[rule.name]
        if df_t.empty:
            continue
        header = list(df_t.columns)
        rs = df_t.values.tolist()
        key_idx = list(range(len(header)))
        set_headers = rule_set_headers.get(rule.name, [])
        if getattr(rule, "row_header_col_specified", True):
            required_prefix = ["工作簿名", "工作表名", "数据日期"] + set_headers + ["行头路径"]
        else:
            required_prefix = ["工作簿名", "工作表名", "数据日期"] + set_headers
        stat = append_to_target(
            rule.target_wb_path, rule.target_sheet,
            header, rs, dedup_key_idx=key_idx,
            required_prefix=required_prefix, allow_extend=True,
            dedup_seq_prefix_idx=None,
        )
        if stat.get("written", 0) == 0 and stat.get("skipped_reason") == "header_mismatch":
            log.warning(f"跳过写入目标 {rule.target_wb_path}::{rule.target_sheet}: 表头不匹配（保护旧数据）")
        else:
            log.info(
                "追加到目标 %s::%s 输入=%s 批内去重后=%s 目标去重后新增=%s",
                rule.target_wb_path, rule.target_sheet,
                stat.get("input_rows", 0),
                stat.get("batch_dedup_rows", 0),
                stat.get("existing_filtered_rows", 0),
            )

    return out_path
