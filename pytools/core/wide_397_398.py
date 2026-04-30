from __future__ import annotations
from datetime import datetime
from pathlib import Path
import re
import pandas as pd

from .config_xlsx import GlobalConfig, TimelineRule, PathMapRule
from .rule_engine import extract_cells, _kw_match
from .io_excel import list_source_files, list_sheet_names, write_workbook, append_to_target, clear_sheet_cache
from .logger import get_logger

FIXED_COLS = ["工作簿名", "工作表名", "数据日期", "行头路径"]
FIXED_COLS_NO_ROWPATH = ["工作簿名", "工作表名", "数据日期"]


def _row_no_from_addr(addr: str) -> str:
    # 单元格地址形如 'AB123'，字母在前数字在后。
    # 原先 re.search 每次都触发正则引擎，换成手写尾部数字提取，快 5~10 倍
    if not addr:
        return ""
    i = len(addr) - 1
    while i >= 0 and addr[i].isdigit():
        i -= 1
    return addr[i + 1:] if i < len(addr) - 1 else ""


def _is_wide_rule_compatible(rule: TimelineRule, log) -> bool:
    row_header_specified = getattr(rule, "row_header_col_specified", True)
    if row_header_specified:
        return True
    if rule.data_row_start <= 0 or rule.data_col_start <= 0:
        log.warning(
            "规则[%s] 行头列为空时，数据起始行/数据起始列必须填写；已跳过（数据起始行=%s, 数据起始列=%s）",
            rule.name, rule.data_row_start, rule.data_col_start
        )
        return False
    if rule.required_row_headers:
        log.info("规则[%s] 行头列为空，已自动忽略“必含行头”校验。", rule.name)
    return True


def run_wide_summary(rules: list[TimelineRule], path_maps: list[PathMapRule],
                     g: GlobalConfig, row_suffix_enabled: bool, source_paths: list[Path] | None = None) -> Path | None:
    tag = "3_9_7" if row_suffix_enabled else "3_9_8"
    log = get_logger(f"{tag}_wide", g.log_dir)
    if source_paths is None:
        sources = list_source_files(g.input_dir, g.source_exts)
    else:
        ext_ok = {e.lower() for e in g.source_exts} | {".xls"}
        sources = [p for p in source_paths if p.exists() and p.suffix.lower() in ext_ok]
    log.info(f"开始宽表汇总({tag}): 源 {len(sources)}，规则 {len(rules)}，行头后缀={row_suffix_enabled}")

    # rule_name -> dict[(wb,sheet,date,row_path)] -> dict[col_path] -> value
    by_rule: dict[str, dict[tuple, dict[str, object]]] = {}
    rule_output_rowpath: dict[str, bool] = {}
    # 列顺序按首见
    col_order: dict[str, list[str]] = {}
    col_seen: dict[str, set[str]] = {}
    conflict = 0
    rule_stats: dict[str, dict[str, int]] = {}
    for r in rules:
        rule_stats[r.name] = {"wb_miss": 0, "sheet_miss": 0, "cells": 0}

    for src in sources:
        try:
            for rule in rules:
                if not _kw_match(src.name, rule.wb_keyword):
                    rule_stats[rule.name]["wb_miss"] += 1
                    continue
                if not _is_wide_rule_compatible(rule, log):
                    continue
                hit_sheet = False
                for sn in list_sheet_names(src):
                    if not _kw_match(sn, rule.sheet_keyword):
                        continue
                    hit_sheet = True
                    cells = extract_cells(src, rule, sn, path_maps, row_suffix_enabled)
                    if cells:
                        rule_stats[rule.name]["cells"] += len(cells)
                    for ec in cells:
                        bucket = by_rule.setdefault(ec.rule_name, {})
                        cs = col_seen.setdefault(ec.rule_name, set())
                        co = col_order.setdefault(ec.rule_name, [])
                        if ec.col_path not in cs:
                            cs.add(ec.col_path)
                            co.append(ec.col_path)
                        # 3.9.8（行头不加后缀）时，行头文本可能重复；
                        # 这里内部增加源行号，避免不同数据行被错误压并为同一行。
                        row_no = _row_no_from_addr(ec.cell_addr)
                        include_row_path = getattr(rule, "row_header_col_specified", True)
                        rule_output_rowpath.setdefault(ec.rule_name, include_row_path)
                        if row_suffix_enabled:
                            row_unique = ec.row_path
                        else:
                            row_unique = f"{ec.row_path}#R{row_no}" if row_no else ec.row_path
                        row_path_out = ec.row_path if include_row_path else ""
                        key = (ec.source_file, ec.sheet_name, ec.data_date, row_path_out, row_unique)
                        row_dict = bucket.setdefault(key, {})
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
        except Exception as e:
            log.warning(f"处理 {src.name} 失败: {e}")
            if g.error_policy == "fail_fast":
                raise
        finally:
            clear_sheet_cache()

    sheets: dict[str, pd.DataFrame] = {}
    total_rows = 0
    for rule_name, bucket in by_rule.items():
        include_row_path = rule_output_rowpath.get(rule_name, True)
        fixed_cols = FIXED_COLS if include_row_path else FIXED_COLS_NO_ROWPATH
        cols = fixed_cols + col_order[rule_name]
        rows: list[list] = []
        for (wb, sn, dt, rp, _ru), m in bucket.items():
            if include_row_path:
                row = [wb, sn, dt, rp] + [m.get(cp, "") for cp in col_order[rule_name]]
            else:
                row = [wb, sn, dt] + [m.get(cp, "") for cp in col_order[rule_name]]
            rows.append(row)
        df = pd.DataFrame(rows, columns=cols)
        sheets[rule_name or "未命名规则"] = df
        total_rows += len(df)

    if not sheets or total_rows == 0:
        for rule_name, st in rule_stats.items():
            log.info("规则[%s] wb_miss=%s sheet_miss=%s cells=%s",
                     rule_name, st["wb_miss"], st["sheet_miss"], st["cells"])
        log.info("完成: 无匹配结果，不生成输出文件")
        return None
    if row_suffix_enabled:
        menu_name = "宽表规则汇总（行头加后缀）"
    else:
        menu_name = "宽表规则汇总（行头不加后缀）"
    out_path = g.output_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{menu_name}.xlsx"
    write_workbook(out_path, sheets)
    log.info(f"完成: 规则数 {len(by_rule)}，输出行 {total_rows}，冲突 {conflict} → {out_path}")

    # 写目标簿
    for rule in rules:
        if not rule.target_write_enabled:
            continue
        if not (rule.target_wb_path and rule.target_sheet):
            continue
        if rule.name not in sheets:
            continue
        df = sheets[rule.name]
        if df.empty:
            continue
        header = list(df.columns)
        rs = df.values.tolist()
        # 宽表写目标按整行全列去重，避免仅按前四列导致误判。
        key_idx = list(range(len(header)))
        required_prefix = FIXED_COLS if getattr(rule, "row_header_col_specified", True) else FIXED_COLS_NO_ROWPATH
        stat = append_to_target(rule.target_wb_path, rule.target_sheet,
                                header, rs, dedup_key_idx=key_idx,
                                required_prefix=required_prefix, allow_extend=True,
                                dedup_seq_prefix_idx=None)
        if stat.get("written", 0) == 0 and stat.get("skipped_reason") == "header_mismatch":
            log.warning(f"跳过写入目标 {rule.target_wb_path}::{rule.target_sheet}: 表头不匹配（保护旧数据）")
        else:
            log.info(
                "追加到目标 %s::%s 输入=%s 批内去重后=%s 目标去重后新增=%s",
                rule.target_wb_path, rule.target_sheet,
                stat.get("input_rows", 0),
                stat.get("batch_dedup_rows", 0),
                stat.get("existing_filtered_rows", 0)
            )

    return out_path
