from __future__ import annotations
from datetime import datetime
from pathlib import Path
import pandas as pd

from .config_xlsx import GlobalConfig, TimelineRule
from .io_excel import list_sheet_names, read_sheet_2d, write_workbook
from .rule_engine import match_all_keywords, _val_str
from .logger import get_logger

# 输出 12 列（前 10 列结构对齐 路径标准化映射 Sheet）
RESULT_COLS = [
    "是否启用", "映射名称", "适用规则名",
    "工作簿关键字", "工作表关键字",
    "作用对象", "匹配方式",
    "原始路径", "标准路径", "备注",
    "源文件", "工作表",
]


def _resolve_range(df, rule: TimelineRule):
    """返回 (rh_col, ch_rows, r_start, r_end, c_start, c_end) 全部 0-based。"""
    rows, cols = df.shape
    rh_col = rule.row_header_col - 1
    ch_rows = [r - 1 for r in rule.col_header_rows if r > 0]
    r_start = (rule.data_row_start - 1) if rule.data_row_start else (max(ch_rows) + 1 if ch_rows else 0)
    r_end = (rule.data_row_end - 1) if rule.data_row_end else (rows - 1)
    c_start = (rule.data_col_start - 1) if rule.data_col_start else (rh_col + 1)
    c_end = (rule.data_col_end - 1) if rule.data_col_end else (cols - 1)
    return rh_col, ch_rows, r_start, min(r_end, rows - 1), c_start, min(c_end, cols - 1)


def _build_col_path(df, c: int, ch_rows: list[int]) -> str:
    parts = []
    for r in ch_rows:
        if 0 <= r < df.shape[0] and 0 <= c < df.shape[1]:
            v = _val_str(df.iat[r, c])
            if v:
                parts.append(v)
    return "_".join(parts)


def _build_row_path(df, r: int, rh_col: int) -> str:
    if 0 <= r < df.shape[0] and 0 <= rh_col < df.shape[1]:
        return _val_str(df.iat[r, rh_col])
    return ""


def _collect_path_set(df, rule: TimelineRule) -> tuple[set[str], set[str]]:
    rh_col, ch_rows, r_start, r_end, c_start, c_end = _resolve_range(df, rule)
    rows = set()
    for r in range(r_start, r_end + 1):
        p = _build_row_path(df, r, rh_col)
        if p:
            rows.add(p)
    cols = set()
    for c in range(c_start, c_end + 1):
        p = _build_col_path(df, c, ch_rows)
        if p:
            cols.add(p)
    return rows, cols


def _position_diff_rows(tdf, sdf, rule: TimelineRule) -> list[tuple[str, str, str]]:
    """返回 [(target='行头/列头', source_path, template_path), ...]，仅差异。"""
    out: list[tuple[str, str, str]] = []
    rh_col, ch_rows, _, _, c_start, _ = _resolve_range(tdf, rule)
    # 用最大边界（含 source）
    t_rows, t_cols = tdf.shape
    s_rows, s_cols = sdf.shape

    r_start = (rule.data_row_start - 1) if rule.data_row_start else (max(ch_rows) + 1 if ch_rows else 0)
    r_end_t = (rule.data_row_end - 1) if rule.data_row_end else (t_rows - 1)
    r_end_s = (rule.data_row_end - 1) if rule.data_row_end else (s_rows - 1)
    r_end = max(min(r_end_t, t_rows - 1), min(r_end_s, s_rows - 1))

    c_end_t = (rule.data_col_end - 1) if rule.data_col_end else (t_cols - 1)
    c_end_s = (rule.data_col_end - 1) if rule.data_col_end else (s_cols - 1)
    c_end = max(min(c_end_t, t_cols - 1), min(c_end_s, s_cols - 1))

    # 列方向
    for c in range(c_start, c_end + 1):
        tp = _build_col_path(tdf, c, ch_rows)
        sp = _build_col_path(sdf, c, ch_rows)
        if tp != sp:
            out.append(("列头", sp, tp))
    # 行方向
    for r in range(r_start, r_end + 1):
        tp = _build_row_path(tdf, r, rh_col)
        sp = _build_row_path(sdf, r, rh_col)
        if tp != sp:
            out.append(("行头", sp, tp))
    return out


def _path_diff_rows(tdf, sdf, rule: TimelineRule) -> list[tuple[str, str, str]]:
    t_rows, t_cols = _collect_path_set(tdf, rule)
    s_rows, s_cols = _collect_path_set(sdf, rule)
    out: list[tuple[str, str, str]] = []
    # 行头：源独有 / 模板独有
    for p in sorted(s_rows - t_rows):
        out.append(("行头", p, ""))
    for p in sorted(t_rows - s_rows):
        out.append(("行头", "", p))
    for p in sorted(s_cols - t_cols):
        out.append(("列头", p, ""))
    for p in sorted(t_cols - s_cols):
        out.append(("列头", "", p))
    return out


def _resolve_template_sheet(template_sheets: list[str], source_sheet_name: str) -> str | None:
    """同名优先；否则若仅 1 张匹配则用之；否则 None。"""
    for ts in template_sheets:
        if ts == source_sheet_name:
            return ts
    if len(template_sheets) == 1:
        return template_sheets[0]
    return None


def run_compare(template_path: Path, source_paths: list[Path],
                rules: list[TimelineRule], g: GlobalConfig) -> Path:
    """3.9.3 表头路径比对（混合：位置后路径）。
    - 工作簿/工作表关键字按 ; 分，全部包含才命中（AND）
    - 输出 12 列，与 路径标准化映射 Sheet 前 10 列结构一致
    - 仅输出差异行；位置差异 备注=按位置；该 sheet 有位置差异时再补做路径差异 备注=按路径
    """
    log = get_logger("3_9_3_compare", g.log_dir)
    log.info(f"模板: {template_path.name}, 源 {len(source_paths)} 个, 规则 {len(rules)} 条")

    if not template_path.exists():
        raise FileNotFoundError(f"模板不存在: {template_path}")
    template_sheets_all = list_sheet_names(template_path)

    # 护栏：剔除与模板同路径的源文件，避免误把模板当源重复比对
    tmpl_resolved = template_path.resolve()
    filtered_sources: list[Path] = []
    skipped_same_as_template = 0
    for p in source_paths:
        try:
            if p.resolve() == tmpl_resolved:
                skipped_same_as_template += 1
                log.warning(f"跳过源文件（与模板同路径）: {p}")
                continue
        except Exception:
            pass
        filtered_sources.append(p)
    if skipped_same_as_template:
        log.info(f"已跳过 {skipped_same_as_template} 个与模板同路径的源文件")

    summary_rows: list[list] = []
    diff_rows: list[list] = []
    total_pos = total_path = 0

    for src in filtered_sources:
        try:
            src_sheet_names = list_sheet_names(src)
        except Exception as e:
            log.warning(f"打开源失败 {src.name}: {e}")
            summary_rows.append([src.name, "", 0, 0, f"打开失败: {e}"])
            continue

        per_src_pos = per_src_path = 0
        for rule in rules:
            # 工作簿关键字 AND 匹配
            if not match_all_keywords(src.name, rule.wb_keyword):
                continue
            # 源端匹配的 sheet 列表
            src_match_sheets = [s for s in src_sheet_names
                                if match_all_keywords(s, rule.sheet_keyword)]
            if not src_match_sheets:
                continue
            # 模板端匹配的 sheet 列表
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
                tdf = read_sheet_2d(template_path, t_sheet)
                sdf = read_sheet_2d(src, s_sheet)

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

                # 仅当本表有位置差异时，再补做路径比对（混合模式）
                if pos:
                    path = _path_diff_rows(tdf, sdf, rule)
                    for target, sp, tp in path:
                        diff_rows.append([
                            "", "", rule.name,
                            rule.wb_keyword, rule.sheet_keyword,
                            target, "精确",
                            sp, tp, "按路径",
                            src.name, s_sheet,
                        ])
                    per_src_path += len(path)

        total_pos += per_src_pos
        total_path += per_src_path
        status = "通过" if (per_src_pos == 0 and per_src_path == 0) else "存在差异"
        summary_rows.append([src.name, status, per_src_pos, per_src_path, ""])

    out_path = g.output_dir / f"compare_3_9_3_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    diff_df = pd.DataFrame(diff_rows, columns=RESULT_COLS)
    summary_df = pd.DataFrame(summary_rows,
                              columns=["源文件", "状态", "按位置差异数", "按路径差异数", "备注"])
    write_workbook(out_path, {"表头比对结果": diff_df, "汇总": summary_df})
    log.info(f"完成: 源 {len(source_paths)}, 按位置 {total_pos}, 按路径 {total_path} → {out_path}")
    return out_path
