"""6-1 / 6-2 汇总工具 — COM 版本

数据读取走 Excel COM；模板批注解析仍用 openpyxl（metadata 解析很轻）；
输出仍用 openpyxl 写出 xlsx，与既有 COM 模块风格一致。
若模板/源为 .xls，运行时用 COM 临时另存为 .xlsx 后再喂给 openpyxl 解析。
"""
from __future__ import annotations

import shutil
import tempfile
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from .com_engine import (
    ArrayLike,
    open_readonly,
    read_sheet_values,
    safe_close,
    safe_quit,
    start_excel_app,
)
from .logger import get_logger
from .summary_321_322 import (
    TemplateSheetSpec,
    _col2num,
    _disambiguate_headers,
    _extract_comments,
    _extract_regions,
    _extract_set_info,
    _build_col_headers,
    _split_addr,
    _supported,
)


def _supported_com(path: Path) -> bool:
    """COM 分支放宽：除 xlsx/xlsm 外，额外支持 xls（COM 引擎可直接打开）。"""
    return path.suffix.lower() in (".xlsx", ".xlsm", ".xls")


_XLSX_FILE_FORMAT = 51  # xlOpenXMLWorkbook


def _ensure_openpyxl_compatible(path: Path, app, cache: dict[str, Path],
                                 temp_dir: Path) -> Path:
    """xlsx/xlsm 直接返回；xls 用 COM 转存为 temp_dir 下的 .xlsx，按原路径缓存复用。

    供 openpyxl-only 的解析路径（模板批注、合并区域）使用。
    """
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        return path
    key = str(path.resolve())
    cached = cache.get(key)
    if cached is not None:
        return cached
    temp_dir.mkdir(parents=True, exist_ok=True)
    tag = abs(hash(key)) & 0xFFFFFF
    out = temp_dir / f"{path.stem}_{tag:06x}.xlsx"
    wb = app.Workbooks.Open(str(path.resolve()), ReadOnly=True, UpdateLinks=0)
    try:
        wb.SaveAs(str(out.resolve()), FileFormat=_XLSX_FILE_FORMAT)
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
    cache[key] = out
    return out


def _gcv_from_array(df: ArrayLike, merge_ranges, r: int, c: int):
    """1-based (r, c)。落在合并区域则返回左上角值，否则返回单元格自身值。
    与 summary_321_322._gcv 的 openpyxl 行为对齐。"""
    r0, c0 = r - 1, c - 1
    rows_n, cols_n = df.shape
    if r0 < 0 or c0 < 0 or r0 >= rows_n or c0 >= cols_n:
        return None
    for min_r, max_r, min_c, max_c in merge_ranges:
        if min_r <= r0 <= max_r and min_c <= c0 <= max_c:
            return df.iat[min_r, min_c]
    return df.iat[r0, c0]


def run_summary_by_usedrange_com(
    source_paths: list[Path], output_dir: Path, log_dir: Path
) -> dict[str, object]:
    """6-1 按使用区域汇总（COM 引擎）。"""
    log = get_logger("3_2_1_summary_usedrange_com", log_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    valid = [p for p in source_paths if p.exists() and _supported_com(p)]
    if not valid:
        return {"saved": False, "rows": 0, "path": ""}

    out_wb = Workbook()
    ws_out = out_wb.active
    ws_out.title = "汇总"

    max_cols = 0
    rows_written = 0
    row_idx = 2
    files_hit = 0

    print(f"→ [COM] 开始汇总，共 {len(valid)} 个文件...")
    app = start_excel_app()
    try:
        for idx, p in enumerate(valid, start=1):
            src_wb = None
            try:
                src_wb = open_readonly(app, p)
            except Exception as e:
                log.warning("打开失败 source=%s err=%s", p, e)
                continue
            print(f"  [{idx}/{len(valid)}] 打开: {p.name}")
            try:
                files_hit += 1
                wb_name = p.stem
                for ws in src_wb.Worksheets:
                    ur = ws.UsedRange
                    sr = int(ur.Row)
                    sc = int(ur.Column)
                    er = sr + int(ur.Rows.Count) - 1
                    ec = sc + int(ur.Columns.Count) - 1
                    if er < sr or ec < sc:
                        continue
                    print(f"     - 处理工作表: {ws.Name}")
                    # 一次性整块读取（从 A 列到 ec 列，左侧空白自动 None）
                    rng = ws.Range(ws.Cells(sr, 1), ws.Cells(er, ec))
                    v = rng.Value2
                    if v is None:
                        continue
                    if not isinstance(v, tuple):
                        v = ((v,),)
                    for offset, row in enumerate(v):
                        r = sr + offset
                        # row 长度 = ec（因为 Range 起点是第 1 列）
                        for c_off, val in enumerate(row, start=1):
                            ws_out.cell(row_idx, 3 + c_off).value = val
                        ws_out.cell(row_idx, 1).value = wb_name
                        ws_out.cell(row_idx, 2).value = ws.Name
                        ws_out.cell(row_idx, 3).value = r
                        row_idx += 1
                        rows_written += 1
                    if ec > max_cols:
                        max_cols = ec
                print(f"     完成: {p.name}")
            finally:
                safe_close(src_wb)
    finally:
        safe_quit(app)

    if rows_written == 0:
        out_wb.close()
        return {"saved": False, "rows": 0, "path": ""}

    # 数据写完后再写表头
    ws_out.cell(1, 1).value = "工作簿"
    ws_out.cell(1, 2).value = "工作表"
    ws_out.cell(1, 3).value = "行号"
    for c in range(1, max_cols + 1):
        ws_out.cell(1, 3 + c).value = f"列{c}"
    for c in range(1, 4 + max_cols):
        ws_out.cell(1, c).font = ws_out.cell(1, c).font.copy(bold=True)
    ws_out.freeze_panes = "A2"

    out_path = output_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_按使用区域汇总（全量）_COM.xlsx"
    print("→ 正在保存结果文件...")
    out_wb.save(out_path)
    out_wb.close()
    log.info("[COM] done wb=%s rows=%s output=%s", files_hit, rows_written, out_path)
    return {"saved": True, "rows": rows_written, "path": str(out_path)}


def _parse_template_specs(template_path: Path) -> dict[str, TemplateSheetSpec]:
    """模板批注解析仍用 openpyxl（metadata 读取轻）。"""
    tmpl_wb = load_workbook(template_path, data_only=False, read_only=False,
                            keep_vba=(template_path.suffix.lower() == ".xlsm"))
    specs: dict[str, TemplateSheetSpec] = {}
    try:
        for ws in tmpl_wb.worksheets:
            cm = _extract_comments(ws)
            if not cm:
                continue
            row_regs = _extract_regions(cm, "行区域")
            col_regs = _extract_regions(cm, "列区域")
            set_names, set_addrs = _extract_set_info(cm)
            if not row_regs or not col_regs:
                continue
            headers = _build_col_headers(ws, col_regs)
            specs[ws.title] = TemplateSheetSpec(
                name=ws.title,
                comment_map=cm,
                row_regions=row_regs,
                col_regions=col_regs,
                set_names=set_names,
                set_addrs=set_addrs,
                col_headers=headers,
            )
    finally:
        tmpl_wb.close()
    return specs


def _read_src_sheet_via_com(src_ws, accessed_rows: set[int],
                             accessed_cols: set[int]) -> tuple[ArrayLike, tuple]:
    """读源 sheet 到 ArrayLike，并通过 openpyxl 取合并区域。
    accessed_* 用于决定 Range 边界，避免读多余数据。"""
    if not accessed_rows or not accessed_cols:
        return ArrayLike([]), tuple()
    ur = src_ws.UsedRange
    last_used_row = int(ur.Row) + int(ur.Rows.Count) - 1
    last_used_col = int(ur.Column) + int(ur.Columns.Count) - 1
    max_r = max(max(accessed_rows), last_used_row)
    max_c = max(max(accessed_cols), last_used_col)
    if max_r < 1 or max_c < 1:
        return ArrayLike([]), tuple()
    rng = src_ws.Range(src_ws.Cells(1, 1), src_ws.Cells(max_r, max_c))
    v = rng.Value2
    if v is None:
        return ArrayLike([]), tuple()
    if not isinstance(v, tuple):
        v = ((v,),)
    return ArrayLike([list(row) for row in v]), tuple()


def _load_merge_for_sheet(src_path: Path, sheet_name: str) -> tuple:
    """惰性加载某 sheet 的合并区域（openpyxl + LRU 缓存）。"""
    from .rule_engine import _load_merge_ranges
    return _load_merge_ranges(str(src_path), sheet_name)


def run_summary_by_comment_com(
    template_path: Path, source_paths: list[Path], output_dir: Path, log_dir: Path
) -> dict[str, object]:
    """6-2 按批注汇总（COM 引擎）。"""
    log = get_logger("3_2_2_summary_comment_com", log_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not template_path.exists() or not _supported_com(template_path):
        raise ValueError("模板必须为 xlsx/xlsm/xls 且必须存在。")

    xls_temp_dir = Path(tempfile.mkdtemp(prefix="3_2_2_xls_"))
    xls_cache: dict[str, Path] = {}
    app = start_excel_app()
    try:
        template_for_parse = _ensure_openpyxl_compatible(
            template_path, app, xls_cache, xls_temp_dir
        )
        specs = _parse_template_specs(template_for_parse)
        if not specs:
            raise ValueError(
                "模板未识别到可用批注区域（需包含 行区域N/#N 与 列区域N/#N）。"
            )

        # 按 (spec, 位置) 展开，重复名追加 _N 后缀
        set_layout: list[tuple[str, int]] = []
        col_layout: list[tuple[str, int]] = []
        raw_set_headers: list[str] = []
        raw_col_headers: list[str] = []
        for spec_key, sp in specs.items():
            for pos, n in enumerate(sp.set_names):
                set_layout.append((spec_key, pos))
                raw_set_headers.append(n)
            for pos, h in enumerate(sp.col_headers):
                col_layout.append((spec_key, pos))
                raw_col_headers.append(h)
        set_headers = _disambiguate_headers(raw_set_headers)
        col_headers_flat = _disambiguate_headers(raw_col_headers)

        out_wb = Workbook()
        ws_out = out_wb.active
        ws_out.title = "汇总"
        headers = ["工作簿", "工作表"] + set_headers + col_headers_flat + ["行号"]
        for i, h in enumerate(headers, 1):
            ws_out.cell(1, i).value = h
            ws_out.cell(1, i).font = ws_out.cell(1, i).font.copy(bold=True)

        set_loc: dict[tuple[str, int], int] = {
            key: 3 + i for i, key in enumerate(set_layout)
        }
        col_loc: dict[tuple[str, int], int] = {
            key: 3 + len(set_layout) + i for i, key in enumerate(col_layout)
        }

        fallback_spec = specs.get("模板")
        row_idx = 2
        rows_written = 0
        files_hit = 0

        valid_sources = [p for p in source_paths if p.exists() and _supported_com(p)]
        for p in valid_sources:
            src_wb = None
            try:
                src_wb = open_readonly(app, p)
            except Exception as e:
                log.warning("打开失败 source=%s err=%s", p, e)
                continue
            try:
                files_hit += 1
                wb_name = p.stem
                # .xls 走 COM 转 .xlsx 后供 openpyxl 取合并区域；其他直接用原路径
                merge_src_path = _ensure_openpyxl_compatible(
                    p, app, xls_cache, xls_temp_dir
                )
                for src_ws in src_wb.Worksheets:
                    sname = str(src_ws.Name)
                    spec = specs.get(sname)
                    matched_template_name = sname
                    if spec is None:
                        spec = fallback_spec
                        matched_template_name = "模板"
                    if spec is None:
                        log.info("skip wb=%s sheet=%s reason=no_template_match", wb_name, sname)
                        continue
                    if specs.get(sname) is None and fallback_spec is not None:
                        log.info("fallback wb=%s sheet=%s template=模板", wb_name, sname)

                    # 收集本 spec 涉及的行/列范围，按需读取
                    accessed_rows: set[int] = set()
                    accessed_cols: set[int] = set()
                    for addr in spec.set_addrs:
                        scol, srow = _split_addr(addr)
                        accessed_rows.add(srow)
                        accessed_cols.add(_col2num(scol))
                    for sr, er, _, _ in spec.row_regions:
                        for r in range(sr, er + 1):
                            accessed_rows.add(r)
                    for _, _, cs, ce in spec.col_regions:
                        for c in range(cs, ce + 1):
                            accessed_cols.add(c)

                    df, _ = _read_src_sheet_via_com(src_ws, accessed_rows, accessed_cols)
                    if df.shape[0] == 0:
                        continue
                    merge_ranges = _load_merge_for_sheet(merge_src_path, sname)

                    for sr, er, _, _ in spec.row_regions:
                        for data_row in range(sr, er + 1):
                            ws_out.cell(row_idx, 1).value = wb_name
                            ws_out.cell(row_idx, 2).value = sname

                            # set 区（按 spec 内位置定位输出列）
                            for i, addr in enumerate(spec.set_addrs):
                                if i >= len(spec.set_names):
                                    break
                                sc, rr = _split_addr(addr)
                                target_col = set_loc.get((matched_template_name, i))
                                if target_col is not None:
                                    ws_out.cell(row_idx, target_col).value = (
                                        _gcv_from_array(df, merge_ranges, rr, _col2num(sc))
                                    )

                            # 列区域（按 spec 内位置定位输出列）
                            cidx = 0
                            for _, _, cs, ce in spec.col_regions:
                                for dc in range(cs, ce + 1):
                                    if cidx >= len(spec.col_headers):
                                        break
                                    target_col = col_loc.get((matched_template_name, cidx))
                                    if target_col is not None:
                                        ws_out.cell(row_idx, target_col).value = (
                                            _gcv_from_array(df, merge_ranges, data_row, dc)
                                        )
                                    cidx += 1

                            ws_out.cell(row_idx, len(headers)).value = data_row
                            row_idx += 1
                            rows_written += 1
            finally:
                safe_close(src_wb)
    finally:
        safe_quit(app)
        shutil.rmtree(xls_temp_dir, ignore_errors=True)

    if rows_written == 0:
        out_wb.close()
        return {"saved": False, "rows": 0, "path": ""}

    ws_out.freeze_panes = "A2"
    out_path = output_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_按批注汇总（模板+源文件）_COM.xlsx"
    out_wb.save(out_path)
    out_wb.close()
    log.info("[COM] done files=%s rows=%s output=%s", files_hit, rows_written, out_path)
    return {"saved": True, "rows": rows_written, "path": str(out_path)}
