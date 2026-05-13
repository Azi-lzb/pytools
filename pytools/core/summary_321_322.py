from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .logger import get_logger


_REG_NUM = re.compile(r"(\d+)")


def _norm(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _supported(path: Path) -> bool:
    return path.suffix.lower() in (".xlsx", ".xlsm")


def _safe_sheet_name(name: str, used: set[str]) -> str:
    s = re.sub(r"[\\/?*\[\]:]", "_", _norm(name) or "Sheet")[:31] or "Sheet"
    if s not in used:
        used.add(s)
        return s
    base = s[:28]
    i = 2
    while True:
        cand = f"{base}_{i}"[:31]
        if cand not in used:
            used.add(cand)
            return cand
        i += 1


def _disambiguate_headers(names: list[str]) -> list[str]:
    """同名列追加 _2/_3/... 后缀，第一次出现保持原名。"""
    seen: dict[str, int] = {}
    out: list[str] = []
    for n in names:
        cnt = seen.get(n, 0) + 1
        seen[n] = cnt
        out.append(n if cnt == 1 else f"{n}_{cnt}")
    return out


def _gcv(ws: Worksheet, r: int, c: int):
    cell = ws.cell(r, c)
    if cell is None:
        return None
    for rg in ws.merged_cells.ranges:
        if cell.coordinate in rg:
            return ws.cell(rg.min_row, rg.min_col).value
    return cell.value


def _used_bounds(ws: Worksheet) -> tuple[int, int, int, int]:
    try:
        dim = ws.calculate_dimension()
        min_col, min_row, max_col, max_row = range_boundaries(dim)
        if max_row < min_row or max_col < min_col:
            return 1, 1, 1, 1
        return min_row, min_col, max_row, max_col
    except Exception:
        return 1, 1, max(1, ws.max_row), max(1, ws.max_column)


def _extract_comments(ws: Worksheet) -> dict[str, str]:
    out: dict[str, str] = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.comment is not None:
                out[cell.coordinate] = _norm(cell.comment.text)
    return out


def _split_addr(addr: str) -> tuple[str, int]:
    a = addr.replace("$", "")
    col = ""
    row = ""
    for ch in a:
        if ch.isalpha():
            col += ch.upper()
        elif ch.isdigit():
            row += ch
    return col, int(row or 0)


def _col2num(col: str) -> int:
    n = 0
    for ch in col.upper():
        if not ("A" <= ch <= "Z"):
            return 0
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def _extract_regions(comment_map: dict[str, str], keyword: str) -> list[tuple[int, int, int, int]]:
    nums: set[int] = set()
    for txt in comment_map.values():
        if keyword not in txt:
            continue
        m = _REG_NUM.search(txt.replace("#", ""))
        if m:
            nums.add(int(m.group(1)))
    out: list[tuple[int, int, int, int]] = []
    for n in sorted(nums):
        sa = ""
        ea = ""
        for addr, txt in comment_map.items():
            t = txt.replace(" ", "")
            if f"{keyword}{n}" in t and f"{keyword}#{n}" not in t:
                sa = addr
            if f"{keyword}#{n}" in t:
                ea = addr
        if not sa or not ea:
            continue
        sc, sr = _split_addr(sa)
        ec, er = _split_addr(ea)
        scn, ecn = _col2num(sc), _col2num(ec)
        out.append((min(sr, er), max(sr, er), min(scn, ecn), max(scn, ecn)))
    return out


def _extract_set_info(comment_map: dict[str, str]) -> tuple[list[str], list[str]]:
    names: list[str] = []
    addrs: list[str] = []
    for addr, txt in comment_map.items():
        m = re.search(r"set[\(（]([^)）]+)[)）]", txt, flags=re.IGNORECASE)
        if m:
            names.append(_norm(m.group(1)))
            addrs.append(addr)
    return names, addrs


def _build_col_headers(tmpl_ws: Worksheet, col_regions: list[tuple[int, int, int, int]]) -> list[str]:
    headers: list[str] = []
    for sr, er, sc, ec in col_regions:
        if sr < er:
            for c in range(sc, ec + 1):
                v1 = _gcv(tmpl_ws, sr, c)
                v2 = _gcv(tmpl_ws, er, c)
                s1 = _norm(v1)
                s2 = _norm(v2)
                if s1 and s2 and s1 != s2:
                    headers.append(f"{s1}_{s2}")
                elif s1:
                    headers.append(s1)
                elif s2:
                    headers.append(s2)
                else:
                    headers.append(get_column_letter(c))
        else:
            for c in range(sc, ec + 1):
                sv = _norm(_gcv(tmpl_ws, sr, c))
                headers.append(sv or get_column_letter(c))
    return headers


@dataclass
class TemplateSheetSpec:
    name: str
    comment_map: dict[str, str]
    row_regions: list[tuple[int, int, int, int]]
    col_regions: list[tuple[int, int, int, int]]
    set_names: list[str]
    set_addrs: list[str]
    col_headers: list[str]


def run_summary_by_usedrange(
    source_paths: list[Path], output_dir: Path, log_dir: Path
) -> dict[str, object]:
    log = get_logger("3_2_1_summary_usedrange", log_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    out_wb = Workbook()
    ws_out = out_wb.active
    ws_out.title = "汇总"

    max_cols = 0
    rows_written = 0
    wb_hit = 0
    row_idx = 2

    print("→ 正在扫描列宽上限...")
    # 先扫一次最大列数
    for p in source_paths:
        if not p.exists() or not _supported(p):
            continue
        wb = load_workbook(p, data_only=True, read_only=False, keep_vba=(p.suffix.lower() == ".xlsm"))
        try:
            for ws in wb.worksheets:
                _, _, _, ec = _used_bounds(ws)
                if ec > max_cols:
                    max_cols = ec
        finally:
            wb.close()

    if max_cols <= 0:
        return {"saved": False, "rows": 0, "path": ""}

    ws_out.cell(1, 1).value = "工作簿"
    ws_out.cell(1, 2).value = "工作表"
    ws_out.cell(1, 3).value = "行号"
    for c in range(1, max_cols + 1):
        ws_out.cell(1, 3 + c).value = f"列{c}"
    for c in range(1, 4 + max_cols):
        ws_out.cell(1, c).font = ws_out.cell(1, c).font.copy(bold=True)

    print(f"→ 开始汇总，共 {len(source_paths)} 个文件...")
    for idx, p in enumerate(source_paths, start=1):
        if not p.exists() or not _supported(p):
            log.warning("skip source=%s reason=not_supported_or_missing", p)
            continue
        print(f"  [{idx}/{len(source_paths)}] 打开: {p.name}")
        wb = load_workbook(p, data_only=True, read_only=False, keep_vba=(p.suffix.lower() == ".xlsm"))
        wb_hit += 1
        try:
            wb_name = p.stem
            for ws in wb.worksheets:
                print(f"     - 处理工作表: {ws.title}")
                sr, sc, er, ec = _used_bounds(ws)
                if er < sr or ec < sc:
                    continue
                # 批量读取 + 批量写入，避免逐单元格 COM 风格慢循环
                matrix = []
                for r in range(sr, er + 1):
                    row_vals = [wb_name, ws.title, r]
                    for c in range(1, max_cols + 1):
                        if sc <= c <= ec:
                            row_vals.append(ws.cell(r, c).value)
                        else:
                            row_vals.append(None)
                    matrix.append(row_vals)
                if matrix:
                    for row_vals in matrix:
                        ws_out.append(row_vals)
                    row_idx += len(matrix)
                    rows_written += len(matrix)
        finally:
            wb.close()
        print(f"     完成: {p.name}")

    if rows_written == 0:
        out_wb.close()
        return {"saved": False, "rows": 0, "path": ""}

    ws_out.freeze_panes = "A2"
    out_path = output_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_按使用区域汇总（全量）.xlsx"
    print("→ 正在保存结果文件...")
    out_wb.save(out_path)
    out_wb.close()
    log.info("done wb=%s rows=%s output=%s", wb_hit, rows_written, out_path)
    return {"saved": True, "rows": rows_written, "path": str(out_path)}


def run_summary_by_comment(
    template_path: Path, source_paths: list[Path], output_dir: Path, log_dir: Path
) -> dict[str, object]:
    log = get_logger("3_2_2_summary_comment", log_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not template_path.exists() or not _supported(template_path):
        raise ValueError("模板仅支持 xlsx/xlsm 且必须存在。")

    tmpl_wb = load_workbook(template_path, data_only=False, read_only=False, keep_vba=(template_path.suffix.lower() == ".xlsm"))
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

    if not specs:
        raise ValueError("模板未识别到可用批注区域（需包含 行区域N/#N 与 列区域N/#N）。")

    out_wb = Workbook()
    default_ws = out_wb.active
    out_wb.remove(default_ws)
    used_sheet_names: set[str] = set()
    sheet_ctx: dict[str, dict[str, object]] = {}
    for spec_key, sp in specs.items():
        ws = out_wb.create_sheet(_safe_sheet_name(spec_key, used_sheet_names))
        set_headers = _disambiguate_headers(sp.set_names)
        col_headers = _disambiguate_headers(sp.col_headers)
        headers = ["工作簿", "工作表"] + set_headers + col_headers + ["行号"]
        for i, h in enumerate(headers, 1):
            ws.cell(1, i).value = h
            ws.cell(1, i).font = ws.cell(1, i).font.copy(bold=True)
        sheet_ctx[spec_key] = {
            "ws": ws,
            "row_idx": 2,
            "set_base": 3,
            "col_base": 3 + len(set_headers),
            "last_col": len(headers),
        }

    rows_written = 0
    files_hit = 0
    fallback_spec = specs.get("模板")
    for p in source_paths:
        if not p.exists() or not _supported(p):
            log.warning("skip source=%s reason=not_supported_or_missing", p)
            continue
        src_wb = load_workbook(p, data_only=True, read_only=False, keep_vba=(p.suffix.lower() == ".xlsm"))
        try:
            files_hit += 1
            wb_name = p.stem
            for src_ws in src_wb.worksheets:
                sname = src_ws.title
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

                ctx = sheet_ctx[matched_template_name]
                ws_out = ctx["ws"]
                row_idx = int(ctx["row_idx"])
                set_base = int(ctx["set_base"])
                col_base = int(ctx["col_base"])
                last_col = int(ctx["last_col"])

                # 本行数据先构建字典，再映射到并集列
                for sr, er, _, _ in spec.row_regions:
                    for data_row in range(sr, er + 1):
                        ws_out.cell(row_idx, 1).value = wb_name
                        ws_out.cell(row_idx, 2).value = sname

                        # set 区（按 spec 内位置定位输出列）
                        for i, addr in enumerate(spec.set_addrs):
                            if i >= len(spec.set_names):
                                break
                            sc, rr = _split_addr(addr)
                            ws_out.cell(row_idx, set_base + i).value = _gcv(src_ws, rr, _col2num(sc))

                        # 列区域（按 spec 内位置定位输出列）
                        cidx = 0
                        for _, _, cs, ce in spec.col_regions:
                            for dc in range(cs, ce + 1):
                                if cidx >= len(spec.col_headers):
                                    break
                                ws_out.cell(row_idx, col_base + cidx).value = _gcv(src_ws, data_row, dc)
                                cidx += 1

                        ws_out.cell(row_idx, last_col).value = data_row
                        row_idx += 1
                        rows_written += 1
                ctx["row_idx"] = row_idx
        finally:
            src_wb.close()

    if rows_written == 0:
        out_wb.close()
        return {"saved": False, "rows": 0, "path": ""}

    for ctx in sheet_ctx.values():
        ws = ctx["ws"]
        if int(ctx["row_idx"]) > 2:
            ws.freeze_panes = "A2"
    out_path = output_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_按批注汇总（模板+源文件）.xlsx"
    out_wb.save(out_path)
    out_wb.close()
    log.info("done files=%s rows=%s output=%s", files_hit, rows_written, out_path)
    return {"saved": True, "rows": rows_written, "path": str(out_path)}
