from __future__ import annotations

from copy import copy
from datetime import datetime
from pathlib import Path
import re
from time import perf_counter

from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.worksheet import Worksheet

from .logger import get_logger


_RE_PRINT_AREA = re.compile(r"打印区域\s*#?\s*(\d+)", re.IGNORECASE)


def _norm(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _next_available_output_path(base_path: Path) -> Path:
    if not base_path.exists():
        return base_path
    stem = base_path.stem
    suffix = base_path.suffix
    parent = base_path.parent
    i = 1
    while True:
        cand = parent / f"{stem}_{i}{suffix}"
        if not cand.exists():
            return cand
        i += 1


def _save_with_suffix_fallback(wb: Workbook, preferred_path: Path) -> Path:
    target = _next_available_output_path(preferred_path)
    for _ in range(1000):
        try:
            wb.save(target)
            return target
        except PermissionError:
            # 文件被占用时继续尝试下一个后缀
            target = _next_available_output_path(target.with_name(f"{target.stem}_1{target.suffix}"))
    raise PermissionError(f"无法保存输出文件（可能持续被占用）: {preferred_path}")


def _is_supported(path: Path) -> bool:
    return path.suffix.lower() in (".xlsx", ".xlsm")


def _open_wb(path: Path, read_only: bool = False):
    return load_workbook(path, read_only=read_only, keep_vba=(path.suffix.lower() == ".xlsm"))


def _used_bounds(ws: Worksheet) -> tuple[int, int, int, int]:
    try:
        dim = ws.calculate_dimension()
        min_col, min_row, max_col, max_row = range_boundaries(dim)
        if max_row < min_row or max_col < min_col:
            return 1, 1, 1, 1
        return min_row, min_col, max_row, max_col
    except Exception:
        max_row = max(1, int(ws.max_row or 1))
        max_col = max(1, int(ws.max_column or 1))
        return 1, 1, max_row, max_col


def _parse_orientation(text: str) -> str | None:
    t = _norm(text)
    if not t:
        return None
    has_landscape = "横向" in t
    has_portrait = "纵向" in t
    if has_landscape and has_portrait:
        return None
    if has_landscape:
        return "landscape"
    if has_portrait:
        return "portrait"
    return None


def _auto_orientation(rows_count: int, cols_count: int) -> str:
    return "portrait" if rows_count >= cols_count else "landscape"


def _range_size_score(ws: Worksheet, bounds: tuple[int, int, int, int]) -> tuple[float, float]:
    min_row, min_col, max_row, max_col = bounds
    default_h = float(ws.sheet_format.defaultRowHeight or 15.0)
    default_w = float(ws.sheet_format.defaultColWidth or 8.43)
    total_h = 0.0
    total_w = 0.0

    for r in range(min_row, max_row + 1):
        h = ws.row_dimensions[r].height
        total_h += float(h if h is not None else default_h)
    for c in range(min_col, max_col + 1):
        w = ws.column_dimensions[get_column_letter(c)].width
        total_w += float(w if w is not None else default_w)
    return total_h, total_w


def _auto_orientation_by_range(ws: Worksheet, bounds: tuple[int, int, int, int]) -> str:
    # 默认纵向；仅当宽度显著大于高度时横向
    h, w = _range_size_score(ws, bounds)
    if w > h:
        return "landscape"
    return "portrait"


def _apply_print_setup(
    ws: Worksheet,
    total_rows: int,
    total_cols: int,
    fit_wide: int = 1,
    fit_tall: int = 1,
    orientation: str | None = None,
    center_h: bool = False,
    center_v: bool = False,
) -> None:
    rows_count = max(total_rows, 1)
    cols_count = max(total_cols, 1)
    final_orientation = orientation or _auto_orientation(rows_count, cols_count)
    ws.page_setup.orientation = final_orientation
    if ws.sheet_properties.pageSetUpPr is None:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    else:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = max(1, int(fit_wide or 1))
    # fitToHeight=0 表示不限制页高，按实际内容分页
    ws.page_setup.fitToHeight = max(0, int(fit_tall if fit_tall is not None else 1))
    ws.page_setup.scale = None
    ws.page_setup.paperSize = 9  # A4
    ws.print_area = f"A1:{get_column_letter(cols_count)}{rows_count}"
    ws.print_options.horizontalCentered = bool(center_h)
    ws.print_options.verticalCentered = bool(center_v)
    ws.oddFooter.center.text = "第 &P 页 / 共 &N 页"


def _apply_sheet_order_footer(out_wb: Workbook) -> None:
    sheets = list(out_wb.worksheets)
    total = len(sheets)
    if total <= 0:
        return
    for idx, ws in enumerate(sheets, start=1):
        ws.oddFooter.center.text = f"第 {idx} 页 / 共 {total} 页"


def _parse_ranges_text(range_text: str) -> list[tuple[int, int, int, int]]:
    txt = _norm(range_text)
    if not txt:
        return []
    segs = [p.strip() for p in txt.replace("；", ";").split(";") if p.strip()]
    out: list[tuple[int, int, int, int]] = []
    for seg in segs:
        min_col, min_row, max_col, max_row = range_boundaries(seg)
        out.append((min_row, min_col, max_row, max_col))
    return out


def _find_comment_ranges(ws: Worksheet) -> list[tuple[int, int, int, int]]:
    min_row, min_col, max_row, max_col = _used_bounds(ws)
    starts: dict[int, tuple[int, int]] = {}
    ends: dict[int, tuple[int, int]] = {}

    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cm = ws.cell(r, c).comment
            if cm is None:
                continue
            txt = _norm(cm.text)
            if "打印区域" not in txt:
                continue
            m = _RE_PRINT_AREA.search(txt)
            if not m:
                continue
            idx = int(m.group(1))
            if "打印区域#" in txt.replace(" ", ""):
                ends[idx] = (r, c)
            else:
                starts[idx] = (r, c)

    ranges: list[tuple[int, int, int, int]] = []
    for idx in sorted(set(starts.keys()) & set(ends.keys())):
        sr, sc = starts[idx]
        er, ec = ends[idx]
        ranges.append((min(sr, er), min(sc, ec), max(sr, er), max(sc, ec)))
    return ranges


def _copy_block(
    src_ws: Worksheet,
    dst_ws: Worksheet,
    bounds: tuple[int, int, int, int],
    dst_start_row: int,
    mode: int,
    value_ws: Worksheet | None = None,
    hide_zero: bool = False,
    drop_comments: bool = False,
) -> tuple[int, int]:
    min_row, min_col, max_row, max_col = bounds
    rows_count = max_row - min_row + 1
    cols_count = max_col - min_col + 1

    for r in range(rows_count):
        src_r = min_row + r
        dst_r = dst_start_row + r
        src_h = src_ws.row_dimensions[src_r].height
        if src_h is not None and mode == 1:
            dst_ws.row_dimensions[dst_r].height = src_h
        for c in range(cols_count):
            src_c = min_col + c
            dst_c = 1 + c
            src_cell = src_ws.cell(src_r, src_c)
            dst_cell = dst_ws.cell(dst_r, dst_c)
            raw_value = None
            if value_ws is not None:
                raw_value = value_ws.cell(src_r, src_c).value
            else:
                raw_value = src_cell.value
            dst_cell.value = _normalize_out_value(raw_value, hide_zero)
            if mode in (1, 2):
                if src_cell.has_style:
                    # 跨工作簿不能直接赋 _style（会引用源样式索引），需逐属性复制
                    dst_cell.font = copy(src_cell.font)
                    dst_cell.fill = copy(src_cell.fill)
                    dst_cell.border = copy(src_cell.border)
                    dst_cell.alignment = copy(src_cell.alignment)
                    dst_cell.protection = copy(src_cell.protection)
                    dst_cell.number_format = src_cell.number_format
                if (not drop_comments) and src_cell.comment is not None:
                    dst_cell.comment = copy(src_cell.comment)
            else:
                # 快速复制：仅保留基础样式（不复制边框/保护/批注）
                if src_cell.has_style:
                    dst_cell.font = copy(src_cell.font)
                    dst_cell.fill = copy(src_cell.fill)
                    dst_cell.alignment = copy(src_cell.alignment)
                    dst_cell.number_format = src_cell.number_format

    if mode in (1, 2):
        for c in range(cols_count):
            src_c = min_col + c
            dst_c = 1 + c
            src_w = src_ws.column_dimensions[get_column_letter(src_c)].width
            if src_w is not None:
                dst_ws.column_dimensions[get_column_letter(dst_c)].width = src_w
        for merged in list(src_ws.merged_cells.ranges):
            if (
                merged.min_row >= min_row
                and merged.max_row <= max_row
                and merged.min_col >= min_col
                and merged.max_col <= max_col
            ):
                dr1 = dst_start_row + (merged.min_row - min_row)
                dr2 = dst_start_row + (merged.max_row - min_row)
                dc1 = 1 + (merged.min_col - min_col)
                dc2 = 1 + (merged.max_col - min_col)
                try:
                    dst_ws.merge_cells(
                        f"{get_column_letter(dc1)}{dr1}:{get_column_letter(dc2)}{dr2}"
                    )
                except Exception:
                    # 个别源文件样式损坏时，跳过该合并，不中断整页输出
                    pass
    else:
        # 快速复制也保留合并结构
        for merged in list(src_ws.merged_cells.ranges):
            if (
                merged.min_row >= min_row
                and merged.max_row <= max_row
                and merged.min_col >= min_col
                and merged.max_col <= max_col
            ):
                dr1 = dst_start_row + (merged.min_row - min_row)
                dr2 = dst_start_row + (merged.max_row - min_row)
                dc1 = 1 + (merged.min_col - min_col)
                dc2 = 1 + (merged.max_col - min_col)
                try:
                    dst_ws.merge_cells(
                        f"{get_column_letter(dc1)}{dr1}:{get_column_letter(dc2)}{dr2}"
                    )
                except Exception:
                    pass

    if mode == 2:
        _autofit_block_columns(dst_ws, dst_start_row, rows_count, cols_count)
        _compact_block_row_heights(dst_ws, dst_start_row, rows_count)
    return rows_count, cols_count


def _is_zero_like(v) -> bool:
    if v is None:
        return False
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return abs(float(v)) == 0.0
    s = str(v).strip().replace(",", "")
    if s == "":
        return False
    try:
        return abs(float(s)) == 0.0
    except Exception:
        return False


def _normalize_out_value(v, hide_zero: bool):
    if hide_zero and _is_zero_like(v):
        return None
    return v


def _autofit_block_columns(dst_ws: Worksheet, start_row: int, rows: int, cols: int) -> None:
    """mode=2：紧凑适配列宽，尽量避免 #### 且不过度变宽。"""
    if rows <= 0 or cols <= 0:
        return
    for c in range(1, cols + 1):
        max_len = 0
        for r in range(start_row, start_row + rows):
            v = dst_ws.cell(r, c).value
            if v is None:
                continue
            n = len(str(v).strip())
            if n > max_len:
                max_len = n
        if max_len <= 0:
            continue
        # 紧凑口径：给少量冗余，且限制最大宽度，避免“挤成一页时字体过小”
        target_w = min(18.0, max(6.5, max_len * 0.95 + 1.2))
        key = get_column_letter(c)
        dst_ws.column_dimensions[key].width = float(target_w)


def _compact_block_row_heights(dst_ws: Worksheet, start_row: int, rows: int) -> None:
    """mode=2：行高压回标准紧凑值。"""
    if rows <= 0:
        return
    compact_h = 15.0
    for r in range(start_row, start_row + rows):
        dst_ws.row_dimensions[r].height = compact_h


def run_print_fast_by_comment(
    source_paths: list[Path], output_dir: Path, log_dir: Path, hide_zero: bool = False
) -> dict[str, int]:
    log = get_logger("3_10_3_print_fast", log_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    wb_hit = 0
    sheet_hit = 0
    saved_files = 0
    skipped = 0

    for src_path in source_paths:
        if not src_path.exists() or not _is_supported(src_path):
            skipped += 1
            log.warning("skip source=%s reason=not_found_or_unsupported", src_path)
            continue
        src_wb = _open_wb(src_path)
        src_val_wb = load_workbook(
            src_path,
            read_only=False,
            keep_vba=(src_path.suffix.lower() == ".xlsm"),
            data_only=True,
        )
        out_wb = Workbook()
        default_ws = out_wb.active
        if default_ws is not None and default_ws.title == "Sheet":
            out_wb.remove(default_ws)

        book_sheet_hit = 0
        for src_ws in src_wb.worksheets:
            a1_txt = _norm(src_ws["A1"].comment.text if src_ws["A1"].comment else "")
            if "打印" not in a1_txt and "输出" not in a1_txt:
                continue

            ranges = _find_comment_ranges(src_ws)
            if not ranges:
                ranges = [_used_bounds(src_ws)]
            if not ranges:
                continue

            book_sheet_hit += 1
            for i, rg in enumerate(ranges, start=1):
                name = src_ws.title if len(ranges) == 1 else f"{src_ws.title}_{i}"
                dst_ws = out_wb.create_sheet(title=name[:31])
                src_val_ws = src_val_wb[src_ws.title] if src_ws.title in src_val_wb.sheetnames else None
                rows, cols = _copy_block(
                    src_ws, dst_ws, rg, 1, mode=3, value_ws=src_val_ws, hide_zero=hide_zero
                )
                orientation = _parse_orientation(a1_txt) or _auto_orientation(rows, cols)
                _apply_print_setup(dst_ws, rows, cols, 1, 0, orientation, False, False)
                sheet_hit += 1

        if book_sheet_hit > 0:
            wb_hit += 1
            _apply_sheet_order_footer(out_wb)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = output_dir / f"{ts}_按批注打印（快速复制）.xlsx"
            out_path = _save_with_suffix_fallback(out_wb, out_path)
            saved_files += 1
            log.info("saved source=%s output=%s sheet=%s", src_path.name, out_path, book_sheet_hit)
        else:
            log.info("skip source=%s reason=no_marked_sheet", src_path.name)

        src_wb.close()
        src_val_wb.close()
        out_wb.close()

    return {
        "workbooks_hit": wb_hit,
        "sheets_hit": sheet_hit,
        "saved_files": saved_files,
        "skipped": skipped,
    }


def run_print_keep_by_comment(
    source_paths: list[Path], output_dir: Path, log_dir: Path, hide_zero: bool = False
) -> dict[str, int]:
    log = get_logger("3_10_1_print_keep", log_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    wb_hit = 0
    sheet_hit = 0
    saved_files = 0
    skipped = 0

    for src_path in source_paths:
        if not src_path.exists() or not _is_supported(src_path):
            skipped += 1
            log.warning("skip source=%s reason=not_found_or_unsupported", src_path)
            continue
        src_wb = _open_wb(src_path)
        src_val_wb = load_workbook(
            src_path,
            read_only=False,
            keep_vba=(src_path.suffix.lower() == ".xlsm"),
            data_only=True,
        )
        out_wb = Workbook()
        default_ws = out_wb.active
        if default_ws is not None and default_ws.title == "Sheet":
            out_wb.remove(default_ws)

        book_sheet_hit = 0
        for src_ws in src_wb.worksheets:
            a1_txt = _norm(src_ws["A1"].comment.text if src_ws["A1"].comment else "")
            if "打印" not in a1_txt and "输出" not in a1_txt:
                continue

            ranges = _find_comment_ranges(src_ws)
            if not ranges:
                ranges = [_used_bounds(src_ws)]
            if not ranges:
                continue

            book_sheet_hit += 1
            for i, rg in enumerate(ranges, start=1):
                name = src_ws.title if len(ranges) == 1 else f"{src_ws.title}_{i}"
                dst_ws = out_wb.create_sheet(title=name[:31])
                src_val_ws = src_val_wb[src_ws.title] if src_ws.title in src_val_wb.sheetnames else None
                rows, cols = _copy_block(
                    src_ws, dst_ws, rg, 1, mode=1, value_ws=src_val_ws, hide_zero=hide_zero
                )
                orientation = _parse_orientation(a1_txt) or _auto_orientation_by_range(src_ws, rg)
                _apply_print_setup(dst_ws, rows, cols, 1, 0, orientation, False, False)
                sheet_hit += 1

        if book_sheet_hit > 0:
            wb_hit += 1
            _apply_sheet_order_footer(out_wb)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = output_dir / f"{ts}_按批注打印（保留源格式）.xlsx"
            out_path = _save_with_suffix_fallback(out_wb, out_path)
            saved_files += 1
            log.info("saved source=%s output=%s sheet=%s", src_path.name, out_path, book_sheet_hit)
        else:
            log.info("skip source=%s reason=no_marked_sheet", src_path.name)

        src_wb.close()
        src_val_wb.close()
        out_wb.close()

    return {
        "workbooks_hit": wb_hit,
        "sheets_hit": sheet_hit,
        "saved_files": saved_files,
        "skipped": skipped,
    }


def run_print_config_precheck(tasks: list[dict], log_dir: Path) -> dict[str, int]:
    log = get_logger("3_10_8_print_config_precheck", log_dir)
    ok = warn = fail = 0
    source_wb_cache: dict[str, object] = {}

    for t in tasks:
        if not t.get("enabled", False):
            continue
        row_no = int(t.get("row_no", 0))
        mode = int(t.get("mode", 0) or 0)
        src_wb = t.get("source_wb")
        src_ws = _norm(t.get("source_ws"))
        tgt_wb = t.get("target_wb")
        tgt_ws = _norm(t.get("target_ws"))
        rg_text = _norm(t.get("source_ranges"))
        orientation = _norm(t.get("orientation"))

        if mode not in (1, 2, 3):
            fail += 1
            log.error("row=%s fail=invalid_mode mode=%s", row_no, mode)
            continue
        if src_wb is None or (not Path(src_wb).exists()) or (not _is_supported(Path(src_wb))):
            fail += 1
            log.error("row=%s fail=invalid_source_wb", row_no)
            continue
        if not src_ws:
            fail += 1
            log.error("row=%s fail=source_ws_empty", row_no)
            continue
        if tgt_wb is None or not _norm(str(tgt_wb)):
            fail += 1
            log.error("row=%s fail=target_wb_empty", row_no)
            continue
        if Path(tgt_wb).suffix.lower() not in (".xlsx", ".xlsm"):
            fail += 1
            log.error("row=%s fail=target_wb_unsupported_ext", row_no)
            continue
        if not tgt_ws:
            fail += 1
            log.error("row=%s fail=target_ws_empty", row_no)
            continue

        if rg_text:
            try:
                _parse_ranges_text(rg_text)
            except Exception as e:
                fail += 1
                log.error("row=%s fail=invalid_range err=%s", row_no, e)
                continue

        fw = int(t.get("fit_wide", 1) or 1)
        ft = int(t.get("fit_tall", 1) or 1)
        if fw <= 0 or ft <= 0:
            warn += 1
            log.warning("row=%s warn=fit_non_positive reset_to_1x1", row_no)

        if orientation and _parse_orientation(orientation) is None:
            warn += 1
            log.warning("row=%s warn=invalid_orientation_use_auto", row_no)

        wb_key = str(src_wb)
        wb = source_wb_cache.get(wb_key)
        if wb is None:
            try:
                wb = _open_wb(Path(src_wb), read_only=True)
                source_wb_cache[wb_key] = wb
            except Exception as e:
                fail += 1
                log.error("row=%s fail=open_source err=%s", row_no, e)
                continue
        if src_ws not in wb.sheetnames:
            fail += 1
            log.error("row=%s fail=source_ws_not_found sheet=%s", row_no, src_ws)
            continue

        ok += 1

    for wb in source_wb_cache.values():
        wb.close()

    return {"ok": ok, "warn": warn, "fail": fail}


def run_print_config_all_modes(tasks: list[dict], log_dir: Path) -> dict[str, object]:
    log = get_logger("3_10_7_print_config_run_all", log_dir)
    source_cache: dict[str, object] = {}
    source_value_cache: dict[str, object] = {}
    target_cache: dict[str, object] = {}
    cleared_target_ws: set[tuple[str, str]] = set()
    changed_targets: set[str] = set()
    written_sheet_count = 0
    written_row_count = 0
    task_ok = 0
    task_skip = 0
    total_enabled = sum(
        1 for t in tasks
        if t.get("enabled", False) and int(t.get("mode", 0) or 0) in (1, 2, 3)
    )
    progress_idx = 0
    started_ts = perf_counter()
    slow_records: list[dict[str, object]] = []

    for mode in (1, 2, 3):
        mode_started = perf_counter()
        mode_ok_before = task_ok
        mode_skip_before = task_skip
        mode_rows_before = written_row_count
        for t in tasks:
            if not t.get("enabled", False):
                continue
            if int(t.get("mode", 0) or 0) != mode:
                continue

            row_no = int(t.get("row_no", 0))
            src_wb_path = t.get("source_wb")
            src_ws_name = _norm(t.get("source_ws"))
            tgt_wb_path = t.get("target_wb")
            tgt_ws_name = _norm(t.get("target_ws"))
            progress_idx += 1
            task_started = perf_counter()
            src_disp = f"{Path(str(src_wb_path)).name if src_wb_path else '?'}|{src_ws_name or '?'}"
            tgt_disp = f"{Path(str(tgt_wb_path)).name if tgt_wb_path else '?'}|{tgt_ws_name or '?'}"
            print(
                f"[3-3进度] {progress_idx}/{total_enabled} mode={mode} row={row_no} src={src_disp} -> tgt={tgt_disp}",
                flush=True,
            )

            if (
                src_wb_path is None
                or not Path(src_wb_path).exists()
                or not _is_supported(Path(src_wb_path))
                or not src_ws_name
                or tgt_wb_path is None
                or not tgt_ws_name
            ):
                task_skip += 1
                log.warning("skip row=%s reason=config_invalid", row_no)
                elapsed = perf_counter() - task_started
                print(f"[3-3完成] row={row_no} skip=config_invalid 耗时={elapsed:.2f}s", flush=True)
                slow_records.append({"row_no": row_no, "mode": mode, "target": tgt_disp, "elapsed_sec": elapsed, "status": "skip:config_invalid"})
                continue
            if Path(tgt_wb_path).suffix.lower() not in (".xlsx", ".xlsm"):
                task_skip += 1
                log.warning("skip row=%s reason=target_ext_unsupported", row_no)
                elapsed = perf_counter() - task_started
                print(f"[3-3完成] row={row_no} skip=target_ext_unsupported 耗时={elapsed:.2f}s", flush=True)
                slow_records.append({"row_no": row_no, "mode": mode, "target": tgt_disp, "elapsed_sec": elapsed, "status": "skip:target_ext_unsupported"})
                continue

            src_key = str(src_wb_path)
            src_wb = source_cache.get(src_key)
            if src_wb is None:
                try:
                    src_wb = _open_wb(Path(src_wb_path))
                    source_cache[src_key] = src_wb
                    source_value_cache[src_key] = load_workbook(
                        Path(src_wb_path),
                        read_only=False,
                        keep_vba=(Path(src_wb_path).suffix.lower() == ".xlsm"),
                        data_only=True,
                    )
                except Exception as e:
                    task_skip += 1
                    log.warning("skip row=%s reason=open_source_failed err=%s", row_no, e)
                    elapsed = perf_counter() - task_started
                    print(f"[3-3完成] row={row_no} skip=open_source_failed 耗时={elapsed:.2f}s", flush=True)
                    slow_records.append({"row_no": row_no, "mode": mode, "target": tgt_disp, "elapsed_sec": elapsed, "status": "skip:open_source_failed"})
                    continue

            if src_ws_name not in src_wb.sheetnames:
                task_skip += 1
                log.warning("skip row=%s reason=source_sheet_not_found sheet=%s", row_no, src_ws_name)
                elapsed = perf_counter() - task_started
                print(f"[3-3完成] row={row_no} skip=source_sheet_not_found 耗时={elapsed:.2f}s", flush=True)
                slow_records.append({"row_no": row_no, "mode": mode, "target": tgt_disp, "elapsed_sec": elapsed, "status": "skip:source_sheet_not_found"})
                continue
            src_ws = src_wb[src_ws_name]
            src_val_wb = source_value_cache.get(src_key)
            src_val_ws = None
            if src_val_wb is not None and src_ws_name in src_val_wb.sheetnames:
                src_val_ws = src_val_wb[src_ws_name]

            ranges_text = _norm(t.get("source_ranges"))
            if ranges_text:
                try:
                    ranges = _parse_ranges_text(ranges_text)
                except Exception:
                    task_skip += 1
                    log.warning("skip row=%s reason=invalid_range", row_no)
                    elapsed = perf_counter() - task_started
                    print(f"[3-3完成] row={row_no} skip=invalid_range 耗时={elapsed:.2f}s", flush=True)
                    slow_records.append({"row_no": row_no, "mode": mode, "target": tgt_disp, "elapsed_sec": elapsed, "status": "skip:invalid_range"})
                    continue
            else:
                ranges = [_used_bounds(src_ws)]

            tgt_key = str(tgt_wb_path)
            tgt_wb = target_cache.get(tgt_key)
            if tgt_wb is None:
                if Path(tgt_wb_path).exists() and _is_supported(Path(tgt_wb_path)):
                    tgt_wb = _open_wb(Path(tgt_wb_path))
                else:
                    tgt_wb = Workbook()
                    dws = tgt_wb.active
                    if dws is not None and dws.title == "Sheet":
                        tgt_wb.remove(dws)
                target_cache[tgt_key] = tgt_wb

            if tgt_ws_name in tgt_wb.sheetnames:
                tgt_ws = tgt_wb[tgt_ws_name]
            else:
                tgt_ws = tgt_wb.create_sheet(title=tgt_ws_name[:31])

            clear_key = (tgt_key, tgt_ws.title)
            if clear_key not in cleared_target_ws:
                if tgt_ws.max_row > 1 or _norm(tgt_ws.cell(1, 1).value):
                    tgt_ws.delete_rows(1, tgt_ws.max_row)
                cleared_target_ws.add(clear_key)

            out_row = 1
            max_cols = 1
            for rg in ranges:
                rows, cols = _copy_block(
                    src_ws,
                    tgt_ws,
                    rg,
                    out_row,
                    mode=mode,
                    value_ws=src_val_ws,
                    hide_zero=bool(t.get("hide_zero", False)),
                    drop_comments=bool(t.get("drop_comments", False)),
                )
                out_row += rows + 1  # 段间空 1 行
                max_cols = max(max_cols, cols)
                written_row_count += rows

            orientation = _parse_orientation(_norm(t.get("orientation")))
            if orientation is None and ranges:
                orientation = _auto_orientation_by_range(src_ws, ranges[0])
            fw = int(t.get("fit_wide", 1) or 1)
            ft = int(t.get("fit_tall", 1) or 1)
            _apply_print_setup(
                tgt_ws,
                max(out_row - 1, 1),
                max_cols,
                fit_wide=max(1, fw),
                fit_tall=max(1, ft),
                orientation=orientation,
                center_h=bool(t.get("center_h", False)),
                center_v=bool(t.get("center_v", False)),
            )

            changed_targets.add(tgt_key)
            task_ok += 1
            written_sheet_count += 1
            log.info("ok row=%s mode=%s target=%s|%s", row_no, mode, Path(tgt_key).name, tgt_ws.title)
            elapsed = perf_counter() - task_started
            print(f"[3-3完成] row={row_no} ok 耗时={elapsed:.2f}s", flush=True)
            slow_records.append({"row_no": row_no, "mode": mode, "target": f"{Path(tgt_key).name}|{tgt_ws.title}", "elapsed_sec": elapsed, "status": "ok"})

        mode_elapsed = perf_counter() - mode_started
        print(
            f"[3-3阶段] mode={mode} ok+{task_ok - mode_ok_before} skip+{task_skip - mode_skip_before} "
            f"rows+{written_row_count - mode_rows_before} 耗时={mode_elapsed:.2f}s",
            flush=True,
        )

    for key, wb in target_cache.items():
        if key in changed_targets:
            _apply_sheet_order_footer(wb)
            p = Path(key)
            p.parent.mkdir(parents=True, exist_ok=True)
            wb.save(p)
        wb.close()
    for wb in source_cache.values():
        wb.close()
    for wb in source_value_cache.values():
        wb.close()

    elapsed_sec = perf_counter() - started_ts
    slow_top = sorted(slow_records, key=lambda x: float(x.get("elapsed_sec", 0.0)), reverse=True)[:3]

    return {
        "task_ok": task_ok,
        "task_skip": task_skip,
        "written_sheets": written_sheet_count,
        "written_rows": written_row_count,
        "elapsed_sec": elapsed_sec,
        "task_total": total_enabled,
        "slow_top": slow_top,
    }
