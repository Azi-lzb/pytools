"""3.10.x 打印工具 — COM 版本

利用 Excel 原生 Range.Copy + PasteSpecial 一次性保留所有样式/合并/列宽/批注，
对比 openpyxl 逐属性 copy 既快又更保真。
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from .com_engine import (
    open_readonly,
    safe_close,
    safe_quit,
    start_excel_app,
)
from .logger import get_logger
from .print_310 import (
    _RE_PRINT_AREA,
    _is_supported,
    _next_available_output_path,
    _norm,
    _parse_orientation,
    _parse_ranges_text,
)


def _is_supported_com(path: Path) -> bool:
    """COM 分支放宽：源文件除 xlsx/xlsm 外额外支持 xls。"""
    return path.suffix.lower() in (".xlsx", ".xlsm", ".xls")


# Excel 常量
XL_PASTE_ALL = -4104
XL_PASTE_VALUES = -4163
XL_PASTE_FORMATS = -4122
XL_ORIENT_LANDSCAPE = 2
XL_ORIENT_PORTRAIT = 1
XL_PAPER_A4 = 9
XL_OPEN_XML_WORKBOOK = 51


def _col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _used_bounds_com(ws) -> tuple[int, int, int, int]:
    ur = ws.UsedRange
    sr = int(ur.Row)
    sc = int(ur.Column)
    er = sr + int(ur.Rows.Count) - 1
    ec = sc + int(ur.Columns.Count) - 1
    return sr, sc, er, ec


def _a1_comment_text(ws) -> str:
    try:
        cm = ws.Cells(1, 1).Comment
        if cm is None:
            return ""
        return _norm(str(cm.Text()))
    except Exception:
        return ""


def _find_comment_ranges_com(ws) -> list[tuple[int, int, int, int]]:
    """走 ws.Comments 集合（只迭代被批注的格子，O(K)）。"""
    starts: dict[int, tuple[int, int]] = {}
    ends: dict[int, tuple[int, int]] = {}
    try:
        cnt = int(ws.Comments.Count)
    except Exception:
        return []
    for i in range(1, cnt + 1):
        try:
            cm = ws.Comments(i)
            txt = _norm(str(cm.Text()))
            if "打印区域" not in txt:
                continue
            m = _RE_PRINT_AREA.search(txt)
            if not m:
                continue
            idx = int(m.group(1))
            cell = cm.Parent
            r = int(cell.Row)
            c = int(cell.Column)
            if "打印区域#" in txt.replace(" ", ""):
                ends[idx] = (r, c)
            else:
                starts[idx] = (r, c)
        except Exception:
            continue
    out: list[tuple[int, int, int, int]] = []
    for idx in sorted(set(starts.keys()) & set(ends.keys())):
        sr, sc = starts[idx]
        er, ec = ends[idx]
        out.append((min(sr, er), min(sc, ec), max(sr, er), max(sc, ec)))
    return out


def _auto_orientation_by_dim(rows: int, cols: int) -> str:
    return "landscape" if cols > rows else "portrait"


def _auto_orientation_by_range_com(src_ws, bounds) -> str:
    sr, sc, er, ec = bounds
    try:
        h = sum(float(src_ws.Cells(r, sc).RowHeight or 15.0) for r in range(sr, er + 1))
        w = sum(float(src_ws.Cells(sr, c).ColumnWidth or 9.0) for c in range(sc, ec + 1))
        return "landscape" if w > h else "portrait"
    except Exception:
        return _auto_orientation_by_dim(er - sr + 1, ec - sc + 1)


def _copy_block_com(src_ws, dst_ws, src_bounds, dst_start_row: int, mode: int) -> tuple[int, int]:
    """复制 src_bounds 整块到 dst_ws，左上角落到 (dst_start_row, 1)。
    mode=1/2: 值口径粘贴（值+格式+合并+边框+批注；公式按源结果落地）
    mode=3:   值 + 格式（不带合并/批注；走 PasteValues + PasteFormats）
    """
    sr, sc, er, ec = src_bounds
    rows = er - sr + 1
    cols = ec - sc + 1
    src_rng = src_ws.Range(src_ws.Cells(sr, sc), src_ws.Cells(er, ec))
    src_rng.Copy()
    dst_anchor = dst_ws.Cells(dst_start_row, 1)
    if mode == 3:
        dst_anchor.PasteSpecial(Paste=XL_PASTE_VALUES)
        dst_anchor.PasteSpecial(Paste=XL_PASTE_FORMATS)
    else:
        dst_anchor.PasteSpecial(Paste=XL_PASTE_ALL)
        # 把刚粘过来的公式就地替换为源结果值，避免跨簿引用断裂/公式链污染
        dst_rng = dst_ws.Range(
            dst_ws.Cells(dst_start_row, 1),
            dst_ws.Cells(dst_start_row + rows - 1, cols),
        )
        try:
            dst_rng.Value = src_rng.Value
        except Exception:
            pass
    # 列宽（PasteSpecial 不带列宽，得单独搬）
    for c in range(1, cols + 1):
        try:
            dst_ws.Cells(1, c).ColumnWidth = float(src_ws.Cells(sr, sc + c - 1).ColumnWidth or 9.0)
        except Exception:
            pass
    if mode == 2:
        try:
            dst_rng = dst_ws.Range(
                dst_ws.Cells(dst_start_row, 1),
                dst_ws.Cells(dst_start_row + rows - 1, cols),
            )
            dst_rng.EntireColumn.AutoFit()
            # 紧凑口径：限制列宽上限，避免打印缩放后内容过小
            for c in range(1, cols + 1):
                cw = float(dst_ws.Cells(1, c).ColumnWidth or 9.0)
                dst_ws.Cells(1, c).ColumnWidth = max(6.5, min(18.0, cw))
            # 行高压回标准紧凑值
            for r in range(dst_start_row, dst_start_row + rows):
                dst_ws.Rows(r).RowHeight = 15.0
        except Exception:
            pass
    return rows, cols


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


def _clear_zero_in_block(dst_ws, start_row: int, rows: int, cols: int) -> None:
    if rows <= 0 or cols <= 0:
        return
    try:
        rng = dst_ws.Range(dst_ws.Cells(start_row, 1), dst_ws.Cells(start_row + rows - 1, cols))
        data = rng.Value
        if data is None:
            return
        if not isinstance(data, tuple):
            rng.Value = None if _is_zero_like(data) else data
            return
        payload = []
        for row in data:
            new_row = []
            for v in row:
                new_row.append(None if _is_zero_like(v) else v)
            payload.append(tuple(new_row))
        rng.Value = tuple(payload)
    except Exception:
        pass


def _clear_comments_in_block(dst_ws, start_row: int, rows: int, cols: int) -> None:
    if rows <= 0 or cols <= 0:
        return
    try:
        rng = dst_ws.Range(dst_ws.Cells(start_row, 1), dst_ws.Cells(start_row + rows - 1, cols))
        try:
            rng.ClearComments()
            return
        except Exception:
            pass
        for r in range(start_row, start_row + rows):
            for c in range(1, cols + 1):
                try:
                    cell = dst_ws.Cells(r, c)
                    if cell.Comment is not None:
                        cell.Comment.Delete()
                except Exception:
                    continue
    except Exception:
        pass


def _apply_print_setup_com(ws, rows: int, cols: int, fit_wide: int, fit_tall: int,
                           orientation: str | None,
                           center_h: bool = False,
                           center_v: bool = False) -> None:
    ps = ws.PageSetup
    ps.Orientation = (
        XL_ORIENT_LANDSCAPE if orientation == "landscape" else XL_ORIENT_PORTRAIT
    )
    try:
        ps.Zoom = False  # 必须先关 Zoom 才能让 FitToPages 生效
    except Exception:
        pass
    try:
        ps.FitToPagesWide = max(1, int(fit_wide or 1))
        # FitToPagesTall=0 表示不限制页高，按实际内容分页
        ps.FitToPagesTall = max(0, int(fit_tall if fit_tall is not None else 1))
    except Exception:
        pass
    try:
        ps.PaperSize = XL_PAPER_A4
    except Exception:
        pass
    try:
        ps.PrintArea = f"A1:{_col_letter(max(cols, 1))}{max(rows, 1)}"
        ps.CenterFooter = "第 &P 页 / 共 &N 页"
    except Exception:
        pass
    try:
        ps.CenterHorizontally = bool(center_h)
        ps.CenterVertically = bool(center_v)
    except Exception:
        pass


def _apply_sheet_order_footer_com(wb) -> None:
    try:
        total = int(wb.Worksheets.Count)
    except Exception:
        return
    if total <= 0:
        return
    for idx in range(1, total + 1):
        try:
            ws = wb.Worksheets(idx)
            ws.PageSetup.CenterFooter = f"第 {idx} 页 / 共 {total} 页"
        except Exception:
            continue


def _add_clean_workbook(app):
    """新建 wb 并把默认 Sheet1 留下，调用方可后续删除。"""
    old = None
    try:
        old = app.SheetsInNewWorkbook
        app.SheetsInNewWorkbook = 1
    except Exception:
        pass
    wb = app.Workbooks.Add()
    if old is not None:
        try:
            app.SheetsInNewWorkbook = old
        except Exception:
            pass
    return wb


def _drop_default_sheet(wb, default_name: str | None) -> None:
    """如果默认空 sheet 还在，删除之。"""
    if not default_name:
        return
    try:
        for ws in list(wb.Worksheets):
            if str(ws.Name) == default_name:
                ws.Delete()
                break
    except Exception:
        pass


def _save_xlsx_with_fallback(app, wb, preferred_path: Path) -> Path:
    target = _next_available_output_path(preferred_path)
    for _ in range(1000):
        try:
            wb.SaveAs(str(target.resolve()), FileFormat=XL_OPEN_XML_WORKBOOK)
            return target
        except Exception:
            target = _next_available_output_path(
                target.with_name(f"{target.stem}_1{target.suffix}")
            )
    raise PermissionError(f"无法保存输出文件: {preferred_path}")


def _ensure_sheet(wb, name: str):
    """确保 wb 里有指定名称 sheet（不存在则末尾创建）。"""
    safe = (name or "Sheet")[:31]
    for ws in wb.Worksheets:
        if str(ws.Name) == safe:
            return ws
    new_ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    try:
        new_ws.Name = safe
    except Exception:
        pass
    return new_ws


def _clear_sheet_contents(ws) -> None:
    try:
        ws.Cells.Clear()
    except Exception:
        pass


# ============== 3.10.1 / 3.10.3：按批注打印 ==============

def _run_print_by_comment_com(
    source_paths: list[Path],
    output_dir: Path,
    log_dir: Path,
    *,
    mode: int,
    menu_name: str,
    feature_tag: str,
    hide_zero: bool = False,
) -> dict[str, int]:
    log = get_logger(feature_tag, log_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    wb_hit = sheet_hit = saved_files = skipped = 0

    valid = [p for p in source_paths if p.exists() and _is_supported_com(p)]
    if not valid:
        return {"workbooks_hit": 0, "sheets_hit": 0, "saved_files": 0,
                "skipped": len(source_paths) - len(valid)}

    app = start_excel_app()
    try:
        for src_path in valid:
            src_wb = None
            out_wb = None
            try:
                try:
                    src_wb = open_readonly(app, src_path)
                except Exception as e:
                    skipped += 1
                    log.warning("skip source=%s reason=open_failed err=%s", src_path, e)
                    continue

                out_wb = _add_clean_workbook(app)
                default_name = str(out_wb.Worksheets(1).Name)

                book_sheet_hit = 0
                for src_ws in src_wb.Worksheets:
                    a1 = _a1_comment_text(src_ws)
                    if "打印" not in a1 and "输出" not in a1:
                        continue

                    ranges = _find_comment_ranges_com(src_ws)
                    if not ranges:
                        ranges = [_used_bounds_com(src_ws)]
                    if not ranges:
                        continue

                    book_sheet_hit += 1
                    src_title = str(src_ws.Name)
                    for i, rg in enumerate(ranges, start=1):
                        name = src_title if len(ranges) == 1 else f"{src_title}_{i}"
                        dst_ws = wb_create_sheet(out_wb, name[:31])
                        rows, cols = _copy_block_com(src_ws, dst_ws, rg, 1, mode=mode)
                        if hide_zero:
                            _clear_zero_in_block(dst_ws, 1, rows, cols)
                        # 自动方向：mode=3 用尺寸快速判定；其他用行列实际宽高
                        if mode == 3:
                            orientation = (
                                _parse_orientation(a1)
                                or _auto_orientation_by_dim(rows, cols)
                            )
                        else:
                            orientation = (
                                _parse_orientation(a1)
                                or _auto_orientation_by_range_com(src_ws, rg)
                            )
                        _apply_print_setup_com(dst_ws, rows, cols, 1, 0, orientation, False, False)
                        sheet_hit += 1

                if book_sheet_hit > 0:
                    wb_hit += 1
                    _drop_default_sheet(out_wb, default_name)
                    _apply_sheet_order_footer_com(out_wb)
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    out_path = output_dir / f"{ts}_{menu_name}_COM.xlsx"
                    saved = _save_xlsx_with_fallback(app, out_wb, out_path)
                    saved_files += 1
                    log.info("saved source=%s output=%s sheet=%s",
                             src_path.name, saved, book_sheet_hit)
                else:
                    log.info("skip source=%s reason=no_marked_sheet", src_path.name)
            finally:
                safe_close(out_wb)
                safe_close(src_wb)
    finally:
        safe_quit(app)

    return {
        "workbooks_hit": wb_hit,
        "sheets_hit": sheet_hit,
        "saved_files": saved_files,
        "skipped": skipped,
    }


def wb_create_sheet(wb, title: str):
    """末尾追加一张 sheet 并设标题（同名时 Excel 会自动加后缀）。"""
    new_ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    try:
        new_ws.Name = title[:31]
    except Exception:
        pass
    return new_ws


def run_print_keep_by_comment_com(
    source_paths: list[Path], output_dir: Path, log_dir: Path, hide_zero: bool = False
) -> dict[str, int]:
    """3.10.1 按批注打印（保留源格式）— COM。"""
    return _run_print_by_comment_com(
        source_paths, output_dir, log_dir,
        mode=1,
        menu_name="按批注打印（保留源格式）",
        feature_tag="3_10_1_print_keep_com",
        hide_zero=hide_zero,
    )


def run_print_fast_by_comment_com(
    source_paths: list[Path], output_dir: Path, log_dir: Path, hide_zero: bool = False
) -> dict[str, int]:
    """3.10.3 按批注打印（快速复制：值+格式）— COM。"""
    return _run_print_by_comment_com(
        source_paths, output_dir, log_dir,
        mode=3,
        menu_name="按批注打印（快速复制）",
        feature_tag="3_10_3_print_fast_com",
        hide_zero=hide_zero,
    )


# ============== 3.10.7：按配置打印（执行全部模式） ==============

def run_print_config_all_modes_com(tasks: list[dict], log_dir: Path) -> dict[str, int]:
    """3.10.7 按配置打印（mode=1/2/3 顺序执行）— COM。"""
    log = get_logger("3_10_7_print_config_run_all_com", log_dir)

    source_cache: dict[str, object] = {}
    target_cache: dict[str, object] = {}
    target_default_sheet: dict[str, str] = {}
    cleared_target_ws: set[tuple[str, str]] = set()
    changed_targets: set[str] = set()
    written_sheet_count = 0
    written_row_count = 0
    task_ok = 0
    task_skip = 0

    app = start_excel_app()
    try:
        for current_mode in (1, 2, 3):
            for t in tasks:
                if not t.get("enabled", False):
                    continue
                if int(t.get("mode", 0) or 0) != current_mode:
                    continue

                row_no = int(t.get("row_no", 0))
                src_wb_path = t.get("source_wb")
                src_ws_name = _norm(t.get("source_ws"))
                tgt_wb_path = t.get("target_wb")
                tgt_ws_name = _norm(t.get("target_ws"))

                if (
                    src_wb_path is None
                    or not Path(src_wb_path).exists()
                    or not _is_supported_com(Path(src_wb_path))
                    or not src_ws_name
                    or tgt_wb_path is None
                    or not tgt_ws_name
                ):
                    task_skip += 1
                    log.warning("skip row=%s reason=config_invalid", row_no)
                    continue
                if Path(tgt_wb_path).suffix.lower() not in (".xlsx", ".xlsm"):
                    task_skip += 1
                    log.warning("skip row=%s reason=target_ext_unsupported", row_no)
                    continue

                # 打开/复用 source workbook
                src_key = str(Path(src_wb_path).resolve())
                src_wb = source_cache.get(src_key)
                if src_wb is None:
                    try:
                        src_wb = open_readonly(app, Path(src_wb_path))
                        source_cache[src_key] = src_wb
                    except Exception as e:
                        task_skip += 1
                        log.warning("skip row=%s reason=open_source_failed err=%s", row_no, e)
                        continue

                # 找 source sheet
                src_ws = None
                for ws in src_wb.Worksheets:
                    if str(ws.Name) == src_ws_name:
                        src_ws = ws
                        break
                if src_ws is None:
                    task_skip += 1
                    log.warning("skip row=%s reason=source_sheet_not_found sheet=%s",
                                row_no, src_ws_name)
                    continue

                # 解析 ranges
                ranges_text = _norm(t.get("source_ranges"))
                if ranges_text:
                    try:
                        ranges = _parse_ranges_text(ranges_text)
                    except Exception:
                        task_skip += 1
                        log.warning("skip row=%s reason=invalid_range", row_no)
                        continue
                else:
                    ranges = [_used_bounds_com(src_ws)]

                # 打开/创建 target workbook（写权限，不只读）
                tgt_key = str(Path(tgt_wb_path).resolve())
                tgt_wb = target_cache.get(tgt_key)
                if tgt_wb is None:
                    try:
                        tgt_p = Path(tgt_wb_path)
                        if tgt_p.exists() and _is_supported(tgt_p):
                            tgt_wb = app.Workbooks.Open(
                                str(tgt_p.resolve()),
                                ReadOnly=False, UpdateLinks=0,
                            )
                            target_default_sheet[tgt_key] = ""
                        else:
                            tgt_wb = _add_clean_workbook(app)
                            target_default_sheet[tgt_key] = str(tgt_wb.Worksheets(1).Name)
                        target_cache[tgt_key] = tgt_wb
                    except Exception as e:
                        task_skip += 1
                        log.warning("skip row=%s reason=open_target_failed err=%s",
                                    row_no, e)
                        continue

                tgt_ws = _ensure_sheet(tgt_wb, tgt_ws_name)
                clear_key = (tgt_key, str(tgt_ws.Name))
                if clear_key not in cleared_target_ws:
                    _clear_sheet_contents(tgt_ws)
                    cleared_target_ws.add(clear_key)

                out_row = 1
                max_cols = 1
                for rg in ranges:
                    rows, cols = _copy_block_com(src_ws, tgt_ws, rg, out_row, mode=current_mode)
                    if bool(t.get("hide_zero", False)):
                        _clear_zero_in_block(tgt_ws, out_row, rows, cols)
                    if bool(t.get("drop_comments", False)):
                        _clear_comments_in_block(tgt_ws, out_row, rows, cols)
                    out_row += rows + 1  # 段间留 1 空行
                    max_cols = max(max_cols, cols)
                    written_row_count += rows

                orientation = _parse_orientation(_norm(t.get("orientation")))
                if orientation is None and ranges:
                    orientation = _auto_orientation_by_range_com(src_ws, ranges[0])
                fw = int(t.get("fit_wide", 1) or 1)
                ft = int(t.get("fit_tall", 1) or 1)
                _apply_print_setup_com(
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
                log.info("ok row=%s mode=%s target=%s|%s",
                         row_no, current_mode, Path(tgt_key).name, str(tgt_ws.Name))

        # 保存所有改动的 target
        for key, wb in target_cache.items():
            if key not in changed_targets:
                continue
            try:
                _drop_default_sheet(wb, target_default_sheet.get(key, ""))
                _apply_sheet_order_footer_com(wb)
                p = Path(key)
                p.parent.mkdir(parents=True, exist_ok=True)
                if p.exists():
                    wb.Save()
                else:
                    wb.SaveAs(str(p.resolve()), FileFormat=XL_OPEN_XML_WORKBOOK)
            except Exception as e:
                log.warning("save target failed key=%s err=%s", key, e)
    finally:
        for wb in target_cache.values():
            safe_close(wb)
        for wb in source_cache.values():
            safe_close(wb)
        safe_quit(app)

    return {
        "task_ok": task_ok,
        "task_skip": task_skip,
        "written_sheets": written_sheet_count,
        "written_rows": written_row_count,
    }


# ============== 3.10.x：按配置打印 → 导出 PDF（用作打印预览） ==============

XL_TYPE_PDF = 0


def run_print_config_to_pdf_com(tasks: list[dict], output_dir: Path,
                                log_dir: Path) -> dict:
    """与 run_print_config_all_modes_com 同口径处理任务，但**不写 xlsx**，
    而是把每个 target 工作簿用 Excel 原生 ExportAsFixedFormat 导出为 PDF。
    PDF 文件名按 target_wb 的 stem 命名，落到 output_dir。
    用法：跑一次拿到 PDF，PDF 阅读器缩略图扫一遍页面是否切版/方向/缩放是否对。
    """
    log = get_logger("3_10_pdf_preview_com", log_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    source_cache: dict[str, object] = {}
    target_cache: dict[str, object] = {}
    target_default_sheet: dict[str, str] = {}
    cleared_target_ws: set[tuple[str, str]] = set()
    populated_targets: set[str] = set()
    task_ok = 0
    task_skip = 0
    written_sheet_count = 0
    written_row_count = 0
    pdf_paths: list[str] = []

    app = start_excel_app()
    try:
        for current_mode in (1, 2, 3):
            for t in tasks:
                if not t.get("enabled", False):
                    continue
                if int(t.get("mode", 0) or 0) != current_mode:
                    continue

                row_no = int(t.get("row_no", 0))
                src_wb_path = t.get("source_wb")
                src_ws_name = _norm(t.get("source_ws"))
                tgt_wb_path = t.get("target_wb")
                tgt_ws_name = _norm(t.get("target_ws"))

                if (
                    src_wb_path is None
                    or not Path(src_wb_path).exists()
                    or not _is_supported_com(Path(src_wb_path))
                    or not src_ws_name
                    or tgt_wb_path is None
                    or not tgt_ws_name
                ):
                    task_skip += 1
                    log.warning("skip row=%s reason=config_invalid", row_no)
                    continue

                src_key = str(Path(src_wb_path).resolve())
                src_wb = source_cache.get(src_key)
                if src_wb is None:
                    try:
                        src_wb = open_readonly(app, Path(src_wb_path))
                        source_cache[src_key] = src_wb
                    except Exception as e:
                        task_skip += 1
                        log.warning("skip row=%s reason=open_source_failed err=%s",
                                    row_no, e)
                        continue

                src_ws = None
                for ws in src_wb.Worksheets:
                    if str(ws.Name) == src_ws_name:
                        src_ws = ws
                        break
                if src_ws is None:
                    task_skip += 1
                    log.warning("skip row=%s reason=source_sheet_not_found sheet=%s",
                                row_no, src_ws_name)
                    continue

                ranges_text = _norm(t.get("source_ranges"))
                if ranges_text:
                    try:
                        ranges = _parse_ranges_text(ranges_text)
                    except Exception:
                        task_skip += 1
                        log.warning("skip row=%s reason=invalid_range", row_no)
                        continue
                else:
                    ranges = [_used_bounds_com(src_ws)]

                # 永远在内存里建 target wb（不读盘上既有 xlsx，避免污染）
                tgt_key = str(Path(tgt_wb_path).resolve())
                tgt_wb = target_cache.get(tgt_key)
                if tgt_wb is None:
                    tgt_wb = _add_clean_workbook(app)
                    target_default_sheet[tgt_key] = str(tgt_wb.Worksheets(1).Name)
                    target_cache[tgt_key] = tgt_wb

                tgt_ws = _ensure_sheet(tgt_wb, tgt_ws_name)
                clear_key = (tgt_key, str(tgt_ws.Name))
                if clear_key not in cleared_target_ws:
                    _clear_sheet_contents(tgt_ws)
                    cleared_target_ws.add(clear_key)

                out_row = 1
                max_cols = 1
                for rg in ranges:
                    rows, cols = _copy_block_com(src_ws, tgt_ws, rg, out_row, mode=current_mode)
                    if bool(t.get("hide_zero", False)):
                        _clear_zero_in_block(tgt_ws, out_row, rows, cols)
                    if bool(t.get("drop_comments", False)):
                        _clear_comments_in_block(tgt_ws, out_row, rows, cols)
                    out_row += rows + 1
                    max_cols = max(max_cols, cols)
                    written_row_count += rows

                orientation = _parse_orientation(_norm(t.get("orientation")))
                if orientation is None and ranges:
                    orientation = _auto_orientation_by_range_com(src_ws, ranges[0])
                fw = int(t.get("fit_wide", 1) or 1)
                ft = int(t.get("fit_tall", 1) or 1)
                _apply_print_setup_com(
                    tgt_ws,
                    max(out_row - 1, 1),
                    max_cols,
                    fit_wide=max(1, fw),
                    fit_tall=max(1, ft),
                    orientation=orientation,
                    center_h=bool(t.get("center_h", False)),
                    center_v=bool(t.get("center_v", False)),
                )

                populated_targets.add(tgt_key)
                task_ok += 1
                written_sheet_count += 1
                log.info("ok row=%s mode=%s target=%s|%s",
                         row_no, current_mode, Path(tgt_key).name, str(tgt_ws.Name))

        # 每个 target 导出一份 PDF
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        for key, wb in target_cache.items():
            if key not in populated_targets:
                continue
            try:
                _drop_default_sheet(wb, target_default_sheet.get(key, ""))
                _apply_sheet_order_footer_com(wb)
                stem = Path(key).stem
                pdf_path = output_dir / f"{ts}_按配置打印导出PDF（打印预览）_{stem}.pdf"
                pdf_path = _next_available_output_path(pdf_path)
                wb.ExportAsFixedFormat(
                    Type=XL_TYPE_PDF,
                    Filename=str(pdf_path.resolve()),
                    Quality=0,            # 0=标准
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False,
                )
                pdf_paths.append(str(pdf_path))
                log.info("pdf saved: %s", pdf_path)
            except Exception as e:
                log.warning("pdf export failed key=%s err=%s", key, e)
    finally:
        for wb in target_cache.values():
            safe_close(wb)
        for wb in source_cache.values():
            safe_close(wb)
        safe_quit(app)

    return {
        "task_ok": task_ok,
        "task_skip": task_skip,
        "written_sheets": written_sheet_count,
        "written_rows": written_row_count,
        "pdf_files": pdf_paths,
    }
