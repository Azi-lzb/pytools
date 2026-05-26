from __future__ import annotations

from pathlib import Path
import os
import re
import gc

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string

from .config_xlsx import SHEET_CONFIG_RENAME
from .logger import get_logger


def _norm(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _split_tokens(v: str) -> list[str]:
    s = _norm(v).replace("；", ";").replace(",", ";")
    return [x.strip() for x in s.split(";") if x.strip()]


def _workbook_supported(path: Path) -> bool:
    return path.suffix.lower() in (".xlsx", ".xlsm")


def _excel_like(path: Path) -> bool:
    return path.suffix.lower() in (".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm", ".csv")


def _load_config_rename_df(cfg_path: Path) -> pd.DataFrame:
    return pd.read_excel(cfg_path, SHEET_CONFIG_RENAME, header=0, dtype=object)


def _build_rename_maps(cfg_path: Path) -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    df = _load_config_rename_df(cfg_path)
    short_to_full: dict[str, str] = {}
    full_to_code: dict[str, str] = {}
    kv: dict[str, str] = {}

    for _, row in df.iterrows():
        a = _norm(row.get("简称"))
        b = _norm(row.get("全称"))
        d = _norm(row.get("代码"))
        e = _norm(row.get("全称(代码映射)"))
        g = _norm(row.get("键"))
        h = _norm(row.get("值"))
        if a and b:
            short_to_full[a] = b
        if e and d:
            full_to_code[e] = d
        if g and h:
            kv[g] = h
    return short_to_full, full_to_code, kv


def _next_path(p: Path) -> Path:
    if not p.exists():
        return p
    i = 1
    while True:
        cand = p.with_name(f"{p.stem}_{i}{p.suffix}")
        if not cand.exists():
            return cand
        i += 1


def _dispatch_first(progids: list[str]):
    try:
        import win32com.client  # type: ignore
    except Exception as e:
        raise RuntimeError(f"缺少 pywin32，无法执行 COM 转换: {e}")
    last_err = None
    for pid in progids:
        try:
            app = win32com.client.DispatchEx(pid)
            return app, pid
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"无法创建 COM 应用（已尝试: {', '.join(progids)}）: {last_err}")


def run_batch_rename_33(cfg_path: Path, files: list[Path], log_dir: Path) -> dict[str, int]:
    log = get_logger("3_3_batch_rename", log_dir)
    short_to_full, full_to_code, kv = _build_rename_maps(cfg_path)
    date_text = kv.get("数据日期", "")
    report_name = kv.get("报表名称", "")
    if not short_to_full or not full_to_code or not date_text or not report_name:
        raise ValueError("重命名配置 缺少必要配置：简称/全称、全称(代码映射)/代码、数据日期、报表名称。")

    ok = skip = 0
    for p in files:
        if not p.exists():
            skip += 1
            log.warning("skip file=%s reason=not_exists", p)
            continue
        name = p.name
        matched_short = ""
        for s in short_to_full.keys():
            if s.lower() in name.lower():
                matched_short = s
                break
        if not matched_short:
            skip += 1
            log.info("skip file=%s reason=short_name_not_matched", name)
            continue
        full_name = short_to_full.get(matched_short, "")
        code = full_to_code.get(full_name, "")
        if not code:
            skip += 1
            log.info("skip file=%s reason=full_name_no_code full=%s", name, full_name)
            continue

        new_name = f"{date_text} {code} {full_name} {report_name}{p.suffix}"
        tgt = p.with_name(new_name)
        if tgt.resolve() == p.resolve():
            skip += 1
            log.info("skip file=%s reason=name_same", name)
            continue
        if tgt.exists():
            tgt = _next_path(tgt)
        try:
            os.rename(p, tgt)
            ok += 1
            log.info("ok old=%s new=%s", name, tgt.name)
        except Exception as e:
            skip += 1
            log.warning("skip file=%s reason=rename_failed err=%s", name, e)
    return {"ok": ok, "skip": skip}


_EXCEL_FMT_MAP = {
    "xls": 56, "xlsx": 51, "xlsm": 52, "csv": 6, "xlt": 17, "xltx": 54, "xltm": 53, "xlsb": 50,
}


def _start_excel_app():
    app, engine = _dispatch_first([
        "Excel.Application",      # Microsoft Excel
        "ket.Application",        # WPS 表格（常见）
        "KET.Application",        # WPS 表格（大小写变体）
    ])
    app.Visible = False
    app.DisplayAlerts = False
    for attr, val in (("ScreenUpdating", False), ("EnableEvents", False),
                      ("AskToUpdateLinks", False), ("AlertBeforeOverwriting", False)):
        try:
            setattr(app, attr, val)
        except Exception:
            pass
    try:
        app.Calculation = -4135  # xlCalculationManual
    except Exception:
        pass
    return app, engine


def _excel_convert_one(app, src: Path, out: Path, target_ext: str) -> None:
    wb = app.Workbooks.Open(str(src), UpdateLinks=0, ReadOnly=True)
    try:
        wb.SaveAs(str(out), FileFormat=_EXCEL_FMT_MAP[target_ext], CreateBackup=False)
    finally:
        wb.Close(SaveChanges=False)


def run_excel_convert_34(
    files: list[Path], target_format: str, log_dir: Path
) -> dict[str, int]:
    log = get_logger("3_4_excel_convert", log_dir)
    ext = _norm(target_format).lower().lstrip(".") or "xlsx"
    if ext not in _EXCEL_FMT_MAP:
        ext = "xlsx"

    # 预筛有效文件，避免为 0 个有效文件也启动 Excel
    valid: list[tuple[Path, Path]] = []
    ok = skip = 0
    for p in files:
        if not p.exists() or not _excel_like(p):
            skip += 1
            log.warning("skip file=%s reason=not_excel_like", p)
            continue
        out_dir = p.parent / ext
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{p.stem}.{ext}"
        if out_path.resolve() == p.resolve():
            skip += 1
            log.info("skip file=%s reason=source_equals_target", p.name)
            continue
        if out_path.exists():
            out_path = _next_path(out_path)
        valid.append((p, out_path))

    if not valid:
        return {"ok": 0, "skip": skip, "target_ext": ext, "engine": ""}

    app, engine = _start_excel_app()
    try:
        for src, out_path in valid:
            try:
                _excel_convert_one(app, src, out_path, ext)
                ok += 1
                log.info("ok src=%s out=%s engine=%s", src, out_path, engine)
            except Exception as e:
                skip += 1
                log.warning("skip file=%s reason=convert_failed err=%s", src, e)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
    return {"ok": ok, "skip": skip, "target_ext": ext, "engine": engine}


_WORD_FMT_MAP = {"doc": 0, "docx": 12}


def _start_word_app():
    app, engine = _dispatch_first([
        "Word.Application",       # Microsoft Word
        "kwps.Application",       # WPS 文字（常见）
        "KWPS.Application",       # WPS 文字（大小写变体）
        "wps.Application",        # 兼容别名（部分环境）
    ])
    app.Visible = False
    try:
        app.DisplayAlerts = 0
    except Exception:
        pass
    return app, engine


def _word_convert_one(app, src: Path, out: Path, target_ext: str) -> None:
    doc = app.Documents.Open(str(src), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
    try:
        try:
            doc.SaveAs2(str(out), FileFormat=_WORD_FMT_MAP[target_ext])
        except Exception:
            doc.SaveAs(str(out), FileFormat=_WORD_FMT_MAP[target_ext])
    finally:
        doc.Close(SaveChanges=False)


def run_word_convert_36(
    files: list[Path], target_format: str, log_dir: Path
) -> dict[str, int]:
    log = get_logger("3_6_word_convert", log_dir)
    ext = _norm(target_format).lower().lstrip(".") or "docx"
    if ext not in _WORD_FMT_MAP:
        ext = "docx"

    valid: list[tuple[Path, Path]] = []
    ok = skip = 0
    for p in files:
        if not p.exists() or p.suffix.lower() not in (".doc", ".docx"):
            skip += 1
            log.warning("skip file=%s reason=not_word_doc", p)
            continue
        out_dir = p.parent / ext
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{p.stem}.{ext}"
        if out_path.resolve() == p.resolve():
            skip += 1
            log.info("skip file=%s reason=source_equals_target", p.name)
            continue
        if out_path.exists():
            out_path = _next_path(out_path)
        valid.append((p, out_path))

    if not valid:
        return {"ok": 0, "skip": skip, "target_ext": ext, "engine": ""}

    app, engine = _start_word_app()
    try:
        for src, out_path in valid:
            try:
                _word_convert_one(app, src, out_path, ext)
                ok += 1
                log.info("ok src=%s out=%s engine=%s", src, out_path, engine)
            except Exception as e:
                skip += 1
                log.warning("skip file=%s reason=convert_failed err=%s", src, e)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
    return {"ok": ok, "skip": skip, "target_ext": ext, "engine": engine}


def _load_sheet_rename_map(cfg_path: Path) -> dict[str, str]:
    df = _load_config_rename_df(cfg_path)
    mp: dict[str, str] = {}
    for _, row in df.iterrows():
        old_name = _norm(row.get("原表名"))
        new_name = _norm(row.get("新表名"))
        if old_name and new_name:
            mp[old_name] = new_name
    return mp


def _load_file_local_rename_map(cfg_path: Path) -> list[tuple[str, str]]:
    df = _load_config_rename_df(cfg_path)
    out: list[tuple[str, str]] = []
    for _, row in df.iterrows():
        src = _norm(row.get("原文件名片段"))
        dst = _norm(row.get("新文件名片段"))
        if src:
            out.append((src, dst))
    return out


def _apply_local_name_map(name: str, mapping: list[tuple[str, str]]) -> str:
    stem = Path(name).stem
    suffix = Path(name).suffix
    new_stem = stem
    for src, dst in mapping:
        new_stem = new_stem.replace(src, dst)
    return f"{new_stem}{suffix}"


def _load_content_rename_spec(cfg_path: Path) -> str:
    wb = load_workbook(cfg_path, read_only=True, data_only=True)
    try:
        if SHEET_CONFIG_RENAME not in wb.sheetnames:
            return ""
        ws = wb[SHEET_CONFIG_RENAME]
        v = ws["L2"].value
        return _norm(v)
    finally:
        wb.close()


def _sanitize_name_part(text: str) -> str:
    s = _norm(text)
    if not s:
        return ""
    s = re.sub(r'[<>:"/\\|?*]', "_", s)
    s = re.sub(r"\s+", " ", s).strip().strip(".")
    return s


def _content_prefix_items(vals: list[str]) -> list[str]:
    return [x for x in (_sanitize_name_part(v) for v in vals) if x]


def _already_renamed_by_content(file_name: str, prefix_items: list[str]) -> bool:
    if not prefix_items:
        return False
    stem = Path(file_name).stem
    first_part = stem.split("_", 1)[0]
    haystacks = (first_part, stem)
    return all(any(item in h for h in haystacks) for item in prefix_items)


def _content_prefix(prefix_items: list[str]) -> str:
    prefix = "_".join(prefix_items)
    return f"{prefix}_" if prefix else ""


def _remove_content_prefix(file_name: str, prefix_items: list[str]) -> str | None:
    prefix = _content_prefix(prefix_items)
    if not prefix:
        return None
    stem = Path(file_name).stem
    suffix = Path(file_name).suffix
    changed = False
    while stem.startswith(prefix):
        stem = stem[len(prefix):]
        changed = True
    item_set = set(prefix_items)
    while "_" in stem:
        head, rest = stem.split("_", 1)
        if head not in item_set:
            break
        stem = rest
        changed = True
    if not changed or not stem:
        return None
    return f"{stem}{suffix}"


def _parse_sheet_cell_spec(spec: str) -> list[tuple[str, str]]:
    parts = [x.strip() for x in spec.replace("；", ";").split(";") if x and x.strip()]
    out: list[tuple[str, str]] = []
    for p in parts:
        if "@" not in p:
            raise ValueError(f"L2 配置格式错误: {p}（应为 sheet@A1）")
        sheet_name, addr = p.split("@", 1)
        sheet_name = _norm(sheet_name)
        addr = _norm(addr).upper()
        if not sheet_name or not addr:
            raise ValueError(f"L2 配置格式错误: {p}（sheet 或地址为空）")
        try:
            coordinate_from_string(addr)
        except Exception:
            raise ValueError(f"L2 地址非法: {p}") from None
        out.append((sheet_name, addr))
    if not out:
        raise ValueError("L2 为空，无法执行根据文件内容重命名。")
    return out


def _read_cells_openpyxl(path: Path, items: list[tuple[str, str]]) -> list[str]:
    wb = load_workbook(path, data_only=True, keep_vba=(path.suffix.lower() == ".xlsm"))
    try:
        vals: list[str] = []
        for sheet_name, addr in items:
            if sheet_name not in wb.sheetnames:
                vals.append(f"{sheet_name}!{addr}")
                continue
            ws = wb[sheet_name]
            v = ws[addr].value
            txt = _sanitize_name_part("" if v is None else str(v))
            vals.append(txt or f"{sheet_name}!{addr}")
        return vals
    finally:
        wb.close()
        del wb
        gc.collect()


def _read_cells_com(app, path: Path, items: list[tuple[str, str]]) -> list[str]:
    wb = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
    try:
        vals: list[str] = []
        for sheet_name, addr in items:
            try:
                ws = wb.Worksheets(sheet_name)
            except Exception:
                vals.append(f"{sheet_name}!{addr}")
                continue
            try:
                v = ws.Range(addr).Value2
            except Exception:
                v = None
            txt = _sanitize_name_part("" if v is None else str(v))
            vals.append(txt or f"{sheet_name}!{addr}")
        return vals
    finally:
        wb.Close(SaveChanges=False)


def run_batch_rename_sheet_37(cfg_path: Path, files: list[Path], log_dir: Path) -> dict[str, int]:
    log = get_logger("3_7_batch_rename_sheet", log_dir)
    mapping = _load_sheet_rename_map(cfg_path)
    if not mapping:
        raise ValueError("重命名配置 缺少原表名/新表名映射。")

    wb_count = rename_ok = skip = 0
    for p in files:
        if not p.exists() or not _workbook_supported(p):
            skip += 1
            log.warning("skip file=%s reason=not_supported_or_missing", p)
            continue
        wb = load_workbook(p, keep_vba=(p.suffix.lower() == ".xlsm"))
        changed = False
        try:
            wb_count += 1
            for ws in wb.worksheets:
                old_name = ws.title
                lookup_name = old_name if old_name in mapping else old_name.strip()
                if lookup_name not in mapping:
                    continue
                new_name = mapping[lookup_name]
                if old_name == new_name:
                    skip += 1
                    continue
                if re.search(r"[:\\/?*\[\]]", new_name) or len(new_name) > 31 or (new_name in wb.sheetnames):
                    skip += 1
                    log.warning("skip wb=%s sheet=%s reason=invalid_or_duplicate_new_name", p.name, old_name)
                    continue
                ws.title = new_name
                rename_ok += 1
                changed = True
                log.info("ok wb=%s old=%s new=%s", p.name, old_name, new_name)
            if changed:
                wb.save(p)
        finally:
            wb.close()
    return {"workbooks": wb_count, "rename_ok": rename_ok, "skip": skip}


def run_batch_rename_by_content_46(cfg_path: Path, files: list[Path], log_dir: Path) -> dict[str, int]:
    log = get_logger("4_6_batch_rename_by_content", log_dir)
    spec_text = _load_content_rename_spec(cfg_path)
    items = _parse_sheet_cell_spec(spec_text)

    ok = skip = 0
    com_candidates: list[Path] = []
    for p in files:
        if not p.exists() or not _excel_like(p):
            skip += 1
            log.warning("skip file=%s reason=not_excel_like_or_missing", p)
            continue
        if p.suffix.lower() in (".xls", ".xlt"):
            com_candidates.append(p)
            continue
        try:
            vals = _read_cells_openpyxl(p, items)
            prefix_items = _content_prefix_items(vals)
            if _already_renamed_by_content(p.name, prefix_items):
                skip += 1
                log.info(
                    "skip file=%s reason=already_renamed_by_content items=%s",
                    p.name, "|".join(prefix_items),
                )
                continue
            prefix = "_".join(prefix_items)[:180]
            new_name = f"{prefix}_{p.name}" if prefix else p.name
            tgt = p.with_name(new_name)
            if tgt.resolve() == p.resolve():
                skip += 1
                log.info("skip file=%s reason=name_same", p.name)
                continue
            if tgt.exists():
                tgt = _next_path(tgt)
            os.rename(p, tgt)
            ok += 1
            log.info("ok old=%s new=%s", p.name, tgt.name)
        except Exception as e:
            skip += 1
            log.warning("skip file=%s reason=rename_failed err=%s", p.name, e)

    if com_candidates:
        app = None
        try:
            app, engine = _start_excel_app()
            for p in com_candidates:
                try:
                    vals = _read_cells_com(app, p, items)
                    prefix_items = _content_prefix_items(vals)
                    if _already_renamed_by_content(p.name, prefix_items):
                        skip += 1
                        log.info(
                            "skip file=%s reason=already_renamed_by_content items=%s engine=%s",
                            p.name, "|".join(prefix_items), engine,
                        )
                        continue
                    prefix = "_".join(prefix_items)[:180]
                    new_name = f"{prefix}_{p.name}" if prefix else p.name
                    tgt = p.with_name(new_name)
                    if tgt.resolve() == p.resolve():
                        skip += 1
                        log.info("skip file=%s reason=name_same", p.name)
                        continue
                    if tgt.exists():
                        tgt = _next_path(tgt)
                    os.rename(p, tgt)
                    ok += 1
                    log.info("ok old=%s new=%s engine=%s", p.name, tgt.name, engine)
                except Exception as e:
                    skip += 1
                    log.warning("skip file=%s reason=com_rename_failed err=%s", p.name, e)
        except Exception as e:
            skip += len(com_candidates)
            log.warning("skip com_files=%s reason=start_excel_failed err=%s", len(com_candidates), e)
        finally:
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
    return {"ok": ok, "skip": skip, "spec": spec_text}


def run_batch_unrename_by_content_47(cfg_path: Path, files: list[Path], log_dir: Path) -> dict[str, int]:
    log = get_logger("4_7_batch_unrename_by_content", log_dir)
    spec_text = _load_content_rename_spec(cfg_path)
    items = _parse_sheet_cell_spec(spec_text)

    ok = skip = 0
    com_candidates: list[Path] = []
    for p in files:
        if not p.exists() or not _excel_like(p):
            skip += 1
            log.warning("skip file=%s reason=not_excel_like_or_missing", p)
            continue
        if p.suffix.lower() in (".xls", ".xlt"):
            com_candidates.append(p)
            continue
        try:
            vals = _read_cells_openpyxl(p, items)
            prefix_items = _content_prefix_items(vals)
            new_name = _remove_content_prefix(p.name, prefix_items)
            if not new_name:
                skip += 1
                log.info(
                    "skip file=%s reason=prefix_not_matched items=%s",
                    p.name, "|".join(prefix_items),
                )
                continue
            tgt = p.with_name(new_name)
            if tgt.exists():
                tgt = _next_path(tgt)
            os.rename(p, tgt)
            ok += 1
            log.info("ok old=%s new=%s", p.name, tgt.name)
        except Exception as e:
            skip += 1
            log.warning("skip file=%s reason=unrename_failed err=%s", p.name, e)

    if com_candidates:
        app = None
        try:
            app, engine = _start_excel_app()
            for p in com_candidates:
                try:
                    vals = _read_cells_com(app, p, items)
                    prefix_items = _content_prefix_items(vals)
                    new_name = _remove_content_prefix(p.name, prefix_items)
                    if not new_name:
                        skip += 1
                        log.info(
                            "skip file=%s reason=prefix_not_matched items=%s engine=%s",
                            p.name, "|".join(prefix_items), engine,
                        )
                        continue
                    tgt = p.with_name(new_name)
                    if tgt.exists():
                        tgt = _next_path(tgt)
                    os.rename(p, tgt)
                    ok += 1
                    log.info("ok old=%s new=%s engine=%s", p.name, tgt.name, engine)
                except Exception as e:
                    skip += 1
                    log.warning("skip file=%s reason=com_unrename_failed err=%s", p.name, e)
        except Exception as e:
            skip += len(com_candidates)
            log.warning("skip com_files=%s reason=start_excel_failed err=%s", len(com_candidates), e)
        finally:
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
    return {"ok": ok, "skip": skip, "spec": spec_text}


def run_batch_rename_local_map_48(cfg_path: Path, files: list[Path], log_dir: Path) -> dict[str, int]:
    log = get_logger("4_8_batch_rename_local_map", log_dir)
    mapping = _load_file_local_rename_map(cfg_path)
    if not mapping:
        raise ValueError("重命名配置 缺少 M/N 列局部文件名映射。")

    ok = skip = 0
    for p in files:
        if not p.exists() or not p.is_file():
            skip += 1
            log.warning("skip file=%s reason=not_file_or_missing", p)
            continue
        try:
            new_name = _apply_local_name_map(p.name, mapping)
            if new_name == p.name:
                skip += 1
                log.info("skip file=%s reason=no_fragment_matched", p.name)
                continue
            tgt = p.with_name(new_name)
            if tgt.exists():
                tgt = _next_path(tgt)
            os.rename(p, tgt)
            ok += 1
            log.info("ok old=%s new=%s", p.name, tgt.name)
        except Exception as e:
            skip += 1
            log.warning("skip file=%s reason=local_map_rename_failed err=%s", p.name, e)
    return {"ok": ok, "skip": skip, "rules": len(mapping)}


def _sheet_name_invalid(new_name: str, existing_names: set[str], old_name: str) -> bool:
    if old_name == new_name:
        return True
    if re.search(r"[:\\/?*\[\]]", new_name):
        return True
    if len(new_name) > 31:
        return True
    if new_name in existing_names:
        return True
    return False


def run_batch_rename_sheet_37_com(cfg_path: Path, files: list[Path], log_dir: Path) -> dict[str, int]:
    log = get_logger("3_7_batch_rename_sheet_com", log_dir)
    mapping = _load_sheet_rename_map(cfg_path)
    if not mapping:
        raise ValueError("重命名配置 缺少原表名/新表名映射。")

    valid_files: list[Path] = []
    skip = 0
    for p in files:
        if not p.exists() or not _excel_like(p):
            skip += 1
            log.warning("skip file=%s reason=not_excel_like_or_missing", p)
            continue
        valid_files.append(p)

    if not valid_files:
        return {"workbooks": 0, "rename_ok": 0, "skip": skip, "engine": ""}

    app, engine = _start_excel_app()
    wb_count = 0
    rename_ok = 0
    try:
        for p in valid_files:
            wb = None
            changed = False
            try:
                wb = app.Workbooks.Open(str(p), UpdateLinks=0, ReadOnly=False)
                wb_count += 1
                sheet_count = wb.Worksheets.Count
                existing_names = set()
                for idx in range(1, sheet_count + 1):
                    existing_names.add(str(wb.Worksheets(idx).Name))

                for idx in range(1, sheet_count + 1):
                    ws = wb.Worksheets(idx)
                    old_name = str(ws.Name)
                    lookup_name = old_name if old_name in mapping else old_name.strip()
                    if lookup_name not in mapping:
                        continue
                    new_name = mapping[lookup_name]
                    names_without_old = {x for x in existing_names if x != old_name}
                    if _sheet_name_invalid(new_name, names_without_old, old_name):
                        skip += 1
                        log.warning(
                            "skip wb=%s sheet=%s reason=invalid_or_duplicate_new_name",
                            p.name, old_name
                        )
                        continue
                    ws.Name = new_name
                    existing_names.discard(old_name)
                    existing_names.add(new_name)
                    rename_ok += 1
                    changed = True
                    log.info("ok wb=%s old=%s new=%s", p.name, old_name, new_name)

                if changed:
                    wb.Save()
            except Exception as e:
                skip += 1
                log.warning("skip wb=%s reason=com_rename_failed err=%s", p.name, e)
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
    finally:
        try:
            app.Quit()
        except Exception:
            pass

    return {"workbooks": wb_count, "rename_ok": rename_ok, "skip": skip, "engine": engine}
