from __future__ import annotations
import sys
import os
import time
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pytools.core.config_xlsx import (
    validate_sheets_for_feature, load_global,
    load_timeline_rules, load_path_maps, load_dedup_tasks, load_print_tasks, load_global_value,
    load_institution_mapping, load_foreign_banks, load_extract_tasks,
)
from pytools.core.precheck_399 import run_precheck
from pytools.core.compare_393 import run_compare
from pytools.core.compare_393_com import run_compare_com
from pytools.core.timeline_396 import run_timeline_slim
from pytools.core.timeline_396_com import run_timeline_slim_com
from pytools.core.wide_397_398 import run_wide_summary
from pytools.core.wide_397_398_com import run_wide_summary_com
from pytools.core.dedup_311 import (
    DedupTask,
    run_check_by_comment,
    run_delete_by_comment,
    run_check_by_config,
    run_delete_by_config,
    run_append_by_config,
)
from pytools.core.config_init import initialize_or_repair_config
from pytools.core.print_310 import (
    run_print_keep_by_comment,
    run_print_fast_by_comment,
    run_print_config_all_modes,
    run_print_config_precheck,
)
from pytools.core.print_310_com import (
    run_print_keep_by_comment_com,
    run_print_fast_by_comment_com,
    run_print_config_all_modes_com,
    run_print_config_to_pdf_com,
)
from pytools.core.summary_321_322 import (
    run_summary_by_usedrange,
    run_summary_by_comment,
)
from pytools.core.summary_321_322_com import (
    run_summary_by_usedrange_com,
    run_summary_by_comment_com,
)
from pytools.core.survey_stats import run_survey_stats
from pytools.core.convert_33_34_35_37 import (
    run_batch_rename_33,
    run_excel_convert_34,
    run_word_convert_36,
    run_batch_rename_sheet_37,
    run_batch_rename_sheet_37_com,
)
from pytools.core.feature_1_1_split_village_bank import run_split_village_bank
from pytools.core.feature_1_2_normalize_institution import run_normalize_institution
from pytools.core.feature_1_3_remove_foreign_bank import run_remove_foreign_bank
from pytools.core.feature_1_4_fx_header import run_fx_header_fix
from pytools.core.feature_1_5_region_sum_check import run_region_sum_check
from pytools.core.feature_1_6_extract_sheets import run_extract_sheets
from pytools.core.feature_1_8_adjust_rural_loan import run_adjust_rural_loan
from pytools.core.feature_1_4_fx_header_com import run_fx_header_fix_com
from pytools.core.feature_1_6_extract_sheets_com import run_extract_sheets_com
from pytools.core.feature_1_1_split_village_bank_com import run_split_village_bank_com
from pytools.core.feature_1_2_normalize_institution_com import run_normalize_institution_com
from pytools.core.feature_1_3_remove_foreign_bank_com import run_remove_foreign_bank_com
from pytools.core.feature_1_5_region_sum_check_com import run_region_sum_check_com
from pytools.core.feature_1_8_adjust_rural_loan_com import run_adjust_rural_loan_com

def _app_base_dir() -> Path:

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CFG_PATH = _app_base_dir() / "config" / "config.xlsx"

MAIN_MENU = """
======== pytools 菜单 ========
  1  时序工具
  2  去重工具
  3  打印工具
  4  转换工具
  5  配置工具
  6  汇总工具
  7  一二批处理工具
  8  调研选项统计
  9  快捷工具
  0  退出
==============================
"""

PENDING_MENU = """
------ 一二批处理工具 ------
  1  [自动路由COM] 拆分村镇银行数据
  2  [自动路由COM] 机构名称标准化
  3  [自动路由COM] 删除分机构表的外资行
  4  [自动路由COM] 外汇页眉修改
  5  [自动路由COM] 地区总分校验
  6  [自动路由COM] 提取工作表数据
  7  [自动路由COM] 涉农贷款比上月修正
  0  返回
------------------------
"""

TIMELINE_MENU = """
------ 时序工具 ------
  1  表头路径比对（混合）
  2  时序提取（快速精简）
  3  宽表汇总（行头加后缀）
  4  宽表汇总（行头不加后缀）
  5  预校验
  6  [自动路由COM] 表头路径比对（混合）
  7  [自动路由COM] 时序提取（快速精简）
  8  [自动路由COM] 宽表汇总（行头加后缀）
  9  [自动路由COM] 宽表汇总（行头不加后缀）
  0  返回
----------------------
"""

DEDUP_MENU = """
------ 去重工具 ------
  1  按批注检查重复
  2  按批注删除重复
  3  按配置检查重复
  4  按配置删除重复
  5  按配置去重追加到目标
  6  按配置预校验
  0  返回
----------------------
"""

CONFIG_MENU = """
------ 配置工具 ------
  1  初始化/修复配置表
  2  查看当前配置（重新读取磁盘）
  0  返回
----------------------
"""

CONVERT_MENU = """
------ 转换工具 ------
  1  批量重命名文件
  2  [COM] 批量Excel格式转换
  3  [COM] 批量Word格式转换
  4  批量修改Sheet名
  5  [自动路由COM] 批量修改Sheet名
  0  返回
----------------------
"""

SUMMARY_MENU = """
------ 汇总工具 ------
  1  按使用区域汇总（全量）
  2  按批注汇总（模板+源文件）
  3  [自动路由COM] 按使用区域汇总（全量）
  4  [自动路由COM] 按批注汇总（模板+源文件）
  0  返回
----------------------
"""

SURVEY_MENU = """
------ 调研选项统计 ------
  1  按题目选项比例统计 + 展示版（按 config 中『问卷统计配置』执行）
  0  返回
----------------------
"""

QUICK_MENU = """
------ 快捷工具 ------
  1  打开配置文件（config.xlsx）
  2  打开 output 最新文件
  0  返回
----------------------
"""

PRINT_MENU = """
------ 打印工具 ------
  1  按批注打印（保留源格式）
  2  按批注打印（快速复制）
  3  按配置打印（执行全部模式）
  4  按配置打印预校验
  5  [自动路由COM] 按批注打印（保留源格式）
  6  [自动路由COM] 按批注打印（快速复制）
  7  [自动路由COM] 按配置打印（执行全部模式）
  8  [自动路由COM] 按配置打印 → 导出 PDF（打印预览）
  0  返回
----------------------
"""


def _pick_files_for_compare() -> tuple[Path | None, list[Path]]:
    """弹文件选择框：先模板，再多选源文件。"""
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    print("→ 请在弹出窗口中选择 [模板文件]")
    tmpl = filedialog.askopenfilename(
        title="选择模板文件",
        filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("所有", "*.*")],
    )
    if not tmpl:
        root.destroy()
        return None, []
    print(f"  模板: {tmpl}")
    print("→ 请在弹出窗口中选择 [源文件]（可多选）")
    srcs = filedialog.askopenfilenames(
        title="选择源文件（可多选）",
        filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("所有", "*.*")],
    )
    root.destroy()
    src_paths = [Path(p) for p in srcs]
    print(f"  源文件 {len(src_paths)} 个")
    return Path(tmpl), src_paths


def _pick_source_files(title: str = "选择源文件（可多选）") -> list[Path]:
    print(f"→ 正在打开文件选择窗口：{title}")
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        root.lift()
        try:
            root.focus_force()
        except Exception:
            pass
        root.update()
        srcs = filedialog.askopenfilenames(
            title=title,
            filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.xls *.csv"), ("所有", "*.*")],
        )
        root.destroy()
        picked = [Path(p) for p in srcs]
        if picked:
            print(f"→ 已选择 {len(picked)} 个文件")
            return picked
    except Exception as e:
        print(f"[提示] 文件弹窗不可用：{e}")

    print("→ 未通过弹窗选到文件，可手动输入路径（分号分隔），或输入 0 取消：")
    raw = input("文件路径> ").strip()
    if raw in ("", "0"):
        return []
    parts = [x.strip().strip('"').strip("'") for x in raw.replace("；", ";").split(";") if x.strip()]
    out: list[Path] = []
    for p in parts:
        pp = Path(p)
        if pp.exists():
            out.append(pp)
    return out


def _split_sources_for_auto_route(paths: list[Path]) -> tuple[list[Path], list[Path]]:
    """自动路由：.xls 走 COM；其余走非 COM。"""
    com_list: list[Path] = []
    normal_list: list[Path] = []
    for p in paths:
        if p.suffix.lower() == ".xls":
            com_list.append(p)
        else:
            normal_list.append(p)
    return normal_list, com_list


def _open_path(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(str(path))
    os.startfile(str(path))


def _latest_file_in_dir(dir_path: Path) -> Path | None:
    if not dir_path.exists():
        return None
    files = [p for p in dir_path.rglob("*") if p.is_file() and not p.name.startswith("~$")]
    if not files:
        return None
    return max(files, key=lambda p: p.stat().st_mtime)


def _split_print_tasks_for_auto_route(tasks: list[dict]) -> tuple[list[dict], list[dict]]:
    """打印配置自动路由：.xls 走 COM；其余走非 COM。"""
    normal_list: list[dict] = []
    com_list: list[dict] = []
    for t in tasks:
        p = t.get("source_wb")
        suffix = str(getattr(p, "suffix", "")).lower() if p is not None else ""
        if suffix == ".xls":
            com_list.append(t)
        else:
            normal_list.append(t)
    return normal_list, com_list


def _show_main_menu() -> str:
    print(MAIN_MENU)
    while True:
        try:
            s = input("主菜单选择> ").strip()
        except EOFError:
            return "0"
        if s in ("1", "2", "3", "4", "5", "6", "7", "8", "9", "0"):
            return s
        print("仅接受 1/2/3/4/5/6/7/8/9/0")


def _show_sub_menu(title: str, valid: tuple[str, ...]) -> str:
    print(title)
    while True:
        try:
            s = input("子菜单选择> ").strip()
        except EOFError:
            return "0"
        if s in valid:
            return s
        print("输入无效，请重试。")


def _pick_one_workbook(title: str) -> Path | None:
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    p = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("所有", "*.*")],
    )
    root.destroy()
    if not p:
        return None
    return Path(p)


def _to_dedup_tasks(raw_tasks: list[dict]) -> list[DedupTask]:
    tasks: list[DedupTask] = []
    for t in raw_tasks:
        if t.get("source_wb") is None:
            continue
        tasks.append(DedupTask(
            row_no=int(t.get("row_no") or 0),
            enabled=bool(t.get("enabled")),
            source_wb=t["source_wb"],
            source_ws=str(t.get("source_ws") or ""),
            key_cols=list(t.get("key_cols") or []),
            target_wb=t.get("target_wb"),
            target_ws=t.get("target_ws"),
            write_mode=str(t.get("write_mode") or ""),
            task_name=str(t.get("task_name") or ""),
        ))
    return tasks


def _view_config() -> None:
    import pandas as pd
    print(f"配置文件: {CFG_PATH}")
    if not CFG_PATH.exists():
        print("  [缺失] 请先 5-1 初始化/修复配置表")
        return
    try:
        mtime = CFG_PATH.stat().st_mtime
        from datetime import datetime
        print(f"最后修改: {datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')}")
    except Exception:
        pass

    try:
        xls = pd.ExcelFile(CFG_PATH)
    except Exception as e:
        print(f"  [读取失败] {e}")
        return
    print(f"Sheet 列表（{len(xls.sheet_names)}）:")
    for s in xls.sheet_names:
        try:
            df = pd.read_excel(xls, s, header=0, dtype=object)
            print(f"  - {s}: {len(df)} 行 × {df.shape[1]} 列")
        except Exception as e:
            print(f"  - {s}: 读取失败 {e}")

    try:
        g = load_global(CFG_PATH)
        print("全局配置解析:")
        print(f"  输入目录: {g.input_dir}")
        print(f"  输出目录: {g.output_dir}")
        print(f"  日志目录: {g.log_dir}")
        print(f"  源文件扩展名: {g.source_exts}")
    except Exception as e:
        print(f"  [全局配置解析失败] {e}")

    try:
        rules = load_timeline_rules(CFG_PATH)
        print(f"时序提取规则（启用）: {len(rules)} 条")
        for r in rules:
            print(f"  · {r.name}  wb={r.wb_keyword}  sheet={r.sheet_keyword}")
    except Exception:
        pass

    try:
        mapping = load_institution_mapping(CFG_PATH)
        foreign = load_foreign_banks(CFG_PATH)
        print(f"机构映射: {len(mapping)} 条；外资行: {len(foreign)} 个")
    except Exception:
        pass

    try:
        tasks = load_extract_tasks(CFG_PATH)
        print(f"工作表提取任务（启用）: {len(tasks)} 条")
    except Exception:
        pass


def _check(feature: str) -> bool:
    errs = validate_sheets_for_feature(feature, CFG_PATH)
    if errs:
        print("[配置缺失，已退出本功能]")
        for e in errs:
            print("  - " + e)
        return False
    return True


def dispatch(choice: str) -> None:
    if choice == "c":
        stat = initialize_or_repair_config(CFG_PATH)
        print(
            f"[完成] 配置修复: 新建Sheet={stat.get('created_sheets', 0)} "
            f"修复Sheet={stat.get('repaired_sheets', 0)} 文件={CFG_PATH}"
        )
        for line in stat.get("details", []) or []:
            print(f"  - {line}")
        return
    if choice == "cv":
        _view_config()
        return
    if choice == "q1":
        try:
            _open_path(CFG_PATH)
            print(f"[完成] 已打开配置文件: {CFG_PATH}")
        except Exception as e:
            print(f"[失败] 打开配置文件失败: {e}")
        return
    if choice == "q2":
        g = load_global(CFG_PATH)
        latest = _latest_file_in_dir(g.output_dir)
        if latest is None:
            print(f"[提示] 未找到输出文件: {g.output_dir}")
            return
        try:
            _open_path(latest)
            print(f"[完成] 已打开最新输出文件: {latest}")
        except Exception as e:
            print(f"[失败] 打开输出文件失败: {e}")
        return
    if choice == "5":
        run_precheck(CFG_PATH)
        return
    if not _check(choice):
        return
    g = load_global(CFG_PATH)
    g.output_dir.mkdir(parents=True, exist_ok=True)
    g.log_dir.mkdir(parents=True, exist_ok=True)
    hide_zero_global = str(
        load_global_value(CFG_PATH, "打印零值不输出", "")
        or load_global_value(CFG_PATH, "打印零值不显示", "")
    ).strip().lower() in ("是", "1", "true", "y", "yes", "on")

    if choice == "1":
        rules = load_timeline_rules(CFG_PATH)
        if not rules:
            print("[无启用的时序提取规则]")
            return
        tmpl, srcs = _pick_files_for_compare()
        if not tmpl or not srcs:
            print("[已取消]")
            return
        out = run_compare(tmpl, srcs, rules, g)
        print(f"[完成] 输出: {out}")
    elif choice == "1com":
        rules = load_timeline_rules(CFG_PATH)
        if not rules:
            print("[无启用的时序提取规则]")
            return
        tmpl, srcs = _pick_files_for_compare()
        if not tmpl or not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        out_normal = None
        out_com = None
        if normal_srcs:
            out_normal = run_compare(tmpl, normal_srcs, rules, g)
        if com_srcs:
            try:
                out_com = run_compare_com(tmpl, com_srcs, rules, g)
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
        if out_normal is None and out_com is None:
            print("[完成] 无匹配结果，未生成文件")
        else:
            if out_normal is not None:
                print(f"[完成] 非COM输出: {out_normal}")
            if out_com is not None:
                print(f"[完成] COM输出: {out_com}")
    elif choice == "2com":
        rules = load_timeline_rules(CFG_PATH)
        path_maps = load_path_maps(CFG_PATH)
        if not rules:
            print("[无启用的时序提取规则]")
            return
        srcs = _pick_source_files("选择时序提取源文件（自动路由：xls走COM，其它走非COM）")
        if not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        out_normal = None
        out_com = None
        if normal_srcs:
            out_normal = run_timeline_slim(rules, path_maps, g, source_paths=normal_srcs)
        if com_srcs:
            try:
                out_com = run_timeline_slim_com(rules, path_maps, g, source_paths=com_srcs)
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
        if out_normal is None and out_com is None:
            print("[完成] 无匹配结果，未生成文件")
        else:
            if out_normal is not None:
                print(f"[完成] 非COM输出: {out_normal}")
            if out_com is not None:
                print(f"[完成] COM输出: {out_com}")
    elif choice in ("3com", "4com"):
        rules = load_timeline_rules(CFG_PATH)
        path_maps = load_path_maps(CFG_PATH)
        if not rules:
            print("[无启用的时序提取规则]")
            return
        srcs = _pick_source_files("选择宽表汇总源文件（自动路由：xls走COM，其它走非COM）")
        if not srcs:
            print("[已取消]")
            return
        suffix = (choice == "3com")
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        out_normal = None
        out_com = None
        if normal_srcs:
            out_normal = run_wide_summary(
                rules, path_maps, g, row_suffix_enabled=suffix, source_paths=normal_srcs
            )
        if com_srcs:
            try:
                out_com = run_wide_summary_com(
                    rules, path_maps, g, row_suffix_enabled=suffix, source_paths=com_srcs
                )
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
        if out_normal is None and out_com is None:
            print("[完成] 无匹配结果，未生成文件")
        else:
            if out_normal is not None:
                print(f"[完成] 非COM输出: {out_normal}")
            if out_com is not None:
                print(f"[完成] COM输出: {out_com}")
    elif choice == "2":
        rules = load_timeline_rules(CFG_PATH)
        path_maps = load_path_maps(CFG_PATH)
        if not rules:
            print("[无启用的时序提取规则]")
            return
        srcs = _pick_source_files("选择时序提取源文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        out = run_timeline_slim(rules, path_maps, g, source_paths=srcs)
        if out is None:
            print("[完成] 无匹配结果，未生成文件")
        else:
            print(f"[完成] 输出: {out}")
    elif choice in ("3", "4"):
        rules = load_timeline_rules(CFG_PATH)
        path_maps = load_path_maps(CFG_PATH)
        if not rules:
            print("[无启用的时序提取规则]")
            return
        srcs = _pick_source_files("选择宽表汇总源文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        suffix = (choice == "3")
        t0 = time.perf_counter()
        out = run_wide_summary(rules, path_maps, g, row_suffix_enabled=suffix, source_paths=srcs)
        elapsed = time.perf_counter() - t0
        if out is None:
            print(f"[完成] 无匹配结果，未生成文件（耗时 {elapsed:.2f}s）")
        else:
            print(f"[完成] 输出: {out}（耗时 {elapsed:.2f}s）")
    elif choice == "6":
        picked = _pick_one_workbook("选择要按批注检查重复的工作簿")
        if not picked:
            print("[已取消]")
            return
        stat = run_check_by_comment(picked, g.log_dir)
        print(f"[完成] 命中Sheet={stat['hit_sheets']} 标红行={stat['marked_rows']}")
    elif choice == "7":
        picked = _pick_one_workbook("选择要按批注删除重复的工作簿")
        if not picked:
            print("[已取消]")
            return
        stat = run_delete_by_comment(picked, g.log_dir)
        print(f"[完成] 命中Sheet={stat['hit_sheets']} 删除行={stat['deleted_rows']}")
    elif choice == "8":
        tasks = _to_dedup_tasks(load_dedup_tasks(CFG_PATH))
        stat = run_check_by_config(tasks, g.log_dir)
        print(
            f"[完成] 任务成功={stat['task_ok']} 跳过={stat['task_skip']} "
            f"标红行={stat['marked_rows']} 耗时={float(stat.get('elapsed_sec', 0.0)):.2f}s"
        )
    elif choice == "9":
        tasks = _to_dedup_tasks(load_dedup_tasks(CFG_PATH))
        stat = run_delete_by_config(tasks, g.log_dir)
        print(
            f"[完成] 任务成功={stat['task_ok']} 跳过={stat['task_skip']} "
            f"删除行={stat['deleted_rows']} 耗时={float(stat.get('elapsed_sec', 0.0)):.2f}s"
        )
    elif choice == "a":
        tasks = _to_dedup_tasks(load_dedup_tasks(CFG_PATH))
        stat = run_append_by_config(tasks, g.log_dir)
        print(
            f"[完成] 任务成功={stat['task_ok']} 跳过={stat['task_skip']} "
            f"追加={stat['appended_rows']} 去重删除={stat['deleted_rows']} "
            f"耗时={float(stat.get('elapsed_sec', 0.0)):.2f}s"
        )
    elif choice == "b":
        tasks = _to_dedup_tasks(load_dedup_tasks(CFG_PATH))
        stat = run_precheck_by_config(tasks, g.log_dir)
        print(f"[完成] 预校验 OK={stat['ok']} WARN={stat['warn']} FAIL={stat['fail']}")
    elif choice == "p1":
        srcs = _pick_source_files("选择按批注打印（保留源格式）源文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        stat = run_print_keep_by_comment(srcs, g.output_dir, g.log_dir, hide_zero=hide_zero_global)
        print(
            f"[完成] 命中工作簿={stat['workbooks_hit']} 命中sheet={stat['sheets_hit']} "
            f"输出文件={stat['saved_files']} 跳过={stat['skipped']}"
        )
    elif choice == "p3":
        srcs = _pick_source_files("选择 3.10.3 源文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        stat = run_print_fast_by_comment(srcs, g.output_dir, g.log_dir, hide_zero=hide_zero_global)
        print(
            f"[完成] 命中工作簿={stat['workbooks_hit']} 命中sheet={stat['sheets_hit']} "
            f"输出文件={stat['saved_files']} 跳过={stat['skipped']}"
        )
    elif choice == "p7":
        tasks = load_print_tasks(CFG_PATH)
        stat = run_print_config_all_modes(tasks, g.log_dir)
        print(
            f"[完成] 任务成功={stat['task_ok']} 跳过={stat['task_skip']} "
            f"写入sheet={stat['written_sheets']} 写入行={stat['written_rows']}"
        )
        total = int(stat.get("task_total", 0) or 0)
        elapsed = float(stat.get("elapsed_sec", 0.0) or 0.0)
        avg = (elapsed / total) if total > 0 else 0.0
        print(f"[耗时] 总耗时={elapsed:.2f}s 平均每任务={avg:.2f}s (任务总数={total})")
        slow_top = stat.get("slow_top", []) or []
        if slow_top:
            print("[耗时] 最慢任务TOP3:")
            for i, r in enumerate(slow_top, start=1):
                print(
                    f"  {i}. row={r.get('row_no')} mode={r.get('mode')} "
                    f"target={r.get('target')} 耗时={float(r.get('elapsed_sec', 0.0)):.2f}s "
                    f"status={r.get('status')}"
                )
    elif choice == "p8":
        tasks = load_print_tasks(CFG_PATH)
        stat = run_print_config_precheck(tasks, g.log_dir)
        print(f"[完成] 预校验 OK={stat['ok']} WARN={stat['warn']} FAIL={stat['fail']}")
    elif choice == "p1com":
        srcs = _pick_source_files("选择 [COM] 按批注打印（保留源格式）源文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        if normal_srcs:
            stat = run_print_keep_by_comment(
                normal_srcs, g.output_dir, g.log_dir, hide_zero=hide_zero_global
            )
            print(
                f"[完成] 非COM 命中工作簿={stat['workbooks_hit']} 命中sheet={stat['sheets_hit']} "
                f"输出文件={stat['saved_files']} 跳过={stat['skipped']}"
            )
        if com_srcs:
            try:
                stat = run_print_keep_by_comment_com(
                    com_srcs, g.output_dir, g.log_dir, hide_zero=hide_zero_global
                )
                print(
                    f"[完成] COM 命中工作簿={stat['workbooks_hit']} 命中sheet={stat['sheets_hit']} "
                    f"输出文件={stat['saved_files']} 跳过={stat['skipped']}"
                )
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "p3com":
        srcs = _pick_source_files("选择 [COM] 按批注打印（快速复制）源文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        if normal_srcs:
            stat = run_print_fast_by_comment(
                normal_srcs, g.output_dir, g.log_dir, hide_zero=hide_zero_global
            )
            print(
                f"[完成] 非COM 命中工作簿={stat['workbooks_hit']} 命中sheet={stat['sheets_hit']} "
                f"输出文件={stat['saved_files']} 跳过={stat['skipped']}"
            )
        if com_srcs:
            try:
                stat = run_print_fast_by_comment_com(
                    com_srcs, g.output_dir, g.log_dir, hide_zero=hide_zero_global
                )
                print(
                    f"[完成] COM 命中工作簿={stat['workbooks_hit']} 命中sheet={stat['sheets_hit']} "
                    f"输出文件={stat['saved_files']} 跳过={stat['skipped']}"
                )
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "p7com":
        tasks = load_print_tasks(CFG_PATH)
        normal_tasks, com_tasks = _split_print_tasks_for_auto_route(tasks)
        print(f"→ 自动路由：非COM任务={len(normal_tasks)}，COM任务(.xls)={len(com_tasks)}")
        if normal_tasks:
            stat = run_print_config_all_modes(normal_tasks, g.log_dir)
            print(
                f"[完成] 非COM 任务成功={stat['task_ok']} 跳过={stat['task_skip']} "
                f"写入sheet={stat['written_sheets']} 写入行={stat['written_rows']}"
            )
        if com_tasks:
            try:
                stat = run_print_config_all_modes_com(com_tasks, g.log_dir)
                print(
                    f"[完成] COM 任务成功={stat['task_ok']} 跳过={stat['task_skip']} "
                    f"写入sheet={stat['written_sheets']} 写入行={stat['written_rows']}"
                )
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "ppdf":
        tasks = load_print_tasks(CFG_PATH)
        normal_tasks, com_tasks = _split_print_tasks_for_auto_route(tasks)
        print(f"→ 自动路由：非COM任务={len(normal_tasks)}，COM任务(.xls)={len(com_tasks)}")
        if normal_tasks:
            print("[提示] 非COM任务不支持直接导出PDF，已按非COM常规打印执行。")
            stat = run_print_config_all_modes(normal_tasks, g.log_dir)
            print(
                f"[完成] 非COM 任务成功={stat['task_ok']} 跳过={stat['task_skip']} "
                f"写入sheet={stat['written_sheets']} 写入行={stat['written_rows']}"
            )
        if com_tasks:
            try:
                stat = run_print_config_to_pdf_com(com_tasks, g.output_dir, g.log_dir)
                print(
                    f"[完成] COM(PDF) 任务成功={stat['task_ok']} 跳过={stat['task_skip']} "
                    f"写入sheet={stat['written_sheets']} 写入行={stat['written_rows']}"
                )
                for p in stat["pdf_files"]:
                    print(f"  PDF -> {p}")
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "t3":
        srcs = _pick_source_files("选择要批量重命名的文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        stat = run_batch_rename_33(CFG_PATH, srcs, g.log_dir)
        print(f"[完成] 成功={stat['ok']} 跳过={stat['skip']}")
    elif choice == "t4":
        srcs = _pick_source_files("选择要批量转换格式的文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        target_fmt = (
            load_global_value(CFG_PATH, "excel目的格式", "")
            or load_global_value(CFG_PATH, "3.4目的格式", "xlsx")
            or "xlsx"
        )
        stat = run_excel_convert_34(srcs, target_fmt, g.log_dir)
        print(
            f"[完成] 目标格式={stat['target_ext']} 引擎={stat.get('engine','')} "
            f"成功={stat['ok']} 跳过/失败={stat['skip']}"
        )
    elif choice == "t5":
        srcs = _pick_source_files("选择要批量转换Word格式的文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        target_fmt = (
            load_global_value(CFG_PATH, "word目的格式", "")
            or load_global_value(CFG_PATH, "3.6目的格式", "docx")
            or "docx"
        )
        stat = run_word_convert_36(srcs, target_fmt, g.log_dir)
        print(
            f"[完成] 目标格式={stat['target_ext']} 引擎={stat.get('engine','')} "
            f"成功={stat['ok']} 跳过/失败={stat['skip']}"
        )
    elif choice == "t7":
        srcs = _pick_source_files("选择要批量修改Sheet名的工作簿（可多选）")
        if not srcs:
            print("[已取消]")
            return
        stat = run_batch_rename_sheet_37(CFG_PATH, srcs, g.log_dir)
        print(f"[完成] 工作簿={stat['workbooks']} 重命名={stat['rename_ok']} 跳过={stat['skip']}")
    elif choice == "t7com":
        srcs = _pick_source_files("选择 [COM] 要批量修改Sheet名的工作簿（可多选）")
        if not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        if normal_srcs:
            stat = run_batch_rename_sheet_37(CFG_PATH, normal_srcs, g.log_dir)
            print(
                f"[完成] 非COM 工作簿={stat['workbooks']} "
                f"重命名={stat['rename_ok']} 跳过={stat['skip']}"
            )
        if com_srcs:
            try:
                stat = run_batch_rename_sheet_37_com(CFG_PATH, com_srcs, g.log_dir)
                print(
                    f"[完成] COM 引擎={stat.get('engine','')} "
                    f"工作簿={stat['workbooks']} 重命名={stat['rename_ok']} 跳过={stat['skip']}"
                )
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "s21":
        srcs = _pick_source_files("选择按使用区域汇总的源文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        stat = run_summary_by_usedrange(srcs, g.output_dir, g.log_dir)
        if not stat.get("saved"):
            print("[完成] 无可汇总数据，未生成文件")
        else:
            print(f"[完成] 行数={stat['rows']} 输出={stat['path']}")
    elif choice == "s21com":
        srcs = _pick_source_files("选择 [COM] 按使用区域汇总的源文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        if normal_srcs:
            stat = run_summary_by_usedrange(normal_srcs, g.output_dir, g.log_dir)
            if not stat.get("saved"):
                print("[完成] 非COM 无可汇总数据，未生成文件")
            else:
                print(f"[完成] 非COM 行数={stat['rows']} 输出={stat['path']}")
        if com_srcs:
            try:
                stat = run_summary_by_usedrange_com(com_srcs, g.output_dir, g.log_dir)
                if not stat.get("saved"):
                    print("[完成] COM 无可汇总数据，未生成文件")
                else:
                    print(f"[完成] COM 行数={stat['rows']} 输出={stat['path']}")
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "s22com":
        tmpl = _pick_one_workbook("选择 [COM] 批注汇总模板文件")
        if not tmpl:
            print("[已取消]")
            return
        srcs = _pick_source_files("选择 [COM] 按批注汇总的源文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        if normal_srcs:
            stat = run_summary_by_comment(tmpl, normal_srcs, g.output_dir, g.log_dir)
            if not stat.get("saved"):
                print("[完成] 非COM 无可汇总数据，未生成文件")
            else:
                print(f"[完成] 非COM 行数={stat['rows']} 输出={stat['path']}")
        if com_srcs:
            try:
                stat = run_summary_by_comment_com(tmpl, com_srcs, g.output_dir, g.log_dir)
                if not stat.get("saved"):
                    print(
                        f"[完成] COM 无可汇总数据，未生成文件 "
                        f"(命中sheet={stat.get('sheets_ok',0)} 跳过sheet={stat.get('sheets_skip',0)})"
                    )
                else:
                    print(
                        f"[完成] COM 行数={stat['rows']} 输出={stat['path']} "
                        f"(命中sheet={stat.get('sheets_ok',0)} 跳过sheet={stat.get('sheets_skip',0)})"
                    )
                errs = stat.get("errors", []) or []
                if errs:
                    print(f"[COM告警] 跳过异常sheet {len(errs)} 个（示例前{min(3, len(errs))}条）：")
                    for s in errs[:3]:
                        print(f"  - {s}")
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "sv1":
        src = _pick_one_workbook("选择问卷汇总输入文件（通常是 3.2.2 按批注汇总输出）")
        if not src:
            print("[已取消]")
            return
        try:
            stat = run_survey_stats(src, CFG_PATH, g.log_dir)
        except Exception as e:
            print(f"[运行失败] {e}")
            return
        if not stat.get("saved"):
            print("[完成] 未识别到任何题目，未生成文件")
        else:
            print(
                f"[完成] 题数={stat['questions']} 样本机构数={stat['sample']} "
                f"重命名命中/未命中={stat['rename_hits']}/{stat['rename_miss']}\n"
                f"  明细：{stat['detail']}\n"
                f"  展示：{stat['show']}"
            )
    elif choice == "x11":
        curr = _pick_one_workbook("选择本期文件（含村镇银行数据）")
        if not curr:
            print("[已取消]")
            return
        prev = _pick_one_workbook("选择去年同期文件（含村镇银行数据）")
        if not prev:
            print("[已取消]")
            return
        srcs = _pick_source_files("选择包含'分机构'工作表的目标文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        normal_curr = curr if curr.suffix.lower() != ".xls" else None
        normal_prev = prev if prev.suffix.lower() != ".xls" else None
        com_curr = curr if curr.suffix.lower() == ".xls" else None
        com_prev = prev if prev.suffix.lower() == ".xls" else None
        if com_srcs and (com_curr is None or com_prev is None):
            print("[提示] 目标含 .xls，但本期/上期不是 .xls，COM 分支跳过。")
        if normal_srcs:
            stat = run_split_village_bank(curr, prev, normal_srcs, g.log_dir)
            if stat.get("matrix_empty"):
                print("[已中止] 非COM 村镇银行计算结果为空")
            else:
                print(
                    f"[完成] 非COM 文件={stat['files']} 分机构sheet={stat['branch_sheets']} "
                    f"填入={stat['filled_sheets']} 跳过={stat['skipped_sheets']} 失败={stat['failed_files']}"
                )
        if com_srcs and com_curr is not None and com_prev is not None:
            try:
                stat = run_split_village_bank_com(com_curr, com_prev, com_srcs, g.log_dir)
                if stat.get("matrix_empty"):
                    print("[已中止] COM 村镇银行计算结果为空")
                else:
                    print(
                        f"[完成] COM 文件={stat['files']} 分机构sheet={stat['branch_sheets']} "
                        f"填入={stat['filled_sheets']} 跳过={stat['skipped_sheets']} 失败={stat['failed_files']}"
                    )
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "x12":
        mapping = load_institution_mapping(CFG_PATH)
        if not mapping:
            print("[已中止] 机构映射表为空，请在 config.xlsx 的'机构映射表'中填写 A=原始 B=映射")
            return
        srcs = _pick_source_files("选择包含'分机构'工作表的目标文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        if normal_srcs:
            stat = run_normalize_institution(normal_srcs, mapping, g.log_dir)
            print(
                f"[完成] 非COM 文件={stat['files']} 分机构sheet={stat['branch_sheets']} "
                f"命中sheet={stat['mapped_sheets']} 修改行={stat['mapped_rows']} "
                f"失败={stat['failed_files']}"
            )
        if com_srcs:
            try:
                stat = run_normalize_institution_com(com_srcs, mapping, g.log_dir)
                print(
                    f"[完成] COM 文件={stat['files']} 分机构sheet={stat['branch_sheets']} "
                    f"命中sheet={stat['mapped_sheets']} 修改行={stat['mapped_rows']} "
                    f"失败={stat['failed_files']}"
                )
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "x13":
        foreign_set = load_foreign_banks(CFG_PATH)
        if not foreign_set:
            print("[已中止] 外资行配置为空，请在 config.xlsx 的'机构映射表' C 列将外资行标记为 1")
            return
        srcs = _pick_source_files("选择包含'分机构'工作表的目标文件（可多选，将生成(金融局)副本）")
        if not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        if normal_srcs:
            stat = run_remove_foreign_bank(normal_srcs, foreign_set, g.log_dir)
            print(
                f"[完成] 非COM 文件={stat['files']} 成功={stat['ok_files']} "
                f"分机构sheet={stat['branch_sheets']} 删除行={stat['deleted_rows']} "
                f"失败={stat['failed_files']}"
            )
        if com_srcs:
            try:
                stat = run_remove_foreign_bank_com(com_srcs, foreign_set, g.log_dir)
                print(
                    f"[完成] COM 文件={stat['files']} 成功={stat['ok_files']} "
                    f"分机构sheet={stat['branch_sheets']} 删除行={stat['deleted_rows']} "
                    f"失败={stat['failed_files']}"
                )
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "x14":
        srcs = _pick_source_files("选择要修改外汇页眉的Excel文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        stat = run_fx_header_fix(srcs, g.log_dir)
        print(
            f"[完成] 文件={stat['files']} 保存={stat['saved_files']} "
            f"外汇sheet={stat['fx_sheets']} 修改sheet={stat['modified_sheets']} "
            f"失败={stat['failed_files']}"
        )
    elif choice == "x15":
        srcs = _pick_source_files("选择要做地区总分校验的Excel文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        if normal_srcs:
            stat = run_region_sum_check(normal_srcs, g.output_dir, g.log_dir)
            print(
                f"[完成] 非COM 文件={stat['files']} 工作表={stat['sheets']} "
                f"错误={stat['errors']} 失败={stat['failed_files']} 输出={stat['path']}"
            )
        if com_srcs:
            try:
                stat = run_region_sum_check_com(com_srcs, g.output_dir, g.log_dir)
                print(
                    f"[完成] COM 文件={stat['files']} 工作表={stat['sheets']} "
                    f"错误={stat['errors']} 失败={stat['failed_files']} 输出={stat['path']}"
                )
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "x16":
        tasks = load_extract_tasks(CFG_PATH)
        if not tasks:
            print("[已中止] '工作表提取'配置无启用任务")
            return
        srcs = _pick_source_files("选择要提取的源文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        stat = run_extract_sheets(srcs, tasks, g.output_dir, g.log_dir)
        print(
            f"[完成] 源文件={stat['files']} 提取sheet={stat['extracted_sheets']} "
            f"输出文件={stat['output_files']} 失败={stat['failed_files']}"
        )
        for p in stat["paths"]:
            print(f"  -> {p}")
    elif choice == "x14com":
        srcs = _pick_source_files("选择 [COM] 要修改外汇页眉的Excel文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        if normal_srcs:
            stat = run_fx_header_fix(normal_srcs, g.log_dir)
            print(
                f"[完成] 非COM 文件={stat['files']} 保存={stat['saved_files']} "
                f"外汇sheet={stat['fx_sheets']} 修改sheet={stat['modified_sheets']} "
                f"失败={stat['failed_files']}"
            )
        if com_srcs:
            try:
                stat = run_fx_header_fix_com(com_srcs, g.log_dir)
                print(
                    f"[完成] COM 文件={stat['files']} 保存={stat['saved_files']} "
                    f"外汇sheet={stat['fx_sheets']} 修改sheet={stat['modified_sheets']} "
                    f"失败={stat['failed_files']}"
                )
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "x16com":
        tasks = load_extract_tasks(CFG_PATH)
        if not tasks:
            print("[已中止] '工作表提取'配置无启用任务")
            return
        srcs = _pick_source_files("选择 [COM] 要提取的源文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        normal_srcs, com_srcs = _split_sources_for_auto_route(srcs)
        print(f"→ 自动路由：非COM={len(normal_srcs)}，COM(.xls)={len(com_srcs)}")
        if normal_srcs:
            stat = run_extract_sheets(normal_srcs, tasks, g.output_dir, g.log_dir)
            print(
                f"[完成] 非COM 源文件={stat['files']} 提取sheet={stat['extracted_sheets']} "
                f"输出文件={stat['output_files']} 失败={stat['failed_files']}"
            )
            for p in stat["paths"]:
                print(f"  -> {p}")
        if com_srcs:
            try:
                stat = run_extract_sheets_com(com_srcs, tasks, g.output_dir, g.log_dir)
                print(
                    f"[完成] COM 源文件={stat['files']} 提取sheet={stat['extracted_sheets']} "
                    f"输出文件={stat['output_files']} 失败={stat['failed_files']}"
                )
                for p in stat["paths"]:
                    print(f"  -> {p}")
            except RuntimeError as e:
                print(f"[COM 不可用] {e}")
    elif choice == "x18":
        cur = _pick_one_workbook("选择本期 Excel 文件")
        if not cur:
            print("[已取消]")
            return
        prev = _pick_one_workbook("选择上期 Excel 文件")
        if not prev:
            print("[已取消]")
            return
        normal_files, com_files = _split_sources_for_auto_route([cur, prev])
        print(f"→ 自动路由：非COM={len(normal_files)}，COM(.xls)={len(com_files)}")
        try:
            if cur.suffix.lower() == ".xls" or prev.suffix.lower() == ".xls":
                r = run_adjust_rural_loan_com(cur, prev, g.log_dir)
            else:
                r = run_adjust_rural_loan(cur, prev, g.log_dir)
        except RuntimeError as e:
            print(f"[COM 不可用] {e}")
            return
        if r["error"]:
            print(f"[失败] {r['error']}")
        else:
            print(
                f"[完成] 匹配={r['matched']} 分机构合计={r['fill_total']:.2f} "
                f"分地区总数={r['area_total']:.2f} "
                f"{'校验通过' if r['consistent'] else '校验不一致'}"
            )
            if r["unmatched_current"]:
                print("  本期未匹配到上期: " + "、".join(r["unmatched_current"]))
            if r["previous_only"]:
                print("  上期独有: " + "、".join(r["previous_only"]))
    elif choice == "s22":
        tmpl = _pick_one_workbook("选择批注汇总模板文件")
        if not tmpl:
            print("[已取消]")
            return
        srcs = _pick_source_files("选择按批注汇总的源文件（可多选）")
        if not srcs:
            print("[已取消]")
            return
        stat = run_summary_by_comment(tmpl, srcs, g.output_dir, g.log_dir)
        if not stat.get("saved"):
            print("[完成] 无可汇总数据，未生成文件")
        else:
            print(f"[完成] 行数={stat['rows']} 输出={stat['path']}")


def main() -> None:
    while True:
        main_choice = _show_main_menu()
        if main_choice == "0":
            print("再见。")
            return

        if main_choice == "1":
            sub = _show_sub_menu(TIMELINE_MENU, ("1", "2", "3", "4", "5", "6", "7", "8", "9", "0"))
            if sub == "0":
                continue
            # 新菜单 -> 旧分发代号
            mapped = {"1": "1", "2": "2", "3": "3", "4": "4", "5": "5",
                      "6": "1com", "7": "2com", "8": "3com", "9": "4com"}[sub]
        elif main_choice == "2":
            sub = _show_sub_menu(DEDUP_MENU, ("1", "2", "3", "4", "5", "6", "0"))
            if sub == "0":
                continue
            mapped = {"1": "6", "2": "7", "3": "8", "4": "9", "5": "a", "6": "b"}[sub]
        elif main_choice == "3":
            sub = _show_sub_menu(PRINT_MENU, ("1", "2", "3", "4", "5", "6", "7", "8", "0"))
            if sub == "0":
                continue
            mapped = {"1": "p1", "2": "p3", "3": "p7", "4": "p8",
                      "5": "p1com", "6": "p3com", "7": "p7com",
                      "8": "ppdf"}[sub]
        elif main_choice == "4":
            sub = _show_sub_menu(CONVERT_MENU, ("1", "2", "3", "4", "5", "0"))
            if sub == "0":
                continue
            mapped = {"1": "t3", "2": "t4", "3": "t5", "4": "t7", "5": "t7com"}[sub]
        elif main_choice == "5":
            sub = _show_sub_menu(CONFIG_MENU, ("1", "2", "0"))
            if sub == "0":
                continue
            mapped = {"1": "c", "2": "cv"}[sub]
        elif main_choice == "6":
            sub = _show_sub_menu(SUMMARY_MENU, ("1", "2", "3", "4", "0"))
            if sub == "0":
                continue
            mapped = {"1": "s21", "2": "s22", "3": "s21com", "4": "s22com"}[sub]
        elif main_choice == "7":
            sub = _show_sub_menu(PENDING_MENU, ("1", "2", "3", "4", "5", "6", "7", "0"))
            if sub == "0":
                continue
            mapped = {"1": "x11", "2": "x12", "3": "x13", "4": "x14com",
                      "5": "x15", "6": "x16com", "7": "x18"}[sub]
        elif main_choice == "8":
            sub = _show_sub_menu(SURVEY_MENU, ("1", "0"))
            if sub == "0":
                continue
            mapped = {"1": "sv1"}[sub]
        else:
            sub = _show_sub_menu(QUICK_MENU, ("1", "2", "0"))
            if sub == "0":
                continue
            mapped = {"1": "q1", "2": "q2"}[sub]

        try:
            dispatch(mapped)
        except Exception as e:
            print(f"[运行异常] {e}")
            import traceback; traceback.print_exc()

if __name__ == "__main__":
    main()
