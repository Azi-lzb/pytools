from __future__ import annotations
from pathlib import Path
import pandas as pd

from .config_xlsx import (
    SHEET_GLOBAL, SHEET_TIMELINE_RULE, SHEET_PATH_MAP,
    REQUIRED_GLOBAL_KEYS, SHEET_COL_REQUIRED,
    load_global, _truthy, _to_str, _parse_set_items,
)
from .rule_engine import parse_col_spec


def run_precheck(cfg_path: Path) -> int:
    """对 config.xlsx 当前所需 Sheet + 输入/输出/日志目录做预检查。
    打印报告，返回错误数。"""
    print("\n=== 预校验报告 ===")
    print(f"配置文件: {cfg_path}")
    errs = 0
    warns = 0
    if not cfg_path.exists():
        print(f"  [错误] 配置文件不存在")
        return 1
    try:
        xls = pd.ExcelFile(cfg_path)
    except Exception as e:
        print(f"  [错误] 无法打开配置: {e}")
        return 1

    needed_all = {SHEET_GLOBAL, SHEET_TIMELINE_RULE, SHEET_PATH_MAP}
    for sh in needed_all:
        if sh not in xls.sheet_names:
            print(f"  [错误] 缺少 Sheet: {sh}")
            errs += 1

    if SHEET_GLOBAL in xls.sheet_names:
        df = pd.read_excel(xls, SHEET_GLOBAL, header=0, dtype=object)
        if df.shape[1] < 2:
            print(f"  [错误] {SHEET_GLOBAL} 至少需要 键/值 两列")
            errs += 1
        else:
            keys = {_to_str(k) for k in df.iloc[:, 0].tolist()}
            for k in REQUIRED_GLOBAL_KEYS:
                if k not in keys:
                    print(f"  [错误] {SHEET_GLOBAL} 缺少键: {k}")
                    errs += 1

    for sh, cols in SHEET_COL_REQUIRED.items():
        if sh not in xls.sheet_names:
            continue
        df_head = pd.read_excel(xls, sh, header=0, dtype=object, nrows=0)
        present = set(df_head.columns)
        for c in cols:
            if c not in present:
                print(f"  [错误] {sh} 缺少列: {c}")
                errs += 1

    if errs == 0:
        try:
            g = load_global(cfg_path)
            print(f"  输入目录: {g.input_dir}  存在={g.input_dir.exists()}")
            print(f"  输出目录: {g.output_dir}  (创建={not g.output_dir.exists()})")
            print(f"  日志目录: {g.log_dir}  (创建={not g.log_dir.exists()})")
            if not g.input_dir.exists():
                print(f"  [错误] 输入目录不存在: {g.input_dir}")
                errs += 1
            g.output_dir.mkdir(parents=True, exist_ok=True)
            g.log_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"  [错误] 解析全局配置失败: {e}")
            errs += 1

        # 启用规则数 / 启用任务数 摘要
        if SHEET_TIMELINE_RULE in xls.sheet_names:
            df = pd.read_excel(xls, SHEET_TIMELINE_RULE, header=0, dtype=object)
            n = sum(1 for _, r in df.iterrows() if _truthy(r.get("是否启用")))
            print(f"  时序提取规则: 启用 {n} / 共 {len(df)} 条")

            enabled_names: set[str] = set()
            dup_enabled_names: set[str] = set()
            for i, r in df.iterrows():
                if not _truthy(r.get("是否启用")):
                    continue
                row_no = i + 2
                name = _to_str(r.get("规则名称"))
                wb_kw = _to_str(r.get("工作簿关键字"))
                sh_kw = _to_str(r.get("工作表关键字"))
                row_col_raw = _to_str(r.get("行头列"))
                row_col = parse_col_spec(r.get("行头列"))
                col_rows_raw = _to_str(r.get("列表头行"))
                data_row_start = _to_str(r.get("数据起始行"))
                data_row_end = _to_str(r.get("数据结束行"))
                data_col_start = parse_col_spec(r.get("数据起始列"))
                data_col_end = parse_col_spec(r.get("数据结束列"))
                write_flag = _truthy(r.get("启用目标写入")) if "启用目标写入" in df.columns else True
                target_wb = _to_str(r.get("目标工作簿路径"))
                target_ws = _to_str(r.get("目标工作表"))

                if not name:
                    print(f"  [错误] 时序提取规则 第{row_no}行: 规则名称为空")
                    errs += 1
                else:
                    if name in enabled_names:
                        dup_enabled_names.add(name)
                    enabled_names.add(name)
                if not wb_kw:
                    print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 工作簿关键字为空")
                    errs += 1
                if not sh_kw:
                    print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 工作表关键字为空")
                    errs += 1
                if row_col_raw:
                    if row_col <= 0:
                        print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 行头列非法")
                        errs += 1
                else:
                    # 行头列留空：宽表模式允许，但必须提供起始行/起始列。
                    if not data_row_start:
                        print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 行头列为空时，数据起始行必填")
                        errs += 1
                    if data_col_start <= 0:
                        print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 行头列为空时，数据起始列必填且合法")
                        errs += 1
                    if _to_str(r.get("必含行头")):
                        print(f"  [警告] 时序提取规则 第{row_no}行({name or '未命名'}): 行头列为空时将忽略“必含行头”")
                        warns += 1
                if not col_rows_raw:
                    print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 列表头行为空")
                    errs += 1
                else:
                    vals = [x.strip() for x in col_rows_raw.replace("，", ",").split(",") if x.strip()]
                    if not vals:
                        print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 列表头行非法")
                        errs += 1
                    else:
                        ok = True
                        for v in vals:
                            try:
                                if int(float(v)) <= 0:
                                    ok = False
                                    break
                            except Exception:
                                ok = False
                                break
                        if not ok:
                            print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 列表头行必须为正整数或逗号列表")
                            errs += 1
                if data_row_start:
                    try:
                        if int(float(data_row_start)) <= 0:
                            raise ValueError("<=0")
                    except Exception:
                        print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 数据起始行非法")
                        errs += 1
                if data_row_start and data_row_end:
                    try:
                        if int(float(data_row_end)) < int(float(data_row_start)):
                            print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 数据结束行 < 数据起始行")
                            errs += 1
                    except Exception:
                        print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 数据结束行非法")
                        errs += 1
                if _to_str(r.get("数据起始列")) and data_col_start <= 0:
                    print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 数据起始列非法")
                    errs += 1
                if _to_str(r.get("数据结束列")) and data_col_end <= 0:
                    print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 数据结束列非法")
                    errs += 1
                if data_col_start > 0 and data_col_end > 0 and data_col_end < data_col_start:
                    print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): 数据结束列 < 数据起始列")
                    errs += 1
                if write_flag and (not target_wb or not target_ws):
                    print(f"  [警告] 时序提取规则 第{row_no}行({name or '未命名'}): 启用目标写入=是，但目标工作簿路径/目标工作表未完整填写")
                    warns += 1
                if "set区域" in df.columns:
                    _, set_err = _parse_set_items(r.get("set区域"))
                    if set_err:
                        print(f"  [错误] 时序提取规则 第{row_no}行({name or '未命名'}): {set_err}")
                        errs += 1

            for dup in sorted(dup_enabled_names):
                print(f"  [警告] 时序提取规则: 启用规则名称重复 -> {dup}")
                warns += 1
        if SHEET_PATH_MAP in xls.sheet_names:
            df = pd.read_excel(xls, SHEET_PATH_MAP, header=0, dtype=object)
            n = sum(1 for _, r in df.iterrows() if _truthy(r.get("是否启用")))
            print(f"  路径标准化映射: 启用 {n} / 共 {len(df)} 条")

            timeline_df = pd.read_excel(xls, SHEET_TIMELINE_RULE, header=0, dtype=object) if SHEET_TIMELINE_RULE in xls.sheet_names else pd.DataFrame()
            enabled_rule_names = {
                _to_str(r.get("规则名称"))
                for _, r in timeline_df.iterrows()
                if _truthy(r.get("是否启用")) and _to_str(r.get("规则名称"))
            }
            for i, r in df.iterrows():
                if not _truthy(r.get("是否启用")):
                    continue
                row_no = i + 2
                name = _to_str(r.get("映射名称"))
                apply_rule = _to_str(r.get("适用规则名"))
                target = _to_str(r.get("作用对象")) or "两者"
                mode = _to_str(r.get("匹配方式")) or "精确"
                original = _to_str(r.get("原始路径"))
                standard = _to_str(r.get("标准路径"))

                if not name:
                    print(f"  [错误] 路径标准化映射 第{row_no}行: 映射名称为空")
                    errs += 1
                if target not in ("行头", "列头", "两者"):
                    print(f"  [错误] 路径标准化映射 第{row_no}行({name or '未命名'}): 作用对象必须为 行头/列头/两者")
                    errs += 1
                if mode not in ("精确", "包含"):
                    print(f"  [错误] 路径标准化映射 第{row_no}行({name or '未命名'}): 匹配方式必须为 精确/包含")
                    errs += 1
                if not original:
                    print(f"  [错误] 路径标准化映射 第{row_no}行({name or '未命名'}): 原始路径为空")
                    errs += 1
                if not standard:
                    print(f"  [错误] 路径标准化映射 第{row_no}行({name or '未命名'}): 标准路径为空")
                    errs += 1
                if apply_rule and apply_rule not in ("全部", "*") and apply_rule not in enabled_rule_names:
                    print(f"  [错误] 路径标准化映射 第{row_no}行({name or '未命名'}): 适用规则名[{apply_rule}]未在启用时序规则中找到")
                    errs += 1

    print(f"=== 预校验完成: 错误 {errs} 项，警告 {warns} 项 ===\n")
    return errs
