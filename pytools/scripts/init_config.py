"""生成空白 config.xlsx 模板（含 4 个 Sheet 表头与示例行）。

用法： python pytools/scripts/init_config.py [输出路径]
默认输出 pytools/config/config.xlsx
"""
from __future__ import annotations
import sys
from pathlib import Path
import pandas as pd

# 允许直接脚本运行
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from pytools.core.config_xlsx import (
    SHEET_GLOBAL, SHEET_TIMELINE_RULE, SHEET_PATH_MAP,
    SHEET_DEDUP_TASK,
    TIMELINE_COLS, PATH_MAP_COLS, DEDUP_TASK_COLS, REQUIRED_GLOBAL_KEYS,
)


def build_global_df() -> pd.DataFrame:
    rows = [
        ("输入目录", "input", "源文件根目录（绝对或相对 pytools/）"),
        ("输出目录", "output", "结果输出目录"),
        ("日志目录", "logs", "运行日志目录"),
        ("错误策略", "continue", "continue 或 fail_fast"),
        ("默认编码", "utf-8", "CSV 默认编码"),
        ("源文件扩展名", ".xlsx;.xlsm;.csv", "分号分隔"),
    ]
    return pd.DataFrame(rows, columns=["键", "值", "备注"])


def build_timeline_df() -> pd.DataFrame:
    sample = {c: "" for c in TIMELINE_COLS}
    sample.update({
        "是否启用": "否",
        "规则名称": "示例规则",
        "工作簿关键字": "存款",
        "工作表关键字": "本外币",
        "行头列": 1,
        "列表头行": "2",
        "数据起始行": 3,
        "数据起始列": 2,
        "跳过关键字": "合计;小计",
        "启用目标写入": "否",
    })
    return pd.DataFrame([sample], columns=TIMELINE_COLS)


def build_pathmap_df() -> pd.DataFrame:
    sample = ["否", "示例映射", "示例规则", "", "", "行头", "精确", "原路径示例", "标准路径示例", ""]
    return pd.DataFrame([sample], columns=PATH_MAP_COLS)

def build_dedup_df() -> pd.DataFrame:
    sample = ["N", r"C:\Users\AZI\Desktop\demo_source.xlsx", "源数据", "1;2;5",
              r"C:\Users\AZI\Desktop\demo_target.xlsx", "汇总结果", "1",
              "示例：按第1、2、5列联合去重。留空则默认全列。"]
    return pd.DataFrame([sample], columns=DEDUP_TASK_COLS)


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else (
        Path(__file__).resolve().parent.parent / "config" / "config.xlsx"
    )
    out.parent.mkdir(parents=True, exist_ok=True)
    sheets = {
        SHEET_GLOBAL: build_global_df(),
        SHEET_TIMELINE_RULE: build_timeline_df(),
        SHEET_PATH_MAP: build_pathmap_df(),
        SHEET_DEDUP_TASK: build_dedup_df(),
    }
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        for sn, df in sheets.items():
            df.to_excel(w, sheet_name=sn, index=False)
    print(f"已生成: {out}")
    print(f"必填 {SHEET_GLOBAL} 键: {', '.join(REQUIRED_GLOBAL_KEYS)}")


if __name__ == "__main__":
    main()
