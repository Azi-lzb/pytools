"""端到端最小测试：构造一个 mini 仓库结构，跑通 5 个功能。"""
from __future__ import annotations
import shutil
from pathlib import Path
import pandas as pd
import pytest

from pytools.core.config_xlsx import (
    SHEET_GLOBAL, SHEET_TIMELINE_RULE, SHEET_PATH_MAP,
    TIMELINE_COLS, PATH_MAP_COLS,
    load_global, load_timeline_rules, load_path_maps,
    validate_sheets_for_feature,
)
from pytools.core.precheck_399 import run_precheck
from pytools.core.timeline_396 import run_timeline_slim
from pytools.core.wide_397_398 import run_wide_summary
from pytools.core.compare_393 import run_compare


def _make_source(path: Path, sheet_name: str, df_2d: list[list]) -> None:
    df = pd.DataFrame(df_2d)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=sheet_name, header=False, index=False)


def _make_config(cfg_path: Path, input_dir: Path, output_dir: Path, log_dir: Path) -> None:
    g_df = pd.DataFrame([
        ("输入目录", str(input_dir), ""),
        ("输出目录", str(output_dir), ""),
        ("日志目录", str(log_dir), ""),
        ("错误策略", "continue", ""),
        ("默认编码", "utf-8", ""),
        ("源文件扩展名", ".xlsx", ""),
    ], columns=["键", "值", "备注"])

    rule_row = ["是", "存款规则", "存款", "本外币", 1, "2",
                "本月余额", "活期",
                3, "", 2, "", "合计", "", ""]
    t_df = pd.DataFrame([rule_row], columns=TIMELINE_COLS)

    pm_row = ["是", "活期标准化", "存款规则", "", "", "行头", "精确", "活期存款", "活期", ""]
    p_df = pd.DataFrame([pm_row], columns=PATH_MAP_COLS)

    cfg_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(cfg_path, engine="openpyxl") as w:
        g_df.to_excel(w, sheet_name=SHEET_GLOBAL, index=False)
        t_df.to_excel(w, sheet_name=SHEET_TIMELINE_RULE, index=False)
        p_df.to_excel(w, sheet_name=SHEET_PATH_MAP, index=False)


@pytest.fixture
def env(tmp_path: Path):
    base = tmp_path / "pytools_test"
    inp = base / "input"
    out = base / "output"
    logs = base / "logs"
    cfg_dir = base / "config"
    inp.mkdir(parents=True)
    cfg = cfg_dir / "config.xlsx"

    # 模板表（在 input 外）
    tmpl_dir = base / "tmpl"
    tmpl_dir.mkdir()
    tmpl_path = tmpl_dir / "template.xlsx"

    # source A：本外币存款 2026-04
    sheet_data_a = [
        ["2026年4月本外币存款", "", ""],
        ["项目", "本月余额", "比上月"],
        ["活期存款", 100, 5],
        ["定期存款", 200, -3],
        ["合计", 300, 2],  # "合计" 为跳过
    ]
    src_a = inp / "本外币存款表_202604.xlsx"
    _make_source(src_a, "本外币", sheet_data_a)

    # source B：另一日期 2026-05，重复路径触发后缀
    sheet_data_b = [
        ["2026年5月本外币存款", "", ""],
        ["项目", "本月余额", "比上月"],
        ["活期存款", 110, 10],
        ["活期存款", 120, 1],  # 同名行头 → 后缀
        ["合计", 230, 11],
    ]
    src_b = inp / "本外币存款表_202605.xlsx"
    _make_source(src_b, "本外币", sheet_data_b)

    # template
    _make_source(tmpl_path, "本外币", sheet_data_a)

    _make_config(cfg, inp, out, logs)
    return cfg, inp, out, tmpl_path, src_a


def test_5_precheck_pass(env):
    cfg, *_ = env
    assert run_precheck(cfg) == 0


def test_5_precheck_fail_missing_sheet(tmp_path):
    cfg = tmp_path / "config.xlsx"
    pd.DataFrame([("输入目录", "x")], columns=["键", "值"]).to_excel(
        cfg, sheet_name=SHEET_GLOBAL, index=False
    )
    assert run_precheck(cfg) > 0


def test_5_validate_per_feature(env):
    cfg, *_ = env
    for f in ("1", "2", "3", "4", "5"):
        assert validate_sheets_for_feature(f, cfg) == []


def test_2_timeline_slim(env):
    cfg, _, _, _, _ = env
    g = load_global(cfg)
    rules = load_timeline_rules(cfg)
    pmaps = load_path_maps(cfg)
    out = run_timeline_slim(rules, pmaps, g)
    assert out.exists()
    df = pd.read_excel(out, sheet_name="时序提取结果")
    # 7 列
    assert list(df.columns) == ["源文件", "工作表名", "规则名称", "数据日期",
                                 "行头路径", "列头路径", "数值"]
    # 2026-04: 活期/定期 各 2 列 = 4 行
    apr = df[df["数据日期"] == "2026-04"]
    assert len(apr) == 4
    # 路径标准化生效：行头不应再出现「活期存款」
    assert "活期存款" not in set(apr["行头路径"])
    assert "活期" in set(apr["行头路径"])


def test_3_4_wide_suffix_diff(env):
    cfg, _, _, _, _ = env
    g = load_global(cfg)
    rules = load_timeline_rules(cfg)
    pmaps = load_path_maps(cfg)
    out_with = run_wide_summary(rules, pmaps, g, row_suffix_enabled=True)
    out_no = run_wide_summary(rules, pmaps, g, row_suffix_enabled=False)
    df_with = pd.read_excel(out_with, sheet_name="存款规则")
    df_no = pd.read_excel(out_no, sheet_name="存款规则")
    # 5 月源里两条「活期存款」（标准化后→「活期」）
    may_with = df_with[df_with["数据日期"] == "2026-05"]["行头路径"].tolist()
    may_no = df_no[df_no["数据日期"] == "2026-05"]["行头路径"].tolist()
    # 加后缀模式：保留 2 行（活期_1, 活期_2）
    assert sum(1 for r in may_with if r.startswith("活期")) == 2
    # 不加后缀模式：合并冲突，只剩 1 行（保留首值）
    assert may_no.count("活期") == 1


def test_1_compare_pass_and_diff(env):
    cfg, _, _, tmpl, _ = env
    g = load_global(cfg)
    rules = load_timeline_rules(cfg)
    # 注入一个差异源
    diff_src = g.input_dir / "本外币存款表_202607.xlsx"
    _make_source(diff_src, "本外币", [
        ["2026年7月本外币存款", "", ""],
        ["项目", "本月余额", "比上月XXX"],  # 列头改名
        ["活期", 100, 5],                   # 行头被改
        ["合计", 100, 5],
    ])
    sources = sorted(g.input_dir.glob("*.xlsx"))
    out = run_compare(tmpl, sources, rules, g)
    assert out.exists()
    diff = pd.read_excel(out, sheet_name="差异明细")
    types = set(diff["差异类型"])
    assert "位置" in types
