from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from pytools.core.config_xlsx import SHEET_CONFIG_RENAME
from pytools.core.convert_33_34_35_37 import (
    run_batch_rename_by_content_46,
    run_batch_unrename_by_content_47,
)


def _make_cfg(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_CONFIG_RENAME
    ws["L1"] = "根据文件内容重命名"
    ws["L2"] = "sheet1@A1;sheet2@A2"
    wb.save(path)
    wb.close()


def _make_source(path: Path) -> None:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "sheet1"
    ws1["A1"] = "A"
    ws2 = wb.create_sheet("sheet2")
    ws2["A2"] = "B"
    wb.save(path)
    wb.close()


def test_content_rename_skips_when_all_l2_values_already_in_name(tmp_path):
    cfg = tmp_path / "config.xlsx"
    logs = tmp_path / "logs"
    src = tmp_path / "测试文件.xlsx"
    _make_cfg(cfg)
    _make_source(src)

    stat1 = run_batch_rename_by_content_46(cfg, [src], logs)
    renamed = tmp_path / "A_B_测试文件.xlsx"
    assert stat1["ok"] == 1
    assert renamed.exists()

    stat2 = run_batch_rename_by_content_46(cfg, [renamed], logs)

    assert stat2["ok"] == 0
    assert stat2["skip"] == 1
    assert renamed.exists()
    assert not (tmp_path / "A_B_A_B_测试文件.xlsx").exists()


def test_content_rename_does_not_skip_when_only_some_l2_values_in_name(tmp_path):
    cfg = tmp_path / "config.xlsx"
    logs = tmp_path / "logs"
    src = tmp_path / "A_测试文件.xlsx"
    _make_cfg(cfg)
    _make_source(src)

    stat = run_batch_rename_by_content_46(cfg, [src], logs)

    assert stat["ok"] == 1
    assert (tmp_path / "A_B_A_测试文件.xlsx").exists()


def test_content_unrename_removes_single_prefix(tmp_path):
    cfg = tmp_path / "config.xlsx"
    logs = tmp_path / "logs"
    src = tmp_path / "A_B_测试文件.xlsx"
    _make_cfg(cfg)
    _make_source(src)

    stat = run_batch_unrename_by_content_47(cfg, [src], logs)

    assert stat["ok"] == 1
    assert (tmp_path / "测试文件.xlsx").exists()
    assert not src.exists()


def test_content_unrename_removes_repeated_prefixes(tmp_path):
    cfg = tmp_path / "config.xlsx"
    logs = tmp_path / "logs"
    src = tmp_path / "A_B_A_B_测试文件.xlsx"
    _make_cfg(cfg)
    _make_source(src)

    stat = run_batch_unrename_by_content_47(cfg, [src], logs)

    assert stat["ok"] == 1
    assert (tmp_path / "测试文件.xlsx").exists()
    assert not src.exists()


def test_content_unrename_removes_partial_prefix(tmp_path):
    cfg = tmp_path / "config.xlsx"
    logs = tmp_path / "logs"
    src = tmp_path / "A_测试文件.xlsx"
    _make_cfg(cfg)
    _make_source(src)

    stat = run_batch_unrename_by_content_47(cfg, [src], logs)

    assert stat["ok"] == 1
    assert (tmp_path / "测试文件.xlsx").exists()
    assert not src.exists()


def test_content_unrename_removes_any_l2_items_from_leading_prefix(tmp_path):
    cfg = tmp_path / "config.xlsx"
    logs = tmp_path / "logs"
    src = tmp_path / "B_A_B_测试文件.xlsx"
    _make_cfg(cfg)
    _make_source(src)

    stat = run_batch_unrename_by_content_47(cfg, [src], logs)

    assert stat["ok"] == 1
    assert (tmp_path / "测试文件.xlsx").exists()
    assert not src.exists()


def test_content_unrename_skips_when_no_l2_prefix_item_at_start(tmp_path):
    cfg = tmp_path / "config.xlsx"
    logs = tmp_path / "logs"
    src = tmp_path / "C_D_测试文件.xlsx"
    _make_cfg(cfg)
    _make_source(src)

    stat = run_batch_unrename_by_content_47(cfg, [src], logs)

    assert stat["ok"] == 0
    assert stat["skip"] == 1
    assert src.exists()


def test_content_unrename_uses_next_path_when_target_exists(tmp_path):
    cfg = tmp_path / "config.xlsx"
    logs = tmp_path / "logs"
    src = tmp_path / "A_B_测试文件.xlsx"
    existing = tmp_path / "测试文件.xlsx"
    _make_cfg(cfg)
    _make_source(src)
    _make_source(existing)

    stat = run_batch_unrename_by_content_47(cfg, [src], logs)

    assert stat["ok"] == 1
    assert existing.exists()
    assert (tmp_path / "测试文件_1.xlsx").exists()
