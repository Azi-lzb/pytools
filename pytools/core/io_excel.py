from __future__ import annotations
from pathlib import Path
import re
from functools import lru_cache
import pandas as pd

_INVALID_SHEET_CHARS = re.compile(r"[\\/?*\[\]:]")


def safe_sheet_name(name: str, used: set[str]) -> str:
    """清洗非法字符 + 截断 31 字符 + 处理重复后缀。"""
    s = _INVALID_SHEET_CHARS.sub("_", name or "Sheet")[:31] or "Sheet"
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


def write_workbook(out_path: Path, sheets: dict[str, pd.DataFrame]) -> Path:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        used: set[str] = set()
        for name, df in sheets.items():
            sn = safe_sheet_name(name, used)
            df.to_excel(w, sheet_name=sn, index=False)
    return out_path


@lru_cache(maxsize=16)
def _read_sheet_2d_cached(path_text: str, sheet: str, mtime: float, size: int) -> pd.DataFrame:
    return pd.read_excel(path_text, sheet_name=sheet, header=None, dtype=object)


def read_sheet_2d(path: Path, sheet: str) -> pd.DataFrame:
    """读取整个 sheet 为 header=None 的二维 DataFrame，列名为 0..N-1。
    按 (路径+mtime+size+sheet) 缓存，避免同一批里同一个 sheet 被多条规则重复读取。"""
    try:
        st = path.stat()
    except FileNotFoundError:
        return pd.DataFrame()
    return _read_sheet_2d_cached(str(path), sheet, st.st_mtime, st.st_size)


def clear_sheet_cache() -> None:
    """处理完一个源文件后清空缓存，释放内存。"""
    _read_sheet_2d_cached.cache_clear()
    _list_sheet_names_cached.cache_clear()


@lru_cache(maxsize=64)
def _list_sheet_names_cached(path_text: str, mtime: float, size: int) -> tuple[str, ...]:
    return tuple(pd.ExcelFile(path_text).sheet_names)


def list_sheet_names(path: Path) -> list[str]:
    # 按 (路径+mtime+size) 缓存，避免同一批汇总里 N 条规则重复打开同一个大 xlsx
    try:
        st = path.stat()
        key = (str(path), st.st_mtime, st.st_size)
    except FileNotFoundError:
        return []
    return list(_list_sheet_names_cached(*key))


def list_source_files(input_dir: Path, exts: list[str]) -> list[Path]:
    if not input_dir.exists():
        return []
    return sorted(p for p in input_dir.rglob("*")
                  if p.is_file()
                  and p.suffix.lower() in exts
                  and not p.name.startswith("~$"))


def append_to_target(target_wb: Path, target_sheet: str,
                     header: list[str], rows: list[list],
                     dedup_key_idx: list[int],
                     required_prefix: list[str] | None = None,
                     allow_extend: bool = False,
                     dedup_seq_prefix_idx: list[int] | None = None) -> dict[str, int]:
    """追加到目标簿/表；不存在则新建；按 dedup_key_idx 列做去重。

    返回统计：
    - input_rows: 输入行数
    - batch_dedup_rows: 本批内去重后行数
    - existing_filtered_rows: 与目标既有数据过滤后可写入行数
    - written: 是否实际写入
    - skipped_reason: 跳过原因（如表头不匹配）
    """
    if not rows:
        return {
            "input_rows": 0, "batch_dedup_rows": 0, "existing_filtered_rows": 0,
            "written": 0, "skipped_reason": ""
        }
    new_df = pd.DataFrame(rows, columns=header)
    input_rows = len(new_df)
    cols = [header[i] for i in dedup_key_idx if 0 <= i < len(header)] if dedup_key_idx else []
    seq_prefix_cols = [header[i] for i in (dedup_seq_prefix_idx or []) if 0 <= i < len(header)]
    if cols:
        # 先对本批新增数据去重。若启用行序号去重，口径改为“键列 + 组内序号”，
        # 避免仅按键列去重误删同组内本应保留的多行。
        if seq_prefix_cols and all(c in new_df.columns for c in seq_prefix_cols):
            new_prefix = new_df[seq_prefix_cols].fillna("").astype(str)
            new_seq = new_prefix.groupby(seq_prefix_cols, dropna=False).cumcount()
            new_key_df = new_df[cols].fillna("").astype(str).copy()
            new_key_df["__seq"] = new_seq.values
            keep_mask = ~new_key_df.duplicated(keep="first")
            new_df = new_df[keep_mask]
        else:
            new_df = new_df.drop_duplicates(subset=cols, keep="first")
        batch_dedup_rows = len(new_df)
        if new_df.empty:
            return {
                "input_rows": input_rows, "batch_dedup_rows": 0, "existing_filtered_rows": 0,
                "written": 0, "skipped_reason": ""
            }
    else:
        batch_dedup_rows = len(new_df)

    if target_wb.exists() and target_sheet in list_sheet_names(target_wb):
        old = pd.read_excel(target_wb, sheet_name=target_sheet, dtype=object)
        old_cols = [str(c) for c in old.columns]

        # 表头安全：要求固定前缀列存在，否则跳过写入，避免误把不相干表覆盖掉
        prefix = required_prefix or []
        if prefix and any(c not in old_cols for c in prefix):
            return {
                "input_rows": input_rows,
                "batch_dedup_rows": batch_dedup_rows,
                "existing_filtered_rows": 0,
                "written": 0,
                "skipped_reason": "header_mismatch",
            }

        # 列对齐：宽表允许扩展，时序则按新表头重排
        if allow_extend:
            merged_cols = list(old_cols)
            for c in header:
                if c not in merged_cols:
                    merged_cols.append(c)
            old = old.reindex(columns=merged_cols)
            new_df = new_df.reindex(columns=merged_cols)
            work_cols = [c for c in cols if c in merged_cols]
        else:
            old = old.reindex(columns=header)
            new_df = new_df.reindex(columns=header)
            work_cols = cols
        if work_cols:
            # 行序号去重：按前缀列分组后，用组内序号作为附加键，减少全列键开销
            if seq_prefix_cols and all(c in old.columns for c in seq_prefix_cols) and all(c in new_df.columns for c in seq_prefix_cols):
                old_prefix = old[seq_prefix_cols].fillna("").astype(str)
                old_seq = old_prefix.groupby(seq_prefix_cols, dropna=False).cumcount()
                old_key_df = old[work_cols].fillna("").astype(str).copy()
                old_key_df["__seq"] = old_seq.values

                new_prefix = new_df[seq_prefix_cols].fillna("").astype(str)
                new_seq = new_prefix.groupby(seq_prefix_cols, dropna=False).cumcount()
                new_key_df = new_df[work_cols].fillna("").astype(str).copy()
                new_key_df["__seq"] = new_seq.values
            else:
                old_key_df = old[work_cols].fillna("").astype(str)
                new_key_df = new_df[work_cols].fillna("").astype(str)

            existing = set(tuple(x) for x in old_key_df.itertuples(index=False, name=None))
            keep_mask = []
            for key in new_key_df.itertuples(index=False, name=None):
                if key in existing:
                    keep_mask.append(False)
                else:
                    keep_mask.append(True)
                    existing.add(key)
            new_df = new_df[keep_mask]
            if new_df.empty:
                return {
                    "input_rows": input_rows, "batch_dedup_rows": batch_dedup_rows, "existing_filtered_rows": 0,
                    "written": 0, "skipped_reason": ""
                }
        merged = pd.concat([old, new_df], ignore_index=True)
    else:
        merged = new_df
    # 首次建表不做全表去重，保留传入数据顺序与完整性
    existing_filtered_rows = len(new_df)

    target_wb.parent.mkdir(parents=True, exist_ok=True)
    if target_wb.exists():
        # 保留其它 sheet
        with pd.ExcelWriter(target_wb, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
            merged.to_excel(w, sheet_name=target_sheet, index=False)
    else:
        with pd.ExcelWriter(target_wb, engine="openpyxl") as w:
            merged.to_excel(w, sheet_name=target_sheet, index=False)
    return {
        "input_rows": input_rows,
        "batch_dedup_rows": batch_dedup_rows,
        "existing_filtered_rows": existing_filtered_rows,
        "written": 1,
        "skipped_reason": "",
    }
