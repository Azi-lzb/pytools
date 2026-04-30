from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
import pandas as pd

SHEET_GLOBAL = "全局配置"
SHEET_TIMELINE_RULE = "时序提取规则"
SHEET_PATH_MAP = "路径标准化映射"
SHEET_DEDUP_TASK = "去重追加数据配置"
SHEET_DEDUP_TASK_LEGACY = "数据处理配置"
SHEET_PRINT_CONFIG = "打印配置"
SHEET_CONFIG_RENAME = "重命名配置"
SHEET_INSTITUTION_MAPPING = "机构映射表"
SHEET_EXTRACT_CONFIG = "工作表提取"

REQUIRED_GLOBAL_KEYS = ["输入目录", "输出目录", "日志目录"]
TIMELINE_COLS = [
    "是否启用", "规则名称", "工作簿关键字", "工作表关键字",
    "行头列", "列表头行",
    "必含列头", "必含行头",
    "数据起始行", "数据结束行", "数据起始列", "数据结束列",
    "跳过关键字", "目标工作簿路径", "目标工作表", "启用目标写入",
]
TIMELINE_REQUIRED_COLS = [c for c in TIMELINE_COLS if c != "启用目标写入"]
PATH_MAP_COLS = [
    "是否启用", "映射名称", "适用规则名",
    "工作簿关键字", "工作表关键字",
    "作用对象", "匹配方式", "原始路径", "标准路径", "备注",
]
DEDUP_TASK_COLS = [
    "是否启用", "源数据工作簿", "源数据工作表", "标识列序号",
    "目标工作簿", "目标工作表", "执行模式", "备注",
]
PRINT_CONFIG_COLS = [
    "是否启用",
    "打印模式",
    "源工作簿",
    "源工作表",
    "源工作表打印区域",
    "目标工作簿",
    "目标工作表",
    "FitToPagesWide",
    "FitToPagesTall",
    "打印方向",
    "备注",
]
CONFIG_RENAME_COLS = [
    "简称", "全称", "占位C", "代码", "全称(代码映射)", "占位F", "键", "值", "占位I", "原表名", "新表名",
]
INSTITUTION_MAPPING_COLS = [
    "原始机构名称", "映射后机构名称", "是否为外资行",
]
EXTRACT_CONFIG_COLS = [
    "表A-币种", "表A-地区", "表A-机构", "表A-类型", "表A-名称",
    "是否禁用", "是否整表提取", "提取行", "提取列", "输出文件名",
]

FEATURE_REQUIRED = {
    "1": [SHEET_GLOBAL, SHEET_TIMELINE_RULE],  # 3.9.3 仅依赖规则；模板/源走交互选择
    "1com": [SHEET_GLOBAL, SHEET_TIMELINE_RULE],  # 3.9.3 COM 版
    "2com": [SHEET_GLOBAL, SHEET_TIMELINE_RULE, SHEET_PATH_MAP],  # 3.9.6 COM 版
    "3com": [SHEET_GLOBAL, SHEET_TIMELINE_RULE, SHEET_PATH_MAP],  # 3.9.7 COM 版
    "4com": [SHEET_GLOBAL, SHEET_TIMELINE_RULE, SHEET_PATH_MAP],  # 3.9.8 COM 版
    "2": [SHEET_GLOBAL, SHEET_TIMELINE_RULE, SHEET_PATH_MAP],
    "3": [SHEET_GLOBAL, SHEET_TIMELINE_RULE, SHEET_PATH_MAP],
    "4": [SHEET_GLOBAL, SHEET_TIMELINE_RULE, SHEET_PATH_MAP],
    "5": [SHEET_GLOBAL, SHEET_TIMELINE_RULE, SHEET_PATH_MAP],
    "8": [SHEET_DEDUP_TASK],
    "9": [SHEET_DEDUP_TASK],
    "a": [SHEET_DEDUP_TASK],
    "b": [SHEET_DEDUP_TASK],
    "p1": [SHEET_GLOBAL],
    "p3": [SHEET_GLOBAL],
    "p7": [SHEET_GLOBAL, SHEET_PRINT_CONFIG],
    "p8": [SHEET_GLOBAL, SHEET_PRINT_CONFIG],
    "p1com": [SHEET_GLOBAL],
    "p3com": [SHEET_GLOBAL],
    "p7com": [SHEET_GLOBAL, SHEET_PRINT_CONFIG],
    "ppdf": [SHEET_GLOBAL, SHEET_PRINT_CONFIG],
    "t3": [SHEET_CONFIG_RENAME],
    "t4": [SHEET_GLOBAL],
    "t5": [SHEET_GLOBAL],
    "t7": [SHEET_CONFIG_RENAME],
    "s21": [SHEET_GLOBAL],
    "s22": [SHEET_GLOBAL],
    "s21com": [SHEET_GLOBAL],
    "s22com": [SHEET_GLOBAL],
    "x11": [SHEET_GLOBAL],
    "x12": [SHEET_GLOBAL, SHEET_INSTITUTION_MAPPING],
    "x13": [SHEET_GLOBAL, SHEET_INSTITUTION_MAPPING],
    "x14": [SHEET_GLOBAL],
    "x14com": [SHEET_GLOBAL],
    "x15": [SHEET_GLOBAL],
    "x16": [SHEET_GLOBAL, SHEET_EXTRACT_CONFIG],
    "x16com": [SHEET_GLOBAL, SHEET_EXTRACT_CONFIG],
    "x18": [SHEET_GLOBAL],
}
SHEET_COL_REQUIRED = {
    SHEET_TIMELINE_RULE: TIMELINE_REQUIRED_COLS,
    SHEET_PATH_MAP: PATH_MAP_COLS,
    SHEET_PRINT_CONFIG: PRINT_CONFIG_COLS,
    SHEET_CONFIG_RENAME: CONFIG_RENAME_COLS,
    SHEET_INSTITUTION_MAPPING: INSTITUTION_MAPPING_COLS,
    SHEET_EXTRACT_CONFIG: EXTRACT_CONFIG_COLS,
}


@dataclass
class GlobalConfig:
    input_dir: Path
    output_dir: Path
    log_dir: Path
    error_policy: str = "continue"
    default_encoding: str = "utf-8"
    source_exts: list[str] = field(default_factory=lambda: [".xlsx", ".xlsm", ".csv"])


@dataclass
class TimelineRule:
    enabled: bool
    name: str
    wb_keyword: str
    sheet_keyword: str
    row_header_col: int
    col_header_row: int  # 单行用 "3"，多行用 "2,3"
    row_header_col_specified: bool = True
    col_header_rows: list[int] = field(default_factory=list)
    required_col_headers: list[str] = field(default_factory=list)
    required_row_headers: list[str] = field(default_factory=list)
    data_row_start: int = 0
    data_row_end: int | None = None
    data_col_start: int = 0
    data_col_end: int | None = None
    skip_keywords: list[str] = field(default_factory=list)
    target_wb_path: Path | None = None
    target_sheet: str | None = None
    target_write_enabled: bool = True


@dataclass
class PathMapRule:
    enabled: bool
    name: str
    applicable_rule: str
    wb_keyword: str
    sheet_keyword: str
    target: str  # 行头/列头/两者
    match_mode: str  # 精确/包含
    original: str
    standard: str


def _truthy(v) -> bool:
    if v is None:
        return False
    s = str(v).strip().lower()
    return s in ("是", "1", "true", "y", "yes", "on")


def _split(v, sep=";") -> list[str]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return []
    s = str(v).strip()
    if not s:
        return []
    return [p.strip() for p in s.replace("；", ";").replace(",", ";").split(sep) if p.strip()]


def _to_int(v, default=0) -> int:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    try:
        return int(float(v))
    except (ValueError, TypeError):
        return default


def _to_int_or_none(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
        return None
    try:
        return int(float(v))
    except (ValueError, TypeError):
        return None


def _to_str(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def validate_sheets_for_feature(feature: str, cfg_path: Path) -> list[str]:
    """返回缺失项清单；空列表表示通过。"""
    errors: list[str] = []
    if not cfg_path.exists():
        return [f"配置文件不存在: {cfg_path}"]
    try:
        xls = pd.ExcelFile(cfg_path)
    except Exception as e:
        return [f"无法打开配置文件 {cfg_path}: {e}"]
    needed = FEATURE_REQUIRED.get(feature, [])
    if feature in ("8", "9", "a", "b"):
        try:
            xls = pd.ExcelFile(cfg_path)
            if SHEET_DEDUP_TASK not in xls.sheet_names and SHEET_DEDUP_TASK_LEGACY not in xls.sheet_names:
                return [f"缺少 Sheet: {SHEET_DEDUP_TASK}（或 {SHEET_DEDUP_TASK_LEGACY}）"]
            return []
        except Exception as e:
            return [f"无法打开配置文件 {cfg_path}: {e}"]
    for sh in needed:
        if sh not in xls.sheet_names:
            errors.append(f"缺少 Sheet: {sh}")
            continue
        if sh == SHEET_GLOBAL:
            df = pd.read_excel(xls, sh, header=0, dtype=object)
            if df.shape[1] < 2:
                errors.append(f"{SHEET_GLOBAL} 至少需要 键/值 两列")
                continue
            keys = {_to_str(k) for k in df.iloc[:, 0].tolist()}
            for k in REQUIRED_GLOBAL_KEYS:
                if k not in keys:
                    errors.append(f"{SHEET_GLOBAL} 缺少键: {k}")
        else:
            df = pd.read_excel(xls, sh, header=0, dtype=object, nrows=0)
            cols = list(df.columns)
            for c in SHEET_COL_REQUIRED[sh]:
                if c not in cols:
                    errors.append(f"{sh} 缺少列: {c}")
    return errors


def load_global(cfg_path: Path) -> GlobalConfig:
    df = pd.read_excel(cfg_path, SHEET_GLOBAL, header=0, dtype=object)
    kv = {}
    for _, row in df.iterrows():
        k = _to_str(row.iloc[0])
        v = _to_str(row.iloc[1]) if df.shape[1] > 1 else ""
        if k:
            kv[k] = v
    base = cfg_path.parent.parent  # pytools/
    def _path(key, default):
        v = kv.get(key, "").strip()
        if not v:
            return (base / default).resolve()
        p = Path(v)
        return p if p.is_absolute() else (base / p).resolve()
    exts = _split(kv.get("源文件扩展名", ""))
    if not exts:
        exts = [".xlsx", ".xlsm", ".csv"]
    return GlobalConfig(
        input_dir=_path("输入目录", "input"),
        output_dir=_path("输出目录", "output"),
        log_dir=_path("日志目录", "logs"),
        error_policy=kv.get("错误策略", "continue") or "continue",
        default_encoding=kv.get("默认编码", "utf-8") or "utf-8",
        source_exts=[e.lower() if e.startswith(".") else "." + e.lower() for e in exts],
    )


def load_global_value(cfg_path: Path, key: str, default: str = "") -> str:
    df = pd.read_excel(cfg_path, SHEET_GLOBAL, header=0, dtype=object)
    for _, row in df.iterrows():
        k = _to_str(row.iloc[0]) if df.shape[1] > 0 else ""
        if k == key:
            return _to_str(row.iloc[1]) if df.shape[1] > 1 else default
    return default


def load_timeline_rules(cfg_path: Path) -> list[TimelineRule]:
    df = pd.read_excel(cfg_path, SHEET_TIMELINE_RULE, header=0, dtype=object)
    rules: list[TimelineRule] = []
    for _, row in df.iterrows():
        if not _truthy(row.get("是否启用")):
            continue
        col_header_raw = _to_str(row.get("列表头行"))
        col_header_rows = [_to_int(x) for x in col_header_raw.replace("，", ",").split(",") if x.strip()]
        if not col_header_rows:
            col_header_rows = [_to_int(col_header_raw, 1)]
        target_wb = _to_str(row.get("目标工作簿路径"))
        # 规则级目标写入开关：列不存在时默认 True（兼容旧配置）
        if "启用目标写入" in df.columns:
            target_write_enabled = _truthy(row.get("启用目标写入"))
        else:
            target_write_enabled = True
        from .rule_engine import parse_col_spec
        row_col_raw = _to_str(row.get("行头列"))
        row_col_parsed = parse_col_spec(row.get("行头列"))
        row_col_specified = row_col_parsed > 0
        rules.append(TimelineRule(
            enabled=True,
            name=_to_str(row.get("规则名称")),
            wb_keyword=_to_str(row.get("工作簿关键字")),
            sheet_keyword=_to_str(row.get("工作表关键字")),
            row_header_col=row_col_parsed or 1,
            row_header_col_specified=row_col_specified if row_col_raw else False,
            col_header_row=col_header_rows[0],
            col_header_rows=col_header_rows,
            required_col_headers=_split(row.get("必含列头")),
            required_row_headers=_split(row.get("必含行头")),
            data_row_start=_to_int(row.get("数据起始行"), 0),
            data_row_end=_to_int_or_none(row.get("数据结束行")),
            data_col_start=parse_col_spec(row.get("数据起始列")),
            data_col_end=parse_col_spec(row.get("数据结束列")) or None,
            skip_keywords=_split(row.get("跳过关键字")),
            target_wb_path=Path(target_wb) if target_wb else None,
            target_sheet=_to_str(row.get("目标工作表")) or None,
            target_write_enabled=target_write_enabled,
        ))
    return rules


def load_path_maps(cfg_path: Path) -> list[PathMapRule]:
    if SHEET_PATH_MAP not in pd.ExcelFile(cfg_path).sheet_names:
        return []
    df = pd.read_excel(cfg_path, SHEET_PATH_MAP, header=0, dtype=object)
    out: list[PathMapRule] = []
    for _, row in df.iterrows():
        if not _truthy(row.get("是否启用")):
            continue
        out.append(PathMapRule(
            enabled=True,
            name=_to_str(row.get("映射名称")),
            applicable_rule=_to_str(row.get("适用规则名")),
            wb_keyword=_to_str(row.get("工作簿关键字")),
            sheet_keyword=_to_str(row.get("工作表关键字")),
            target=_to_str(row.get("作用对象")) or "两者",
            match_mode=_to_str(row.get("匹配方式")) or "精确",
            original=_to_str(row.get("原始路径")),
            standard=_to_str(row.get("标准路径")),
        ))
    return out


def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _parse_key_cols(raw) -> list[int]:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return []
    txt = str(raw).strip()
    if not txt:
        return []
    txt = txt.replace("，", ";").replace(",", ";").replace("；", ";").replace(":", ";")
    out: list[int] = []
    for p in txt.split(";"):
        p = p.strip()
        if not p:
            continue
        try:
            v = int(float(p))
            if v > 0:
                out.append(v)
        except Exception:
            pass
    return out


def _resolve_any_path(path_text: str, base: Path) -> Path:
    p = Path(path_text)
    if p.is_absolute():
        return p
    return (base / p).resolve()


def load_dedup_tasks(cfg_path: Path) -> list[dict]:
    xls = pd.ExcelFile(cfg_path)
    if SHEET_DEDUP_TASK in xls.sheet_names:
        sh = SHEET_DEDUP_TASK
    elif SHEET_DEDUP_TASK_LEGACY in xls.sheet_names:
        sh = SHEET_DEDUP_TASK_LEGACY
    else:
        return []

    df = pd.read_excel(cfg_path, sh, header=0, dtype=object)
    base = cfg_path.parent.parent  # pytools/

    col_enabled = _find_col(df, ["是否启用", "启用"])
    col_task_name = _find_col(df, ["任务名", "任务名称"])
    col_src_wb = _find_col(df, ["源数据工作簿", "源工作簿路径", "源工作簿"])
    col_src_ws = _find_col(df, ["源数据工作表", "源工作表名", "源工作表"])
    col_key_cols = _find_col(df, ["标识列序号", "去重列序号", "去重列"])
    col_tgt_wb = _find_col(df, ["目标工作簿", "目标工作簿路径"])
    col_tgt_ws = _find_col(df, ["目标工作表", "目标工作表名"])
    col_write_mode = _find_col(df, ["写入方式", "覆盖方式"])

    tasks: list[dict] = []
    for i, row in df.iterrows():
        src_wb_text = _to_str(row.get(col_src_wb)) if col_src_wb else ""
        src_ws_text = _to_str(row.get(col_src_ws)) if col_src_ws else ""
        tgt_wb_text = _to_str(row.get(col_tgt_wb)) if col_tgt_wb else ""
        tgt_ws_text = _to_str(row.get(col_tgt_ws)) if col_tgt_ws else ""
        task = {
            "row_no": i + 2,
            "enabled": _truthy(row.get(col_enabled)) if col_enabled else True,
            "task_name": _to_str(row.get(col_task_name)) if col_task_name else "",
            "source_wb": _resolve_any_path(src_wb_text, base) if src_wb_text else None,
            "source_ws": src_ws_text,
            "key_cols": _parse_key_cols(row.get(col_key_cols)) if col_key_cols else [],
            "target_wb": _resolve_any_path(tgt_wb_text, base) if tgt_wb_text else None,
            "target_ws": tgt_ws_text,
            "write_mode": _to_str(row.get(col_write_mode)) if col_write_mode else "",
        }
        tasks.append(task)
    return tasks


def load_print_tasks(cfg_path: Path) -> list[dict]:
    xls = pd.ExcelFile(cfg_path)
    if SHEET_PRINT_CONFIG not in xls.sheet_names:
        return []
    df = pd.read_excel(cfg_path, SHEET_PRINT_CONFIG, header=0, dtype=object)
    base = cfg_path.parent.parent

    def _resolve_any_path(path_text: str) -> Path:
        p = Path(path_text)
        if p.is_absolute():
            return p
        return (base / p).resolve()

    out: list[dict] = []
    for i, row in df.iterrows():
        src_wb_text = _to_str(row.get("源工作簿"))
        tgt_wb_text = _to_str(row.get("目标工作簿"))
        out.append({
            "row_no": i + 2,
            "enabled": _truthy(row.get("是否启用")),
            "mode": _to_int(row.get("打印模式"), 0),
            "source_wb": _resolve_any_path(src_wb_text) if src_wb_text else None,
            "source_ws": _to_str(row.get("源工作表")),
            "source_ranges": _to_str(row.get("源工作表打印区域")),
            "target_wb": _resolve_any_path(tgt_wb_text) if tgt_wb_text else None,
            "target_ws": _to_str(row.get("目标工作表")),
            "fit_wide": _to_int(row.get("FitToPagesWide"), 1),
            "fit_tall": _to_int(row.get("FitToPagesTall"), 1),
            "orientation": _to_str(row.get("打印方向")),
            "remark": _to_str(row.get("备注")),
        })
    return out


def load_extract_tasks(cfg_path: Path) -> list[dict]:
    """读取'工作表提取'配置表，返回启用任务列表。"""
    xls = pd.ExcelFile(cfg_path)
    if SHEET_EXTRACT_CONFIG not in xls.sheet_names:
        return []
    df = pd.read_excel(cfg_path, SHEET_EXTRACT_CONFIG, header=0, dtype=object)
    out: list[dict] = []
    for i, row in df.iterrows():
        disabled = _truthy(row.get("是否禁用"))
        if disabled:
            continue
        keywords = [_to_str(row.get(f"表A-{k}")) for k in ("币种", "地区", "机构", "类型", "名称")]
        if not any(k for k in keywords):
            continue
        out.append({
            "row_no": i + 2,
            "keywords": keywords,
            "full_extract": _truthy(row.get("是否整表提取")),
            "rows_spec": _to_str(row.get("提取行")),
            "cols_spec": _to_str(row.get("提取列")),
            "output_name": _to_str(row.get("输出文件名")),
        })
    return out


def load_foreign_banks(cfg_path: Path) -> set[str]:
    """读取机构映射表 B 列机构名 + C 列外资行标志（1=是）。返回外资行名集合。"""
    xls = pd.ExcelFile(cfg_path)
    if SHEET_INSTITUTION_MAPPING not in xls.sheet_names:
        return set()
    df = pd.read_excel(cfg_path, SHEET_INSTITUTION_MAPPING, header=0, dtype=object)
    out: set[str] = set()
    for _, row in df.iterrows():
        if df.shape[1] < 3:
            break
        name = _to_str(row.iloc[1])
        flag = _to_str(row.iloc[2])
        if not name:
            continue
        try:
            if int(float(flag)) == 1:
                out.add(name)
        except (ValueError, TypeError):
            continue
    return out


def load_institution_mapping(cfg_path: Path) -> dict[str, str]:
    """读取机构映射表：A列原始 -> B列映射；空键/空值忽略。"""
    xls = pd.ExcelFile(cfg_path)
    if SHEET_INSTITUTION_MAPPING not in xls.sheet_names:
        return {}
    df = pd.read_excel(cfg_path, SHEET_INSTITUTION_MAPPING, header=0, dtype=object)
    mapping: dict[str, str] = {}
    for _, row in df.iterrows():
        k = _to_str(row.iloc[0]) if df.shape[1] > 0 else ""
        v = _to_str(row.iloc[1]) if df.shape[1] > 1 else ""
        if k and v:
            mapping[k] = v
    return mapping
