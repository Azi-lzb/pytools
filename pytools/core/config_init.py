from __future__ import annotations

from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.comments import Comment

from .config_xlsx import (
    SHEET_GLOBAL,
    SHEET_TIMELINE_RULE,
    SHEET_PATH_MAP,
    SHEET_DEDUP_TASK,
    SHEET_PRINT_CONFIG,
    SHEET_CONFIG_RENAME,
    SHEET_INSTITUTION_MAPPING,
    SHEET_EXTRACT_CONFIG,
    TIMELINE_COLS,
    PATH_MAP_COLS,
    DEDUP_TASK_COLS,
    PRINT_CONFIG_COLS,
    CONFIG_RENAME_COLS,
    INSTITUTION_MAPPING_COLS,
    EXTRACT_CONFIG_COLS,
)


GLOBAL_COLS = ["键", "值", "备注"]

GLOBAL_KEY_FIELDS = {"键", "值"}
GLOBAL_REMARK_FIELDS = {"备注"}
TIMELINE_REMARK_FIELDS = set()
PATHMAP_REMARK_FIELDS = {"备注"}

HEADER_COMMENTS = {
    SHEET_GLOBAL: {
        "键": "关键字段（程序读取）：配置键名，例如 输入目录/输出目录/日志目录。",
        "值": "关键字段（程序读取）：对应配置值。",
        "备注": "备注字段（仅人工说明）：程序不读取。",
    },
    SHEET_TIMELINE_RULE: {
        "是否启用": "关键字段（程序读取）：Y/是/1 表示启用。",
        "规则名称": "关键字段（程序读取）：规则唯一标识，建议不重复。",
        "工作簿关键字": "关键字段（程序读取）：支持分号分隔的关键字匹配。",
        "工作表关键字": "关键字段（程序读取）：支持分号分隔的关键字匹配。",
        "行头列": "关键字段（程序读取）：行头所在列（如 B 或 2）。",
        "列表头行": "关键字段（程序读取）：列表头行号，可多行（如 2,3）。",
        "必含列头": "关键字段（程序读取）：可空；用于前置命中校验。",
        "必含行头": "关键字段（程序读取）：可空；用于前置命中校验。",
        "数据起始行": "关键字段（程序读取）：可空；空则按表头后自动推断。",
        "数据结束行": "关键字段（程序读取）：可空。",
        "数据起始列": "关键字段（程序读取）：可空；空则按行头列后自动推断。",
        "数据结束列": "关键字段（程序读取）：可空。",
        "跳过关键字": "关键字段（程序读取）：可空；命中则跳过该行头。",
        "目标工作簿路径": "关键字段（程序读取）：可空；与目标工作表一起决定是否写入目标。",
        "目标工作表": "关键字段（程序读取）：可空；与目标工作簿路径一起生效。",
        "启用目标写入": "关键字段（程序读取）：Y/是/1 才写目标；空/否不写目标。",
        "set区域": "关键字段（程序读取）：可空；格式 别名@地址;别名@地址（如 机构代码@A3;报表口径@A5），仅宽表汇总使用。",
    },
    SHEET_PATH_MAP: {
        "是否启用": "关键字段（程序读取）：Y/是/1 表示启用映射。",
        "映射名称": "关键字段（程序读取）：映射规则名称。",
        "适用规则名": "关键字段（程序读取）：可填具体规则名/全部/*。",
        "工作簿关键字": "关键字段（程序读取）：可空；用于缩小映射生效范围。",
        "工作表关键字": "关键字段（程序读取）：可空；用于缩小映射生效范围。",
        "作用对象": "关键字段（程序读取）：行头/列头/两者。",
        "匹配方式": "关键字段（程序读取）：精确/包含。",
        "原始路径": "关键字段（程序读取）：待匹配文本。",
        "标准路径": "关键字段（程序读取）：替换后的标准文本。",
        "备注": "备注字段（仅人工说明）：程序不读取。",
    },
    SHEET_DEDUP_TASK: {
        "是否启用": "关键字段（程序读取）：Y/是/1 表示启用。",
        "源数据工作簿": "关键字段（程序读取）：源工作簿完整路径。",
        "源数据工作表": "关键字段（程序读取）：源工作表名。",
        "标识列序号": "关键字段（程序读取）：去重键列，如 1;2;5；留空默认全列。",
        "目标工作簿": "关键字段（程序读取）：目标工作簿完整路径。",
        "目标工作表": "关键字段（程序读取）：目标工作表名。",
        "执行模式": "关键字段（程序读取）：1=正常执行；2=仅校验不写入；3=备份后执行。",
        "备注": "备注字段（仅人工说明）：程序不读取。",
    },
    SHEET_PRINT_CONFIG: {
        "是否启用": "关键字段（程序读取）：Y/是/1 表示启用。",
        "打印模式": "关键字段（程序读取）：1/2/3；3.10.7 会按 1/2/3 全部执行。",
        "源工作簿": "关键字段（程序读取）：源工作簿完整路径。",
        "源工作表": "关键字段（程序读取）：源工作表名（精确匹配）。",
        "源工作表打印区域": "关键字段（程序读取）：可空；支持 A1:H30;A35:H70，多段分号分隔。",
        "目标工作簿": "关键字段（程序读取）：目标工作簿完整路径；不存在会自动创建。",
        "目标工作表": "关键字段（程序读取）：目标工作表名；不存在会自动创建。",
        "FitToPagesWide": "关键字段（程序读取）：分页宽度页数，留空默认 1。",
        "FitToPagesTall": "关键字段（程序读取）：分页高度页数，留空默认 1。",
        "打印方向": "关键字段（程序读取）：横向/纵向；留空按宽高比自动。",
        "备注": "备注字段（仅人工说明）：程序不读取。",
    },
    SHEET_CONFIG_RENAME: {
        "简称": "关键字段（程序读取）：3.3 使用，简称->全称映射（A->B）。",
        "全称": "关键字段（程序读取）：3.3 使用。",
        "代码": "关键字段（程序读取）：3.3 使用，与 E 列全称做全称->代码映射（E->D）。",
        "全称(代码映射)": "关键字段（程序读取）：3.3 使用。",
        "键": "关键字段（程序读取）：3.3 使用（如 数据日期、报表名称）。",
        "值": "关键字段（程序读取）：3.3 使用。",
        "原表名": "关键字段（程序读取）：3.7 使用，原 Sheet 名。",
        "新表名": "关键字段（程序读取）：3.7 使用，新 Sheet 名。",
    },
    SHEET_INSTITUTION_MAPPING: {
        "原始机构名称": "关键字段（程序读取）：1.2 使用，源文件中待标准化的机构名。",
        "映射后机构名称": "关键字段（程序读取）：1.2 使用，标准化后的机构名。",
        "是否为外资行": "关键字段（程序读取）：1.3 使用，1=外资行（将被剔除），0/空=否。",
    },
    SHEET_EXTRACT_CONFIG: {
        "表A-币种": "关键字段（程序读取）：1.6 匹配关键词1，按子串匹配工作表名（可空）。",
        "表A-地区": "关键字段（程序读取）：1.6 匹配关键词2（可空）。",
        "表A-机构": "关键字段（程序读取）：1.6 匹配关键词3（可空）。",
        "表A-类型": "关键字段（程序读取）：1.6 匹配关键词4（可空）。",
        "表A-名称": "关键字段（程序读取）：1.6 匹配关键词5（可空）。",
        "是否禁用": "关键字段（程序读取）：Y/是/1 禁用该行。",
        "是否整表提取": "关键字段（程序读取）：Y/是/1 整表复制；否=按提取行/列。",
        "提取行": "关键字段（程序读取）：可空=全部；支持 2,3,4,5 或 2:5。",
        "提取列": "关键字段（程序读取）：可空=全部；支持 A,B,C 或 A:D 或数字。",
        "输出文件名": "关键字段（程序读取）：输出工作簿基名（不含扩展名）；空则自动命名。",
    },
}

GLOBAL_DEFAULTS = {
    "输入目录": "input",
    "输出目录": "output",
    "日志目录": "logs",
    "源文件扩展名": ".xlsx;.xlsm;.csv",
    "错误策略": "continue",
    "默认编码": "utf-8",
    "excel目的格式": "xlsx",
    "word目的格式": "docx",
}


def _ensure_sheet_headers(ws, headers: list[str]) -> bool:
    """仅维护第1行表头，不改数据区。返回是否发生了变更。"""
    changed = False
    for i, h in enumerate(headers, start=1):
        if ws.cell(1, i).value != h:
            ws.cell(1, i).value = h
            changed = True
    return changed


def _ensure_header_comments(ws, sheet_name: str, headers: list[str]) -> bool:
    changed = False
    cm = HEADER_COMMENTS.get(sheet_name, {})
    for i, h in enumerate(headers, start=1):
        text = cm.get(h, "")
        cell = ws.cell(1, i)
        old = cell.comment.text if cell.comment is not None else ""
        if text and old != text:
            cell.comment = Comment(text, "pytools")
            changed = True
    return changed


def _ensure_global_defaults(ws) -> bool:
    """仅在键不存在或值为空时回填默认值，不覆盖已有非空值。"""
    changed = False
    key_row_map: dict[str, int] = {}
    last_row = max(ws.max_row, 1)
    for r in range(2, last_row + 1):
        key = ws.cell(r, 1).value
        if key is None:
            continue
        key_text = str(key).strip()
        if key_text:
            key_row_map[key_text] = r

    append_row = max(last_row + 1, 2)
    for k, default_v in GLOBAL_DEFAULTS.items():
        if k in key_row_map:
            rr = key_row_map[k]
            cur_v = ws.cell(rr, 2).value
            if cur_v is None or str(cur_v).strip() == "":
                ws.cell(rr, 2).value = default_v
                changed = True
        else:
            ws.cell(append_row, 1).value = k
            ws.cell(append_row, 2).value = default_v
            append_row += 1
            changed = True
    return changed


def initialize_or_repair_config(cfg_path: Path) -> dict[str, int]:
    """初始化或修复 config.xlsx。

    - 缺失 Sheet 自动创建
    - 仅修复第1行表头（不清空、不重排、不改已有配置内容）
    """
    cfg_path.parent.mkdir(parents=True, exist_ok=True)

    created = 0
    repaired = 0

    if cfg_path.exists():
        wb = load_workbook(cfg_path)
    else:
        wb = Workbook()
        # 移除默认空Sheet，避免命名冲突
        default_ws = wb.active
        wb.remove(default_ws)

    spec = [
        (SHEET_GLOBAL, GLOBAL_COLS),
        (SHEET_TIMELINE_RULE, TIMELINE_COLS),
        (SHEET_PATH_MAP, PATH_MAP_COLS),
        (SHEET_DEDUP_TASK, DEDUP_TASK_COLS),
        (SHEET_PRINT_CONFIG, PRINT_CONFIG_COLS),
        (SHEET_CONFIG_RENAME, CONFIG_RENAME_COLS),
        (SHEET_INSTITUTION_MAPPING, INSTITUTION_MAPPING_COLS),
        (SHEET_EXTRACT_CONFIG, EXTRACT_CONFIG_COLS),
    ]

    for sheet_name, headers in spec:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            h_changed = _ensure_sheet_headers(ws, headers)
            c_changed = _ensure_header_comments(ws, sheet_name, headers)
            d_changed = _ensure_global_defaults(ws) if sheet_name == SHEET_GLOBAL else False
            if h_changed or c_changed or d_changed:
                repaired += 1
        else:
            ws = wb.create_sheet(sheet_name)
            _ensure_sheet_headers(ws, headers)
            _ensure_header_comments(ws, sheet_name, headers)
            if sheet_name == SHEET_GLOBAL:
                _ensure_global_defaults(ws)
            created += 1

    wb.save(cfg_path)
    try:
        wb.close()
    except Exception:
        pass

    return {"created_sheets": created, "repaired_sheets": repaired}
