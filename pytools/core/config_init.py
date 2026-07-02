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
    SHEET_ARCHIVE_TYPE_CONFIG,
    SHEET_SUBMISSION_CHECK_CONFIG,
    SHEET_FILE_SORT_CONFIG,
    TIMELINE_COLS,
    PATH_MAP_COLS,
    DEDUP_TASK_COLS,
    PRINT_CONFIG_COLS,
    CONFIG_RENAME_COLS,
    INSTITUTION_MAPPING_COLS,
    EXTRACT_CONFIG_COLS,
    ARCHIVE_TYPE_CONFIG_COLS,
    SUBMISSION_CHECK_CONFIG_COLS,
    FILE_SORT_CONFIG_COLS,
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
        "目标去重列": "关键字段（程序读取）：可空；宽表目标写入去重列，填写输出表列名或列标/列号（如 数据日期;行头路径 或 A;B;C），支持 分号/逗号 分隔。空则默认按 数据日期+set列+行头路径 去重。",
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
        "打印模式": "关键字段（程序读取）：1=保留源格式；2=保留源格式+自动扩列（减少####）；3=快速复制（值+基础格式，速度更快）。3.10.7 会按 1/2/3 全部执行。",
        "源工作簿": "关键字段（程序读取）：源工作簿完整路径。",
        "源工作表": "关键字段（程序读取）：源工作表名（精确匹配）。",
        "源工作表打印区域": "关键字段（程序读取）：可空；支持 A1:H30;A35:H70，多段分号分隔。",
        "目标工作簿": "关键字段（程序读取）：目标工作簿完整路径；不存在会自动创建。",
        "目标工作表": "关键字段（程序读取）：目标工作表名；不存在会自动创建。",
        "FitToPagesWide": "关键字段（程序读取）：分页宽度页数，留空默认 1。",
        "FitToPagesTall": "关键字段（程序读取）：分页高度页数，留空默认 1。",
        "打印方向": "关键字段（程序读取）：横向/纵向；留空按宽高比自动。",
        "水平居中": "关键字段（程序读取）：Y/是/1 时打印页面水平居中。",
        "垂直居中": "关键字段（程序读取）：Y/是/1 时打印页面垂直居中。",
        "不输出批注": "关键字段（程序读取）：Y/是/1 时清除输出文件中的单元格批注。",
        "零值不输出": "关键字段（程序读取）：Y/是/1 时将 0 值输出为空白。",
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
        "内容重命名启用": "关键字段（程序读取）：4.5/4.6 使用，Y/是/1/true 表示启用。只能启用一条。",
        "内容重命名取值规则": "关键字段（程序读取）：4.5/4.6 使用，填写 sheet@单元格 规则，多个用分号分隔。",
        "原文件名片段": "关键字段（程序读取）：4.8 使用，文件名中要替换的片段（M列）。",
        "新文件名片段": "关键字段（程序读取）：4.8 使用，替换后的文件名片段（N列）。",
        "备注": "备注字段（仅人工说明）：程序不读取。建议写适用文件或报表类型。",
    },
    SHEET_ARCHIVE_TYPE_CONFIG: {
        "是否启用": "关键字段（程序读取）：Y/是/1 表示启用该类型规则。",
        "匹配关键词": "关键字段（程序读取）：按类型归档时，在文件名中查找的关键词。",
        "归档文件夹": "关键字段（程序读取）：命中关键词后放入的单层文件夹名称。",
        "备注": "备注字段（仅人工说明）：程序不读取。建议把更具体的关键词放在前面。",
        "需要归档的后缀": "关键字段（程序读取）：E2 填写允许归档的后缀，支持 xlsx;.xls;docx 多个用分号/逗号/空格分隔；空表示不限制。",
        "排除归档的后缀": "关键字段（程序读取）：F2 填写要排除的后缀，支持 exe;dll;tmp 多个用分号/逗号/空格分隔；优先级高于 E2。",
    },
    SHEET_SUBMISSION_CHECK_CONFIG: {
        "是否启用": "关键字段（程序读取）：Y/是/1/true 表示启用该检查项。",
        "日期关键词": "关键字段（程序读取）：可空；非空时必须出现在文件名中。",
        "县区关键词": "关键字段（程序读取）：可空；非空时必须出现在文件名中。",
        "机构关键词": "关键字段（程序读取）：可空；非空时必须出现在文件名中。",
        "备用关键词": "关键字段（程序读取）：可空；非空时必须出现在文件名中，可填报表类型等。",
        "命中次数": "输出字段（程序写入）：匹配到的文件数量。",
        "命中的文件名": "输出字段（程序写入）：匹配到的文件名，多个用分号分隔。",
        "备注": "备注字段（仅人工说明）：程序不读取。",
    },
    SHEET_FILE_SORT_CONFIG: {
        "是否启用": "关键字段（程序读取）：Y/是/1/true 表示启用该排序规则。",
        "排序前缀": "关键字段（程序读取）：支持 1/01/001，程序统一补齐为三位数字。",
        "匹配关键词": "关键字段（程序读取）：文件名命中任一关键词即使用该前缀，多个用分号/逗号分隔。",
        "备注": "备注字段（仅人工说明）：程序不读取。",
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
    "打印零值不输出": "否",
    "时序规则命中策略": "all_match",
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


def _ensure_rename_sheet_extra_cells(ws) -> bool:
    changed = False
    examples = {
        "O2": "Y",
        "P2": "sheet1@A1;sheet2@A2",
        "Q2": "示例：按文件内容单元格生成前缀",
    }
    for addr, value in examples.items():
        if ws[addr].value is None or str(ws[addr].value).strip() == "":
            ws[addr].value = value
            changed = True
    for addr in ("L1", "L2"):
        if ws[addr].comment is not None:
            ws[addr].comment = None
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


def _ensure_archive_type_examples(ws) -> bool:
    """仅在无数据行时写入示例，不覆盖用户配置。"""
    if ws.max_row > 1:
        return False
    examples = [
        ("是", "普惠贷款表", "普惠贷款表", "更具体的关键词放前面"),
        ("是", "科技贷款表", "科技贷款表", ""),
        ("是", "存款表", "存款表", ""),
        ("是", "贷款表", "贷款表", "短关键词放后面"),
    ]
    for row_no, row in enumerate(examples, start=2):
        for col_no, value in enumerate(row, start=1):
            ws.cell(row_no, col_no).value = value
    return True


def _ensure_submission_check_examples(ws) -> bool:
    """仅在无数据行时写入示例，不覆盖用户配置。"""
    if ws.max_row > 1:
        return False
    examples = [
        ["Y", "202506", "惠城区", "建设银行", "存款表", "", "", "示例：四个非空关键词都要出现在同一文件名"],
        ["Y", "202506", "", "工商银行", "贷款表", "", "", "示例：文件名没有县区时，县区关键词可留空"],
    ]
    for row_no, row in enumerate(examples, start=2):
        for col_no, value in enumerate(row, start=1):
            ws.cell(row_no, col_no).value = value
    return True


def _ensure_file_sort_examples(ws) -> bool:
    """仅在无数据行时写入示例，不覆盖用户配置。"""
    if ws.max_row > 1:
        return False
    examples = [
        ["Y", "001", "政策性银行;国开行;农发行", "政策性银行"],
        ["Y", "002", "工商银行;农业银行;中国银行;建设银行;交通银行", "国有银行"],
        ["Y", "003", "招商银行;浦发银行;中信银行", "股份制银行"],
    ]
    for row_no, row in enumerate(examples, start=2):
        for col_no, value in enumerate(row, start=1):
            ws.cell(row_no, col_no).value = value
    return True


def initialize_or_repair_config(cfg_path: Path) -> dict[str, int]:
    """初始化或修复 config.xlsx。

    - 缺失 Sheet 自动创建
    - 仅修复第1行表头（不清空、不重排、不改已有配置内容）
    """
    cfg_path.parent.mkdir(parents=True, exist_ok=True)

    created = 0
    repaired = 0
    details: list[str] = []

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
        (SHEET_ARCHIVE_TYPE_CONFIG, ARCHIVE_TYPE_CONFIG_COLS),
        (SHEET_SUBMISSION_CHECK_CONFIG, SUBMISSION_CHECK_CONFIG_COLS),
        (SHEET_FILE_SORT_CONFIG, FILE_SORT_CONFIG_COLS),
        (SHEET_INSTITUTION_MAPPING, INSTITUTION_MAPPING_COLS),
        (SHEET_EXTRACT_CONFIG, EXTRACT_CONFIG_COLS),
    ]

    for sheet_name, headers in spec:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            h_changed = _ensure_sheet_headers(ws, headers)
            c_changed = _ensure_header_comments(ws, sheet_name, headers)
            r_changed = _ensure_rename_sheet_extra_cells(ws) if sheet_name == SHEET_CONFIG_RENAME else False
            a_changed = _ensure_archive_type_examples(ws) if sheet_name == SHEET_ARCHIVE_TYPE_CONFIG else False
            sub_changed = _ensure_submission_check_examples(ws) if sheet_name == SHEET_SUBMISSION_CHECK_CONFIG else False
            fs_changed = _ensure_file_sort_examples(ws) if sheet_name == SHEET_FILE_SORT_CONFIG else False
            d_changed = _ensure_global_defaults(ws) if sheet_name == SHEET_GLOBAL else False
            if h_changed or c_changed or d_changed or r_changed or a_changed or sub_changed or fs_changed:
                repaired += 1
                changed_items: list[str] = []
                if h_changed:
                    changed_items.append("headers")
                if c_changed:
                    changed_items.append("comments")
                if r_changed:
                    changed_items.append("extra_cells")
                if a_changed:
                    changed_items.append("examples")
                if sub_changed:
                    changed_items.append("examples")
                if fs_changed:
                    changed_items.append("examples")
                if d_changed:
                    changed_items.append("defaults")
                details.append(f"{sheet_name}: repaired ({','.join(changed_items)})")
            else:
                details.append(f"{sheet_name}: unchanged")
        else:
            ws = wb.create_sheet(sheet_name)
            _ensure_sheet_headers(ws, headers)
            _ensure_header_comments(ws, sheet_name, headers)
            if sheet_name == SHEET_CONFIG_RENAME:
                _ensure_rename_sheet_extra_cells(ws)
            if sheet_name == SHEET_ARCHIVE_TYPE_CONFIG:
                _ensure_archive_type_examples(ws)
            if sheet_name == SHEET_SUBMISSION_CHECK_CONFIG:
                _ensure_submission_check_examples(ws)
            if sheet_name == SHEET_FILE_SORT_CONFIG:
                _ensure_file_sort_examples(ws)
            if sheet_name == SHEET_GLOBAL:
                _ensure_global_defaults(ws)
            created += 1
            details.append(f"{sheet_name}: created")

    wb.save(cfg_path)
    try:
        wb.close()
    except Exception:
        pass

    return {"created_sheets": created, "repaired_sheets": repaired, "details": details}
