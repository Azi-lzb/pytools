from __future__ import annotations
import re
from datetime import datetime

_PATTERNS = [
    # 优先匹配明确分隔的完整日期，避免把机构代码里的 201401 误识别为日期
    (re.compile(r"(?<!\d)(20\d{2})[年\-_./](\d{1,2})[月\-_./](\d{1,2})(?:日)?(?!\d)"), "ymd"),
    # 再匹配紧凑的 YYYYMMDD（要求数字边界）
    (re.compile(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])([0-2]\d|3[01])(?!\d)"), "ymd"),
    # 退化到年月
    (re.compile(r"(?<!\d)(20\d{2})[年\-_./](\d{1,2})(?:月)?(?!\d)"), "ym"),
    (re.compile(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?!\d)"), "ym"),
]


def _normalize(y: int, m: int, d: int | None) -> str:
    try:
        if d is not None:
            return datetime(y, m, d).strftime("%Y-%m-%d")
        return datetime(y, m, 1).strftime("%Y-%m")
    except ValueError:
        return ""


def parse_data_date(source_file: str, sheet_a1: str | None = None) -> tuple[str, str]:
    """返回 (日期字符串, 来源)。来源为 '文件名'/'A1'/''."""
    for pat, kind in _PATTERNS:
        m = pat.search(source_file or "")
        if m:
            y = int(m.group(1))
            mo = int(m.group(2))
            d = int(m.group(3)) if kind == "ymd" else None
            r = _normalize(y, mo, d)
            if r:
                return r, "文件名"
    if sheet_a1:
        for pat, kind in _PATTERNS:
            m = pat.search(str(sheet_a1))
            if m:
                y = int(m.group(1))
                mo = int(m.group(2))
                d = int(m.group(3)) if kind == "ymd" else None
                r = _normalize(y, mo, d)
                if r:
                    return r, "A1"
    return "", ""
