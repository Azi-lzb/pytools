from __future__ import annotations
from .config_xlsx import PathMapRule


def _match_kw(kw: str, name: str) -> bool:
    if not kw:
        return True
    src = (name or "").lower()
    raw = str(kw).replace("，", ",").replace("；", ";").replace(",", ";")
    parts = [p.strip().lower() for p in raw.split(";") if p.strip()]
    if not parts:
        return True
    # 多关键字按 AND：全部包含才命中
    for p in parts:
        if p not in src:
            return False
    return True


def standardize(path: str, target: str, rule_name: str,
                wb_name: str, sheet_name: str,
                maps: list[PathMapRule]) -> str:
    """target: '行头' or '列头'"""
    if not path:
        return path
    for m in maps:
        if m.applicable_rule and m.applicable_rule != rule_name:
            continue
        if m.target not in (target, "两者", ""):
            continue
        if not _match_kw(m.wb_keyword, wb_name):
            continue
        if not _match_kw(m.sheet_keyword, sheet_name):
            continue
        if m.match_mode == "包含":
            if m.original and m.original in path:
                return path.replace(m.original, m.standard)
        else:  # 精确
            if path == m.original:
                return m.standard
    return path
