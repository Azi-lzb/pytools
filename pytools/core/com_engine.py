"""Excel COM 共用工具：启动、读 sheet 值为 2D 数组、安全关闭等。

所有 COM 引擎版本的功能模块（compare_393_com / timeline_396_com / wide_397_398_com / ...）
统一依赖本模块，避免代码重复。
"""
from __future__ import annotations

from pathlib import Path
import time
import os


XL_CALC_MANUAL = -4135


def _com_debug_enabled() -> bool:
    return str(os.environ.get("PYTOOLS_DEBUG_COM", "")).strip().lower() in ("1", "true", "yes", "on")


class ArrayLike:
    """2D list 的薄壳：暴露 .shape 和 .iat[r, c]，与 pandas.DataFrame 在 extract_cells 等
    纯计算函数里的使用方式兼容，从而让 COM 分支和 pandas 分支复用同一套下游逻辑。"""

    class _IAt:
        __slots__ = ("_data",)

        def __init__(self, data: list[list]) -> None:
            self._data = data

        def __getitem__(self, key):
            r, c = key
            row = self._data[r]
            if c < len(row):
                return row[c]
            return None

    __slots__ = ("data", "shape", "iat")

    def __init__(self, data: list[list]) -> None:
        self.data = data
        nrow = len(data)
        ncol = max((len(r) for r in data), default=0)
        self.shape = (nrow, ncol)
        self.iat = ArrayLike._IAt(data)

    @property
    def values(self):
        """与 pandas.DataFrame.values 等价语义：支持 values[r][c] 访问。"""
        return self.data


def start_excel_app():
    """启动 Excel，关闭所有告警/事件/刷屏/计算，返回 Application 对象。
    任何异常均抛 RuntimeError，供调用方在 UI 处做"COM 不可用"回退提示。"""
    try:
        import win32com.client  # type: ignore
    except Exception as e:
        raise RuntimeError(f"缺少 pywin32，无法使用 COM 引擎: {e}")
    last_err = None
    app = None
    for pid in ("Excel.Application", "ket.Application", "KET.Application"):
        try:
            app = win32com.client.DispatchEx(pid)
            break
        except Exception as e:
            last_err = e
    if app is None:
        raise RuntimeError(f"未能启动 Excel COM 应用: {last_err}")
    app.Visible = False
    app.DisplayAlerts = False
    for attr, val in (
        ("ScreenUpdating", False),
        ("EnableEvents", False),
        ("AskToUpdateLinks", False),
        ("AlertBeforeOverwriting", False),
    ):
        try:
            setattr(app, attr, val)
        except Exception:
            pass
    try:
        app.Calculation = XL_CALC_MANUAL
    except Exception:
        pass
    return app


def read_sheet_values(ws) -> ArrayLike:
    """读 sheet 从 A1 到 UsedRange 右下角的值矩阵。0-based 索引。"""
    ur = ws.UsedRange
    last_row = int(ur.Row) + int(ur.Rows.Count) - 1
    last_col = int(ur.Column) + int(ur.Columns.Count) - 1
    if last_row < 1 or last_col < 1:
        return ArrayLike([])
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col))
    v = rng.Value2
    if v is None:
        return ArrayLike([])
    if not isinstance(v, tuple):
        return ArrayLike([[v]])
    return ArrayLike([list(row) for row in v])


def read_merged_aware_texts(ws, coords: list[tuple[int, int]]) -> dict[tuple[int, int], str]:
    """按 0-based 坐标批量读取文本；命中合并区域时取左上角文本。

    只读取少量指定坐标，避免 `.xls` 下 `MergeAreas` 不可用时回退全表扫描。
    """
    out: dict[tuple[int, int], str] = {}
    seen: set[tuple[int, int]] = set()
    for r0, c0 in coords:
        if r0 < 0 or c0 < 0:
            continue
        key = (r0, c0)
        if key in seen:
            continue
        seen.add(key)
        text = ""
        try:
            cell = ws.Cells(r0 + 1, c0 + 1)
            try:
                if bool(cell.MergeCells):
                    v = cell.MergeArea.Cells(1, 1).Value
                else:
                    v = cell.Value
            except Exception:
                v = cell.Value
            if v is not None:
                text = str(v).strip()
        except Exception:
            text = ""
        out[key] = text
    return out


def get_merge_ranges_com(ws) -> tuple[tuple[int, int, int, int], ...]:
    """从当前 COM worksheet 读取合并区域，返回 0-based 坐标元组。"""
    ranges: list[tuple[int, int, int, int]] = []
    debug = _com_debug_enabled()
    t0 = time.perf_counter()
    sheet_name = ""
    try:
        sheet_name = str(ws.Name)
    except Exception:
        sheet_name = "<unknown>"
    try:
        used = ws.UsedRange
    except Exception:
        if debug:
            print(f"[COM][merge] {sheet_name}: UsedRange 读取失败")
        return tuple()

    try:
        merge_areas = used.MergeAreas
        for ma in merge_areas:
            ranges.append((
                int(ma.Row) - 1,
                int(ma.Row + ma.Rows.Count - 1) - 1,
                int(ma.Column) - 1,
                int(ma.Column + ma.Columns.Count - 1) - 1,
            ))
        if debug:
            print(
                f"[COM][merge] {sheet_name}: MergeAreas 路径, "
                f"count={len(ranges)}, cost={time.perf_counter() - t0:.2f}s"
            )
    except Exception:
        # 回退逐格扫描在 COM 下非常慢（常见分钟级），默认禁用以保证时效。
        if debug:
            print(
                f"[COM][merge] {sheet_name}: MergeAreas 失败，已跳过逐格回退扫描，"
                f"cost={time.perf_counter() - t0:.2f}s"
            )
    return tuple(sorted(set(ranges)))


def list_sheet_names_com(wb) -> list[str]:
    return [str(ws.Name) for ws in wb.Worksheets]


def find_sheet_com(wb, name: str):
    for ws in wb.Worksheets:
        if str(ws.Name) == name:
            return ws
    return None


def open_readonly(app, path: Path):
    # 显式关闭常见交互项，避免 COM 调用被弹窗阻塞。
    return app.Workbooks.Open(
        str(path.resolve()),
        ReadOnly=True,
        UpdateLinks=0,
        AddToMru=False,
        Notify=False,
        IgnoreReadOnlyRecommended=True,
        CorruptLoad=0,
        Local=True,
    )


def safe_close(wb) -> None:
    if wb is None:
        return
    try:
        wb.Close(SaveChanges=False)
    except Exception:
        pass


def safe_quit(app) -> None:
    if app is None:
        return
    try:
        app.Quit()
    except Exception:
        pass
