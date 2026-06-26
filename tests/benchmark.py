"""
大量資料渲染基準量測（spike）。

單一子程序量測單一資料量級的耗時與峰值記憶體（RSS），可乾淨取得各量級數據。

用法（每個量級各跑一次乾淨子程序）：
    uv run python -m tests.benchmark 10000
    uv run python -m tests.benchmark 100000
    uv run python -m tests.benchmark 500000

輸出 CSV 一行：rows,seconds,peak_rss_mb
"""
from __future__ import annotations

import resource
import sys
import time
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from openpyxl import Workbook

from tests.golden_compare import render_quietly
from tests.render_adapter import do_render


def _build_template(path: Path) -> None:
    """建立最小單表模板：A1 放置表格標籤，由引擎展開資料。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "data"
    ws["A1"] = "#{{big_df}}"
    wb.save(str(path))


def _make_df(rows: int) -> pd.DataFrame:
    return pd.DataFrame({
        "id": range(rows),
        "name": ["列資料"] * rows,
        "amount": [12345] * rows,
        "ratio": [0.1234] * rows,
        "note": ["x" * 8] * rows,
    })


def _peak_rss_mb() -> float:
    # macOS 的 ru_maxrss 單位為 bytes；Linux 為 KB。
    maxrss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    divisor = 1024 * 1024 if sys.platform == "darwin" else 1024
    return maxrss / divisor


def main(rows: int) -> None:
    with TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        template = tmp_path / "tpl.xlsx"
        output = tmp_path / "out.xlsx"
        _build_template(template)
        df = _make_df(rows)

        start = time.perf_counter()
        render_quietly(lambda: do_render(template, output, {"big_df": df}))
        elapsed = time.perf_counter() - start

        print(f"{rows},{elapsed:.2f},{_peak_rss_mb():.1f}")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 10000)
