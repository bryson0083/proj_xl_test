"""渲染結果物件。"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class RenderResult:
    """``render()`` 的回傳結果。

    Attributes:
        output_path: 已寫出的輸出檔路徑。
        report: 可選的渲染報告（僅在 ``with_report=True`` 時產生），預設 ``None``。
        warnings: 渲染過程的非致命警告訊息（例如未解析的標籤）。
    """

    output_path: str
    report: Optional[Any] = None
    warnings: list[str] = field(default_factory=list)
