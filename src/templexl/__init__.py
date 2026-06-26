"""
templexl

基於 openpyxl 的 Excel 模板報表引擎——以 Excel 檔作為模板、將資料渲染成報表。
定位為 xlwings Reports 的免費替代品；採用 {{variable}} 與 #{{dataframe}} 語法，
無需安裝 Excel 即可使用。
"""

import logging

__version__ = "0.1.0"
__author__ = "Bryson Xue"
__email__ = "bryson0083@gmail.com"

# 函式庫慣例：預設掛 NullHandler，呼叫端未設定 logging 時不產生任何輸出。
logging.getLogger(__name__).addHandler(logging.NullHandler())

from .api import render
from .result import RenderResult
from .exceptions import (
    TemplateError,
    TemplateNotFoundError,
    FileFormatError,
    RenderError,
)

__all__ = [
    "render",
    "RenderResult",
    "TemplateError",
    "TemplateNotFoundError",
    "FileFormatError",
    "RenderError",
]
