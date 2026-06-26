"""
Excel Template Renderer

輕量級的Excel模板渲染工具，模仿xlwings pro模塊xlwings Reports功能。
支持簡單變數替換和表格數據渲染，採用{{variable}}和#{{dataframe}}語法。
"""

import logging

__version__ = "0.1.0"
__author__ = "Your Name"
__email__ = "your.email@example.com"

# 函式庫慣例：預設掛 NullHandler，呼叫端未設定 logging 時不產生任何輸出。
logging.getLogger(__name__).addHandler(logging.NullHandler())

from .api import render_template

__all__ = ["render_template"]
