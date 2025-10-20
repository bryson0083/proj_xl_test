"""
Excel Template Renderer - 輔助功能模組

提供各種輔助功能：
- 標籤排序器：優化標籤渲染順序
- 範圍檢查器：檢測範圍重疊
- 圖片物件管理器：處理圖片位移和更新
- 公式處理器：處理公式中的儲存格參照更新
- 程序隔離機制：提供安全的檔案處理環境
"""

from .tag_sorter import TagSorter
from .range_checker import RangeOverlapChecker, OverlapInfo
from .image_manager import ImageObjectManager, ShiftInfo
from .formula_processor import FormulaProcessor, FormulaShiftInfo, FormulaReference
from .process_isolation import ProcessIsolationManager, IsolationConfig, ProcessState

__all__ = [
    # 標籤排序器
    'TagSorter',
    
    # 範圍檢查器
    'RangeOverlapChecker',
    'OverlapInfo',
    
    # 圖片物件管理器
    'ImageObjectManager', 
    'ShiftInfo',
    
    # 公式處理器
    'FormulaProcessor',
    'FormulaShiftInfo',
    'FormulaReference',
    
    # 程序隔離機制
    'ProcessIsolationManager',
    'IsolationConfig',
    'ProcessState',
]
