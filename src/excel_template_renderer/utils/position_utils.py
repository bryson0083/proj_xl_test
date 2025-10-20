#!/usr/bin/env python3
"""
位置工具模組

提供Excel座標與行列位置轉換、範圍計算和重疊檢測等基礎功能
"""
import re
from typing import Tuple, Dict, Any, Optional
from openpyxl.utils import get_column_letter, column_index_from_string


class PositionUtils:
    """位置計算工具類別"""
    
    @staticmethod
    def excel_to_position(coordinate: str) -> Tuple[int, int]:
        """
        Excel座標(如'B5')轉換為行列位置(5, 2)
        
        Args:
            coordinate: Excel座標字串，如 'B5', 'AA10'
            
        Returns:
            tuple: (row, col) 行列位置，從1開始計數
            
        Raises:
            ValueError: 當座標格式錯誤時拋出
        """
        if not coordinate or not isinstance(coordinate, str):
            raise ValueError("座標必須是非空字串")
        
        # 解析座標格式 (如 B5, AA10)
        match = re.match(r'^([A-Z]+)(\d+)$', coordinate.upper())
        if not match:
            raise ValueError(f"無效的Excel座標格式: {coordinate}")
        
        col_str, row_str = match.groups()
        
        try:
            row = int(row_str)
            col = column_index_from_string(col_str)
            return (row, col)
        except Exception as e:
            raise ValueError(f"座標轉換失敗: {coordinate}, 錯誤: {str(e)}")
    
    @staticmethod
    def position_to_excel(row: int, col: int) -> str:
        """
        行列位置(5, 2)轉換為Excel座標('B5')
        
        Args:
            row: 行號，從1開始
            col: 列號，從1開始
            
        Returns:
            str: Excel座標字串
            
        Raises:
            ValueError: 當行列數小於1時拋出
        """
        if row < 1 or col < 1:
            raise ValueError(f"行列數必須大於0，得到: row={row}, col={col}")
        
        try:
            col_letter = get_column_letter(col)
            return f"{col_letter}{row}"
        except Exception as e:
            raise ValueError(f"位置轉換失敗: row={row}, col={col}, 錯誤: {str(e)}")
    
    @staticmethod
    def calculate_range(start_pos: Tuple[int, int], data_shape: Dict[str, int]) -> Dict[str, Any]:
        """
        根據起始位置和數據形狀計算完整範圍
        
        Args:
            start_pos: 起始位置 (row, col)
            data_shape: 數據形狀 {'rows': int, 'cols': int}
            
        Returns:
            dict: 包含範圍資訊的字典
            {
                'start_row': int,
                'start_col': int,  
                'end_row': int,
                'end_col': int,
                'start_coordinate': str,
                'end_coordinate': str,
                'range_string': str,
                'data_shape': dict
            }
        """
        if not isinstance(start_pos, tuple) or len(start_pos) != 2:
            raise ValueError("起始位置必須是包含兩個元素的tuple")
        
        if not isinstance(data_shape, dict) or 'rows' not in data_shape or 'cols' not in data_shape:
            raise ValueError("數據形狀必須包含 'rows' 和 'cols' 鍵")
        
        start_row, start_col = start_pos
        rows, cols = data_shape['rows'], data_shape['cols']
        
        if rows < 1 or cols < 1:
            raise ValueError(f"數據形狀必須大於0，得到: rows={rows}, cols={cols}")
        
        end_row = start_row + rows - 1
        end_col = start_col + cols - 1
        
        start_coordinate = PositionUtils.position_to_excel(start_row, start_col)
        end_coordinate = PositionUtils.position_to_excel(end_row, end_col)
        
        # 單一儲存格的情況
        if start_coordinate == end_coordinate:
            range_string = start_coordinate
        else:
            range_string = f"{start_coordinate}:{end_coordinate}"
        
        return {
            'start_row': start_row,
            'start_col': start_col,
            'end_row': end_row,
            'end_col': end_col,
            'start_coordinate': start_coordinate,
            'end_coordinate': end_coordinate,
            'range_string': range_string,
            'data_shape': data_shape.copy()
        }
    
    @staticmethod
    def is_position_overlap(range1: Dict[str, Any], range2: Dict[str, Any]) -> bool:
        """
        檢查兩個範圍是否重疊
        
        Args:
            range1: 第一個範圍字典
            range2: 第二個範圍字典
            
        Returns:
            bool: True 如果重疊，False 如果不重疊
        """
        required_keys = ['start_row', 'start_col', 'end_row', 'end_col']
        
        for range_dict in [range1, range2]:
            if not all(key in range_dict for key in required_keys):
                raise ValueError(f"範圍字典必須包含: {required_keys}")
        
        # 檢查水平重疊
        horizontal_overlap = not (range1['end_col'] < range2['start_col'] or 
                                range2['end_col'] < range1['start_col'])
        
        # 檢查垂直重疊
        vertical_overlap = not (range1['end_row'] < range2['start_row'] or 
                              range2['end_row'] < range1['start_row'])
        
        return horizontal_overlap and vertical_overlap
    
    @staticmethod
    def get_range_bounds(range_str: str) -> Dict[str, Any]:
        """
        解析範圍字串(如'B5:F10')為邊界資訊
        
        Args:
            range_str: Excel範圍字串，如 'B5:F10' 或 'B5'
            
        Returns:
            dict: 邊界資訊字典
        """
        if not range_str or not isinstance(range_str, str):
            raise ValueError("範圍字串不能為空")
        
        range_str = range_str.strip().upper()
        
        # 處理單一儲存格的情況
        if ':' not in range_str:
            start_pos = PositionUtils.excel_to_position(range_str)
            return PositionUtils.calculate_range(start_pos, {'rows': 1, 'cols': 1})
        
        # 處理範圍的情況
        try:
            start_coord, end_coord = range_str.split(':')
            start_pos = PositionUtils.excel_to_position(start_coord.strip())
            end_pos = PositionUtils.excel_to_position(end_coord.strip())
            
            start_row, start_col = start_pos
            end_row, end_col = end_pos
            
            if end_row < start_row or end_col < start_col:
                raise ValueError(f"結束位置不能小於起始位置: {range_str}")
            
            rows = end_row - start_row + 1
            cols = end_col - start_col + 1
            
            return PositionUtils.calculate_range(start_pos, {'rows': rows, 'cols': cols})
            
        except Exception as e:
            raise ValueError(f"無法解析範圍字串: {range_str}, 錯誤: {str(e)}")
    
    @staticmethod
    def get_displacement(from_range: Dict[str, Any], to_range: Dict[str, Any]) -> Dict[str, int]:
        """
        計算兩個範圍之間的位移
        
        Args:
            from_range: 原始範圍
            to_range: 目標範圍
            
        Returns:
            dict: 位移資訊 {'rows': int, 'cols': int}
        """
        required_keys = ['start_row', 'start_col']
        
        for range_dict, name in [(from_range, 'from_range'), (to_range, 'to_range')]:
            if not all(key in range_dict for key in required_keys):
                raise ValueError(f"{name} 必須包含: {required_keys}")
        
        row_displacement = to_range['start_row'] - from_range['start_row']
        col_displacement = to_range['start_col'] - from_range['start_col']
        
        return {
            'rows': row_displacement,
            'cols': col_displacement
        }
    
    @staticmethod
    def apply_displacement(original_range: Dict[str, Any], displacement: Dict[str, int]) -> Dict[str, Any]:
        """
        對範圍應用位移
        
        Args:
            original_range: 原始範圍
            displacement: 位移量 {'rows': int, 'cols': int}
            
        Returns:
            dict: 應用位移後的新範圍
        """
        required_keys = ['start_row', 'start_col', 'data_shape']
        if not all(key in original_range for key in required_keys):
            raise ValueError(f"原始範圍必須包含: {required_keys}")
        
        if not all(key in displacement for key in ['rows', 'cols']):
            raise ValueError("位移必須包含 'rows' 和 'cols'")
        
        new_start_row = original_range['start_row'] + displacement['rows']
        new_start_col = original_range['start_col'] + displacement['cols']
        
        if new_start_row < 1 or new_start_col < 1:
            raise ValueError(f"應用位移後的位置不能小於1: row={new_start_row}, col={new_start_col}")
        
        return PositionUtils.calculate_range(
            (new_start_row, new_start_col),
            original_range['data_shape']
        )