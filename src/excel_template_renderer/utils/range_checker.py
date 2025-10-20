"""
範圍重疊檢查器
"""
from typing import List, Optional, Tuple
from dataclasses import dataclass

from ..models.container import Container
from ..models.objects import ObjectInfo
from ..models.base import CellPosition, DataShape
from ..exceptions import RangeOverlapError


@dataclass
class OverlapInfo:
    """重疊資訊"""
    obj1_id: str
    obj2_id: str
    obj1_range: str
    obj2_range: str
    overlap_area: str
    severity: str  # "warning", "error"
    message: str


class RangeOverlapChecker:
    """
    範圍重疊檢查器類別
    
    負責檢查物件範圍是否重疊，防止渲染衝突
    """
    
    def check_overlap(self, container: Container) -> List[OverlapInfo]:
        """
        檢查物件範圍重疊情況
        
        Args:
            container: 容器物件
            
        Returns:
            List[OverlapInfo]: 重疊資訊清單
        """
        overlaps = []
        objects = container.objects
        
        # 檢查每對物件的重疊情況
        for i in range(len(objects)):
            for j in range(i + 1, len(objects)):
                obj1, obj2 = objects[i], objects[j]
                overlap_info = self._check_object_pair_overlap(obj1, obj2)
                if overlap_info:
                    overlaps.append(overlap_info)
        
        return overlaps
    
    def _check_object_pair_overlap(self, obj1: ObjectInfo, obj2: ObjectInfo) -> Optional[OverlapInfo]:
        """
        檢查兩個物件是否重疊
        
        Args:
            obj1: 物件1
            obj2: 物件2
            
        Returns:
            Optional[OverlapInfo]: 重疊資訊，如果沒有重疊則返回None
        """
        # 計算物件範圍
        range1 = self._calculate_object_range(obj1)
        range2 = self._calculate_object_range(obj2)
        
        # 檢查是否重疊
        overlap_area = self._calculate_overlap_area(range1, range2)
        if overlap_area:
            severity = self._determine_overlap_severity(obj1, obj2, overlap_area)
            message = self._generate_overlap_message(obj1, obj2, severity)
            
            return OverlapInfo(
                obj1_id=obj1.obj_id,
                obj2_id=obj2.obj_id,
                obj1_range=self._format_range(range1),
                obj2_range=self._format_range(range2),
                overlap_area=self._format_range(overlap_area),
                severity=severity,
                message=message
            )
        
        return None
    
    def _calculate_object_range(self, obj: ObjectInfo) -> Tuple[int, int, int, int]:
        """
        計算物件的範圍
        
        Args:
            obj: 物件資訊
            
        Returns:
            Tuple[int, int, int, int]: (start_row, start_col, end_row, end_col)
        """
        start_row = obj.cell_position.row
        start_col = obj.cell_position.col
        end_row = start_row + obj.data_shape.rows - 1
        end_col = start_col + obj.data_shape.cols - 1
        
        return (start_row, start_col, end_row, end_col)
    
    def _calculate_overlap_area(
        self, 
        range1: Tuple[int, int, int, int], 
        range2: Tuple[int, int, int, int]
    ) -> Optional[Tuple[int, int, int, int]]:
        """
        計算兩個範圍的重疊區域
        
        Args:
            range1: 範圍1 (start_row, start_col, end_row, end_col)
            range2: 範圍2 (start_row, start_col, end_row, end_col)
            
        Returns:
            Optional[Tuple[int, int, int, int]]: 重疊區域，如果沒有重疊則返回None
        """
        r1_start_row, r1_start_col, r1_end_row, r1_end_col = range1
        r2_start_row, r2_start_col, r2_end_row, r2_end_col = range2
        
        # 計算重疊區域的邊界
        overlap_start_row = max(r1_start_row, r2_start_row)
        overlap_start_col = max(r1_start_col, r2_start_col)
        overlap_end_row = min(r1_end_row, r2_end_row)
        overlap_end_col = min(r1_end_col, r2_end_col)
        
        # 檢查是否真的有重疊
        if overlap_start_row <= overlap_end_row and overlap_start_col <= overlap_end_col:
            return (overlap_start_row, overlap_start_col, overlap_end_row, overlap_end_col)
        
        return None
    
    def _determine_overlap_severity(
        self, 
        obj1: ObjectInfo, 
        obj2: ObjectInfo, 
        overlap_area: Tuple[int, int, int, int]
    ) -> str:
        """
        判斷重疊的嚴重程度
        
        Args:
            obj1: 物件1
            obj2: 物件2
            overlap_area: 重疊區域
            
        Returns:
            str: 嚴重程度 ("warning" 或 "error")
        """
        # 如果是表格物件重疊，通常是錯誤
        if (obj1.obj_type.value in ["table", "table_obj"] and 
            obj2.obj_type.value in ["table", "table_obj"]):
            return "error"
        
        # 如果是簡單變數重疊，可能只是警告
        if (obj1.obj_type.value == "simple" and obj2.obj_type.value == "simple"):
            return "warning"
        
        # 混合類型重疊，視為錯誤
        return "error"
    
    def _generate_overlap_message(self, obj1: ObjectInfo, obj2: ObjectInfo, severity: str) -> str:
        """
        生成重疊訊息
        
        Args:
            obj1: 物件1
            obj2: 物件2
            severity: 嚴重程度
            
        Returns:
            str: 重疊訊息
        """
        if severity == "error":
            return f"錯誤：物件 '{obj1.display_name}' 與 '{obj2.display_name}' 範圍重疊，可能導致渲染衝突"
        else:
            return f"警告：物件 '{obj1.display_name}' 與 '{obj2.display_name}' 範圍重疊，請檢查模板設計"
    
    def _format_range(self, range_tuple: Tuple[int, int, int, int]) -> str:
        """
        格式化範圍為字串
        
        Args:
            range_tuple: 範圍元組 (start_row, start_col, end_row, end_col)
            
        Returns:
            str: 格式化的範圍字串
        """
        start_row, start_col, end_row, end_col = range_tuple
        start_cell = self._get_column_letter(start_col) + str(start_row)
        end_cell = self._get_column_letter(end_col) + str(end_row)
        
        if start_cell == end_cell:
            return start_cell
        else:
            return f"{start_cell}:{end_cell}"
    
    def _get_column_letter(self, col_num: int) -> str:
        """
        將列號轉換為Excel列字母
        
        Args:
            col_num: 列號 (1-based)
            
        Returns:
            str: 列字母 (如 A, B, AA, AB...)
        """
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(65 + col_num % 26) + result
            col_num //= 26
        return result
    
    def validate_container_ranges(self, container: Container) -> None:
        """
        驗證容器中的範圍是否有衝突
        
        Args:
            container: 容器物件
            
        Raises:
            RangeOverlapError: 如果發現嚴重的範圍重疊
        """
        overlaps = self.check_overlap(container)
        
        # 檢查是否有錯誤級別的重疊
        error_overlaps = [overlap for overlap in overlaps if overlap.severity == "error"]
        
        if error_overlaps:
            # 抛出第一個錯誤
            first_error = error_overlaps[0]
            raise RangeOverlapError(
                first_error.obj1_id, 
                first_error.obj2_id, 
                container.sheet_name
            )
    
    def get_safe_rendering_order(self, container: Container) -> List[str]:
        """
        取得安全的渲染順序
        
        分析重疊情況，返回避免衝突的渲染順序
        
        Args:
            container: 容器物件
            
        Returns:
            List[str]: 物件ID的安全渲染順序
        """
        overlaps = self.check_overlap(container)
        
        # 如果沒有重疊，按位置順序返回
        if not overlaps:
            sorted_objects = sorted(
                container.objects, 
                key=lambda obj: (obj.cell_position.row, obj.cell_position.col)
            )
            return [obj.obj_id for obj in sorted_objects]
        
        # 如果有重疊，需要更複雜的排序邏輯
        # 暫時實現：優先渲染位置在上方的物件
        sorted_objects = sorted(
            container.objects,
            key=lambda obj: (obj.cell_position.row, obj.cell_position.col)
        )
        return [obj.obj_id for obj in sorted_objects]
