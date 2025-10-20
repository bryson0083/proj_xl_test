#!/usr/bin/env python3
"""
物件註冊表管理器

統一的物件註冊表管理系統，整合現有的容器管理功能
"""
import uuid
from typing import Dict, List, Any, Optional, Tuple
from datetime import datetime

from openpyxl import Workbook
from openpyxl.worksheet.table import Table

from ..models.base import BlockType, ObjectType, CellPosition, DataShape, RangePosition
from ..models.container import Container
from ..models.objects import Block, ObjectInfo, TableObject, ImageObject
from ..models.tag import Tag
from .template_scanner import TemplateScanner
from .position_calculator import PositionCalculator
from ..utils.registry_utils import RegistryUtils
from ..utils.position_utils import PositionUtils


class ObjectRegistry:
    """統一的物件註冊表管理器"""
    
    def __init__(self):
        self.registry = {}  # 物件註冊表
        self.position_map = {}  # 位置映射表
        self.render_plan = {}  # 渲染計劃
        
        # 整合現有模組
        self.template_scanner = TemplateScanner()
        self.position_calculator = PositionCalculator()
        self.registry_utils = RegistryUtils()
        self.position_utils = PositionUtils()
    
    def scan_and_register(self, workbook: Workbook, template_scanner: Optional[TemplateScanner] = None):
        """
        掃描工作簿並註冊所有物件
        
        整合現有的 container.py 功能
        
        Args:
            workbook: Excel工作簿物件
            template_scanner: 可選的模板掃描器
        """
        # 使用舊版的 ContainerManager 來掃描和建立容器
        from .container import ContainerManager
        
        print("正在使用 ContainerManager 掃描模板...")
        container_manager = ContainerManager()
        containers = container_manager.create_containers(workbook)
        
        # 對每個容器進行區塊分類
        for container in containers:
            container_manager.classify_blocks(container)
        
        print(f"掃描完成，共發現 {len(containers)} 個容器")
        
        # 轉換容器為註冊表格式
        self._convert_containers_to_registry(containers, workbook)
        
        return self.registry
    
    def build_complete_registry(self, workbook: Workbook, data_context: Dict[str, Any]):
        """
        建立完整註冊表
        
        1. 掃描所有物件
        2. 計算初始位置
        3. 預測渲染後位置
        4. 生成重定位計劃
        
        Args:
            workbook: Excel工作簿物件
            data_context: 數據上下文
        """
        # 1. 掃描並註冊物件
        self.scan_and_register(workbook)
        
        # 2. 提取物件清單進行位置計算
        all_objects = self._extract_objects_for_calculation()
        
        # 3. 計算預測位置
        predicted_positions = self.position_calculator.predict_render_positions(
            all_objects, data_context
        )
        
        # 4. 生成重定位計劃
        relocation_plan = self.position_calculator.calculate_tag_relocation_plan(
            all_objects, predicted_positions
        )
        
        # 5. 更新註冊表
        self._update_registry_with_predictions(predicted_positions, relocation_plan)
        
        return self.registry
    
    def export_registry(self, output_dir: Optional[str] = None) -> str:
        """
        匯出註冊表為JSON
        
        檔案命名：template_renderer_obj_registry_{YYYYMMDD}_{HHMMSS}.json
        包含完整的渲染前後座標範圍
        
        Args:
            output_dir: 輸出目錄，預設為當前工作目錄
            
        Returns:
            str: 生成的JSON檔案完整路徑
        """
        return self.registry_utils.serialize_registry(self.registry, output_dir)
    
    def load_and_validate(self, json_path: str) -> bool:
        """
        載入並驗證註冊表JSON
        
        Args:
            json_path: JSON檔案路徑
            
        Returns:
            bool: 載入和驗證是否成功
        """
        try:
            loaded_registry = self.registry_utils.load_registry(json_path)
            is_valid, errors = self.registry_utils.validate_registry(loaded_registry)
            
            if is_valid:
                self.registry = loaded_registry
                return True
            else:
                print(f"註冊表驗證失敗: {errors}")
                return False
                
        except Exception as e:
            print(f"載入註冊表失敗: {str(e)}")
            return False
    
    def get_object_by_id(self, obj_id: str) -> Optional[Dict[str, Any]]:
        """
        根據ID取得物件資訊
        
        Args:
            obj_id: 物件ID
            
        Returns:
            dict or None: 物件資訊
        """
        for sheet_name, sheet_data in self.registry.get('worksheets', {}).items():
            for obj in sheet_data.get('objects', []):
                if obj.get('obj_id') == obj_id:
                    return obj
        return None
    
    def get_render_plan(self) -> Dict[str, Any]:
        """
        取得渲染執行計劃
        
        Returns:
            dict: 渲染計劃
        """
        return self.render_plan.copy()
    
    def get_registry_summary(self) -> Dict[str, Any]:
        """
        取得註冊表摘要資訊
        
        Returns:
            dict: 摘要資訊
        """
        summary = {
            'total_worksheets': 0,
            'total_objects': 0,
            'objects_by_type': {},
            'objects_relocated': 0,
            'position_changes': 0,
            'size_changes': 0
        }
        
        worksheets = self.registry.get('worksheets', {})
        summary['total_worksheets'] = len(worksheets)
        
        for sheet_name, sheet_data in worksheets.items():
            objects = sheet_data.get('objects', [])
            summary['total_objects'] += len(objects)
            
            for obj in objects:
                # 統計物件類型
                obj_type = obj.get('obj_type', 'unknown')
                if obj_type not in summary['objects_by_type']:
                    summary['objects_by_type'][obj_type] = 0
                summary['objects_by_type'][obj_type] += 1
                
                # 統計重定位物件
                if self._object_has_relocation(obj):
                    summary['objects_relocated'] += 1
                
                # 統計位置變化
                if self._object_has_position_change(obj):
                    summary['position_changes'] += 1
                
                # 統計大小變化
                if self._object_has_size_change(obj):
                    summary['size_changes'] += 1
        
        return summary
    
    def _convert_containers_to_registry(self, containers: List[Container], workbook: Workbook):
        """將容器物件轉換為註冊表格式"""
        self.registry = self.registry_utils.create_empty_registry()
        self.registry['worksheets'] = {}
        
        for container in containers:
            sheet_data = self._convert_single_container(container, workbook)
            self.registry['worksheets'][container.sheet_name] = sheet_data
        
        # 更新摘要
        self.registry['summary'] = self.get_registry_summary()
    
    def _convert_single_container(self, container: Container, workbook: Workbook) -> Dict[str, Any]:
        """轉換單一容器為註冊表格式"""
        sheet_data = {
            'container_id': container.container_id,
            'sheet_index': workbook.worksheets.index(workbook[container.sheet_name]),
            'objects': [],
            'blocks': [],
            'render_order': []
        }
        
        # 轉換物件
        for obj_info in container.objects:
            registry_obj = self._convert_object_info(obj_info, container.sheet_name)
            sheet_data['objects'].append(registry_obj)
            sheet_data['render_order'].append(registry_obj['obj_id'])
        
        # 轉換區塊
        for block in container.blocks:
            registry_block = self._convert_block(block)
            sheet_data['blocks'].append(registry_block)
        
        return sheet_data
    
    def _convert_object_info(self, obj_info: ObjectInfo, sheet_name: str) -> Dict[str, Any]:
        """轉換物件資訊為註冊表格式"""
        # 生成可重現的物件ID，處理空值情況
        
        # 處理位置資訊為空的情況
        if not hasattr(obj_info, 'cell_position') or obj_info.cell_position is None:
            print(f"警告：物件 {getattr(obj_info, 'obj_name', 'unknown')} 缺少位置資訊，使用預設位置")
            position_str = "1,1"  # 預設位置
        else:
            try:
                position_str = f"{obj_info.cell_position.row},{obj_info.cell_position.col}"
            except AttributeError:
                print(f"警告：物件 {getattr(obj_info, 'obj_name', 'unknown')} 位置格式錯誤，使用預設位置")
                position_str = "1,1"
        
        # 處理物件名稱為空的情況
        obj_name = getattr(obj_info, 'obj_name', None)
        if not obj_name:
            # 使用display_name或生成預設名稱
            obj_name = getattr(obj_info, 'display_name', None)
            if not obj_name:
                obj_name = f"object_at_{position_str}"
        
        # 確保工作表名稱不為空
        if not sheet_name:
            sheet_name = "Sheet1"  # 預設工作表名稱
        
        obj_id = self.registry_utils.generate_object_id(
            sheet_name, position_str, obj_name
        )
        
        # 基本物件資訊
        registry_obj = {
            'obj_id': obj_id,
            'obj_name': obj_name,  # 使用處理過的obj_name
            'display_name': getattr(obj_info, 'display_name', obj_name),
            'obj_type': obj_info.obj_type.value if obj_info.obj_type else 'unknown',
            'block_id': getattr(obj_info, 'block_id', None),
            'is_multi_rows': getattr(obj_info, 'is_multi_rows', False),
            'having_header': getattr(obj_info, 'having_header', True)
        }
        
        # 位置資訊
        position_before = self._convert_position(obj_info.cell_position, obj_info.data_shape)
        registry_obj['position_before'] = position_before
        
        # 初始時position_after與position_before相同，稍後會被預測結果覆蓋
        registry_obj['position_after'] = position_before.copy()
        
        # 數據資訊
        registry_obj['data_info'] = {
            'source': obj_info.obj_name,
            'data_shape': self._convert_data_shape(obj_info.data_shape)
        }
        
        return registry_obj
    
    def _convert_position(self, cell_pos: CellPosition, data_shape: DataShape) -> Dict[str, Any]:
        """轉換位置資訊"""
        coordinate = self.position_utils.position_to_excel(cell_pos.row, cell_pos.col)
        shape_dict = {'rows': data_shape.rows, 'cols': data_shape.cols}
        
        range_info = self.position_utils.calculate_range(
            (cell_pos.row, cell_pos.col), shape_dict
        )
        
        return {
            'row': cell_pos.row,
            'col': cell_pos.col,
            'coordinate': coordinate,
            'range_start': range_info['start_coordinate'],
            'range_end': range_info['end_coordinate'],
            'range_description': range_info['range_string'],
            'data_shape': shape_dict
        }
    
    def _convert_data_shape(self, data_shape: DataShape) -> Dict[str, int]:
        """轉換數據形狀"""
        return {
            'rows': data_shape.rows,
            'cols': data_shape.cols
        }
    
    def _convert_block(self, block: Block) -> Dict[str, Any]:
        """轉換區塊資訊"""
        return {
            'block_id': block.block_id,
            'block_type': block.block_type.value if block.block_type else 'unknown',
            'range_from': {
                'row': block.rng_from.row,
                'col': block.rng_from.col
            },
            'range_to': {
                'row': block.rng_to.row,
                'col': block.rng_to.col
            }
        }
    
    def _extract_objects_for_calculation(self) -> List[Dict[str, Any]]:
        """提取物件清單用於位置計算"""
        all_objects = []
        
        for sheet_name, sheet_data in self.registry.get('worksheets', {}).items():
            for obj in sheet_data.get('objects', []):
                # 轉換為計算器所需格式
                calc_obj = {
                    'obj_id': obj['obj_id'],
                    'obj_name': obj['obj_name'],
                    'obj_type': obj['obj_type'],
                    'position_before': obj['position_before'],
                    'data_shape': obj['position_before']['data_shape'],
                    'having_header': obj.get('having_header', True),
                    '_sheet_name': sheet_name
                }
                all_objects.append(calc_obj)
        
        return all_objects
    
    def _update_registry_with_predictions(self, predicted_positions: Dict[str, Dict[str, Any]], 
                                        relocation_plan: List[Dict[str, Any]]):
        """使用預測結果更新註冊表"""
        # 更新物件位置
        for sheet_name, sheet_data in self.registry.get('worksheets', {}).items():
            for obj in sheet_data.get('objects', []):
                obj_id = obj['obj_id']
                
                if obj_id in predicted_positions:
                    prediction = predicted_positions[obj_id]
                    
                    # 更新position_after
                    obj['position_after'] = {
                        'row': prediction['start_row'],
                        'col': prediction['start_col'],
                        'coordinate': self.position_utils.position_to_excel(
                            prediction['start_row'], prediction['start_col']
                        ),
                        'range_start': prediction['start_coordinate'],
                        'range_end': prediction['end_coordinate'],
                        'range_description': prediction['range_string'],
                        'data_shape': prediction['data_shape']
                    }
                    
                    # 添加位移資訊
                    if 'cumulative_displacement' in prediction:
                        obj['position_after']['displacement'] = {
                            'rows': prediction['cumulative_displacement']['rows'],
                            'cols': prediction['cumulative_displacement']['cols'],
                            'reason': self._get_displacement_reason(prediction)
                        }
                    
                    # 更新數據資訊
                    obj['data_info'].update({
                        'actual_rows': prediction['data_shape']['rows'],
                        'actual_cols': prediction['data_shape']['cols'],
                        'size_changed': prediction.get('size_change', {'rows': 0, 'cols': 0})
                    })
            
            # 添加重定位計劃
            sheet_relocations = [
                item for item in relocation_plan 
                if any(obj['obj_id'] == item['obj_id'] for obj in sheet_data.get('objects', []))
            ]
            
            if sheet_relocations:
                sheet_data['relocation_plan'] = sheet_relocations
    
    def _get_displacement_reason(self, prediction: Dict[str, Any]) -> str:
        """取得位移原因"""
        displacement = prediction.get('cumulative_displacement', {'rows': 0, 'cols': 0})
        size_change = prediction.get('size_change', {'rows': 0, 'cols': 0})
        
        reasons = []
        if displacement['rows'] != 0:
            reasons.append(f"垂直位移{displacement['rows']:+d}行")
        if displacement['cols'] != 0:
            reasons.append(f"水平位移{displacement['cols']:+d}列")
        if size_change['rows'] != 0 or size_change['cols'] != 0:
            reasons.append("數據形狀變化")
        
        return "; ".join(reasons) if reasons else "位置調整"
    
    def _object_has_relocation(self, obj: Dict[str, Any]) -> bool:
        """檢查物件是否需要重定位"""
        pos_before = obj.get('position_before', {})
        pos_after = obj.get('position_after', {})
        
        return (pos_before.get('row') != pos_after.get('row') or
                pos_before.get('col') != pos_after.get('col'))
    
    def _object_has_position_change(self, obj: Dict[str, Any]) -> bool:
        """檢查物件是否有位置變化"""
        return self._object_has_relocation(obj)
    
    def _object_has_size_change(self, obj: Dict[str, Any]) -> bool:
        """檢查物件是否有大小變化"""
        data_info = obj.get('data_info', {})
        size_change = data_info.get('size_changed', {'rows': 0, 'cols': 0})
        
        return size_change.get('rows', 0) != 0 or size_change.get('cols', 0) != 0
    
    def validate_registry_integrity(self) -> Tuple[bool, List[str]]:
        """驗證註冊表完整性"""
        return self.registry_utils.validate_registry(self.registry)
    
    def clear_registry(self):
        """清除註冊表"""
        self.registry = {}
        self.position_map = {}
        self.render_plan = {}
    
    def get_objects_by_sheet(self, sheet_name: str) -> List[Dict[str, Any]]:
        """取得指定工作表的所有物件"""
        worksheets = self.registry.get('worksheets', {})
        sheet_data = worksheets.get(sheet_name, {})
        return sheet_data.get('objects', [])