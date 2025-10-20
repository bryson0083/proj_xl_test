#!/usr/bin/env python3
"""
位置計算引擎

提供物件位置計算、渲染預測、標籤重定位計劃等功能
"""
from typing import Dict, List, Any, Tuple, Optional
from ..utils.position_utils import PositionUtils
from ..utils.dependency_utils import DependencyAnalyzer
from ..models.base import ObjectType


class PositionCalculator:
    """位置計算引擎"""
    
    def __init__(self):
        self.position_utils = PositionUtils()
        self.dependency_analyzer = DependencyAnalyzer()
    
    def calculate_initial_positions(self, objects: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
        """
        計算所有物件的初始位置和範圍
        
        Args:
            objects: 物件清單
            
        Returns:
            dict: 物件ID對應的初始位置資訊
        """
        positions = {}
        
        for obj in objects:
            obj_id = obj.get('obj_id')
            if not obj_id:
                continue
            
            # 從物件中提取位置資訊
            initial_position = self._extract_initial_position(obj)
            if initial_position:
                # 計算完整的範圍資訊
                data_shape = self._get_object_data_shape(obj)
                complete_range = self.position_utils.calculate_range(
                    (initial_position['row'], initial_position['col']),
                    data_shape
                )
                
                positions[obj_id] = {
                    **complete_range,
                    'obj_type': obj.get('obj_type', 'unknown'),
                    'data_shape': data_shape,
                    'original_position': initial_position
                }
        
        return positions
    
    def predict_render_positions(self, objects: List[Dict[str, Any]], 
                               data_context: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
        """
        預測渲染後的位置
        
        Args:
            objects: 物件清單
            data_context: 數據上下文，包含實際要渲染的數據
            
        Returns:
            dict: 物件ID對應的預測位置資訊
        """
        # 1. 計算初始位置
        initial_positions = self.calculate_initial_positions(objects)
        
        # 2. 建立依賴關係圖
        dependency_graph = self.dependency_analyzer.build_dependency_graph(objects)
        
        # 3. 取得渲染順序
        render_order = self.dependency_analyzer.get_render_order(dependency_graph)
        
        # 4. 計算實際數據形狀
        actual_data_shapes = self._calculate_actual_data_shapes(objects, data_context)
        
        # 5. 依順序計算每個物件的新位置
        predicted_positions = {}
        position_adjustments = {}  # 記錄位置調整
        
        for obj_id in render_order:
            obj = self._find_object_by_id(objects, obj_id)
            if not obj:
                continue
            
            initial_pos = initial_positions.get(obj_id)
            if not initial_pos:
                continue
            
            # 計算此物件受到的連鎖位移影響
            cumulative_displacement = self._calculate_cumulative_displacement(
                obj_id, position_adjustments, dependency_graph
            )
            
            # 應用位移到初始位置
            adjusted_initial_pos = self._apply_displacement_to_position(
                initial_pos, cumulative_displacement
            )
            
            # 根據實際數據形狀計算最終位置
            actual_shape = actual_data_shapes.get(obj_id, initial_pos['data_shape'])
            final_position = self.position_utils.calculate_range(
                (adjusted_initial_pos['start_row'], adjusted_initial_pos['start_col']),
                actual_shape
            )
            
            predicted_positions[obj_id] = {
                **final_position,
                'obj_type': obj.get('obj_type', 'unknown'),
                'data_shape': actual_shape,
                'initial_position': initial_pos,
                'cumulative_displacement': cumulative_displacement,
                'size_change': self._calculate_size_change(initial_pos['data_shape'], actual_shape)
            }
            
            # 如果物件大小有變化，記錄影響
            size_change = predicted_positions[obj_id]['size_change']
            if size_change['rows'] != 0 or size_change['cols'] != 0:
                position_adjustments[obj_id] = {
                    'expansion': size_change,
                    'affected_range': final_position
                }
        
        return predicted_positions
    
    def calculate_tag_relocation_plan(self, objects: List[Dict[str, Any]], 
                                    predictions: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        計算標籤重定位計劃
        
        Args:
            objects: 物件清單
            predictions: 預測位置結果
            
        Returns:
            list: 重定位計劃清單
        """
        relocation_plan = []
        
        for obj in objects:
            obj_id = obj.get('obj_id')
            if not obj_id or obj_id not in predictions:
                continue
            
            prediction = predictions[obj_id]
            initial_pos = prediction.get('initial_position', {})
            
            # 檢查是否需要重定位
            needs_relocation = self._needs_relocation(
                initial_pos, prediction, obj
            )
            
            if needs_relocation:
                # 計算重定位詳情
                from_coordinate = initial_pos.get('coordinate', 
                    self.position_utils.position_to_excel(
                        initial_pos.get('row', 1), 
                        initial_pos.get('col', 1)
                    )
                )
                
                to_coordinate = self.position_utils.position_to_excel(
                    prediction['start_row'], 
                    prediction['start_col']
                )
                
                relocation_item = {
                    'obj_id': obj_id,
                    'obj_name': obj.get('obj_name'),
                    'obj_type': obj.get('obj_type'),
                    'from_coordinate': from_coordinate,
                    'to_coordinate': to_coordinate,
                    'from_range': initial_pos.get('range_string', from_coordinate),
                    'to_range': prediction['range_string'],
                    'displacement': prediction.get('cumulative_displacement', {'rows': 0, 'cols': 0}),
                    'priority': self._calculate_relocation_priority(obj, prediction),
                    'reason': self._get_relocation_reason(prediction)
                }
                
                relocation_plan.append(relocation_item)
        
        # 按優先級排序
        relocation_plan.sort(key=lambda x: x['priority'])
        
        return relocation_plan
    
    def validate_positions(self, positions: Dict[str, Dict[str, Any]]) -> Tuple[bool, List[str]]:
        """
        驗證位置計算結果，檢查衝突
        
        Args:
            positions: 位置計算結果
            
        Returns:
            tuple: (is_valid, error_messages)
        """
        errors = []
        
        # 檢查位置重疊
        position_list = list(positions.items())
        for i, (obj_id1, pos1) in enumerate(position_list):
            for obj_id2, pos2 in position_list[i + 1:]:
                if self.position_utils.is_position_overlap(pos1, pos2):
                    errors.append(
                        f"物件 {obj_id1} 和 {obj_id2} 位置重疊: "
                        f"{pos1.get('range_string')} vs {pos2.get('range_string')}"
                    )
        
        # 檢查位置邊界
        for obj_id, pos in positions.items():
            if pos.get('start_row', 1) < 1 or pos.get('start_col', 1) < 1:
                errors.append(f"物件 {obj_id} 位置超出邊界: {pos.get('range_string')}")
            
            # 檢查Excel行列限制 (Excel 2016+: 1048576行, 16384列)
            max_row, max_col = 1048576, 16384
            if (pos.get('end_row', 1) > max_row or pos.get('end_col', 1) > max_col):
                errors.append(
                    f"物件 {obj_id} 超出Excel限制: {pos.get('range_string')} "
                    f"(最大: {self.position_utils.position_to_excel(max_row, max_col)})"
                )
        
        return len(errors) == 0, errors
    
    def _extract_initial_position(self, obj: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """從物件中提取初始位置資訊"""
        # 嘗試多種可能的位置欄位
        position_fields = ['position_before', 'cell_position', 'position']
        
        for field in position_fields:
            if field in obj:
                pos_data = obj[field]
                
                if isinstance(pos_data, dict):
                    if 'row' in pos_data and 'col' in pos_data:
                        return {
                            'row': pos_data['row'],
                            'col': pos_data['col'],
                            'coordinate': self.position_utils.position_to_excel(
                                pos_data['row'], pos_data['col']
                            )
                        }
                elif hasattr(pos_data, 'row') and hasattr(pos_data, 'col'):
                    return {
                        'row': pos_data.row,
                        'col': pos_data.col,
                        'coordinate': self.position_utils.position_to_excel(
                            pos_data.row, pos_data.col
                        )
                    }
        
        return None
    
    def _get_object_data_shape(self, obj: Dict[str, Any]) -> Dict[str, int]:
        """取得物件的數據形狀"""
        # 檢查是否有明確的數據形狀資訊
        if 'data_shape' in obj and isinstance(obj['data_shape'], dict):
            shape = obj['data_shape']
            if 'rows' in shape and 'cols' in shape:
                return {'rows': shape['rows'], 'cols': shape['cols']}
        
        # 根據物件類型推測預設形狀
        obj_type = obj.get('obj_type', 'simple')
        
        if obj_type == 'table':
            # 表格物件預設為2x5 (包含標題)
            return {'rows': 2, 'cols': 5}
        elif obj_type == 'image':
            # 圖片物件預設為5x5
            return {'rows': 5, 'cols': 5}
        else:
            # 簡單物件預設為1x1
            return {'rows': 1, 'cols': 1}
    
    def _calculate_actual_data_shapes(self, objects: List[Dict[str, Any]], 
                                    data_context: Dict[str, Any]) -> Dict[str, Dict[str, int]]:
        """根據實際數據計算物件的實際形狀"""
        actual_shapes = {}
        
        for obj in objects:
            obj_id = obj.get('obj_id')
            obj_name = obj.get('obj_name')
            obj_type = obj.get('obj_type', 'simple')
            
            if not obj_id or not obj_name:
                continue
            
            # 從數據上下文中查找對應的數據
            actual_data = data_context.get(obj_name)
            
            if actual_data is not None:
                if obj_type == 'table' and hasattr(actual_data, 'shape'):
                    # pandas DataFrame
                    rows, cols = actual_data.shape
                    # 表格物件需要加上標題行
                    if obj.get('having_header', True):
                        rows += 1
                    actual_shapes[obj_id] = {'rows': rows, 'cols': cols}
                    
                elif obj_type == 'table' and isinstance(actual_data, list):
                    # 清單數據
                    rows = len(actual_data)
                    cols = len(actual_data[0]) if actual_data and isinstance(actual_data[0], (list, tuple)) else 1
                    if obj.get('having_header', True):
                        rows += 1
                    actual_shapes[obj_id] = {'rows': rows, 'cols': cols}
                    
                else:
                    # 簡單數據，保持原始形狀
                    actual_shapes[obj_id] = self._get_object_data_shape(obj)
            else:
                # 沒有找到對應數據，使用原始形狀
                actual_shapes[obj_id] = self._get_object_data_shape(obj)
        
        return actual_shapes
    
    def _find_object_by_id(self, objects: List[Dict[str, Any]], obj_id: str) -> Optional[Dict[str, Any]]:
        """根據ID查找物件"""
        for obj in objects:
            if obj.get('obj_id') == obj_id:
                return obj
        return None
    
    def _calculate_cumulative_displacement(self, obj_id: str, 
                                         position_adjustments: Dict[str, Dict[str, Any]],
                                         dependency_graph: Dict[str, Any]) -> Dict[str, int]:
        """計算累積位移效應"""
        total_displacement = {'rows': 0, 'cols': 0}
        
        # 找出所有影響此物件的物件
        dependencies = dependency_graph['reverse_adjacency_list'].get(obj_id, [])
        
        for dep_id in dependencies:
            if dep_id in position_adjustments:
                # 取得影響資訊
                adjustment = position_adjustments[dep_id]
                expansion = adjustment['expansion']
                
                # 找出對應的邊資訊以確定影響類型
                edge_info = self.dependency_analyzer._find_edge(dependency_graph, dep_id, obj_id)
                
                if edge_info:
                    # 根據關係類型計算傳播的位移
                    propagated = self.dependency_analyzer._calculate_propagated_displacement(
                        expansion, edge_info
                    )
                    
                    total_displacement['rows'] += propagated['rows']
                    total_displacement['cols'] += propagated['cols']
        
        return total_displacement
    
    def _apply_displacement_to_position(self, position: Dict[str, Any], 
                                      displacement: Dict[str, int]) -> Dict[str, Any]:
        """對位置應用位移"""
        if not displacement or (displacement['rows'] == 0 and displacement['cols'] == 0):
            return position.copy()
        
        return self.position_utils.apply_displacement(position, displacement)
    
    def _calculate_size_change(self, original_shape: Dict[str, int], 
                             new_shape: Dict[str, int]) -> Dict[str, int]:
        """計算大小變化"""
        return {
            'rows': new_shape['rows'] - original_shape['rows'],
            'cols': new_shape['cols'] - original_shape['cols']
        }
    
    def _needs_relocation(self, initial_pos: Dict[str, Any], 
                         prediction: Dict[str, Any], 
                         obj: Dict[str, Any]) -> bool:
        """判斷是否需要重定位"""
        # 如果位置沒有變化，不需要重定位
        if (initial_pos.get('row') == prediction.get('start_row') and 
            initial_pos.get('col') == prediction.get('start_col')):
            return False
        
        # 如果有累積位移，需要重定位
        displacement = prediction.get('cumulative_displacement', {'rows': 0, 'cols': 0})
        if displacement['rows'] != 0 or displacement['cols'] != 0:
            return True
        
        # 如果大小有變化且會影響位置，需要重定位
        size_change = prediction.get('size_change', {'rows': 0, 'cols': 0})
        if size_change['rows'] != 0 or size_change['cols'] != 0:
            return True
        
        return False
    
    def _calculate_relocation_priority(self, obj: Dict[str, Any], 
                                     prediction: Dict[str, Any]) -> int:
        """計算重定位優先級"""
        # 優先級越小越先執行
        base_priority = 100
        
        # 根據物件類型調整優先級
        obj_type = obj.get('obj_type', 'simple')
        if obj_type == 'simple':
            base_priority = 10  # 簡單物件優先
        elif obj_type == 'table':
            base_priority = 50  # 表格物件次之
        elif obj_type == 'image':
            base_priority = 90  # 圖片物件最後
        
        # 根據位移量調整優先級
        displacement = prediction.get('cumulative_displacement', {'rows': 0, 'cols': 0})
        displacement_penalty = abs(displacement['rows']) + abs(displacement['cols'])
        
        return base_priority + displacement_penalty
    
    def _get_relocation_reason(self, prediction: Dict[str, Any]) -> str:
        """取得重定位原因"""
        displacement = prediction.get('cumulative_displacement', {'rows': 0, 'cols': 0})
        size_change = prediction.get('size_change', {'rows': 0, 'cols': 0})
        
        reasons = []
        
        if displacement['rows'] > 0:
            reasons.append(f"向下位移{displacement['rows']}行")
        elif displacement['rows'] < 0:
            reasons.append(f"向上位移{abs(displacement['rows'])}行")
        
        if displacement['cols'] > 0:
            reasons.append(f"向右位移{displacement['cols']}列")
        elif displacement['cols'] < 0:
            reasons.append(f"向左位移{abs(displacement['cols'])}列")
        
        if size_change['rows'] != 0 or size_change['cols'] != 0:
            reasons.append(f"大小變化({size_change['rows']:+d}行, {size_change['cols']:+d}列)")
        
        return "; ".join(reasons) if reasons else "位置調整"
    
    def get_position_summary(self, positions: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
        """取得位置計算摘要資訊"""
        if not positions:
            return {
                'total_objects': 0,
                'relocated_objects': 0,
                'size_changed_objects': 0,
                'total_displacement': {'rows': 0, 'cols': 0}
            }
        
        relocated_count = 0
        size_changed_count = 0
        total_displacement = {'rows': 0, 'cols': 0}
        
        for obj_id, pos in positions.items():
            # 計算重定位物件
            displacement = pos.get('cumulative_displacement', {'rows': 0, 'cols': 0})
            if displacement['rows'] != 0 or displacement['cols'] != 0:
                relocated_count += 1
                total_displacement['rows'] += abs(displacement['rows'])
                total_displacement['cols'] += abs(displacement['cols'])
            
            # 計算大小變化物件
            size_change = pos.get('size_change', {'rows': 0, 'cols': 0})
            if size_change['rows'] != 0 or size_change['cols'] != 0:
                size_changed_count += 1
        
        return {
            'total_objects': len(positions),
            'relocated_objects': relocated_count,
            'size_changed_objects': size_changed_count,
            'total_displacement': total_displacement
        }