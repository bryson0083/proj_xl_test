#!/usr/bin/env python3
"""
依賴分析工具模組

分析物件間的位置依賴關係，計算連鎖位移效應，並提供拓撲排序功能
"""
from typing import Dict, List, Any, Set, Tuple, Optional
from collections import defaultdict, deque
from .position_utils import PositionUtils


class DependencyAnalyzer:
    """物件依賴關係分析器"""
    
    def __init__(self):
        self.position_utils = PositionUtils()
    
    def analyze_position_dependencies(self, objects: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        分析物件間的位置依賴關係
        
        Args:
            objects: 物件清單，每個物件包含位置資訊
            
        Returns:
            dict: 依賴分析結果
            {
                'dependencies': dict,  # 每個物件的依賴清單
                'affects': dict,       # 每個物件影響的物件清單
                'vertical_groups': list,  # 垂直相關的物件群組
                'horizontal_groups': list # 水平相關的物件群組
            }
        """
        if not objects:
            return {
                'dependencies': {},
                'affects': {},
                'vertical_groups': [],
                'horizontal_groups': []
            }
        
        dependencies = defaultdict(list)
        affects = defaultdict(list)
        
        # 為每個物件對分析關係
        for i, obj1 in enumerate(objects):
            for j, obj2 in enumerate(objects):
                if i == j:
                    continue
                
                relationship = self._analyze_object_relationship(obj1, obj2)
                
                if relationship['has_dependency']:
                    obj1_id = obj1.get('obj_id', f"obj_{i}")
                    obj2_id = obj2.get('obj_id', f"obj_{j}")
                    
                    # obj1 依賴於 obj2（obj2 的變化會影響 obj1）
                    dependencies[obj1_id].append({
                        'depends_on': obj2_id,
                        'relationship_type': relationship['type'],
                        'priority': relationship['priority'],
                        'displacement_factor': relationship['displacement_factor']
                    })
                    
                    # obj2 影響 obj1
                    affects[obj2_id].append({
                        'affects': obj1_id,
                        'relationship_type': relationship['type'],
                        'priority': relationship['priority'],
                        'displacement_factor': relationship['displacement_factor']
                    })
        
        # 分析物件群組
        vertical_groups = self._find_vertical_groups(objects)
        horizontal_groups = self._find_horizontal_groups(objects)
        
        return {
            'dependencies': dict(dependencies),
            'affects': dict(affects),
            'vertical_groups': vertical_groups,
            'horizontal_groups': horizontal_groups
        }
    
    def _analyze_object_relationship(self, obj1: Dict[str, Any], obj2: Dict[str, Any]) -> Dict[str, Any]:
        """
        分析兩個物件之間的關係
        
        Args:
            obj1: 第一個物件
            obj2: 第二個物件
            
        Returns:
            dict: 關係分析結果
        """
        # 取得物件位置資訊
        pos1 = self._get_object_position(obj1)
        pos2 = self._get_object_position(obj2)
        
        if not pos1 or not pos2:
            return {'has_dependency': False}
        
        # 分析垂直關係
        vertical_rel = self._analyze_vertical_relationship(pos1, pos2)
        if vertical_rel['has_relationship']:
            return {
                'has_dependency': True,
                'type': 'vertical',
                'direction': vertical_rel['direction'],
                'priority': vertical_rel['priority'],
                'displacement_factor': vertical_rel['displacement_factor']
            }
        
        # 分析水平關係
        horizontal_rel = self._analyze_horizontal_relationship(pos1, pos2)
        if horizontal_rel['has_relationship']:
            return {
                'has_dependency': True,
                'type': 'horizontal',
                'direction': horizontal_rel['direction'],
                'priority': horizontal_rel['priority'],
                'displacement_factor': horizontal_rel['displacement_factor']
            }
        
        return {'has_dependency': False}
    
    def _get_object_position(self, obj: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """從物件中提取位置資訊"""
        # 嘗試多種可能的位置資訊欄位
        position_fields = ['position_before', 'cell_position', 'position']
        
        for field in position_fields:
            if field in obj:
                pos_data = obj[field]
                if isinstance(pos_data, dict):
                    return pos_data
                elif hasattr(pos_data, 'row') and hasattr(pos_data, 'col'):
                    return {
                        'start_row': pos_data.row,
                        'start_col': pos_data.col,
                        'end_row': pos_data.row,
                        'end_col': pos_data.col
                    }
        
        return None
    
    def _analyze_vertical_relationship(self, pos1: Dict[str, Any], pos2: Dict[str, Any]) -> Dict[str, Any]:
        """分析垂直關係"""
        # 檢查欄位重疊
        col_overlap = not (pos1.get('end_col', pos1.get('start_col')) < pos2.get('start_col') or 
                          pos2.get('end_col', pos2.get('start_col')) < pos1.get('start_col'))
        
        if not col_overlap:
            return {'has_relationship': False}
        
        pos1_row = pos1.get('start_row')
        pos2_row = pos2.get('start_row')
        pos2_end_row = pos2.get('end_row', pos2_row)
        
        # pos1 在 pos2 下方
        if pos1_row > pos2_end_row:
            return {
                'has_relationship': True,
                'direction': 'below',
                'priority': 1,  # 高優先級，直接影響
                'displacement_factor': 1.0
            }
        
        return {'has_relationship': False}
    
    def _analyze_horizontal_relationship(self, pos1: Dict[str, Any], pos2: Dict[str, Any]) -> Dict[str, Any]:
        """分析水平關係"""
        # 檢查行重疊
        row_overlap = not (pos1.get('end_row', pos1.get('start_row')) < pos2.get('start_row') or 
                          pos2.get('end_row', pos2.get('start_row')) < pos1.get('start_row'))
        
        if not row_overlap:
            return {'has_relationship': False}
        
        pos1_col = pos1.get('start_col')
        pos2_col = pos2.get('start_col')
        pos2_end_col = pos2.get('end_col', pos2_col)
        
        # pos1 在 pos2 右方
        if pos1_col > pos2_end_col:
            return {
                'has_relationship': True,
                'direction': 'right',
                'priority': 2,  # 較低優先級
                'displacement_factor': 0.8
            }
        
        return {'has_relationship': False}
    
    def _find_vertical_groups(self, objects: List[Dict[str, Any]]) -> List[List[str]]:
        """找出垂直相關的物件群組"""
        groups = []
        processed = set()
        
        for obj in objects:
            obj_id = obj.get('obj_id', obj.get('obj_name', str(id(obj))))
            if obj_id in processed:
                continue
            
            # 找出同一垂直軸上的物件
            group = [obj_id]
            obj_pos = self._get_object_position(obj)
            
            if not obj_pos:
                continue
            
            for other_obj in objects:
                other_id = other_obj.get('obj_id', other_obj.get('obj_name', str(id(other_obj))))
                if other_id == obj_id or other_id in processed:
                    continue
                
                other_pos = self._get_object_position(other_obj)
                if not other_pos:
                    continue
                
                # 檢查是否在同一垂直軸
                col_overlap = not (obj_pos.get('end_col', obj_pos.get('start_col')) < other_pos.get('start_col') or 
                                 other_pos.get('end_col', other_pos.get('start_col')) < obj_pos.get('start_col'))
                
                if col_overlap:
                    group.append(other_id)
            
            if len(group) > 1:
                groups.append(group)
                processed.update(group)
        
        return groups
    
    def _find_horizontal_groups(self, objects: List[Dict[str, Any]]) -> List[List[str]]:
        """找出水平相關的物件群組"""
        groups = []
        processed = set()
        
        for obj in objects:
            obj_id = obj.get('obj_id', obj.get('obj_name', str(id(obj))))
            if obj_id in processed:
                continue
            
            # 找出同一水平軸上的物件
            group = [obj_id]
            obj_pos = self._get_object_position(obj)
            
            if not obj_pos:
                continue
            
            for other_obj in objects:
                other_id = other_obj.get('obj_id', other_obj.get('obj_name', str(id(other_obj))))
                if other_id == obj_id or other_id in processed:
                    continue
                
                other_pos = self._get_object_position(other_obj)
                if not other_pos:
                    continue
                
                # 檢查是否在同一水平軸
                row_overlap = not (obj_pos.get('end_row', obj_pos.get('start_row')) < other_pos.get('start_row') or 
                                 other_pos.get('end_row', other_pos.get('start_row')) < obj_pos.get('start_row'))
                
                if row_overlap:
                    group.append(other_id)
            
            if len(group) > 1:
                groups.append(group)
                processed.update(group)
        
        return groups
    
    def build_dependency_graph(self, objects: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        建立依賴關係圖
        
        Args:
            objects: 物件清單
            
        Returns:
            dict: 依賴關係圖
        """
        analysis = self.analyze_position_dependencies(objects)
        
        # 建立圖結構
        graph = {
            'nodes': {},
            'edges': [],
            'adjacency_list': defaultdict(list),
            'reverse_adjacency_list': defaultdict(list)
        }
        
        # 建立節點
        for obj in objects:
            obj_id = obj.get('obj_id', obj.get('obj_name', str(id(obj))))
            graph['nodes'][obj_id] = {
                'obj_id': obj_id,
                'obj_type': obj.get('obj_type', 'unknown'),
                'position': self._get_object_position(obj),
                'data_shape': obj.get('data_shape', {'rows': 1, 'cols': 1})
            }
        
        # 建立邊
        for obj_id, deps in analysis['dependencies'].items():
            for dep in deps:
                depends_on = dep['depends_on']
                
                edge = {
                    'from': depends_on,
                    'to': obj_id,
                    'type': dep['relationship_type'],
                    'priority': dep['priority'],
                    'displacement_factor': dep['displacement_factor']
                }
                
                graph['edges'].append(edge)
                graph['adjacency_list'][depends_on].append(obj_id)
                graph['reverse_adjacency_list'][obj_id].append(depends_on)
        
        return graph
    
    def calculate_displacement_chain(self, trigger_obj: Dict[str, Any], 
                                   expansion_info: Dict[str, Any],
                                   dependency_graph: Dict[str, Any]) -> Dict[str, Any]:
        """
        計算連鎖位移效應
        
        Args:
            trigger_obj: 觸發變化的物件
            expansion_info: 擴展資訊 {'rows': int, 'cols': int}
            dependency_graph: 依賴關係圖
            
        Returns:
            dict: 連鎖位移計算結果
        """
        trigger_id = trigger_obj.get('obj_id', trigger_obj.get('obj_name'))
        
        if trigger_id not in dependency_graph['adjacency_list']:
            return {'displacements': {}}
        
        displacements = {}
        visited = set()
        
        # 使用BFS計算連鎖位移
        queue = deque([{
            'obj_id': trigger_id,
            'displacement': expansion_info,
            'level': 0
        }])
        
        while queue:
            current = queue.popleft()
            obj_id = current['obj_id']
            displacement = current['displacement']
            level = current['level']
            
            if obj_id in visited:
                continue
            
            visited.add(obj_id)
            
            if obj_id != trigger_id:
                displacements[obj_id] = {
                    'displacement': displacement,
                    'level': level,
                    'reason': f"受 {trigger_id} 擴展影響"
                }
            
            # 處理受影響的物件
            for affected_id in dependency_graph['adjacency_list'].get(obj_id, []):
                if affected_id in visited:
                    continue
                
                # 找出對應的邊資訊
                edge_info = self._find_edge(dependency_graph, obj_id, affected_id)
                if not edge_info:
                    continue
                
                # 計算傳播的位移
                propagated_displacement = self._calculate_propagated_displacement(
                    displacement, edge_info
                )
                
                queue.append({
                    'obj_id': affected_id,
                    'displacement': propagated_displacement,
                    'level': level + 1
                })
        
        return {'displacements': displacements}
    
    def _find_edge(self, graph: Dict[str, Any], from_id: str, to_id: str) -> Optional[Dict[str, Any]]:
        """在圖中找出指定的邊"""
        for edge in graph['edges']:
            if edge['from'] == from_id and edge['to'] == to_id:
                return edge
        return None
    
    def _calculate_propagated_displacement(self, original_displacement: Dict[str, int], 
                                         edge_info: Dict[str, Any]) -> Dict[str, int]:
        """計算傳播的位移"""
        factor = edge_info.get('displacement_factor', 1.0)
        relationship_type = edge_info.get('type', 'vertical')
        
        if relationship_type == 'vertical':
            # 垂直關係：只傳播行位移
            return {
                'rows': int(original_displacement.get('rows', 0) * factor),
                'cols': 0
            }
        elif relationship_type == 'horizontal':
            # 水平關係：只傳播列位移
            return {
                'rows': 0,
                'cols': int(original_displacement.get('cols', 0) * factor)
            }
        else:
            # 其他關係：完全傳播
            return {
                'rows': int(original_displacement.get('rows', 0) * factor),
                'cols': int(original_displacement.get('cols', 0) * factor)
            }
    
    def get_render_order(self, dependency_graph: Dict[str, Any]) -> List[str]:
        """
        使用拓撲排序決定渲染順序
        
        Args:
            dependency_graph: 依賴關係圖
            
        Returns:
            list: 渲染順序的物件ID清單
            
        Raises:
            ValueError: 當存在循環依賴時拋出
        """
        # 計算每個節點的入度
        in_degree = defaultdict(int)
        nodes = set(dependency_graph['nodes'].keys())
        
        # 初始化所有節點的入度為0
        for node_id in nodes:
            in_degree[node_id] = 0
        
        # 計算實際入度
        for edge in dependency_graph['edges']:
            in_degree[edge['to']] += 1
        
        # 找出入度為0的節點
        queue = deque([node_id for node_id in nodes if in_degree[node_id] == 0])
        result = []
        
        while queue:
            current = queue.popleft()
            result.append(current)
            
            # 處理當前節點的所有出邊
            for affected_id in dependency_graph['adjacency_list'].get(current, []):
                in_degree[affected_id] -= 1
                if in_degree[affected_id] == 0:
                    queue.append(affected_id)
        
        # 檢查是否存在循環依賴
        if len(result) != len(nodes):
            remaining_nodes = nodes - set(result)
            raise ValueError(f"檢測到循環依賴，涉及節點: {remaining_nodes}")
        
        return result
    
    def detect_circular_dependencies(self, dependency_graph: Dict[str, Any]) -> List[List[str]]:
        """
        檢測循環依賴
        
        Args:
            dependency_graph: 依賴關係圖
            
        Returns:
            list: 循環依賴的路徑清單
        """
        cycles = []
        visited = set()
        rec_stack = set()
        
        def dfs(node_id: str, path: List[str]):
            if node_id in rec_stack:
                # 找到循環
                cycle_start = path.index(node_id)
                cycles.append(path[cycle_start:] + [node_id])
                return
            
            if node_id in visited:
                return
            
            visited.add(node_id)
            rec_stack.add(node_id)
            path.append(node_id)
            
            for neighbor in dependency_graph['adjacency_list'].get(node_id, []):
                dfs(neighbor, path.copy())
            
            rec_stack.remove(node_id)
        
        for node_id in dependency_graph['nodes']:
            if node_id not in visited:
                dfs(node_id, [])
        
        return cycles