#!/usr/bin/env python3
"""
註冊表工具模組

提供物件ID生成、註冊表序列化、驗證和比較等功能
"""
import json
import hashlib
import os
from datetime import datetime
from typing import Dict, Any, Tuple, List, Optional
from pathlib import Path


class RegistryUtils:
    """註冊表工具類別"""
    
    @staticmethod
    def generate_object_id(sheet_name: str, position: str, obj_name: str) -> str:
        """
        生成可重現的物件ID
        
        使用 hashlib 基於工作表名稱、位置、物件名稱生成唯一且可重現的ID
        
        Args:
            sheet_name: 工作表名稱
            position: 物件位置（如 'B5' 或 'B5:F10'）
            obj_name: 物件名稱
            
        Returns:
            str: 物件ID (格式: obj_xxxxxx...)
            
        Raises:
            ValueError: 當任何參數為空時拋出
        """
        if not all([sheet_name, position, obj_name]):
            raise ValueError("工作表名稱、位置和物件名稱都不能為空")
        
        # 建立用於hash的字串
        id_string = f"{sheet_name}|{position}|{obj_name}"
        
        # 使用 SHA-256 生成 hash
        hash_obj = hashlib.sha256(id_string.encode('utf-8'))
        hash_hex = hash_obj.hexdigest()
        
        # 取前12個字符作為ID
        return f"obj_{hash_hex[:12]}"
    
    @staticmethod
    def serialize_registry(registry: Dict[str, Any], output_dir: Optional[str] = None) -> str:
        """
        序列化註冊表為JSON
        
        檔案命名規範：template_renderer_obj_registry_{YYYYMMDD}_{HHMMSS}.json
        
        Args:
            registry: 物件註冊表字典
            output_dir: 輸出目錄，預設為當前工作目錄
            
        Returns:
            str: 生成的JSON檔案完整路徑
            
        Raises:
            ValueError: 當註冊表為空或格式錯誤時拋出
            IOError: 當無法寫入檔案時拋出
        """
        if not registry or not isinstance(registry, dict):
            raise ValueError("註冊表不能為空且必須是字典格式")
        
        # 產生時間戳記
        timestamp = datetime.now()
        date_str = timestamp.strftime("%Y%m%d")
        time_str = timestamp.strftime("%H%M%S")
        
        # 生成檔案名稱
        filename = f"template_renderer_obj_registry_{date_str}_{time_str}.json"
        
        # 決定輸出路徑
        if output_dir is None:
            output_dir = os.getcwd()
        
        output_path = Path(output_dir) / filename
        
        try:
            # 確保輸出目錄存在
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # 在註冊表中加入元數據
            enhanced_registry = {
                **registry,
                'metadata': {
                    'registry_version': '1.0',
                    'created_at': timestamp.isoformat(),
                    'registry_filename': filename,
                    'generated_by': 'ExcelTemplateRenderer'
                }
            }
            
            # 寫入JSON檔案
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(enhanced_registry, f, ensure_ascii=False, indent=2)
            
            return str(output_path)
            
        except Exception as e:
            raise IOError(f"無法寫入註冊表檔案: {output_path}, 錯誤: {str(e)}")
    
    @staticmethod
    def load_registry(json_path: str) -> Dict[str, Any]:
        """
        從JSON檔案載入註冊表
        
        Args:
            json_path: JSON檔案路徑
            
        Returns:
            dict: 載入的註冊表
            
        Raises:
            FileNotFoundError: 當檔案不存在時拋出
            ValueError: 當JSON格式錯誤時拋出
        """
        if not json_path or not isinstance(json_path, str):
            raise ValueError("JSON檔案路徑不能為空")
        
        json_path = Path(json_path)
        
        if not json_path.exists():
            raise FileNotFoundError(f"註冊表檔案不存在: {json_path}")
        
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                registry = json.load(f)
            
            if not isinstance(registry, dict):
                raise ValueError("註冊表格式錯誤：必須是字典格式")
            
            return registry
            
        except json.JSONDecodeError as e:
            raise ValueError(f"JSON格式錯誤: {str(e)}")
        except Exception as e:
            raise IOError(f"無法讀取註冊表檔案: {json_path}, 錯誤: {str(e)}")
    
    @staticmethod
    def validate_registry(registry: Dict[str, Any]) -> Tuple[bool, List[str]]:
        """
        驗證註冊表完整性
        
        Args:
            registry: 註冊表字典
            
        Returns:
            tuple: (is_valid, errors) 驗證結果和錯誤清單
        """
        if not isinstance(registry, dict):
            return False, ["註冊表必須是字典格式"]
        
        errors = []
        
        # 檢查必要的頂層欄位
        required_top_level_fields = ['worksheets']
        for field in required_top_level_fields:
            if field not in registry:
                errors.append(f"缺少必要欄位: {field}")
        
        if 'worksheets' not in registry:
            return False, errors
        
        # 檢查工作表結構
        worksheets = registry['worksheets']
        if not isinstance(worksheets, dict):
            errors.append("worksheets 必須是字典格式")
            return False, errors
        
        for sheet_name, sheet_data in worksheets.items():
            sheet_errors = RegistryUtils._validate_worksheet(sheet_name, sheet_data)
            errors.extend(sheet_errors)
        
        # 檢查物件ID的唯一性
        all_obj_ids = []
        for sheet_data in worksheets.values():
            if 'objects' in sheet_data and isinstance(sheet_data['objects'], list):
                for obj in sheet_data['objects']:
                    if 'obj_id' in obj:
                        all_obj_ids.append(obj['obj_id'])
        
        # 檢查重複ID
        seen_ids = set()
        for obj_id in all_obj_ids:
            if obj_id in seen_ids:
                errors.append(f"發現重複的物件ID: {obj_id}")
            seen_ids.add(obj_id)
        
        return len(errors) == 0, errors
    
    @staticmethod
    def _validate_worksheet(sheet_name: str, sheet_data: Dict[str, Any]) -> List[str]:
        """驗證單一工作表的數據結構"""
        errors = []
        
        if not isinstance(sheet_data, dict):
            errors.append(f"工作表 '{sheet_name}' 的數據必須是字典格式")
            return errors
        
        # 檢查必要欄位
        required_fields = ['objects']
        for field in required_fields:
            if field not in sheet_data:
                errors.append(f"工作表 '{sheet_name}' 缺少必要欄位: {field}")
        
        # 檢查objects結構
        if 'objects' in sheet_data:
            objects = sheet_data['objects']
            if not isinstance(objects, list):
                errors.append(f"工作表 '{sheet_name}' 的 objects 必須是清單格式")
            else:
                for i, obj in enumerate(objects):
                    obj_errors = RegistryUtils._validate_object(sheet_name, i, obj)
                    errors.extend(obj_errors)
        
        return errors
    
    @staticmethod
    def _validate_object(sheet_name: str, obj_index: int, obj_data: Dict[str, Any]) -> List[str]:
        """驗證單一物件的數據結構"""
        errors = []
        
        if not isinstance(obj_data, dict):
            errors.append(f"工作表 '{sheet_name}' 物件 {obj_index} 必須是字典格式")
            return errors
        
        # 檢查必要欄位
        required_fields = ['obj_id', 'obj_name', 'obj_type']
        for field in required_fields:
            if field not in obj_data:
                errors.append(f"工作表 '{sheet_name}' 物件 {obj_index} 缺少必要欄位: {field}")
        
        # 檢查位置資訊
        position_fields = ['position_before', 'position_after']
        for pos_field in position_fields:
            if pos_field in obj_data:
                pos_errors = RegistryUtils._validate_position(sheet_name, obj_index, pos_field, obj_data[pos_field])
                errors.extend(pos_errors)
        
        return errors
    
    @staticmethod
    def _validate_position(sheet_name: str, obj_index: int, pos_field: str, pos_data: Dict[str, Any]) -> List[str]:
        """驗證位置資訊結構"""
        errors = []
        
        if not isinstance(pos_data, dict):
            errors.append(f"工作表 '{sheet_name}' 物件 {obj_index} 的 {pos_field} 必須是字典格式")
            return errors
        
        # 檢查必要的位置欄位
        required_pos_fields = ['row', 'col']
        for field in required_pos_fields:
            if field not in pos_data:
                errors.append(f"工作表 '{sheet_name}' 物件 {obj_index} 的 {pos_field} 缺少必要欄位: {field}")
            elif not isinstance(pos_data[field], int) or pos_data[field] < 1:
                errors.append(f"工作表 '{sheet_name}' 物件 {obj_index} 的 {pos_field}.{field} 必須是大於0的整數")
        
        return errors
    
    @staticmethod
    def diff_registries(before: Dict[str, Any], after: Dict[str, Any]) -> Dict[str, Any]:
        """
        比較兩個註冊表的差異
        
        Args:
            before: 之前的註冊表
            after: 之後的註冊表
            
        Returns:
            dict: 差異報告
        """
        diff_result = {
            'summary': {
                'total_changes': 0,
                'added_objects': 0,
                'removed_objects': 0,
                'modified_objects': 0,
                'position_changes': 0
            },
            'changes': {
                'added': {},
                'removed': {},
                'modified': {}
            }
        }
        
        # 取得所有物件ID
        before_objects = RegistryUtils._extract_all_objects(before)
        after_objects = RegistryUtils._extract_all_objects(after)
        
        before_ids = set(before_objects.keys())
        after_ids = set(after_objects.keys())
        
        # 找出新增的物件
        added_ids = after_ids - before_ids
        for obj_id in added_ids:
            diff_result['changes']['added'][obj_id] = after_objects[obj_id]
            diff_result['summary']['added_objects'] += 1
        
        # 找出移除的物件
        removed_ids = before_ids - after_ids
        for obj_id in removed_ids:
            diff_result['changes']['removed'][obj_id] = before_objects[obj_id]
            diff_result['summary']['removed_objects'] += 1
        
        # 找出修改的物件
        common_ids = before_ids & after_ids
        for obj_id in common_ids:
            obj_diff = RegistryUtils._compare_objects(before_objects[obj_id], after_objects[obj_id])
            if obj_diff['has_changes']:
                diff_result['changes']['modified'][obj_id] = obj_diff
                diff_result['summary']['modified_objects'] += 1
                if obj_diff['position_changed']:
                    diff_result['summary']['position_changes'] += 1
        
        # 計算總變更數
        diff_result['summary']['total_changes'] = (
            diff_result['summary']['added_objects'] +
            diff_result['summary']['removed_objects'] +
            diff_result['summary']['modified_objects']
        )
        
        return diff_result
    
    @staticmethod
    def _extract_all_objects(registry: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
        """從註冊表中提取所有物件"""
        all_objects = {}
        
        if 'worksheets' not in registry:
            return all_objects
        
        for sheet_name, sheet_data in registry['worksheets'].items():
            if 'objects' not in sheet_data or not isinstance(sheet_data['objects'], list):
                continue
            
            for obj in sheet_data['objects']:
                if 'obj_id' in obj:
                    # 加入工作表資訊
                    obj_copy = obj.copy()
                    obj_copy['_sheet_name'] = sheet_name
                    all_objects[obj['obj_id']] = obj_copy
        
        return all_objects
    
    @staticmethod
    def _compare_objects(before_obj: Dict[str, Any], after_obj: Dict[str, Any]) -> Dict[str, Any]:
        """比較兩個物件的差異"""
        diff = {
            'has_changes': False,
            'position_changed': False,
            'changes': {}
        }
        
        # 比較基本屬性
        basic_fields = ['obj_name', 'obj_type', 'display_name']
        for field in basic_fields:
            before_val = before_obj.get(field)
            after_val = after_obj.get(field)
            if before_val != after_val:
                diff['changes'][field] = {'before': before_val, 'after': after_val}
                diff['has_changes'] = True
        
        # 比較位置資訊
        position_fields = ['position_before', 'position_after']
        for pos_field in position_fields:
            if pos_field in before_obj or pos_field in after_obj:
                before_pos = before_obj.get(pos_field, {})
                after_pos = after_obj.get(pos_field, {})
                
                pos_diff = RegistryUtils._compare_positions(before_pos, after_pos)
                if pos_diff['has_changes']:
                    diff['changes'][pos_field] = pos_diff
                    diff['has_changes'] = True
                    diff['position_changed'] = True
        
        return diff
    
    @staticmethod
    def _compare_positions(before_pos: Dict[str, Any], after_pos: Dict[str, Any]) -> Dict[str, Any]:
        """比較位置資訊"""
        pos_diff = {
            'has_changes': False,
            'changes': {}
        }
        
        # 比較基本位置欄位
        position_fields = ['row', 'col', 'coordinate', 'range_string']
        for field in position_fields:
            before_val = before_pos.get(field)
            after_val = after_pos.get(field)
            if before_val != after_val:
                pos_diff['changes'][field] = {'before': before_val, 'after': after_val}
                pos_diff['has_changes'] = True
        
        return pos_diff
    
    @staticmethod
    def find_latest_registry_file(directory: str = None) -> Optional[str]:
        """
        找出最新的註冊表檔案
        
        Args:
            directory: 搜尋目錄，預設為當前工作目錄
            
        Returns:
            str or None: 最新的註冊表檔案路徑，找不到則返回None
        """
        if directory is None:
            directory = os.getcwd()
        
        directory = Path(directory)
        
        # 搜尋符合命名規範的檔案
        pattern = "template_renderer_obj_registry_*.json"
        registry_files = list(directory.glob(pattern))
        
        if not registry_files:
            return None
        
        # 按修改時間排序，取最新的
        latest_file = max(registry_files, key=lambda f: f.stat().st_mtime)
        return str(latest_file)
    
    @staticmethod
    def create_empty_registry() -> Dict[str, Any]:
        """
        建立空的註冊表結構
        
        Returns:
            dict: 空的註冊表結構
        """
        return {
            'registry_version': '1.0',
            'created_at': datetime.now().isoformat(),
            'worksheets': {},
            'summary': {
                'total_worksheets': 0,
                'total_objects': 0,
                'objects_relocated': 0,
                'position_changes': 0,
                'size_changes': 0
            }
        }