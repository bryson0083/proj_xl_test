"""
主要API接口
"""
import os
import uuid
from typing import Any, Dict, Optional

from openpyxl import load_workbook, Workbook

from .core.parser import TemplateParser
from .core.renderer import TemplateRenderer
from .core.container import ContainerManager
from .core.block_manager import BlockManager
from .exceptions import TemplateNotFoundError, FileFormatError, RenderError
from .context import RenderContext
from .models.container import Container
from .utils.registry_utils import RegistryUtils


def render_template(
    template_path: str,
    output_file_name: str,
    process_id: Optional[str] = None,
    validate_result: bool = False,
    **data: Any
) -> Dict[str, Any]:
    """
    主要的模板渲染函數（內建程序隔離安全機制）

    Args:
        template_path: 模板文件路徑
        output_file_name: 輸出文件名稱
        process_id: 程序識別ID（可選）
                   - 若未提供，系統將自動生成唯一ID
        validate_result: 是否驗證渲染結果（可選）
        **data: 渲染數據，以關鍵字參數形式傳入

    Returns:
        Dict[str, Any]: 渲染結果資訊，包含：
            - success: 渲染是否成功
            - registry_file: 物件註冊表檔案路徑
            - validation_result: 驗證結果（若啟用）

    Raises:
        TemplateNotFoundError: 模板文件不存在
        FileFormatError: 不支援的檔案格式
        RenderError: 渲染過程發生錯誤

    Examples:
        >>> from excel_template_renderer import render_template
        >>> import pandas as pd
        >>>
        >>> data = dict(
        ...     oper_name="OPER_NAME",
        ...     report_df=pd.DataFrame({'姓名': ['Alice', 'Bob'], '部門': ['技術部', '業務部']}),
        ...     date_rng_desc="2025/01/01 - 2025/01/31"
        ... )
        >>> result = render_template('template.xlsx', 'output.xlsx', **data)
        >>> print(f"註冊表檔案: {result['registry_file']}")
    """
    # 生成程序ID
    if process_id is None:
        process_id = str(uuid.uuid4())
    
    registry_file = None
    result = {
        'success': False,
        'registry_file': None,
        'validation_result': None
    }

    try:
        # 驗證模板文件存在
        if not os.path.exists(template_path):
            raise TemplateNotFoundError(template_path)

        # 驗證檔案格式
        if not _is_supported_format(template_path):
            raise FileFormatError(template_path)

        # 載入模板工作簿
        workbook = load_workbook(template_path)

        # 創建渲染上下文
        render_context = RenderContext(
            process_id=process_id,
            template_path=template_path,
            output_path=output_file_name,
            data=data
        )

        # 執行渲染流程並收集容器資訊
        containers = _execute_render_pipeline(workbook, render_context)

        # 在儲存前進行最終的表格範圍同步檢查
        _final_table_autofilter_sync(workbook)

        # 儲存輸出檔案
        workbook.save(output_file_name)

        # 生成物件註冊表
        registry_file = _generate_object_registry(containers, render_context, template_path, output_file_name)

        result['success'] = True
        result['registry_file'] = registry_file

        # 驗證結果（如果要求）
        if validate_result:
            result['validation_result'] = _validate_render_result(output_file_name, registry_file)

        return result

    except Exception as e:
        if isinstance(e, (TemplateNotFoundError, FileFormatError, RenderError)):
            raise
        else:
            raise RenderError(f"渲染過程發生未預期錯誤: {str(e)}")


def _is_supported_format(file_path: str) -> bool:
    """
    檢查檔案格式是否支援
    
    Args:
        file_path: 檔案路徑
        
    Returns:
        bool: 是否支援
    """
    supported_extensions = ['.xlsx', '.xlsm']
    file_ext = os.path.splitext(file_path)[1].lower()
    return file_ext in supported_extensions


def _execute_render_pipeline(workbook: Workbook, render_context: 'RenderContext') -> list:
    """
    執行渲染管道

    Args:
        workbook: Excel工作簿
        render_context: 渲染上下文

    Returns:
        list: 容器清單
    """
    # 使用新的架構：統一的模板掃描與註冊表建立
    container_manager = ContainerManager()
    
    # 重要修正：先掃描原始標籤，保留所有標籤的完整資訊
    from .core.parser import TemplateParser
    parser = TemplateParser()
    original_tags = parser.parse_template(workbook)
    
    # 建立容器
    containers = container_manager.create_containers(workbook)
    
    # 3. 區塊分類
    for container in containers:
        container_manager.classify_blocks(container)
    
    # 4. 建立標籤到物件的映射關係
    # 使用原始標籤資訊建立映射，確保相同名稱的多個標籤都能正確處理
    _build_tag_object_mapping_with_original_tags(containers, original_tags, render_context)
    
    # 5. 創建渲染器和區塊管理器
    renderer = TemplateRenderer()
    block_manager = BlockManager()
    
    # 6. 按區塊順序執行渲染
    for container in containers:
        _render_container(container, workbook, render_context, renderer, block_manager)

    return containers


def _render_container(
    container: Container,
    workbook: Workbook,
    render_context: RenderContext,
    renderer: 'TemplateRenderer', 
    block_manager: 'BlockManager'
) -> None:
    """
    渲染單一容器
    
    Args:
        container: 容器物件
        workbook: Excel工作簿
        render_context: 渲染上下文
        renderer: 渲染器
        block_manager: 區塊管理器
    """
    print(f"DEBUG: 使用新的block分區渲染搬移機制處理容器: {container.sheet_name}")
    print(f"DEBUG: Calling block_manager.process_container_with_block_moving")
    
    try:
        # 使用新的block搬移機制
        block_manager.process_container_with_block_moving(container, workbook, render_context, renderer)
        print(f"DEBUG: Successfully called process_container_with_block_moving")
    except Exception as e:
        print(f"DEBUG: Error in process_container_with_block_moving: {e}")
        import traceback
        traceback.print_exc()
        raise
    
    print(f"DEBUG: Finished processing container: {container.sheet_name}")


def _build_tag_object_mapping_with_original_tags(containers, original_tags, render_context):
    """
    使用原始標籤資訊建立標籤到物件的映射關係
    
    確保相同名稱的多個標籤都能正確對應到各自的物件
    
    Args:
        containers: 容器清單
        original_tags: 原始標籤清單（來自TemplateParser的掃描結果）
        render_context: 渲染上下文
    """
    from .models.base import ObjectType
    
    print(f"DEBUG_MAPPING: Building tag-object mapping with {len(original_tags)} original tags")
    
    # 建立位置到標籤清單的映射（支援同一位置多個標籤）
    position_to_tags = {}
    for tag in original_tags:
        key = (tag.sheet_name, tag.cell_position.row, tag.cell_position.col)
        if key not in position_to_tags:
            position_to_tags[key] = []
        position_to_tags[key].append(tag)
        print(f"DEBUG_MAPPING: Tag {tag.tag_name} at position {key}")
    
    # 為每個物件找到對應的標籤
    for container in containers:
        print(f"DEBUG_MAPPING: Processing container {container.sheet_name} with {len(container.objects)} objects")
        for obj in container.objects:
            # 處理標籤相關的物件（SIMPLE、TABLE和TABLE_OBJ類型）
            # TABLE_OBJ是標籤與表格物件綁定後的類型
            if obj.obj_type in [ObjectType.SIMPLE, ObjectType.TABLE, ObjectType.TABLE_OBJ]:
                key = (obj.sheet_name, obj.cell_position.row, obj.cell_position.col)
                print(f"DEBUG_MAPPING: Object {obj.obj_id} type={obj.obj_type} display_name={obj.display_name} at position {key}")
                # 對於TABLE_OBJ類型，需要特殊處理（因為綁定改變了位置）
                if obj.obj_type == ObjectType.TABLE_OBJ:
                    # TABLE_OBJ的display_name保留了原始標籤名
                    # 需要在所有標籤中查找匹配的標籤名
                    matched_tag = None
                    for pos_key, tags in position_to_tags.items():
                        if pos_key[0] == obj.sheet_name:  # 同一工作表
                            for tag in tags:
                                if tag.tag_name == obj.display_name:
                                    matched_tag = tag
                                    print(f"DEBUG_MAPPING: Found tag {tag.tag_name} for TABLE_OBJ {obj.obj_id}")
                                    break
                            if matched_tag:
                                break
                elif key in position_to_tags:
                    # 獲取該位置的所有標籤
                    candidate_tags = position_to_tags[key]
                    
                    # 嘗試根據物件名稱匹配標籤
                    matched_tag = None
                    for tag in candidate_tags:
                        # 檢查物件ID是否包含標籤名稱
                        if tag.tag_name in obj.obj_id:
                            matched_tag = tag
                            break
                    
                    # 如果沒有匹配到，使用第一個標籤作為備選
                    if not matched_tag and candidate_tags:
                        matched_tag = candidate_tags[0]
                else:
                    matched_tag = None
                    
                if matched_tag:
                    render_context.add_tag_mapping(obj.obj_id, matched_tag)
                    print(f"DEBUG: 映射物件 {obj.obj_id} 到標籤 {matched_tag.tag_name} 在位置 ({matched_tag.cell_position.row}, {matched_tag.cell_position.col})")
                else:
                    print(f"DEBUG: 警告 - 無法找到物件 {obj.obj_id} 的匹配標籤 at position {key}")


def _build_tag_object_mapping_from_containers(containers, render_context):
    """
    從容器中的物件重建標籤到物件的映射關係
    
    注意：此函數已棄用，改用 _build_tag_object_mapping_with_original_tags
    
    Args:
        containers: 容器清單
        render_context: 渲染上下文
    """
    from .models.base import ObjectType
    
    for container in containers:
        for obj in container.objects:
            # 只處理標籤相關的物件（SIMPLE和TABLE類型）
            if obj.obj_type in [ObjectType.SIMPLE, ObjectType.TABLE]:
                # 根據物件資訊重建標籤物件
                from .models.tag import Tag
                from .models.base import TagType, DataType, RenderDirection
                
                tag = Tag(
                    tag_name=obj.display_name,
                    tag_type=TagType.TABLE if obj.obj_type == ObjectType.TABLE else TagType.SIMPLE,
                    has_condition=not obj.having_header,  # 對SIMPLE和TABLE都檢查having_header
                    condition="noheader" if not obj.having_header else None,  # 對SIMPLE和TABLE都適用
                    cell_position=obj.cell_position,
                    data_type=DataType.UNKNOWN,
                    render_direction=RenderDirection.VERTICAL if obj.obj_type == ObjectType.TABLE else RenderDirection.HORIZONTAL,
                    sheet_name=obj.sheet_name
                )
                
                # 建立映射關係
                render_context.add_tag_mapping(obj.obj_id, tag)


def _final_table_autofilter_sync(workbook: Workbook) -> None:
    """
    在儲存前進行最終的表格範圍同步檢查
    確保所有表格的 autoFilter.ref 與 table.ref 保持一致
    
    Args:
        workbook: Excel工作簿
    """
    print("DEBUG_FINAL_SYNC: 開始進行最終的表格範圍同步檢查")
    sync_count = 0
    
    for worksheet in workbook.worksheets:
        print(f"DEBUG_FINAL_SYNC: 檢查工作表 {worksheet.title}")
        
        for table_name in worksheet.tables:
            table = worksheet.tables[table_name]
            table_ref = getattr(table, 'ref', '')
            
            if hasattr(table, 'autoFilter') and table.autoFilter:
                autofilter_ref = table.autoFilter.ref
                
                if table_ref != autofilter_ref:
                    print(f"DEBUG_FINAL_SYNC: 發現不同步 - 表格 {table_name}")
                    print(f"  table.ref: {table_ref}")
                    print(f"  autoFilter.ref: {autofilter_ref}")
                    print(f"  正在同步...")
                    
                    # 同步 autoFilter.ref
                    table.autoFilter.ref = table_ref
                    sync_count += 1
                    
                    print(f"  已同步為: {table.autoFilter.ref}")
                else:
                    print(f"DEBUG_FINAL_SYNC: 表格 {table_name} 範圍已同步: {table_ref}")
            else:
                print(f"DEBUG_FINAL_SYNC: 表格 {table_name} 沒有 autoFilter")
    
    print(f"DEBUG_FINAL_SYNC: 同步檢查完成，共修正 {sync_count} 個表格")


def _build_tag_object_mapping(containers, tags, render_context):
    """
    建立標籤到物件的映射關係
    
    Args:
        containers: 容器清單
        tags: 標籤清單  
        render_context: 渲染上下文
    """
    # 建立位置到標籤的映射
    position_to_tag = {}
    for tag in tags:
        key = (tag.sheet_name, tag.cell_position.row, tag.cell_position.col)
        position_to_tag[key] = tag
    
    # 為每個物件找到對應的標籤
    for container in containers:
        for obj in container.objects:
            key = (obj.sheet_name, obj.cell_position.row, obj.cell_position.col)
            if key in position_to_tag:
                tag = position_to_tag[key]
                render_context.add_tag_mapping(obj.obj_id, tag)


def _generate_object_registry(containers: list, render_context: 'RenderContext', template_path: str, output_path: str) -> str:
    """
    生成物件註冊表

    Args:
        containers: 容器清單
        render_context: 渲染上下文
        template_path: 模板文件路徑
        output_path: 輸出文件路徑

    Returns:
        str: 註冊表檔案路徑
    """
    from .models.base import ObjectType

    # 建立註冊表結構
    registry = RegistryUtils.create_empty_registry()
    registry['template_path'] = template_path
    registry['output_path'] = output_path

    # 統計變數
    total_objects = 0
    objects_relocated = 0
    position_changes = 0
    size_changes = 0

    # 處理每個容器（工作表）
    for container in containers:
        worksheet_data = {
            'container_id': container.container_id,
            'sheet_name': container.sheet_name,
            'total_objects': len(container.objects),
            'total_blocks': len(container.blocks),
            'objects': []
        }

        # 處理每個物件
        for obj in container.objects:
            # 計算渲染前位置
            position_before = {
                'row': obj.cell_position.row,
                'col': obj.cell_position.col,
                'coordinate': f"{chr(64 + obj.cell_position.col)}{obj.cell_position.row}",
                'data_shape': {
                    'rows': obj.data_shape.rows,
                    'cols': obj.data_shape.cols
                }
            }

            # 計算渲染後位置（基於標籤映射和數據）
            position_after = position_before.copy()
            tag = render_context.get_tag_for_object(obj.obj_id)

            if tag and render_context.has_data(tag.tag_name):
                data = render_context.get_data(tag.tag_name)
                if hasattr(data, 'shape') and obj.obj_type in [ObjectType.TABLE, ObjectType.TABLE_OBJ]:
                    # 表格物件可能會改變大小
                    data_rows = data.shape[0]
                    data_cols = data.shape[1]

                    # 檢查是否有標題行
                    has_header = obj.having_header
                    if tag.has_condition and tag.condition == "noheader":
                        has_header = False

                    total_rows = data_rows + (1 if has_header else 0)

                    position_after['data_shape'] = {
                        'rows': total_rows,
                        'cols': data_cols
                    }

                    if total_rows != obj.data_shape.rows or data_cols != obj.data_shape.cols:
                        size_changes += 1

            obj_data = {
                'obj_id': obj.obj_id,
                'obj_name': obj.obj_name if hasattr(obj, 'obj_name') else obj.display_name,
                'display_name': obj.display_name,
                'obj_type': obj.obj_type.value if hasattr(obj.obj_type, 'value') else str(obj.obj_type),
                'sheet_name': obj.sheet_name,
                'is_multi_rows': obj.is_multi_rows,
                'having_header': obj.having_header,
                'position_before': position_before,
                'position_after': position_after,
                'block_id': obj.block_id
            }

            worksheet_data['objects'].append(obj_data)
            total_objects += 1

        registry['worksheets'][container.sheet_name] = worksheet_data

    # 更新摘要統計
    registry['summary'].update({
        'total_worksheets': len(containers),
        'total_objects': total_objects,
        'objects_relocated': objects_relocated,
        'position_changes': position_changes,
        'size_changes': size_changes
    })

    # 序列化並儲存註冊表
    try:
        registry_file = RegistryUtils.serialize_registry(registry)
        print(f"DEBUG: 物件註冊表已生成: {registry_file}")
        return registry_file
    except Exception as e:
        print(f"WARNING: 無法生成物件註冊表: {e}")
        return ""


def _validate_render_result(output_file: str, registry_file: str) -> bool:
    """
    驗證渲染結果

    Args:
        output_file: 輸出檔案路徑
        registry_file: 註冊表檔案路徑

    Returns:
        bool: 驗證是否通過
    """
    try:
        # 基本檔案存在性檢查
        if not os.path.exists(output_file):
            print(f"WARNING: 輸出檔案不存在: {output_file}")
            return False

        if registry_file and not os.path.exists(registry_file):
            print(f"WARNING: 註冊表檔案不存在: {registry_file}")
            return False

        # 如果有註冊表，驗證其格式
        if registry_file:
            registry = RegistryUtils.load_registry(registry_file)
            is_valid, errors = RegistryUtils.validate_registry(registry)

            if not is_valid:
                print(f"WARNING: 註冊表驗證失敗: {errors}")
                return False

        print("DEBUG: 渲染結果驗證通過")
        return True

    except Exception as e:
        print(f"WARNING: 驗證過程發生錯誤: {e}")
        return False
