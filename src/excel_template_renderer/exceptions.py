"""
例外處理類別
"""


class ExcelTemplateRendererError(Exception):
    """基礎例外類別"""
    pass


class TemplateNotFoundError(ExcelTemplateRendererError):
    """模板檔案不存在例外"""
    
    def __init__(self, template_path: str):
        self.template_path = template_path
        super().__init__(f"模板檔案不存在: {template_path}")


class InvalidDataTypeError(ExcelTemplateRendererError):
    """不支援的資料類型例外"""
    
    def __init__(self, data_type: str, tag_name: str):
        self.data_type = data_type
        self.tag_name = tag_name
        super().__init__(f"標籤 '{tag_name}' 不支援資料類型: {data_type}")


class InvalidTagSyntaxError(ExcelTemplateRendererError):
    """無效標籤語法例外"""
    
    def __init__(self, tag_string: str, position: str = ""):
        self.tag_string = tag_string
        self.position = position
        pos_info = f" (位置: {position})" if position else ""
        super().__init__(f"無效的標籤語法: '{tag_string}'{pos_info}")


class RenderError(ExcelTemplateRendererError):
    """渲染過程錯誤例外"""
    
    def __init__(self, message: str, tag_name: str = "", sheet_name: str = ""):
        self.tag_name = tag_name
        self.sheet_name = sheet_name
        location_info = ""
        if sheet_name:
            location_info += f" (Sheet: {sheet_name}"
            if tag_name:
                location_info += f", 標籤: {tag_name}"
            location_info += ")"
        super().__init__(f"渲染錯誤: {message}{location_info}")


class RangeOverlapError(ExcelTemplateRendererError):
    """範圍重疊錯誤例外"""
    
    def __init__(self, obj1_id: str, obj2_id: str, sheet_name: str = ""):
        self.obj1_id = obj1_id
        self.obj2_id = obj2_id
        self.sheet_name = sheet_name
        sheet_info = f" (Sheet: {sheet_name})" if sheet_name else ""
        super().__init__(f"物件範圍重疊: {obj1_id} 與 {obj2_id}{sheet_info}")


class TableObjectConflictError(ExcelTemplateRendererError):
    """表格物件衝突錯誤例外"""
    
    def __init__(self, table_ids: list, sheet_name: str = ""):
        self.table_ids = table_ids
        self.sheet_name = sheet_name
        table_list = ", ".join(table_ids)
        sheet_info = f" (Sheet: {sheet_name})" if sheet_name else ""
        super().__init__(f"表格物件範圍衝突: {table_list}{sheet_info}")


class FileFormatError(ExcelTemplateRendererError):
    """檔案格式錯誤例外"""
    
    def __init__(self, file_path: str, expected_format: str = "xlsx/xlsm"):
        self.file_path = file_path
        self.expected_format = expected_format
        super().__init__(f"不支援的檔案格式: {file_path}，期望格式: {expected_format}")


class MemoryError(ExcelTemplateRendererError):
    """記憶體不足錯誤例外"""
    
    def __init__(self, operation: str = ""):
        self.operation = operation
        op_info = f" ({operation})" if operation else ""
        super().__init__(f"記憶體不足{op_info}")


class ProcessIsolationError(ExcelTemplateRendererError):
    """程序隔離錯誤例外"""
    
    def __init__(self, process_id: str, message: str):
        self.process_id = process_id
        super().__init__(f"程序隔離錯誤 (Process: {process_id}): {message}")
