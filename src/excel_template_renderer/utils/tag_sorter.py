"""
標籤排序器
"""
from typing import Dict, List, Tuple
from operator import attrgetter

from ..models.container import Container
from ..models.tag import Tag
from ..models.base import TagType


class TagSorter:
    """
    標籤排序器類別
    
    負責按照sheet和位置排序標籤，確保渲染順序正確
    """
    
    def sort_tags_by_sheet(self, container: Container, tags: List[Tag]) -> Dict[str, List[Tag]]:
        """
        依照sheet和位置排序標籤
        
        Args:
            container: 容器物件
            tags: 標籤清單
            
        Returns:
            Dict[str, List[Tag]]: 按sheet分組並排序的標籤
        """
        # 按sheet分組
        sheet_tags = {}
        for tag in tags:
            sheet_name = tag.sheet_name
            if sheet_name not in sheet_tags:
                sheet_tags[sheet_name] = []
            sheet_tags[sheet_name].append(tag)
        
        # 對每個sheet的標籤進行排序
        for sheet_name, tag_list in sheet_tags.items():
            sheet_tags[sheet_name] = self._sort_tags_by_position(tag_list)
        
        return sheet_tags
    
    def _sort_tags_by_position(self, tags: List[Tag]) -> List[Tag]:
        """
        按照位置排序標籤
        
        Args:
            tags: 標籤清單
            
        Returns:
            List[Tag]: 排序後的標籤清單
        """
        # 按照行號、列號排序
        return sorted(tags, key=lambda tag: (tag.cell_position.row, tag.cell_position.col))
    
    def sort_tags_by_type_and_position(self, tags: List[Tag]) -> List[Tag]:
        """
        按照類型和位置排序標籤
        
        優先順序：
        1. 簡單標籤 (SIMPLE) 優先於表格標籤 (TABLE)
        2. 同類型標籤按位置排序
        
        Args:
            tags: 標籤清單
            
        Returns:
            List[Tag]: 排序後的標籤清單
        """
        def sort_key(tag: Tag):
            # 類型優先級：SIMPLE = 0, TABLE = 1
            type_priority = 0 if tag.tag_type == TagType.SIMPLE else 1
            return (type_priority, tag.cell_position.row, tag.cell_position.col)
        
        return sorted(tags, key=sort_key)
    
    def get_rendering_order(self, container: Container, tags: List[Tag]) -> List[Tag]:
        """
        取得渲染順序的標籤清單
        
        根據區塊類型和位置確定最佳的渲染順序
        
        Args:
            container: 容器物件
            tags: 標籤清單
            
        Returns:
            List[Tag]: 按渲染順序排列的標籤清單
        """
        # 按區塊分組標籤
        block_tags = self._group_tags_by_block(container, tags)
        
        ordered_tags = []
        
        # 按區塊順序處理：Header -> Gap -> Footer
        for block in container.blocks:
            if block.block_id in block_tags:
                block_tag_list = block_tags[block.block_id]
                # 區塊內按位置排序
                sorted_block_tags = self._sort_tags_by_position(block_tag_list)
                ordered_tags.extend(sorted_block_tags)
        
        return ordered_tags
    
    def _group_tags_by_block(self, container: Container, tags: List[Tag]) -> Dict[str, List[Tag]]:
        """
        按區塊分組標籤
        
        Args:
            container: 容器物件
            tags: 標籤清單
            
        Returns:
            Dict[str, List[Tag]]: 按區塊分組的標籤
        """
        block_tags = {}
        
        for tag in tags:
            # 找到標籤所屬的區塊
            block_id = self._find_tag_block(container, tag)
            if block_id:
                if block_id not in block_tags:
                    block_tags[block_id] = []
                block_tags[block_id].append(tag)
        
        return block_tags
    
    def _find_tag_block(self, container: Container, tag: Tag) -> str:
        """
        找到標籤所屬的區塊
        
        Args:
            container: 容器物件
            tag: 標籤物件
            
        Returns:
            str: 區塊ID，如果找不到則返回空字串
        """
        tag_row = tag.cell_position.row
        
        for block in container.blocks:
            if block.rng_from.row <= tag_row <= block.rng_to.row:
                return block.block_id
        
        return ""
    
    def sort_table_tags_for_rendering(self, table_tags: List[Tag]) -> List[Tag]:
        """
        為表格標籤排序以優化渲染
        
        表格標籤需要特殊的排序邏輯，因為它們會影響後續元素的位置
        
        Args:
            table_tags: 表格標籤清單
            
        Returns:
            List[Tag]: 排序後的表格標籤清單
        """
        # 表格標籤按行號從上到下排序，確保上方的表格先渲染
        # 這樣可以正確計算位移量
        return sorted(table_tags, key=attrgetter('cell_position.row'))
    
    def separate_simple_and_table_tags(self, tags: List[Tag]) -> Tuple[List[Tag], List[Tag]]:
        """
        分離簡單標籤和表格標籤
        
        Args:
            tags: 標籤清單
            
        Returns:
            tuple[List[Tag], List[Tag]]: (簡單標籤清單, 表格標籤清單)
        """
        simple_tags = [tag for tag in tags if tag.tag_type == TagType.SIMPLE]
        table_tags = [tag for tag in tags if tag.tag_type == TagType.TABLE]
        
        return simple_tags, table_tags
