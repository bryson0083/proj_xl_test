"""
圖片物件管理器
"""
import logging

logger = logging.getLogger(__name__)

from typing import Dict, List, Optional
from dataclasses import dataclass

from openpyxl.worksheet.worksheet import Worksheet

from ..models.container import Container
from ..models.objects import ImageObject, ObjectInfo
from ..models.base import ObjectType, RangePosition


@dataclass
class ShiftInfo:
    """位移資訊"""
    sheet_name: str
    start_row: int
    shift_amount: int
    affected_objects: List[str]


class ImageObjectManager:
    """
    圖片物件管理器類別
    
    負責管理Excel中的圖片物件，包括位置更新和推移處理
    """
    
    def update_image_positions(self, container: Container, shift_info: ShiftInfo) -> None:
        """
        更新圖片物件位置
        
        根據表格渲染產生的位移，調整圖片物件的位置
        
        Args:
            container: 容器物件
            shift_info: 位移資訊
        """
        # 找出需要調整的圖片物件
        image_objects = [obj for obj in container.objects if obj.obj_type == ObjectType.IMAGE_OBJ]
        
        for image_obj in image_objects:
            if self._should_shift_image(image_obj, shift_info):
                self._apply_shift_to_image(image_obj, shift_info)
    
    def _should_shift_image(self, image_obj: ObjectInfo, shift_info: ShiftInfo) -> bool:
        """
        判斷圖片是否需要位移
        
        Args:
            image_obj: 圖片物件
            shift_info: 位移資訊
            
        Returns:
            bool: 是否需要位移
        """
        # 如果圖片位置在位移起始行之後，則需要位移
        return image_obj.cell_position.row >= shift_info.start_row
    
    def _apply_shift_to_image(self, image_obj: ObjectInfo, shift_info: ShiftInfo) -> None:
        """
        對圖片物件應用位移
        
        Args:
            image_obj: 圖片物件
            shift_info: 位移資訊
        """
        # 更新圖片物件的位置
        image_obj.cell_position.row += shift_info.shift_amount
    
    def scan_image_objects(self, worksheet: Worksheet, sheet_name: str) -> List[ImageObject]:
        """
        掃描工作表中的圖片物件
        
        Args:
            worksheet: 工作表物件
            sheet_name: 工作表名稱
            
        Returns:
            List[ImageObject]: 圖片物件清單
        """
        image_objects = []
        
        # 掃描工作表中的圖片 - 使用正確的openpyxl屬性
        images = getattr(worksheet, '_images', [])
        for idx, image in enumerate(images):
            image_obj = self._create_image_object_from_image(image, idx, sheet_name)
            if image_obj:
                image_objects.append(image_obj)
        
        return image_objects
    
    def _create_image_object_from_image(
        self, 
        image, 
        index: int, 
        sheet_name: str
    ) -> Optional[ImageObject]:
        """
        從Excel圖片物件創建ImageObject
        
        Args:
            image: Excel圖片物件
            index: 圖片索引
            sheet_name: 工作表名稱
            
        Returns:
            Optional[ImageObject]: 圖片物件，如果創建失敗則返回None
        """
        try:
            # 取得圖片錨點資訊
            anchor = image.anchor
            
            # 嘗試獲取from位置 - 支援多種屬性名稱
            from_pos = None
            for from_attr in ['_from', 'from']:
                if hasattr(anchor, from_attr):
                    from_pos = getattr(anchor, from_attr)
                    break
            
            if from_pos:
                from_row = getattr(from_pos, 'row', 0) + 1  # 轉換為1-based
                from_col = getattr(from_pos, 'col', 0) + 1
                from_row_off = getattr(from_pos, 'rowOff', 0)
                from_col_off = getattr(from_pos, 'colOff', 0)
                
                from_position = RangePosition(
                    row=from_row,
                    col=from_col,
                    row_off=from_row_off,
                    col_off=from_col_off
                )
            else:
                from_position = RangePosition(row=1, col=1, row_off=0, col_off=0)
            
            # 嘗試獲取to位置 - 支援多種屬性名稱
            to_position = None
            to_pos = None
            for to_attr in ['_to', 'to']:
                if hasattr(anchor, to_attr):
                    to_pos = getattr(anchor, to_attr)
                    break
            
            if to_pos:
                to_row = getattr(to_pos, 'row', 0) + 1  # 轉換為1-based
                to_col = getattr(to_pos, 'col', 0) + 1
                to_row_off = getattr(to_pos, 'rowOff', 0)
                to_col_off = getattr(to_pos, 'colOff', 0)
                
                to_position = RangePosition(
                    row=to_row,
                    col=to_col,
                    row_off=to_row_off,
                    col_off=to_col_off
                )
                
                logger.debug(f"DEBUG: ImageManager掃描到TwoCellAnchor - from:({from_row},{from_col}) to:({to_row},{to_col})")
            else:
                logger.debug(f"DEBUG: ImageManager掃描到OneCellAnchor - from:({from_row},{from_col})")
            
            # 確定錨點類型
            anchor_type = "TwoCellAnchor" if to_position else "OneCellAnchor"
            
            return ImageObject(
                image_id=f"image_{sheet_name}_{index}",
                anchor_type=anchor_type,
                from_position=from_position,
                to_position=to_position
            )
            
        except Exception:
            # 如果無法解析圖片資訊，返回None
            return None
    
    def update_image_anchors(
        self, 
        worksheet: Worksheet, 
        image_objects: List[ImageObject], 
        shift_info: ShiftInfo
    ) -> None:
        """
        更新工作表中的圖片錨點
        
        Args:
            worksheet: 工作表物件
            image_objects: 圖片物件清單
            shift_info: 位移資訊
        """
        images = getattr(worksheet, '_images', [])
        if not images:
            logger.debug("DEBUG: 工作表沒有圖片物件")
            return
        
        logger.debug(f"DEBUG: 開始更新 {len(images)} 個圖片錨點")
        
        # 更新每個圖片的錨點
        for idx, image in enumerate(images):
            # 對工作表3這樣的情況，直接檢查圖片是否需要位移
            # 這裡簡化邏輯：對於footer區塊的圖片，直接應用位移
            logger.debug(f"DEBUG: 處理圖片 {idx+1}")
            self._update_image_anchor_direct(image, shift_info)
    
    def _update_image_anchor_direct(self, image, shift_info: ShiftInfo) -> None:
        """
        直接更新圖片錨點（基於舊專案的正確做法）

        Args:
            image: Excel圖片物件
            shift_info: 位移資訊
        """
        try:
            anchor = image.anchor
            anchor_type = type(anchor).__name__

            logger.debug(f"DEBUG: 更新圖片錨點，位移量: {shift_info.shift_amount}")
            logger.debug(f"DEBUG: 錨點類型: {anchor_type}")

            # 檢查原始錨點位置 - 使用正確的屬性名稱
            original_from_row = None
            original_to_row = None

            # 讀取原始位置 - 使用 _from 屬性（正確的openpyxl屬性名稱）
            if hasattr(anchor, '_from') and anchor._from is not None:
                if hasattr(anchor._from, 'row') and anchor._from.row is not None:
                    original_from_row = anchor._from.row
                    logger.debug(f"DEBUG: 原始_from位置: {original_from_row} (0-based)")

            # 讀取原始to位置 - 使用 to 屬性
            if hasattr(anchor, 'to') and anchor.to is not None:
                if hasattr(anchor.to, 'row') and anchor.to.row is not None:
                    original_to_row = anchor.to.row
                    logger.debug(f"DEBUG: 原始to位置: {original_to_row} (0-based)")

            # 判斷是否需要位移（如果圖片位於表格之後）
            should_shift = False
            if original_from_row is not None:
                # openpyxl使用0-based索引，shift_info.start_row是1-based
                # 需要轉換比較：0-based圖片位置 >= (1-based起始行 - 1)
                if original_from_row >= (shift_info.start_row - 1):
                    should_shift = True
                    logger.debug(f"DEBUG: 圖片需要位移（原始from位置 {original_from_row} >= 起始行 {shift_info.start_row-1}）")

            if not should_shift:
                logger.debug("DEBUG: 圖片不需要位移")
                return

            # 基於舊專案的正確做法進行更新
            if anchor_type == "TwoCellAnchor":
                self._adjust_two_cell_anchor_correctly(anchor, shift_info.shift_amount)
            elif anchor_type == "OneCellAnchor":
                self._adjust_one_cell_anchor_correctly(anchor, shift_info.shift_amount)
            else:
                logger.debug(f"DEBUG: 未知錨點類型: {anchor_type}")

        except Exception as e:
            logger.debug(f"DEBUG: 圖片錨點更新失敗: {e}")
            import traceback
            logger.debug("例外堆疊", exc_info=True)

    def _adjust_two_cell_anchor_correctly(self, anchor, row_offset: int) -> None:
        """
        正確調整 TwoCellAnchor 類型的圖片位置（基於舊專案邏輯）

        Args:
            anchor: TwoCellAnchor 錨點物件
            row_offset: 行偏移量
        """
        if not hasattr(anchor, '_from') or not anchor._from:
            logger.debug(f"DEBUG: TwoCellAnchor 圖片缺少 _from 屬性")
            return

        if not hasattr(anchor, 'to') or not anchor.to:
            logger.debug(f"DEBUG: TwoCellAnchor 圖片缺少 to 屬性")
            return

        # 保存原始的所有屬性值
        original_from_row = anchor._from.row
        original_from_col = anchor._from.col
        original_from_rowOff = getattr(anchor._from, 'rowOff', 0)
        original_from_colOff = getattr(anchor._from, 'colOff', 0)

        original_to_row = anchor.to.row
        original_to_col = anchor.to.col
        original_to_rowOff = getattr(anchor.to, 'rowOff', 0)
        original_to_colOff = getattr(anchor.to, 'colOff', 0)

        # 調整 _from 位置
        anchor._from.row = original_from_row + row_offset
        # 保持 _from 的其他屬性不變
        if hasattr(anchor._from, 'rowOff'):
            anchor._from.rowOff = original_from_rowOff
        if hasattr(anchor._from, 'colOff'):
            anchor._from.colOff = original_from_colOff

        # 調整 to 位置
        anchor.to.row = original_to_row + row_offset
        # 保持 to 的其他屬性不變
        if hasattr(anchor.to, 'rowOff'):
            anchor.to.rowOff = original_to_rowOff
        if hasattr(anchor.to, 'colOff'):
            anchor.to.colOff = original_to_colOff

        logger.debug(f"DEBUG: TwoCellAnchor 圖片位置已調整：")
        logger.debug(f"DEBUG:   _from: 第{original_from_row}行 -> 第{anchor._from.row}行 (0-based)")
        logger.debug(f"DEBUG:   to: 第{original_to_row}行 -> 第{anchor.to.row}行 (0-based)")
        logger.debug(f"DEBUG:   rowOff 保持: _from({original_from_rowOff}), to({original_to_rowOff})")
        logger.debug(f"DEBUG:   colOff 保持: _from({original_from_colOff}), to({original_to_colOff})")

    def _adjust_one_cell_anchor_correctly(self, anchor, row_offset: int) -> None:
        """
        正確調整 OneCellAnchor 類型的圖片位置（基於舊專案邏輯）

        Args:
            anchor: OneCellAnchor 錨點物件
            row_offset: 行偏移量
        """
        if not hasattr(anchor, '_from') or not anchor._from:
            logger.debug(f"DEBUG: OneCellAnchor 圖片缺少 _from 屬性")
            return

        # 保存原始的 rowOff 和 colOff 值
        original_from_row = anchor._from.row
        original_from_col = anchor._from.col
        original_from_rowOff = getattr(anchor._from, 'rowOff', 0)
        original_from_colOff = getattr(anchor._from, 'colOff', 0)

        # 調整 _from 位置
        anchor._from.row = original_from_row + row_offset
        # 保持 _from 的其他屬性不變
        if hasattr(anchor._from, 'rowOff'):
            anchor._from.rowOff = original_from_rowOff
        if hasattr(anchor._from, 'colOff'):
            anchor._from.colOff = original_from_colOff

        logger.debug(f"DEBUG: OneCellAnchor 圖片位置已調整：")
        logger.debug(f"DEBUG:   _from: 第{original_from_row}行 -> 第{anchor._from.row}行 (0-based)")
        logger.debug(f"DEBUG:   rowOff 保持: {original_from_rowOff}")
        logger.debug(f"DEBUG:   colOff 保持: {original_from_colOff}")
    
    def _should_shift_image_object(self, image_obj: ImageObject, shift_info: ShiftInfo) -> bool:
        """
        判斷圖片物件是否需要位移
        
        Args:
            image_obj: 圖片物件
            shift_info: 位移資訊
            
        Returns:
            bool: 是否需要位移
        """
        return image_obj.from_position.row >= shift_info.start_row
    
    def get_images_in_range(
        self, 
        image_objects: List[ImageObject], 
        start_row: int, 
        end_row: int
    ) -> List[ImageObject]:
        """
        取得指定範圍內的圖片物件
        
        Args:
            image_objects: 圖片物件清單
            start_row: 開始行
            end_row: 結束行
            
        Returns:
            List[ImageObject]: 範圍內的圖片物件
        """
        result = []
        
        for image_obj in image_objects:
            image_row = image_obj.from_position.row
            if start_row <= image_row <= end_row:
                result.append(image_obj)
        
        return result
    
    def calculate_image_shift_impact(
        self, 
        container: Container, 
        shift_info: ShiftInfo
    ) -> Dict[str, int]:
        """
        計算圖片位移的影響
        
        Args:
            container: 容器物件
            shift_info: 位移資訊
            
        Returns:
            Dict[str, int]: 圖片ID到新位置的映射
        """
        impact = {}
        
        image_objects = [obj for obj in container.objects if obj.obj_type == ObjectType.IMAGE_OBJ]
        
        for image_obj in image_objects:
            if self._should_shift_image(image_obj, shift_info):
                new_row = image_obj.cell_position.row + shift_info.shift_amount
                impact[image_obj.obj_id] = new_row
        
        return impact
