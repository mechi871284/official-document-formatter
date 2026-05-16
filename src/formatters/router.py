"""
公文格式化工具 - 格式检测与路由模块
自动检测文件格式并路由到对应的处理器
"""

import os
import logging
from typing import Optional, List, Tuple

from .base_handler import FormatHandler
from .docx_handler import DocxFormatHandler
from .doc_handler import DocFormatHandler
from .wps_handler import WpsFormatHandler

logger = logging.getLogger(__name__)


class FormatRouter:
    """格式路由器。
    
    根据文件扩展名和内容自动选择合适的格式处理器。
    """
    
    def __init__(self):
        self._handlers: List[FormatHandler] = [
            DocxFormatHandler(),
            DocFormatHandler(),
            WpsFormatHandler(),
        ]
    
    def get_handler(self, file_path: str) -> Optional[FormatHandler]:
        """根据文件路径获取合适的格式处理器。
        
        Args:
            file_path: 文件路径
            
        Returns:
            格式处理器实例，如果没有合适的处理器则返回None
        """
        for handler in self._handlers:
            if handler.can_handle(file_path):
                return handler
        return None
    
    def get_supported_extensions(self) -> list:
        """获取所有支持的扩展名。
        
        Returns:
            支持的扩展名列表
        """
        extensions = []
        for handler in self._handlers:
            extensions.extend(handler.supported_extensions)
        return extensions
    
    def validate_file(self, file_path: str) -> Tuple[bool, str, Optional[FormatHandler]]:
        """验证文件格式并返回处理器。
        
        Args:
            file_path: 文件路径
            
        Returns:
            (是否有效, 提示信息, 处理器实例)
        """
        if not os.path.exists(file_path):
            return False, f"文件不存在：{file_path}", None
        
        handler = self.get_handler(file_path)
        if handler is None:
            ext = os.path.splitext(file_path)[1].lower()
            supported = self.get_supported_extensions()
            return False, f"不支持的文件格式：{ext}，支持的格式：{', '.join(supported)}", None
        
        is_valid, message = handler.validate_format(file_path)
        return is_valid, message, handler
    
    def format_file(self, input_path: str, output_path: str,
                   backup_path: Optional[str] = None) -> Optional[str]:
        """格式化文件，自动选择处理器。
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            backup_path: 备份文件路径（可选）
            
        Returns:
            实际备份文件路径
            
        Raises:
            FileNotFoundError: 输入文件不存在
            ValueError: 不支持的文件格式或处理失败
            OSError: 备份或保存失败
        """
        is_valid, message, handler = self.validate_file(input_path)
        if not is_valid:
            raise ValueError(message)
        
        return handler.format_document(input_path, output_path, backup_path)
    
    def get_handler_info(self) -> list:
        """获取所有处理器信息。
        
        Returns:
            处理器信息列表
        """
        return [handler.get_format_info() for handler in self._handlers]


# 全局路由器实例
_router = None


def get_router() -> FormatRouter:
    """获取全局格式路由器实例。
    
    Returns:
        FormatRouter实例
    """
    global _router
    if _router is None:
        _router = FormatRouter()
    return _router
