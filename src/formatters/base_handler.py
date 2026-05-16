"""
公文格式化工具 - 格式处理器基类
定义所有格式处理器的统一接口
"""

from abc import ABC, abstractmethod
from typing import Optional, Tuple
import logging

logger = logging.getLogger(__name__)


class FormatHandler(ABC):
    """格式处理器基类。
    
    所有具体格式处理器必须继承此类并实现抽象方法。
    """
    
    @property
    @abstractmethod
    def supported_extensions(self) -> list:
        """返回支持的扩展名列表（小写，带点号）。
        
        Returns:
            支持的扩展名列表，如 ['.docx']
        """
        pass
    
    @abstractmethod
    def can_handle(self, file_path: str) -> bool:
        """检查是否可以处理指定文件。
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否可以处理该文件
        """
        pass
    
    @abstractmethod
    def validate_format(self, file_path: str) -> Tuple[bool, str]:
        """验证文档格式兼容性。
        
        Args:
            file_path: 文件路径
            
        Returns:
            (是否兼容, 提示信息)
        """
        pass
    
    @abstractmethod
    def format_document(self, input_path: str, output_path: str, 
                       backup_path: Optional[str] = None) -> Optional[str]:
        """格式化文档。
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            backup_path: 备份文件路径（可选）
            
        Returns:
            实际备份文件路径，如果未备份则返回None
            
        Raises:
            FileNotFoundError: 输入文件不存在
            OSError: 备份或保存文件失败
            ValueError: 文档处理错误
        """
        pass
    
    def get_format_info(self) -> dict:
        """获取格式处理器信息。
        
        Returns:
            格式信息字典
        """
        return {
            'extensions': self.supported_extensions,
            'handler_name': self.__class__.__name__,
        }
