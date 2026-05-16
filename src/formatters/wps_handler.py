"""
公文格式化工具 - WPS格式处理器
通过win32com将.wps转换为.docx后处理
"""

import os
import logging
from typing import Optional, Tuple

from .base_handler import FormatHandler
from .docx_handler import DocxFormatHandler

logger = logging.getLogger(__name__)


class WpsFormatHandler(FormatHandler):
    """WPS格式处理器。
    
    通过win32com将WPS .wps格式转换为.docx，然后使用DOCX处理器进行格式化。
    需要安装pywin32库和WPS Office或Microsoft Word。
    """
    
    def __init__(self):
        self._docx_handler = DocxFormatHandler()
        self._win32_available = self._check_win32()
    
    def _check_win32(self) -> bool:
        """检查win32com是否可用。"""
        try:
            import win32com.client
            return True
        except ImportError:
            return False
    
    @property
    def supported_extensions(self) -> list:
        return ['.wps']
    
    def can_handle(self, file_path: str) -> bool:
        return file_path.lower().endswith('.wps') and self._win32_available
    
    def validate_format(self, file_path: str) -> Tuple[bool, str]:
        """验证WPS文件格式。"""
        if not os.path.exists(file_path):
            return False, f"文件不存在：{file_path}"
        if not self._win32_available:
            return False, "需要安装pywin32库和WPS Office/Word才能处理.wps文件"
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(file_path))
            doc.Close()
            word.Quit()
            return True, "WPS格式验证通过"
        except Exception as e:
            return False, f"WPS格式验证失败：{str(e)}"
    
    def format_document(self, input_path: str, output_path: str,
                       backup_path: Optional[str] = None) -> Optional[str]:
        """格式化WPS文档。
        
        处理流程：
        1. 备份原始.wps文件
        2. 使用Word/WPS将.wps转换为临时.docx
        3. 使用DOCX处理器格式化临时.docx
        4. 将格式化后的.docx另存为.wps（覆盖原文件）
        5. 删除临时文件
        """
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"输入文件不存在：{input_path}")
        
        if not self._win32_available:
            raise ImportError("需要安装pywin32库：pip install pywin32")
        
        import win32com.client
        
        file_dir = os.path.dirname(input_path)
        file_name = os.path.splitext(os.path.basename(input_path))[0]
        temp_docx = os.path.join(file_dir, f"{file_name}_temp.docx")
        
        actual_backup_path = None
        if backup_path:
            from .docx_handler import get_unique_backup_path
            actual_backup_path = get_unique_backup_path(backup_path)
            try:
                import shutil
                shutil.copy2(input_path, actual_backup_path)
                logger.info(f"已备份原文件：{actual_backup_path}")
            except OSError as e:
                raise OSError(f"备份文件失败：{e}")
        
        word = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            
            doc = word.Documents.Open(os.path.abspath(input_path))
            doc.SaveAs2(os.path.abspath(temp_docx), FileFormat=16)
            doc.Close()
            
            docx_handler = DocxFormatHandler()
            docx_handler.format_document(temp_docx, temp_docx)
            
            doc = word.Documents.Open(os.path.abspath(temp_docx))
            doc.SaveAs2(os.path.abspath(input_path), FileFormat=0)
            doc.Close()
            
            logger.info(f"WPS格式化完成：{input_path}")
            
        except Exception as e:
            if actual_backup_path and os.path.exists(actual_backup_path):
                try:
                    import shutil
                    shutil.copy2(actual_backup_path, input_path)
                    logger.info(f"已从备份恢复：{actual_backup_path}")
                except OSError:
                    logger.error(f"恢复备份失败：{actual_backup_path}")
            raise ValueError(f"WPS文档处理失败：{e}")
        finally:
            if word:
                word.Quit()
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except OSError:
                    pass
        
        return actual_backup_path
