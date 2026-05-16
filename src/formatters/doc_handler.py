"""
公文格式化工具 - DOC格式处理器
原生支持Microsoft Word .doc格式文档
"""

import os
import logging
from typing import Optional, Tuple

from .base_handler import FormatHandler
from .docx_handler import DocxFormatHandler

logger = logging.getLogger(__name__)

# Word文件格式常量
WD_FORMAT_DOC = 0      # .doc格式
WD_FORMAT_DOCX = 16    # .docx格式


class DocFormatHandler(FormatHandler):
    """DOC格式处理器。
    
    原生支持Microsoft Word .doc格式，通过COM接口直接处理。
    需要安装pywin32库和Microsoft Word或WPS Office。
    """
    
    def __init__(self):
        self._docx_handler = DocxFormatHandler()
        self._win32_available = self._check_win32()
        self._word_app = None
    
    def _check_win32(self) -> bool:
        """检查win32com是否可用。"""
        try:
            import win32com.client
            return True
        except ImportError:
            return False
    
    def _get_word_app(self):
        """获取Word应用程序实例，优先使用WPS，其次使用Word。"""
        if self._word_app is not None:
            return self._word_app
        
        import win32com.client
        
        # 尝试连接WPS Office
        try:
            self._word_app = win32com.client.Dispatch("KWPS.Application")
            logger.info("使用WPS Office处理文档")
            return self._word_app
        except Exception:
            pass
        
        # 尝试连接Microsoft Word
        try:
            self._word_app = win32com.client.Dispatch("Word.Application")
            logger.info("使用Microsoft Word处理文档")
            return self._word_app
        except Exception as e:
            raise ImportError(f"需要安装Microsoft Word或WPS Office才能处理.doc文件：{e}")
    
    @property
    def supported_extensions(self) -> list:
        return ['.doc']
    
    def can_handle(self, file_path: str) -> bool:
        return file_path.lower().endswith('.doc') and self._win32_available
    
    def validate_format(self, file_path: str) -> Tuple[bool, str]:
        """验证DOC文件格式。"""
        if not os.path.exists(file_path):
            return False, f"文件不存在：{file_path}"
        if not self._win32_available:
            return False, "需要安装pywin32库：pip install pywin32"
        try:
            word = self._get_word_app()
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(file_path))
            doc.Close()
            return True, "DOC格式验证通过"
        except Exception as e:
            return False, f"DOC格式验证失败：{str(e)}"
    
    def format_document(self, input_path: str, output_path: str,
                       backup_path: Optional[str] = None) -> Optional[str]:
        """格式化DOC文档。
        
        处理流程：
        1. 备份原始.doc文件
        2. 使用Word/WPS将.doc转换为临时.docx
        3. 使用DOCX处理器格式化临时.docx
        4. 将格式化后的.docx另存为.doc（覆盖原文件）
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
            word = self._get_word_app()
            word.Visible = False
            word.DisplayAlerts = 0
            
            # 打开.doc文件并转换为.docx
            doc = word.Documents.Open(os.path.abspath(input_path))
            doc.SaveAs2(os.path.abspath(temp_docx), FileFormat=WD_FORMAT_DOCX)
            doc.Close()
            
            # 使用DOCX处理器格式化
            docx_handler = DocxFormatHandler()
            docx_handler.format_document(temp_docx, temp_docx)
            
            # 将格式化后的.docx另存为.doc
            doc = word.Documents.Open(os.path.abspath(temp_docx))
            doc.SaveAs2(os.path.abspath(input_path), FileFormat=WD_FORMAT_DOC)
            doc.Close()
            
            logger.info(f"DOC格式化完成：{input_path}")
            
        except Exception as e:
            if actual_backup_path and os.path.exists(actual_backup_path):
                try:
                    import shutil
                    shutil.copy2(actual_backup_path, input_path)
                    logger.info(f"已从备份恢复：{actual_backup_path}")
                except OSError:
                    logger.error(f"恢复备份失败：{actual_backup_path}")
            raise ValueError(f"DOC文档处理失败：{e}")
        finally:
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except OSError:
                    pass
        
        return actual_backup_path
