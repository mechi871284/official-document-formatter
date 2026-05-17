"""
公文格式化工具 - DOC格式处理器（跨平台版）

Windows: 使用 COM 接口 (win32com + Word/WPS)
Linux/macOS: 使用 LibreOffice 命令行转换

支持格式: Microsoft Word .doc
"""

import os
import logging
from typing import Optional, Tuple

from .base_handler import FormatHandler
from .docx_handler import DocxFormatHandler
from .cross_platform import (
    get_conversion_engine,
    PlatformInfo,
    IS_WINDOWS,
    IS_LINUX,
    IS_MACOS,
)

logger = logging.getLogger(__name__)

# Word文件格式常量
WD_FORMAT_DOC = 0      # .doc格式
WD_FORMAT_DOCX = 16    # .docx格式


class DocFormatHandler(FormatHandler):
    """DOC格式处理器（跨平台）。
    
    Windows: 原生支持Microsoft Word .doc格式，通过COM接口直接处理。
    Linux/macOS: 使用LibreOffice转换为.docx，处理后转回.doc。
    """
    
    def __init__(self):
        self._docx_handler = DocxFormatHandler()
        self._engine = get_conversion_engine()
        
        if IS_WINDOWS:
            self._win32_available = self._check_win32()
            self._word_app = None
        else:
            self._win32_available = False
            self._word_app = None
    
    def _check_win32(self) -> bool:
        """Windows-only: 检查win32com是否可用。"""
        try:
            import win32com.client
            return True
        except ImportError:
            return False
    
    def _get_word_app(self):
        """Windows-only: 获取Word应用程序实例。"""
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
    
    def _get_conversion_method(self) -> str:
        """获取当前转换方式。"""
        if IS_WINDOWS and self._win32_available:
            return 'COM (Windows Native)'
        elif self._engine.is_libreoffice_available():
            return 'LibreOffice (Cross-platform)'
        else:
            return 'None (Dependencies missing)'
    
    @property
    def supported_extensions(self) -> list:
        return ['.doc']
    
    def can_handle(self, file_path: str) -> bool:
        """检查是否可以处理指定文件。"""
        if not file_path.lower().endswith('.doc'):
            return False
        if IS_WINDOWS:
            return self._win32_available
        # Linux/macOS: 检查LibreOffice是否可用
        return self._engine.is_libreoffice_available()
    
    def validate_format(self, file_path: str) -> Tuple[bool, str]:
        """验证DOC文件格式。"""
        if not os.path.exists(file_path):
            return False, f"文件不存在：{file_path}"
        
        method = self._get_conversion_method()
        
        if IS_WINDOWS:
            if not self._win32_available:
                return False, "需要安装pywin32库：pip install pywin32"
            try:
                word = self._get_word_app()
                word.Visible = False
                doc = word.Documents.Open(os.path.abspath(file_path))
                doc.Close()
                return True, f"DOC格式验证通过（使用{method}）"
            except Exception as e:
                return False, f"DOC格式验证失败：{str(e)}"
        else:
            # Linux/macOS: 检查LibreOffice
            if not self._engine.is_libreoffice_available():
                hint = self._engine.get_install_hint()
                return False, f"需要安装LibreOffice才能处理.doc格式\n{hint}"
            return True, f"DOC格式支持（使用{method}）"
    
    def format_document(self, input_path: str, output_path: str,
                       backup_path: Optional[str] = None) -> Optional[str]:
        """格式化DOC文档（跨平台实现）。"""
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"输入文件不存在：{input_path}")
        
        # 根据平台选择处理方式
        if IS_WINDOWS and self._win32_available:
            return self._format_with_com(input_path, output_path, backup_path)
        else:
            return self._format_with_libreoffice(input_path, output_path, backup_path)
    
    def _format_with_com(self, input_path: str, output_path: str,
                         backup_path: Optional[str] = None) -> Optional[str]:
        """Windows-only: 使用COM接口处理。"""
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
            raise ValueError(f"DOC文档处理失败(COM)：{e}")
        finally:
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except OSError:
                    pass
        
        return actual_backup_path
    
    def _format_with_libreoffice(self, input_path: str, output_path: str,
                                  backup_path: Optional[str] = None) -> Optional[str]:
        """跨平台: 使用LibreOffice转换处理。"""
        if not self._engine.is_libreoffice_available():
            hint = self._engine.get_install_hint()
            raise ImportError(f"需要安装LibreOffice才能处理.doc格式\n{hint}")
        
        file_dir = os.path.dirname(input_path)
        file_name = os.path.splitext(os.path.basename(input_path))[0]
        
        # 创建临时目录
        import tempfile
        with tempfile.TemporaryDirectory(prefix='doc_conv_') as temp_dir:
            
            temp_docx = os.path.join(temp_dir, f"{file_name}.docx")
            actual_backup_path = None
            
            # 备份原始文件
            if backup_path:
                from .docx_handler import get_unique_backup_path
                actual_backup_path = get_unique_backup_path(backup_path)
                try:
                    import shutil
                    shutil.copy2(input_path, actual_backup_path)
                    logger.info(f"已备份原文件：{actual_backup_path}")
                except OSError as e:
                    raise OSError(f"备份文件失败：{e}")
            
            try:
                # 步骤1: .doc -> .docx 转换
                logger.info(f"LibreOffice: .doc -> .docx ({input_path})")
                converted_docx = self._engine.libreoffice_convert(
                    input_path, 'docx', temp_dir
                )
                
                if not converted_docx or not os.path.exists(converted_docx):
                    raise ValueError("LibreOffice转换失败(.doc->.docx)")
                
                logger.info(f"转换成功: {converted_docx}")
                
                # 步骤2: 使用DOCX处理器格式化
                docx_handler = DocxFormatHandler()
                docx_handler.format_document(converted_docx, converted_docx)
                
                # 步骤3: .docx -> .doc 转换（覆盖原文件）
                logger.info(f"LibreOffice: .docx -> .doc ({converted_docx})")
                converted_doc = self._engine.libreoffice_convert(
                    converted_docx, 'doc', temp_dir
                )
                
                if not converted_doc or not os.path.exists(converted_doc):
                    logger.warning("LibreOffice转换回.doc失败，直接使用.docx...")
                    # 如果转回.doc失败，尝试用LibreOffice再次转换
                    converted_doc2 = self._engine.libreoffice_convert(
                        converted_docx, 'doc', file_dir
                    )
                    if converted_doc2 and os.path.exists(converted_doc2):
                        converted_doc = converted_doc2
                    else:
                        # 最终方案：直接复制docx为doc（Word兼容打开）
                        final_doc = os.path.join(file_dir, f"{file_name}_final.doc")
                        shutil.copy2(converted_docx, final_doc)
                        converted_doc = final_doc
                
                # 复制结果覆盖原文件
                if os.path.exists(converted_doc):
                    import shutil
                    shutil.copy2(converted_doc, input_path)
                    logger.info(f"DOC格式化完成：{input_path}")
                else:
                    raise ValueError("转换后文件不存在")
                
            except Exception as e:
                if actual_backup_path and os.path.exists(actual_backup_path):
                    try:
                        import shutil
                        shutil.copy2(actual_backup_path, input_path)
                        logger.info(f"已从备份恢复：{actual_backup_path}")
                    except OSError:
                        logger.error(f"恢复备份失败：{actual_backup_path}")
                raise ValueError(f"DOC文档处理失败(LibreOffice)：{e}")
            
            return actual_backup_path
