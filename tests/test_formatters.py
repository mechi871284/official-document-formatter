"""
公文格式化工具 - 多格式处理器测试
"""

import os
import pytest
from unittest.mock import patch, MagicMock

from src.formatters.base_handler import FormatHandler
from src.formatters.docx_handler import DocxFormatHandler
from src.formatters.doc_handler import DocFormatHandler
from src.formatters.wps_handler import WpsFormatHandler
from src.formatters.router import FormatRouter, get_router


class TestDocxFormatHandler:
    """DOCX格式处理器测试。"""
    
    def setup_method(self):
        self.handler = DocxFormatHandler()
    
    def test_supported_extensions(self):
        assert self.handler.supported_extensions == ['.docx']
    
    def test_can_handle_docx(self):
        assert self.handler.can_handle('test.docx') is True
        assert self.handler.can_handle('test.DOCX') is True
        assert self.handler.can_handle('test.doc') is False
        assert self.handler.can_handle('test.wps') is False
    
    def test_validate_format_not_exists(self):
        is_valid, msg = self.handler.validate_format('nonexistent.docx')
        assert is_valid is False
        assert '不存在' in msg
    
    def test_validate_format_invalid(self, tmp_path):
        invalid_file = tmp_path / 'invalid.docx'
        invalid_file.write_text('not a docx file')
        is_valid, msg = self.handler.validate_format(str(invalid_file))
        assert is_valid is False
    
    def test_format_document_not_exists(self):
        with pytest.raises(FileNotFoundError):
            self.handler.format_document('nonexistent.docx', 'output.docx')
    
    def test_get_format_info(self):
        info = self.handler.get_format_info()
        assert info['extensions'] == ['.docx']
        assert info['handler_name'] == 'DocxFormatHandler'


class TestDocFormatHandler:
    """DOC格式处理器测试。"""
    
    def setup_method(self):
        self.handler = DocFormatHandler()
    
    def test_supported_extensions(self):
        assert self.handler.supported_extensions == ['.doc']
    
    def test_can_handle_doc(self):
        assert self.handler.can_handle('test.doc') is True
        assert self.handler.can_handle('test.DOC') is True
        assert self.handler.can_handle('test.docx') is False
    
    def test_validate_format_not_exists(self):
        is_valid, msg = self.handler.validate_format('nonexistent.doc')
        assert is_valid is False
        assert '不存在' in msg
    
    def test_format_document_not_exists(self):
        with pytest.raises(FileNotFoundError):
            self.handler.format_document('nonexistent.doc', 'output.doc')


class TestWpsFormatHandler:
    """WPS格式处理器测试。"""
    
    def setup_method(self):
        self.handler = WpsFormatHandler()
    
    def test_supported_extensions(self):
        assert self.handler.supported_extensions == ['.wps']
    
    def test_can_handle_wps(self):
        assert self.handler.can_handle('test.wps') is True
        assert self.handler.can_handle('test.WPS') is True
        assert self.handler.can_handle('test.docx') is False
    
    def test_validate_format_not_exists(self):
        is_valid, msg = self.handler.validate_format('nonexistent.wps')
        assert is_valid is False
        assert '不存在' in msg
    
    def test_format_document_not_exists(self):
        with pytest.raises(FileNotFoundError):
            self.handler.format_document('nonexistent.wps', 'output.wps')


class TestFormatRouter:
    """格式路由器测试。"""
    
    def setup_method(self):
        self.router = FormatRouter()
    
    def test_get_handler_docx(self):
        handler = self.router.get_handler('test.docx')
        assert isinstance(handler, DocxFormatHandler)
    
    def test_get_handler_doc(self):
        handler = self.router.get_handler('test.doc')
        if handler is not None:
            assert isinstance(handler, DocFormatHandler)
    
    def test_get_handler_wps(self):
        handler = self.router.get_handler('test.wps')
        if handler is not None:
            assert isinstance(handler, WpsFormatHandler)
    
    def test_get_handler_unsupported(self):
        handler = self.router.get_handler('test.pdf')
        assert handler is None
    
    def test_get_supported_extensions(self):
        extensions = self.router.get_supported_extensions()
        assert '.docx' in extensions
    
    def test_validate_file_not_exists(self):
        is_valid, msg, handler = self.router.validate_file('nonexistent.docx')
        assert is_valid is False
        assert '不存在' in msg
        assert handler is None
    
    def test_validate_file_unsupported(self, tmp_path):
        unsupported_file = tmp_path / 'test.pdf'
        unsupported_file.touch()
        is_valid, msg, handler = self.router.validate_file(str(unsupported_file))
        assert is_valid is False
        assert '不支持' in msg
        assert handler is None
    
    def test_format_file_unsupported(self, tmp_path):
        unsupported_file = tmp_path / 'test.pdf'
        unsupported_file.touch()
        with pytest.raises(ValueError) as exc_info:
            self.router.format_file(str(unsupported_file), 'output.pdf')
        assert '不支持' in str(exc_info.value)
    
    def test_get_handler_info(self):
        info_list = self.router.get_handler_info()
        assert len(info_list) >= 1
        for info in info_list:
            assert 'extensions' in info
            assert 'handler_name' in info
    
    def test_get_router_singleton(self):
        router1 = get_router()
        router2 = get_router()
        assert router1 is router2


class TestFormatHandlerInterface:
    """格式处理器接口测试。"""
    
    def test_abstract_methods(self):
        """测试抽象方法必须实现。"""
        with pytest.raises(TypeError):
            FormatHandler()


class TestDocxFormatting:
    """DOCX格式化功能测试。"""
    
    def test_is_fullwidth_chinese(self):
        from src.formatters.docx_handler import is_fullwidth
        assert is_fullwidth('中') is True
        assert is_fullwidth('文') is True
    
    def test_is_fullwidth_ascii(self):
        from src.formatters.docx_handler import is_fullwidth
        assert is_fullwidth('a') is False
        assert is_fullwidth('1') is False
    
    def test_is_fullwidth_punctuation(self):
        from src.formatters.docx_handler import is_fullwidth
        assert is_fullwidth('，') is True
        assert is_fullwidth('。') is True
    
    def test_char_width_fullwidth(self):
        from src.formatters.docx_handler import char_width
        assert char_width('中') == 1.0
    
    def test_char_width_ascii(self):
        from src.formatters.docx_handler import char_width
        assert char_width('a') == 0.5
    
    def test_pt_from_chars(self):
        from src.formatters.docx_handler import pt_from_chars
        from docx.shared import Pt
        result = pt_from_chars(16, 2)
        assert result == Pt(32)
