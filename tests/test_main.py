#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""公文格式化工具单元测试"""

import os
import sys
import tempfile
import shutil
import pytest
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from formatters.docx_handler import (
    is_fullwidth,
    char_width,
    pt_from_chars,
    find_title_index,
    get_unique_backup_path,
    RE_H1,
    RE_H2,
    RE_ATTACH,
    RE_MAIN_RECEIVER,
    RE_DATE_PATTERN,
    RE_SIGNATURE_EXCLUDE,
)


class TestIsFullwidth:
    """测试全角字符判断函数"""

    def test_ascii_char(self):
        assert is_fullwidth('a') is False
        assert is_fullwidth('Z') is False
        assert is_fullwidth('0') is False

    def test_chinese_char(self):
        assert is_fullwidth('中') is True
        assert is_fullwidth('文') is True
        assert is_fullwidth('公') is True

    def test_punctuation(self):
        assert is_fullwidth('。') is True
        assert is_fullwidth('，') is True
        assert is_fullwidth('、') is True

    def test_cache_consistency(self):
        result1 = is_fullwidth('中')
        result2 = is_fullwidth('中')
        assert result1 == result2


class TestCharWidth:
    """测试字符宽度计算函数"""

    def test_ascii_width(self):
        assert char_width('a') == 0.5
        assert char_width('Z') == 0.5
        assert char_width('0') == 0.5

    def test_fullwidth_width(self):
        assert char_width('中') == 1.0
        assert char_width('文') == 1.0

    def test_other_width(self):
        assert char_width('。') == 1.0
        assert char_width('，') == 1.0


class TestPtFromChars:
    """测试字符到磅值转换函数"""

    def test_one_char(self):
        result = pt_from_chars(16, 1)
        assert result.pt == 16.0

    def test_two_chars(self):
        result = pt_from_chars(16, 2)
        assert result.pt == 32.0


class TestFindTitleIndex:
    """测试标题索引查找函数"""

    def test_empty_paragraphs_returns_none(self):
        mock_paragraphs = []
        assert find_title_index(mock_paragraphs) is None

    def test_find_centered_title(self):
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt

        mock_p1 = MagicMock()
        mock_p1.text = "  "
        mock_p1.paragraph_format.alignment = None

        mock_p2 = MagicMock()
        mock_p2.text = "关于xxx的通知"
        mock_p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        mock_run = MagicMock()
        mock_run.font.size = Pt(22)
        mock_p2.runs = [mock_run]

        mock_p3 = MagicMock()
        mock_p3.text = "正文内容"
        mock_p3.paragraph_format.alignment = None

        result = find_title_index([mock_p1, mock_p2, mock_p3])
        assert result == 1

    def test_no_title_returns_first_non_empty(self):
        mock_p1 = MagicMock()
        mock_p1.text = "  "
        mock_p1.paragraph_format.alignment = None

        mock_p2 = MagicMock()
        mock_p2.text = "正文内容"
        mock_p2.paragraph_format.alignment = None
        mock_p2.runs = []

        result = find_title_index([mock_p1, mock_p2])
        assert result == 1


class TestGetUniqueBackupPath:
    """测试备份文件路径生成函数"""

    def test_no_existing_file_returns_original(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "test.docx")
            result = get_unique_backup_path(path)
            assert result == path

    def test_existing_file_gets_numbered(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "test-旧.docx")
            with open(path, 'w') as f:
                f.write("test")
            result = get_unique_backup_path(path)
            assert result == os.path.join(tmpdir, "test-旧(1).docx")

    def test_multiple_existing_files_increments(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            base_path = os.path.join(tmpdir, "test-旧.docx")
            with open(base_path, 'w') as f:
                f.write("test")
            with open(os.path.join(tmpdir, "test-旧(1).docx"), 'w') as f:
                f.write("test")
            result = get_unique_backup_path(base_path)
            assert result == os.path.join(tmpdir, "test-旧(2).docx")

    def test_preserves_extension(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "测试-旧.docx")
            with open(path, 'w') as f:
                f.write("test")
            result = get_unique_backup_path(path)
            assert result.endswith('.docx')
            assert '(1)' in result


class TestBackupPath:
    """测试备份功能"""

    def test_format_and_backup(self):
        from docx import Document
        from formatters.docx_handler import DocxFormatHandler

        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, "test.docx")
            output_path = os.path.join(tmpdir, "output.docx")
            backup_path = os.path.join(tmpdir, "test-旧.docx")

            doc = Document()
            doc.add_paragraph("测试内容")
            doc.save(input_path)

            handler = DocxFormatHandler()
            actual_backup = handler.format_document(input_path, output_path, backup_path)

            assert os.path.exists(output_path)
            assert os.path.exists(actual_backup)


class TestRegexPatterns:
    """测试正则表达式模式"""

    def test_h1_pattern(self):
        assert RE_H1.match("一、标题") is not None
        assert RE_H1.match("二、内容") is not None
        assert RE_H1.match("正文内容") is None

    def test_h2_pattern(self):
        assert RE_H2.match("（一）标题") is not None
        assert RE_H2.match("（二）内容") is not None
        assert RE_H2.match("正文内容") is None

    def test_attach_pattern(self):
        assert RE_ATTACH.match("附件：xxx") is not None
        assert RE_ATTACH.match("附件1：xxx") is not None
        assert RE_ATTACH.match("正文内容") is None

    def test_date_pattern(self):
        assert RE_DATE_PATTERN.search("2024年1月1日") is not None
        assert RE_DATE_PATTERN.search("二〇二四年一月一日") is not None

    def test_main_receiver_pattern(self):
        assert RE_MAIN_RECEIVER.match("各有关单位：") is not None
        assert RE_MAIN_RECEIVER.match("正文内容") is None

    def test_signature_exclude_pattern(self):
        assert RE_SIGNATURE_EXCLUDE.match("附件") is not None
        assert RE_SIGNATURE_EXCLUDE.match("抄送") is not None
        assert RE_SIGNATURE_EXCLUDE.match("发文单位") is None


class TestMain:
    """测试主函数"""

    def test_no_args_shows_usage(self, capsys):
        with patch.object(sys, 'argv', ['main.py']):
            with patch('builtins.input', return_value=''):
                from main import main
                main()
                captured = capsys.readouterr()
                assert "公文自动格式化工具" in captured.out
