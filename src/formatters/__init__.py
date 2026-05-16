# 公文格式化工具 - 格式处理器

from .base_handler import FormatHandler
from .docx_handler import DocxFormatHandler
from .doc_handler import DocFormatHandler
from .wps_handler import WpsFormatHandler
from .router import FormatRouter, get_router

__all__ = [
    'FormatHandler',
    'DocxFormatHandler',
    'DocFormatHandler',
    'WpsFormatHandler',
    'FormatRouter',
    'get_router',
]
