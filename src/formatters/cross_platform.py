#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公文格式化工具 - 平台抽象层
提供跨平台一致的API接口，自动适配 Windows/Linux/macOS

Windows: 使用 COM 接口 (win32com)
Linux/macOS: 使用 LibreOffice 命令行转换
"""

import os
import sys
import platform
import subprocess
import shutil
import logging
from typing import Optional, Tuple
from pathlib import Path

logger = logging.getLogger(__name__)

# 系统检测
SYSTEM = platform.system()
IS_WINDOWS = SYSTEM == 'Windows'
IS_LINUX = SYSTEM == 'Linux'
IS_MACOS = SYSTEM == 'Darwin'


class PlatformInfo:
    """平台信息查询。"""
    
    @staticmethod
    def name() -> str:
        return SYSTEM
    
    @staticmethod
    def is_windows() -> bool:
        return IS_WINDOWS
    
    @staticmethod
    def is_linux() -> bool:
        return IS_LINUX
    
    @staticmethod
    def is_macos() -> bool:
        return IS_MACOS
    
    @staticmethod
    def shell_quote(path: str) -> str:
        """Shell路径引号处理。"""
        if IS_WINDOWS:
            return f'"{path}"' if ' ' in path else path
        return path.replace('"', '\\"')
    
    @staticmethod
    def get_temp_dir() -> str:
        """获取临时目录。"""
        import tempfile
        return tempfile.gettempdir()


class LibreOfficeEngine:
    """LibreOffice转换引擎（Linux/macOS）。
    
    替代Windows下的COM接口，通过命令行调用LibreOffice实现格式转换。
    """
    
    # 支持的输入输出格式映射
    FORMAT_MAP = {
        '.doc': 'doc',
        '.docx': 'docx',
        '.wps': 'doc',
        '.odt': 'odt',
        '.rtf': 'rtf',
        '.txt': 'txt',
    }
    
    def __init__(self):
        self._libreoffice_path = self._find_libreoffice()
        self._available = self._libreoffice_path is not None
    
    def _find_libreoffice(self) -> Optional[str]:
        """查找LibreOffice可执行文件路径。"""
        candidates = []
        
        if IS_LINUX:
            candidates = [
                'libreoffice',
                'libreoffice7.0',
                'libreoffice7.1',
                'libreoffice7.2',
                'libreoffice7.3',
                'libreoffice7.4',
                'libreoffice7.5',
                '/usr/bin/libreoffice',
                '/usr/local/bin/libreoffice',
                '/opt/libreoffice/program/soffice',
            ]
        elif IS_MACOS:
            candidates = [
                '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                'libreoffice',
            ]
        
        for candidate in candidates:
            try:
                if os.path.isabs(candidate):
                    if os.path.exists(candidate) and os.access(candidate, os.X_OK):
                        return candidate
                else:
                    result = shutil.which(candidate)
                    if result:
                        return result
            except:
                continue
        
        return None
    
    def is_available(self) -> bool:
        """检查LibreOffice是否可用。"""
        return self._available
    
    def get_version(self) -> Optional[str]:
        """获取LibreOffice版本。"""
        if not self._available:
            return None
        
        try:
            result = subprocess.run(
                [self._libreoffice_path, '--version'],
                capture_output=True,
                text=True,
                timeout=5
            )
            return result.stdout.strip() or result.stderr.strip()
        except:
            return None
    
    def convert(self, input_path: str, output_format: str, 
                output_dir: Optional[str] = None) -> Optional[str]:
        """转换文档格式。
        
        Args:
            input_path: 输入文件路径
            output_format: 目标格式扩展名（如'docx', 'doc'）
            output_dir: 输出目录（默认与输入文件同目录）
            
        Returns:
            转换后的文件路径，失败返回None
        """
        if not self._available:
            raise RuntimeError("LibreOffice not found. Please install LibreOffice first.")
        
        input_path = os.path.abspath(input_path)
        
        if output_dir is None:
            output_dir = os.path.dirname(input_path) or os.getcwd()
        
        output_dir = os.path.abspath(output_dir)
        
        # 构建输出文件名
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_file = os.path.join(output_dir, f"{base_name}.{output_format.lower()}")
        
        # 如果输出文件已存在，先删除
        if os.path.exists(output_file):
            try:
                os.remove(output_file)
            except:
                pass
        
        try:
            # 调用LibreOffice转换
            cmd = [
                self._libreoffice_path,
                '--headless',
                '--invisible',
                '--nocrashreport',
                '--nodefault',
                '--nofirststartwizard',
                '--nolockcheck',
                '--nologo',
                '--norestore',
                f'--convert-to', output_format.lower(),
                f'--outdir', output_dir,
                input_path
            ]
            
            logger.debug(f"Running: {' '.join(cmd)}")
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60  # 最长60秒超时
            )
            
            if result.returncode != 0:
                logger.warning(f"LibreOffice return code: {result.returncode}")
                logger.debug(f"stdout: {result.stdout}")
                logger.debug(f"stderr: {result.stderr}")
            
            # 检查输出文件是否生成
            if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
                return output_file
            
            # 尝试查找可能生成的文件
            for f in os.listdir(output_dir):
                if f.lower().startswith(base_name.lower()) and f.endswith('.' + output_format.lower()):
                    found = os.path.join(output_dir, f)
                    if os.path.getsize(found) > 0:
                        logger.info(f"Found converted file: {f}")
                        return found
            
            return None
            
        except subprocess.TimeoutExpired:
            logger.error("LibreOffice conversion timeout")
            return None
        except Exception as e:
            logger.exception(f"Conversion error: {e}")
            return None


class ConversionEngine:
    """统一转换引擎接口。
    
    自动选择最佳转换引擎：
    - Windows: COM (优先)
    - Linux/macOS: LibreOffice
    """
    
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._init()
        return cls._instance
    
    def _init(self):
        if IS_WINDOWS:
            self._engine_type = 'com'
            # COM初始化在具体处理器中完成
        else:
            self._engine_type = 'libreoffice'
            self._libreoffice = LibreOfficeEngine()
    
    def engine_type(self) -> str:
        """返回当前使用的引擎类型。"""
        return self._engine_type
    
    def is_libreoffice_available(self) -> bool:
        """检查LibreOffice是否可用。"""
        if hasattr(self, '_libreoffice'):
            return self._libreoffice.is_available()
        return False
    
    def get_libreoffice_version(self) -> Optional[str]:
        """获取LibreOffice版本。"""
        if hasattr(self, '_libreoffice'):
            return self._libreoffice.get_version()
        return None
    
    def libreoffice_convert(self, input_path: str, output_format: str, 
                           output_dir: Optional[str] = None) -> Optional[str]:
        """使用LibreOffice转换格式。"""
        if not hasattr(self, '_libreoffice'):
            self._libreoffice = LibreOfficeEngine()
        return self._libreoffice.convert(input_path, output_format, output_dir)
    
    def get_install_hint(self) -> str:
        """获取依赖安装提示。"""
        if IS_LINUX:
            return """
  Ubuntu/Debian: sudo apt install libreoffice libreoffice-writer
  RHEL/CentOS: sudo dnf install libreoffice libreoffice-writer
  Fedora: sudo dnf install libreoffice libreoffice-writer
"""
        elif IS_MACOS:
            return """
  1. 从 https://www.libreoffice.org/download/download-libreoffice/ 下载
  2. 安装 LibreOffice.app 到 /Applications
"""
        else:
            return "Windows系统请安装Microsoft Word或WPS Office"


def get_conversion_engine() -> ConversionEngine:
    """获取全局转换引擎实例。"""
    return ConversionEngine()


# 便捷函数
def print_system_info():
    """打印系统信息。"""
    print("=" * 50)
    print(f"  System: {SYSTEM} {platform.machine()}")
    print(f"  Python: {platform.python_version()}")
    engine = get_conversion_engine()
    print(f"  Engine: {engine.engine_type()}")
    if not IS_WINDOWS:
        if engine.is_libreoffice_available():
            print(f"  LibreOffice: {engine.get_libreoffice_version() or 'Installed'}")
        else:
            print(f"  LibreOffice: ❌ Not found")
            print(engine.get_install_hint())
    print("=" * 50)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    print_system_info()
