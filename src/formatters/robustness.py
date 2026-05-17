"""
公文格式化工具 - 健壮性增强模块
提供参数校验、异常处理、边界条件检测等功能
"""

import os
import re
from typing import Optional, Any, Callable, TypeVar, Tuple
from functools import wraps
from pathlib import Path

# 类型变量
T = TypeVar('T')

# 文件大小限制（默认100MB）
MAX_FILE_SIZE = 100 * 1024 * 1024

# 有效扩展名
VALID_EXTENSIONS = {'.docx', '.doc', '.wps'}


def validate_file_path(file_path: str, 
                       check_exists: bool = True,
                       check_ext: bool = True,
                       max_size: int = MAX_FILE_SIZE) -> Tuple[bool, str]:
    """校验文件路径的合法性。
    
    Args:
        file_path: 文件路径
        check_exists: 是否检查文件存在
        check_ext: 是否检查扩展名
        max_size: 最大文件大小限制
        
    Returns:
        (是否有效, 提示信息)
    """
    if not file_path or not isinstance(file_path, str):
        return False, "文件路径不能为空或无效类型"
    
    file_path = file_path.strip()
    if not file_path:
        return False, "文件路径不能为空字符串"
    
    if check_exists and not os.path.exists(file_path):
        return False, f"文件不存在：{file_path}"
    
    if check_exists and os.path.isdir(file_path):
        return False, f"路径指向目录而非文件：{file_path}"
    
    if check_ext:
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in VALID_EXTENSIONS:
            return False, f"不支持的文件格式：{ext}，支持格式：{', '.join(VALID_EXTENSIONS)}"
    
    if check_exists:
        try:
            size = os.path.getsize(file_path)
            if size > max_size:
                return False, f"文件过大：{size / (1024*1024):.1f}MB，限制：{max_size / (1024*1024):.1f}MB"
            if size == 0:
                return False, "文件为空（0字节）"
        except OSError as e:
            return False, f"无法访问文件：{str(e)}"
    
    return True, "文件路径有效"


def validate_paragraphs(paragraphs: Any) -> Tuple[bool, str]:
    """校验段落列表的合法性。
    
    Args:
        paragraphs: 段落列表
        
    Returns:
        (是否有效, 提示信息)
    """
    if paragraphs is None:
        return False, "段落列表不能为None"
    
    try:
        length = len(paragraphs)
        if length == 0:
            return True, "段落列表为空（有效）"
        return True, f"段落列表有效，共{length}段"
    except TypeError:
        return False, "段落列表不支持len()操作"


def safe_regex_match(pattern: re.Pattern, text: Optional[str], 
                      default: Any = None) -> Any:
    """安全的正则表达式匹配，避免None或异常。
    
    Args:
        pattern: 编译后的正则表达式
        text: 待匹配文本
        default: 默认返回值
        
    Returns:
        匹配结果或default
    """
    if text is None:
        return default
    if not isinstance(text, str):
        return default
    try:
        return pattern.match(text)
    except (re.error, TypeError):
        return default


def timeout_handler(timeout_seconds: int = 30, default_return: Any = None):
    """超时处理装饰器（Windows下使用线程方式）。
    
    Args:
        timeout_seconds: 超时秒数
        default_return: 超时时返回的默认值
    """
    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @wraps(func)
        def wrapper(*args, **kwargs) -> T:
            import threading
            result_container = []
            exception_container = []
            
            def target():
                try:
                    result_container.append(func(*args, **kwargs))
                except Exception as e:
                    exception_container.append(e)
            
            thread = threading.Thread(target=target, daemon=True)
            thread.start()
            thread.join(timeout=timeout_seconds)
            
            if thread.is_alive():
                raise TimeoutError(f"函数执行超时（{timeout_seconds}秒）")
            
            if exception_container:
                raise exception_container[0]
            
            return result_container[0] if result_container else default_return
        return wrapper
    return decorator


def robust_operation(operation_name: str, 
                     default_return: Any = None,
                     logger: Any = None):
    """健壮操作装饰器：捕获并记录所有异常。
    
    Args:
        operation_name: 操作名称（用于日志）
        default_return: 异常时返回的默认值
        logger: 日志记录器实例
    """
    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @wraps(func)
        def wrapper(*args, **kwargs) -> T:
            try:
                return func(*args, **kwargs)
            except Exception as e:
                error_msg = f"{operation_name}失败: {type(e).__name__}: {str(e)}"
                if logger:
                    logger.exception(error_msg)
                else:
                    import logging
                    logging.exception(error_msg)
                return default_return
        return wrapper
    return decorator


class FuzzyMatcher:
    """模糊匹配器，增强文本识别能力。
    
    支持：
    - 全角/半角字符归一化
    - 多余空白字符忽略
    - 常见变体字符匹配（如"："与":"，"（"与"("等）
    """
    
    # 全角转半角映射（标点符号）
    FULL2HALF_PUNCT = {
        '：': ':', '；': ';', '，': ',', '。': '.',
        '！': '!', '？': '?', '（': '(', '）': ')',
        '【': '[', '】': ']', '「': '[', '」': ']',
        '『': '[', '』': ']', '《': '<', '》': '>',
        '\u3000': ' ', '\u00A0': ' ',
    }
    
    # 常见数字/序号模式增强
    DIGIT_VARIANTS = (
        (r'[①②③④⑤⑥⑦⑧⑨⑩]', lambda m: str('①②③④⑤⑥⑦⑧⑨⑩'.index(m.group()) + 1)),
        (r'[⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽]', lambda m: str('⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽'.index(m.group()) + 1)),
        (r'[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]', lambda m: str('ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ'.index(m.group()) + 1)),
    )
    
    @classmethod
    def normalize(cls, text: str) -> str:
        """文本归一化，便于模糊匹配。
        
        Args:
            text: 原始文本
            
        Returns:
            归一化后的文本
        """
        if not text:
            return ''
        
        result = []
        for ch in text:
            # 全角转半角（字母数字）
            if '\uFF01' <= ch <= '\uFF5E':
                result.append(chr(ord(ch) - 0xFEE0))
            # 全角标点转半角
            elif ch in cls.FULL2HALF_PUNCT:
                result.append(cls.FULL2HALF_PUNCT[ch])
            else:
                result.append(ch)
        
        return ''.join(result)
    
    @classmethod
    def normalize_whitespace(cls, text: str) -> str:
        """空白字符归一化。
        
        Args:
            text: 原始文本
            
        Returns:
            空白归一化后的文本
        """
        if not text:
            return ''
        return ' '.join(text.split())
    
    @classmethod
    def fuzzy_match(cls, text: str, patterns: list, 
                    normalize: bool = True) -> Optional[int]:
        """尝试用多个模式模糊匹配文本。
        
        Args:
            text: 待匹配文本
            patterns: 正则表达式模式列表（预编译）
            normalize: 是否先归一化文本
            
        Returns:
            匹配成功的模式索引，失败返回None
        """
        if normalize:
            text = cls.normalize(text)
        
        for idx, pattern in enumerate(patterns):
            if pattern.match(text):
                return idx
        return None


def get_file_safety(path: str) -> dict:
    """获取文件安全信息，用于风险评估。
    
    Args:
        path: 文件路径
        
    Returns:
        安全信息字典
    """
    info = {
        'safe': True,
        'warnings': [],
        'size': 0,
        'extension': '',
        'exists': False,
    }
    
    try:
        path_obj = Path(path)
        info['exists'] = path_obj.exists()
        
        if info['exists']:
            info['size'] = path_obj.stat().st_size
            info['extension'] = path_obj.suffix.lower()
            
            # 检测潜在风险
            if info['size'] > 50 * 1024 * 1024:
                info['warnings'].append('文件较大(>50MB)，处理可能较慢')
            if info['extension'] not in VALID_EXTENSIONS:
                info['warnings'].append(f'非标准扩展名：{info["extension"]}')
            
            info['safe'] = len(info['warnings']) == 0
            
    except Exception as e:
        info['safe'] = False
        info['warnings'].append(f'检测失败: {str(e)}')
    
    return info
