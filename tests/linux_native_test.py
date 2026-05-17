#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公文格式化工具 - 跨平台完整测试验证脚本

运行测试:
    python3 tests/linux_native_test.py
"""

import os
import sys

# 确保模块可被找到
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
src_dir = os.path.join(parent_dir, 'src')
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)


def test_platform_detection():
    """测试平台检测模块。"""
    print("[测试1] 平台抽象层检测...")
    from formatters.cross_platform import (
        SYSTEM,
        IS_WINDOWS,
        IS_LINUX,
        IS_MACOS,
        PlatformInfo,
        get_conversion_engine,
        print_system_info
    )
    
    print(f"  系统: {SYSTEM}")
    print(f"  IS_WINDOWS: {IS_WINDOWS}")
    print(f"  IS_LINUX: {IS_LINUX}")
    print(f"  IS_MACOS: {IS_MACOS}")
    print(f"  平台名称: {PlatformInfo.name()}")
    
    engine = get_conversion_engine()
    print(f"  转换引擎: {engine.engine_type()}")
    
    if not IS_WINDOWS:
        lo_version = engine.get_libreoffice_version()
        print(f"  LibreOffice: {'✅ 已安装' if engine.is_libreoffice_available() else '❌ 未找到'}")
        if lo_version:
            print(f"    版本: {lo_version}")
    print("  ✅ 通过")
    return True


def test_format_routers():
    """测试格式路由。"""
    print("\n[测试2] 格式处理器检测...")
    from formatters.router import get_router
    
    router = get_router()
    supported = router.get_supported_extensions()
    print(f"  报告支持格式: {supported}")
    
    handlers = router.get_handler_info()
    for h in handlers:
        print(f"  - {h['handler_name']}: {h['extensions']}")
    
    # 实际检测can_handle
    print("\n  实际can_handle检测:")
    for ext in ['.docx', '.doc', '.wps']:
        handler = router.get_handler(f'test{ext}')
        if handler:
            print(f"    {ext}: ✅ {handler.__class__.__name__}")
        else:
            print(f"    {ext}: ❌ 不可用")
    print("  ✅ 通过")
    return True


def test_validation():
    """测试文件验证。"""
    print("\n[测试3] 文件验证...")
    from formatters.router import get_router
    
    router = get_router()
    
    # 测试不存在文件
    is_valid, msg, handler = router.validate_file('/nonexistent/file.docx')
    assert not is_valid, "不存在文件应该验证失败"
    print(f"  不存在文件: {msg}")
    
    # 测试不支持格式
    is_valid, msg, handler = router.validate_file('test.txt')
    assert not is_valid, "不支持格式应该验证失败"
    print(f"  不支持格式: {msg}")
    
    print("  ✅ 通过")
    return True


def test_linux_doc_handler():
    """测试Linux DOC处理器初始化。"""
    print("\n[测试4] 跨平台DOC处理器...")
    from formatters.cross_platform import IS_LINUX, get_conversion_engine
    
    try:
        from formatters.doc_handler import DocFormatHandler
        handler = DocFormatHandler()
        print(f"  处理器初始化: ✅")
        print(f"  支持扩展名: {handler.supported_extensions}")
        
        # 检查实际的转换引擎状态
        from formatters.platform import IS_WINDOWS
        if not IS_WINDOWS:
            engine = get_conversion_engine()
            if engine.is_libreoffice_available():
                can_handle = handler.can_handle('test.doc')
                print(f"  can_handle(.doc): {can_handle}")
            else:
                print(f"  LibreOffice未安装，提示用户安装")
        print("  ✅ 通过")
        return True
    except Exception as e:
        print(f"  ❌ 异常: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_linux_wps_handler():
    """测试Linux WPS处理器初始化。"""
    print("\n[测试5] 跨平台WPS处理器...")
    try:
        from formatters.wps_handler import WpsFormatHandler
        handler = WpsFormatHandler()
        print(f"  处理器初始化: ✅")
        print(f"  支持扩展名: {handler.supported_extensions}")
        print("  ✅ 通过")
        return True
    except Exception as e:
        print(f"  ❌ 异常: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """运行所有测试。"""
    print("=" * 60)
    print("  公文格式化工具 - Linux原生支持完整测试验证")
    print("=" * 60)
    
    all_passed = True
    tests = [
        test_platform_detection,
        test_format_routers,
        test_validation,
        test_linux_doc_handler,
        test_linux_wps_handler,
    ]
    
    for test in tests:
        try:
            if not test():
                all_passed = False
        except Exception as e:
            print(f"\n  ❌ 测试异常: {e}")
            import traceback
            traceback.print_exc()
            all_passed = False
    
    print("\n" + "=" * 60)
    if all_passed:
        print("  ✅ 所有测试通过！")
        print("\n  Linux支持状态:")
        from formatters.platform import get_conversion_engine
        engine = get_conversion_engine()
        if engine.is_libreoffice_available():
            print("    - .docx: ✅")
            print("    - .doc:  ✅ (需要LibreOffice)")
            print("    - .wps:  ✅ (需要LibreOffice)")
        else:
            print("    - .docx: ✅")
            print("    - .doc:  ⚠️  请安装LibreOffice: sudo apt install libreoffice")
            print("    - .wps:  ⚠️  请安装LibreOffice: sudo apt install libreoffice")
    else:
        print("  ❌ 部分测试失败")
    print("=" * 60)
    
    return 0 if all_passed else 1


if __name__ == "__main__":
    sys.exit(main())
