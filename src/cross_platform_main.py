#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公文自动格式化工具 - 跨平台启动器
支持 Windows, Linux, macOS

功能特性：
- Windows: 完整支持 .docx, .doc, .wps 格式
- Linux/macOS: 支持 .docx 格式，.doc/.wps 友好提示
- 自动检测运行环境，提供对应功能说明
"""

import os
import sys
import platform

# 跨平台路径处理
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from formatters.router import get_router


def detect_system() -> dict:
    """检测当前系统环境信息。
    
    Returns:
        系统信息字典
    """
    system = platform.system()
    machine = platform.machine()
    release = platform.release()
    
    # 支持的格式
    if system == 'Windows':
        supported = ['.docx', '.doc', '.wps']
        features = '完整功能：.docx, .doc, .wps'
    else:
        supported = ['.docx']
        features = '跨平台兼容：仅支持.docx格式（.doc/.wps需Windows环境）'
    
    return {
        'system': system,
        'machine': machine,
        'release': release,
        'supported_formats': supported,
        'features': features,
        'is_windows': system == 'Windows',
        'is_linux': system == 'Linux',
        'is_macos': system == 'Darwin',
    }


def print_welcome(system_info: dict) -> None:
    """打印欢迎信息，根据系统显示不同内容。"""
    print("=" * 55)
    print("  公文自动格式化工具 v5.1.0 (跨平台版)")
    print(f"  系统: {system_info['system']} {system_info['machine']}")
    print(f"  {system_info['features']}")
    print("=" * 55)
    print()
    
    if not system_info['is_windows']:
        print("📝 Linux/macOS 环境提示:")
        print("   - ✅ 完整支持 .docx 格式")
        print("   - ⚠️  .doc/.wps 格式依赖 Windows COM 组件")
        print("   - 💡 如需处理 .doc/.wps，可先在 Windows 下转换为 .docx")
        print()


def main():
    """跨平台主函数入口。"""
    import logging
    
    # 配置日志
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    logger = logging.getLogger(__name__)
    
    # 系统检测
    system_info = detect_system()
    print_welcome(system_info)
    
    # 获取路由器
    try:
        router = get_router()
    except Exception as e:
        logger.exception("无法初始化格式处理器")
        print(f"\n❌ 初始化失败: {e}")
        input("按回车键退出...")
        return
    
    supported = router.get_supported_extensions()
    
    if len(sys.argv) < 2:
        print(f"使用方法：将文档文件拖到此图标上，或命令行：")
        print(f"  python {os.path.basename(__file__)} 文件1.docx ...")
        print(f"支持格式：{', '.join(supported)}")
        print()
        input("按回车键退出...")
        return
    
    success = 0
    fail = 0
    total = len(sys.argv) - 1
    
    for i, input_file in enumerate(sys.argv[1:], 1):
        print(f"正在处理第 {i}/{total} 个文件：{os.path.basename(input_file)}")
        
        # 检查扩展名支持
        ext = os.path.splitext(input_file)[1].lower()
        if ext not in supported:
            if system_info['is_windows']:
                print(f"  ❌ 不支持的文件格式：{ext}")
            else:
                print(f"  ❌ 当前系统({system_info['system']})不支持：{ext}")
                print(f"     提示：如需处理此格式，请在Windows系统运行本工具")
            fail += 1
            continue
        
        file_dir = os.path.dirname(input_file)
        name, ext = os.path.splitext(os.path.basename(input_file))
        backup_file = os.path.join(file_dir, f"{name}-未修改{ext}")
        output_file = input_file
        
        try:
            actual_backup = router.format_file(input_file, output_file, backup_file)
            print(f"  ✅ 格式化完成！")
            print(f"     📁 原文件已替换为格式化版本")
            if actual_backup:
                print(f"     💾 备份文件：{os.path.basename(actual_backup)}")
            success += 1
        except FileNotFoundError as e:
            print(f"  ❌ 文件不存在：{str(e)}")
            fail += 1
        except ValueError as e:
            print(f"  ❌ 格式错误：{str(e)}")
            fail += 1
        except OSError as e:
            print(f"  ❌ 系统错误：{str(e)}")
            fail += 1
        except ImportError as e:
            print(f"  ❌ 依赖缺失：{str(e)}")
            if not system_info['is_windows']:
                print(f"     提示：该功能需要Windows环境支持")
            fail += 1
        except Exception as e:
            logger.exception(f"处理失败：{input_file}")
            print(f"  ❌ 未知错误：{str(e)}")
            fail += 1
    
    print()
    print("=" * 55)
    print(f"处理完成！成功 {success} 个，失败 {fail} 个")
    print("=" * 55)
    input("按回车键退出...")


if __name__ == "__main__":
    main()
