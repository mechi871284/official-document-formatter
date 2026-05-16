#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
公文自动格式化工具（GB/T 9704-2012）
作者：茗记
版本：5.0.0

改进说明：
- v5.0.0: 多格式支持（.docx, .doc, .wps）
- 落款支持手动偏移量校正，解决字体宽度差异导致的居中偏差
- 所有段落强制清除首行缩进干扰
- 落款数字与汉字统一使用仿宋体
- v4.0.0: 全面代码质量优化（类型注解、错误处理、性能提升、安全性增强）
"""

import sys
import os
import logging

# 确保formatters模块可被找到
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# 使用相对导入避免PyInstaller打包问题
from formatters.router import get_router

# ------------------------------ 日志配置 ---------------------------------------
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)


def main():
    """主函数：处理命令行参数并执行格式化。"""
    print("=" * 50)
    print("公文自动格式化工具 v5.0.0")
    print("作者：茗记")
    print("支持格式：.docx, .doc, .wps")
    print("=" * 50)
    print()
    
    if len(sys.argv) < 2:
        router = get_router()
        supported = router.get_supported_extensions()
        print(f"使用方法：将文档文件拖到此图标上，或命令行：python 工具名.py 文件1.docx ...")
        print(f"支持格式：{', '.join(supported)}")
        print()
        input("按回车键退出...")
        return

    router = get_router()
    success = 0
    fail = 0
    total = len(sys.argv) - 1
    
    for i, input_file in enumerate(sys.argv[1:], 1):
        print(f"正在处理第 {i}/{total} 个文件：{os.path.basename(input_file)}")
        
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
            fail += 1
        except Exception as e:
            logger.exception(f"处理失败：{input_file}")
            print(f"  ❌ 未知错误：{str(e)}")
            fail += 1
    
    print("=" * 50)
    print(f"处理完成！成功 {success} 个，失败 {fail} 个")
    print("=" * 50)
    input("按回车键退出...")


if __name__ == "__main__":
    main()
