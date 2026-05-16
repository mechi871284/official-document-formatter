#!/bin/bash
#
# 公文格式化工具 - Linux/macOS 启动脚本
# ==========================================
# 功能：在 Linux 或 macOS 系统上启动公文格式化工具
# 支持格式：.docx（完整支持），.doc/.wps 提示需要 Windows
#
# 使用方法:
#   1. 直接运行:        ./run.sh
#   2. 处理单个文件:    ./run.sh 文档.docx
#   3. 批量处理:        ./run.sh 文档1.docx 文档2.docx
#

# 获取脚本所在目录
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Python 检查
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误: 未找到 python3"
    echo "   请安装 Python 3.8+"
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo "   macOS: brew install python3"
    else
        echo "   Ubuntu/Debian: sudo apt install python3 python3-pip"
        echo "   RHEL/CentOS: sudo dnf install python3 python3-pip"
    fi
    exit 1
fi

PYTHON_VERSION=$(python3 --version | awk '{print $2}')
echo "🐍 使用 Python: $PYTHON_VERSION"

# 依赖检查
echo "📦 检查依赖..."
cd "$SCRIPT_DIR"

# 检查并安装依赖
python3 -c "import docx" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "⚠️  缺少 python-docx，正在安装..."
    pip3 install python-docx -i https://pypi.tuna.tsinghua.edu.cn/simple
fi

# 运行工具
echo ""
echo "🚀 启动公文格式化工具..."
echo ""

if [ $# -gt 0 ]; then
    # 有参数的情况
    python3 "$SCRIPT_DIR/src/cross_platform_main.py" "$@"
else
    # 无参数的情况
    python3 "$SCRIPT_DIR/src/cross_platform_main.py"
fi
