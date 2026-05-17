#!/bin/bash
#
# 公文格式化工具 - Linux 一键安装脚本
# ====================================
# 功能：自动安装应用到 Linux 系统
# 包括：生成 .desktop 文件、安装依赖、创建快捷方式
#
# 使用方法:
#   用户级安装:  ./install_linux.sh
#   系统级安装:  sudo ./install_linux.sh --system
#

set -e

# ======================== 常量定义 ========================
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
APP_NAME="公文格式化工具"
APP_VERSION="5.1.0"
DESKTOP_FILE_NAME="公文格式化工具-linux.desktop"

# 颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

info()    { echo -e "${BLUE}[INFO]${NC} $1"; }
success() { echo -e "${GREEN}[SUCCESS]${NC} $1"; }
warn()    { echo -e "${YELLOW}[WARN]${NC} $1"; }
error()   { echo -e "${RED}[ERROR]${NC} $1"; }

# ======================== 参数解析 ========================
INSTALL_MODE="user"

parse_args() {
    while [[ $# -gt 0 ]]; do
        case $1 in
            --system)
                INSTALL_MODE="system"
                shift
                ;;
            --user)
                INSTALL_MODE="user"
                shift
                ;;
            -h|--help)
                echo "用法: $0 [--system|--user]"
                echo "  --system  安装到系统目录 (需要 root)"
                echo "  --user    安装到用户目录 (默认)"
                exit 0
                ;;
            *)
                error "未知参数: $1"
                exit 1
                ;;
        esac
    done
}

# ======================== 系统检查 ========================
check_system() {
    info "检查系统环境..."
    
    # 检查操作系统
    if [ "$(uname -s)" != "Linux" ]; then
        error "此脚本仅支持 Linux 系统"
        exit 1
    fi
    
    # 检查 root 权限 (系统模式)
    if [ "$INSTALL_MODE" = "system" ] && [ "$EUID" -ne 0 ]; then
        error "系统级安装需要 root 权限，请使用 sudo 运行"
        exit 1
    fi
    
    success "系统检查通过"
}

# ======================== 依赖安装 ========================
install_dependencies() {
    info "检查并安装依赖..."
    
    # 检查 Python3
    if ! command -v python3 &> /dev/null; then
        warn "Python3 未安装，正在安装..."
        if command -v apt &> /dev/null; then
            sudo apt update && sudo apt install -y python3 python3-pip
        elif command -v dnf &> /dev/null; then
            sudo dnf install -y python3 python3-pip
        elif command -v pacman &> /dev/null; then
            sudo pacman -S python python-pip
        else
            error "无法自动安装 Python3，请手动安装"
            exit 1
        fi
    fi
    
    success "Python3 已安装: $(python3 --version)"
    
    # 检查 python-docx
    if ! python3 -c "import docx" 2>/dev/null; then
        info "安装 python-docx..."
        pip3 install python-docx -i https://pypi.tuna.tsinghua.edu.cn/simple
        success "python-docx 安装完成"
    else
        success "python-docx 已安装"
    fi
    
    # 检查 desktop-file-utils
    if ! command -v update-desktop-database &> /dev/null; then
        warn "desktop-file-utils 未安装"
        if command -v apt &> /dev/null; then
            sudo apt install -y desktop-file-utils
        elif command -v dnf &> /dev/null; then
            sudo dnf install -y desktop-file-utils
        fi
    fi
}

# ======================== 设置权限 ========================
setup_permissions() {
    info "设置文件权限..."
    
    chmod +x "$PROJECT_DIR/run.sh"
    chmod +x "$PROJECT_DIR/src/cross_platform_main.py"
    chmod +x "$SCRIPT_DIR/generate_desktop_file.sh"
    
    success "文件权限已设置"
}

# ======================== 生成 .desktop 文件 ========================
generate_desktop_file() {
    info "生成 .desktop 文件..."
    
    local output_dir=""
    if [ "$INSTALL_MODE" = "system" ]; then
        output_dir="/usr/share/applications"
    else
        output_dir="$HOME/.local/share/applications"
    fi
    
    mkdir -p "$output_dir"
    
    # 获取可执行路径和图标路径
    local exec_path="$PROJECT_DIR/run.sh %f"
    local icon_path=""
    
    # 查找图标
    icon_path=$(find "$PROJECT_DIR/assets" -name "*.png" 2>/dev/null | head -1)
    if [ -z "$icon_path" ]; then
        icon_path="document-save"
    fi
    
    # 生成 .desktop 文件
    cat > "$output_dir/$DESKTOP_FILE_NAME" << EOF
[Desktop Entry]
Version=$APP_VERSION
Name=$APP_NAME
GenericName=Document Formatter
Comment=公文自动格式化工具 -linux (支持 .docx/.doc/.wps)
Exec=$exec_path
Icon=$icon_path
Terminal=true
Type=Application
Categories=Office;Utility;TextTools;
MimeType=application/msword;application/vnd.openxmlformats-officedocument.wordprocessingml.document;
Keywords=word;doc;docx;wps;format;公文;格式化;
StartupNotify=true
StartupWMClass=公文格式化工具
Actions=FormatSingle;FormatBatch;

[Desktop Action FormatSingle]
Name=格式化单个文档
Exec=$exec_path

[Desktop Action FormatBatch]
Name=批量格式化
Exec=$PROJECT_DIR/run.sh %F
EOF

    chmod +x "$output_dir/$DESKTOP_FILE_NAME"
    
    success ".desktop 文件已生成: $output_dir/$DESKTOP_FILE_NAME"
}

# ======================== 更新桌面数据库 ========================
update_desktop_database() {
    info "更新桌面数据库..."
    
    local output_dir=""
    if [ "$INSTALL_MODE" = "system" ]; then
        output_dir="/usr/share/applications"
        sudo update-desktop-database "$output_dir"
    else
        output_dir="$HOME/.local/share/applications"
        update-desktop-database "$output_dir"
    fi
    
    success "桌面数据库已更新"
}

# ======================== 创建桌面快捷方式 ========================
create_desktop_shortcut() {
    local output_dir=""
    if [ "$INSTALL_MODE" = "system" ]; then
        output_dir="/usr/share/applications"
    else
        output_dir="$HOME/.local/share/applications"
    fi
    
    if [ -d "$HOME/Desktop" ]; then
        cp "$output_dir/$DESKTOP_FILE_NAME" "$HOME/Desktop/$DESKTOP_FILE_NAME"
        chmod +x "$HOME/Desktop/$DESKTOP_FILE_NAME"
        success "桌面快捷方式已创建: $HOME/Desktop/$DESKTOP_FILE_NAME"
    else
        warn "桌面目录不存在，跳过快捷方式创建"
    fi
}

# ======================== 显示安装结果 ========================
show_result() {
    echo ""
    echo "=========================================="
    echo "  安装完成！"
    echo "=========================================="
    echo ""
    echo "应用名称: $APP_NAME"
    echo "版本: $APP_VERSION"
    echo "安装模式: $INSTALL_MODE"
    echo ""
    echo "使用方式:"
    echo "  1. 在应用菜单中搜索 '$APP_NAME'"
    echo "  2. 将文件拖拽到应用图标上"
    echo "  3. 右键文档 -> 打开方式 -> $APP_NAME"
    echo "  4. 命令行运行: $PROJECT_DIR/run.sh 文档.docx"
    echo ""
    echo "卸载方法:"
    if [ "$INSTALL_MODE" = "system" ]; then
        echo "  sudo rm /usr/share/applications/$DESKTOP_FILE_NAME"
    else
        echo "  rm ~/.local/share/applications/$DESKTOP_FILE_NAME"
    fi
    echo ""
}

# ======================== 主函数 ========================
main() {
    echo ""
    echo "=========================================="
    echo "  $APP_NAME - Linux 安装程序"
    echo "  版本: $APP_VERSION"
    echo "=========================================="
    echo ""
    
    parse_args "$@"
    
    check_system
    install_dependencies
    setup_permissions
    generate_desktop_file
    update_desktop_database
    create_desktop_shortcut
    
    show_result
    
    success "安装完成！"
    exit 0
}

main "$@"
