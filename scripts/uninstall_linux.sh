#!/bin/bash
#
# 公文格式化工具 - Linux 卸载脚本
# ================================
# 功能：安全卸载应用及所有相关文件
#
# 使用方法:
#   用户级卸载:  ./uninstall_linux.sh
#   系统级卸载:  sudo ./uninstall_linux.sh --system
#

set -e

# ======================== 常量定义 ========================
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
APP_NAME="公文格式化工具"
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
            --remove-data)
                REMOVE_DATA=true
                shift
                ;;
            -h|--help)
                echo "用法: $0 [--system|--user] [--remove-data]"
                echo "  --system      卸载系统级安装 (需要 root)"
                echo "  --user        卸载用户级安装 (默认)"
                echo "  --remove-data 同时删除项目文件"
                exit 0
                ;;
            *)
                error "未知参数: $1"
                exit 1
                ;;
        esac
    done
}

# ======================== 确认卸载 ========================
confirm_uninstall() {
    echo ""
    echo "=========================================="
    echo "  卸载 $APP_NAME"
    echo "=========================================="
    echo ""
    echo "即将执行以下操作:"
    echo "  1. 删除 .desktop 文件"
    echo "  2. 删除桌面快捷方式"
    echo "  3. 更新桌面数据库"
    
    if [ "${REMOVE_DATA:-false}" = "true" ]; then
        echo "  4. 删除项目文件: $PROJECT_DIR"
    fi
    
    echo ""
    read -p "是否继续? (y/N): " confirm
    
    if [[ ! "$confirm" =~ ^[Yy]$ ]]; then
        info "卸载已取消"
        exit 0
    fi
}

# ======================== 卸载 .desktop 文件 ========================
uninstall_desktop_file() {
    info "卸载 .desktop 文件..."
    
    local desktop_path=""
    if [ "$INSTALL_MODE" = "system" ]; then
        desktop_path="/usr/share/applications/$DESKTOP_FILE_NAME"
        if [ -f "$desktop_path" ]; then
            sudo rm -f "$desktop_path"
            success "已删除: $desktop_path"
        else
            warn "文件不存在: $desktop_path"
        fi
    else
        desktop_path="$HOME/.local/share/applications/$DESKTOP_FILE_NAME"
        if [ -f "$desktop_path" ]; then
            rm -f "$desktop_path"
            success "已删除: $desktop_path"
        else
            warn "文件不存在: $desktop_path"
        fi
    fi
}

# ======================== 删除桌面快捷方式 ========================
remove_desktop_shortcut() {
    info "删除桌面快捷方式..."
    
    local shortcut="$HOME/Desktop/$DESKTOP_FILE_NAME"
    if [ -f "$shortcut" ]; then
        rm -f "$shortcut"
        success "已删除: $shortcut"
    else
        warn "快捷方式不存在: $shortcut"
    fi
}

# ======================== 更新桌面数据库 ========================
update_desktop_database() {
    info "更新桌面数据库..."
    
    if command -v update-desktop-database &> /dev/null; then
        if [ "$INSTALL_MODE" = "system" ]; then
            sudo update-desktop-database /usr/share/applications
        else
            update-desktop-database "$HOME/.local/share/applications"
        fi
        success "桌面数据库已更新"
    else
        warn "update-desktop-database 未找到"
    fi
}

# ======================== 删除项目文件 ========================
remove_project_files() {
    if [ "${REMOVE_DATA:-false}" = "true" ]; then
        info "删除项目文件..."
        warn "即将删除: $PROJECT_DIR"
        read -p "确认删除? (y/N): " confirm
        
        if [[ "$confirm" =~ ^[Yy]$ ]]; then
            cd "$PROJECT_DIR/.."
            rm -rf "$PROJECT_DIR"
            success "项目文件已删除"
        else
            info "跳过项目文件删除"
        fi
    fi
}

# ======================== 显示卸载结果 ========================
show_result() {
    echo ""
    echo "=========================================="
    echo "  卸载完成！"
    echo "=========================================="
    echo ""
    echo "已删除:"
    echo "  - .desktop 文件"
    echo "  - 桌面快捷方式"
    echo ""
    
    if [ "${REMOVE_DATA:-false}" = "true" ]; then
        echo "  - 项目文件 (可选)"
    fi
    
    echo ""
    echo "注意: Python 依赖包未自动删除"
    echo "  如需删除: pip3 uninstall python-docx"
    echo ""
}

# ======================== 主函数 ========================
main() {
    parse_args "$@"
    
    # 检查 root 权限 (系统模式)
    if [ "$INSTALL_MODE" = "system" ] && [ "$EUID" -ne 0 ]; then
        error "系统级卸载需要 root 权限，请使用 sudo 运行"
        exit 1
    fi
    
    confirm_uninstall
    uninstall_desktop_file
    remove_desktop_shortcut
    update_desktop_database
    remove_project_files
    
    show_result
    
    success "卸载完成！"
    exit 0
}

main "$@"
