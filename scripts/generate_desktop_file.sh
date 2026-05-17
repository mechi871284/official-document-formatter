#!/bin/bash
#
# 公文格式化工具 - Linux .desktop 文件生成脚本
# ==========================================
# 功能：生成符合 FreeDesktop 规范的 .desktop 文件
# 支持：自动检测可执行路径、图标位置，或手动指定
# 兼容：GNOME / KDE / XFCE / Cinnamon / MATE
#
# 使用方法:
#   1. 自动模式:  ./generate_desktop_file.sh
#   2. 手动模式:  ./generate_desktop_file.sh --name "应用名" --exec "/path/to/exec" --icon "/path/to/icon"
#   3. 完整参数:  ./generate_desktop_file.sh --help
#

set -e

# ======================== 常量定义 ========================
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
DESKTOP_FILE_NAME="公文格式化工具-linux.desktop"
OUTPUT_DIR=""
INSTALL_MODE="user"  # user 或 system

# 默认值
APP_NAME="公文格式化工具"
APP_COMMENT="公文自动格式化工具 -linux (支持 .docx/.doc/.wps)"
APP_EXEC=""
APP_ICON=""
APP_TYPE="Application"
APP_CATEGORIES="Office;Utility;TextTools;"
APP_MIMETYPE="application/msword;application/vnd.openxmlformats-officedocument.wordprocessingml.document;"
APP_KEYWORDS="word;doc;docx;wps;format;公文;格式化;"
APP_TERMINAL="true"
APP_VERSION="5.1.0"

# ======================== 颜色输出 ========================
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

info()    { echo -e "${BLUE}[INFO]${NC} $1"; }
success() { echo -e "${GREEN}[SUCCESS]${NC} $1"; }
warn()    { echo -e "${YELLOW}[WARN]${NC} $1"; }
error()   { echo -e "${RED}[ERROR]${NC} $1"; }

# ======================== 帮助信息 ========================
show_help() {
    cat << EOF
公文格式化工具 - .desktop 文件生成器

用法:
  $0 [选项]

选项:
  -n, --name NAME         应用程序名称 (默认: 公文格式化工具)
  -c, --comment COMMENT   应用程序描述 (默认: 自动包含 -linux 标识)
  -e, --exec PATH         可执行文件路径 (默认: 自动检测)
  -i, --icon PATH         图标文件路径 (默认: 自动检测)
  -o, --output DIR        输出目录 (默认: 根据 --system/--user 决定)
  -t, --terminal BOOL     是否显示终端 (true/false, 默认: true)
  -k, --keywords KW       搜索关键词 (默认: 自动设置)
  -m, --mimetype TYPE     MIME 类型 (默认: 自动设置)
  -g, --categories CAT    应用分类 (默认: Office;Utility;)
      --user              安装到用户目录 (默认)
      --system            安装到系统目录 (需要 root)
  -h, --help              显示此帮助信息

示例:
  # 自动检测并生成
  $0

  # 手动指定参数
  $0 --name "公文工具" --exec "/opt/app/run.sh" --icon "/opt/app/icon.png"

  # 生成到系统目录
  $0 --system

  # 生成到指定目录
  $0 --output ~/Desktop

EOF
    exit 0
}

# ======================== 参数解析 ========================
parse_args() {
    while [[ $# -gt 0 ]]; do
        case $1 in
            -n|--name)
                APP_NAME="$2"
                shift 2
                ;;
            -c|--comment)
                APP_COMMENT="$2"
                shift 2
                ;;
            -e|--exec)
                APP_EXEC="$2"
                shift 2
                ;;
            -i|--icon)
                APP_ICON="$2"
                shift 2
                ;;
            -o|--output)
                OUTPUT_DIR="$2"
                shift 2
                ;;
            -t|--terminal)
                APP_TERMINAL="$2"
                shift 2
                ;;
            -k|--keywords)
                APP_KEYWORDS="$2"
                shift 2
                ;;
            -m|--mimetype)
                APP_MIMETYPE="$2"
                shift 2
                ;;
            -g|--categories)
                APP_CATEGORIES="$2"
                shift 2
                ;;
            --user)
                INSTALL_MODE="user"
                shift
                ;;
            --system)
                INSTALL_MODE="system"
                shift
                ;;
            -h|--help)
                show_help
                ;;
            *)
                error "未知参数: $1"
                show_help
                ;;
        esac
    done
}

# ======================== 自动检测函数 ========================

detect_exec_path() {
    info "检测可执行文件路径..."
    
    # 优先级 1: 项目目录下的 run.sh
    if [ -f "$PROJECT_DIR/run.sh" ]; then
        APP_EXEC="$PROJECT_DIR/run.sh %f"
        success "找到 run.sh: $APP_EXEC"
        return 0
    fi
    
    # 优先级 2: 项目目录下的主程序
    if [ -f "$PROJECT_DIR/src/cross_platform_main.py" ]; then
        APP_EXEC="python3 $PROJECT_DIR/src/cross_platform_main.py %f"
        success "找到 cross_platform_main.py"
        return 0
    fi
    
    # 优先级 3: 查找 python 脚本
    local main_py=$(find "$PROJECT_DIR/src" -name "main.py" -o -name "*.py" 2>/dev/null | head -1)
    if [ -n "$main_py" ]; then
        APP_EXEC="python3 $main_py %f"
        success "找到主程序: $main_py"
        return 0
    fi
    
    warn "未找到可执行文件，请手动指定 --exec 参数"
    return 1
}

detect_icon_path() {
    info "检测图标文件路径..."
    
    # 优先级 1: PNG 图标
    local png_icon=$(find "$PROJECT_DIR/assets" -name "*.png" 2>/dev/null | head -1)
    if [ -n "$png_icon" ]; then
        APP_ICON="$png_icon"
        success "找到 PNG 图标: $APP_ICON"
        return 0
    fi
    
    # 优先级 2: SVG 图标
    local svg_icon=$(find "$PROJECT_DIR/assets" -name "*.svg" 2>/dev/null | head -1)
    if [ -n "$svg_icon" ]; then
        APP_ICON="$svg_icon"
        success "找到 SVG 图标: $APP_ICON"
        return 0
    fi
    
    # 优先级 3: ICO 图标
    local ico_icon=$(find "$PROJECT_DIR/assets" -name "*.ico" 2>/dev/null | head -1)
    if [ -n "$ico_icon" ]; then
        APP_ICON="$ico_icon"
        success "找到 ICO 图标: $APP_ICON"
        return 0
    fi
    
    # 默认图标
    APP_ICON="document-save"
    warn "未找到图标文件，使用系统默认图标"
    return 0
}

# ======================== 输出目录确定 ========================
determine_output_dir() {
    if [ -n "$OUTPUT_DIR" ]; then
        return 0
    fi
    
    if [ "$INSTALL_MODE" = "system" ]; then
        OUTPUT_DIR="/usr/share/applications"
        info "系统模式: 输出到 $OUTPUT_DIR"
    else
        OUTPUT_DIR="$HOME/.local/share/applications"
        info "用户模式: 输出到 $OUTPUT_DIR"
    fi
    
    mkdir -p "$OUTPUT_DIR"
}

# ======================== 生成 .desktop 文件 ========================
generate_desktop_file() {
    local desktop_path="$OUTPUT_DIR/$DESKTOP_FILE_NAME"
    
    info "生成 .desktop 文件..."
    info "输出路径: $desktop_path"
    
    cat > "$desktop_path" << EOF
[Desktop Entry]
Version=$APP_VERSION
Name=$APP_NAME
GenericName=Document Formatter
Comment=$APP_COMMENT
Exec=$APP_EXEC
Icon=$APP_ICON
Terminal=$APP_TERMINAL
Type=$APP_TYPE
Categories=$APP_CATEGORIES
MimeType=$APP_MIMETYPE
Keywords=$APP_KEYWORDS
StartupNotify=true
StartupWMClass=公文格式化工具
Actions=FormatSingle;FormatBatch;

[Desktop Action FormatSingle]
Name=格式化单个文档
Exec=$APP_EXEC

[Desktop Action FormatBatch]
Name=批量格式化
Exec=$APP_EXEC %F
EOF

    success ".desktop 文件已生成: $desktop_path"
    return 0
}

# ======================== 验证 .desktop 文件 ========================
validate_desktop_file() {
    local desktop_path="$OUTPUT_DIR/$DESKTOP_FILE_NAME"
    
    info "验证 .desktop 文件..."
    
    # 检查必要字段
    local required_fields=("Name" "Exec" "Icon" "Type" "Categories")
    local missing_fields=()
    
    for field in "${required_fields[@]}"; do
        if ! grep -q "^${field}=" "$desktop_path"; then
            missing_fields+=("$field")
        fi
    done
    
    if [ ${#missing_fields[@]} -gt 0 ]; then
        error "缺少必要字段: ${missing_fields[*]}"
        return 1
    fi
    
    # 检查 -linux 标识
    if grep -q "\-linux" "$desktop_path"; then
        success "包含 -linux 标识符"
    else
        warn "未找到 -linux 标识符"
    fi
    
    # 检查 Exec 路径
    local exec_line=$(grep "^Exec=" "$desktop_path" | cut -d'=' -f2 | cut -d' ' -f1)
    if [ -f "$exec_line" ] || command -v "$exec_line" &> /dev/null; then
        success "可执行路径有效: $exec_line"
    else
        warn "可执行路径可能无效: $exec_line"
    fi
    
    # 检查 Icon 路径
    local icon_line=$(grep "^Icon=" "$desktop_path" | cut -d'=' -f2)
    if [ -f "$icon_line" ]; then
        success "图标路径有效: $icon_line"
    else
        warn "图标路径可能无效: $icon_line"
    fi
    
    # 使用 desktop-file-validate (如果可用)
    if command -v desktop-file-validate &> /dev/null; then
        info "使用 desktop-file-validate 进行验证..."
        local validation_output=$(desktop-file-validate "$desktop_path" 2>&1)
        if [ -z "$validation_output" ]; then
            success "desktop-file-validate 验证通过"
        else
            warn "desktop-file-validate 发现以下问题:"
            echo "$validation_output"
        fi
    else
        warn "desktop-file-validate 未安装，跳过深度验证"
        warn "安装方法: sudo apt install desktop-file-utils"
    fi
    
    return 0
}

# ======================== 更新桌面数据库 ========================
update_desktop_database() {
    info "更新桌面数据库..."
    
    if command -v update-desktop-database &> /dev/null; then
        if [ "$INSTALL_MODE" = "system" ]; then
            sudo update-desktop-database "$OUTPUT_DIR"
        else
            update-desktop-database "$OUTPUT_DIR"
        fi
        success "桌面数据库已更新"
    else
        warn "update-desktop-database 未找到"
        warn "安装方法: sudo apt install desktop-file-utils"
    fi
    
    # 更新 MIME 数据库
    if command -v update-mime-database &> /dev/null; then
        if [ "$INSTALL_MODE" = "system" ]; then
            sudo update-mime-database ~/.local/share/mime 2>/dev/null || true
        fi
    fi
}

# ======================== 创建桌面快捷方式 ========================
create_desktop_shortcut() {
    local desktop_path="$OUTPUT_DIR/$DESKTOP_FILE_NAME"
    local desktop_shortcut="$HOME/Desktop/$DESKTOP_FILE_NAME"
    
    info "创建桌面快捷方式..."
    
    if [ -d "$HOME/Desktop" ]; then
        cp "$desktop_path" "$desktop_shortcut"
        chmod +x "$desktop_shortcut"
        success "桌面快捷方式已创建: $desktop_shortcut"
    else
        warn "桌面目录不存在: $HOME/Desktop"
        warn "快捷方式未创建"
    fi
}

# ======================== 显示生成结果 ========================
show_result() {
    local desktop_path="$OUTPUT_DIR/$DESKTOP_FILE_NAME"
    
    echo ""
    echo "=========================================="
    echo "  .desktop 文件生成完成"
    echo "=========================================="
    echo ""
    echo "文件路径: $desktop_path"
    echo "安装模式: $INSTALL_MODE"
    echo ""
    echo "内容预览:"
    echo "------------------------------------------"
    cat "$desktop_path"
    echo "------------------------------------------"
    echo ""
    echo "后续操作:"
    echo "  1. 在应用菜单中搜索 '$APP_NAME'"
    echo "  2. 将文件拖拽到应用图标上进行处理"
    echo "  3. 右键点击文档 -> 打开方式 -> $APP_NAME"
    echo ""
    
    if [ "$INSTALL_MODE" = "system" ]; then
        echo "注意: 系统级安装需要 root 权限"
        echo "  卸载: sudo rm $desktop_path"
    else
        echo "用户级安装，无需 root 权限"
        echo "  卸载: rm $desktop_path"
    fi
    
    echo ""
}

# ======================== 主函数 ========================
main() {
    echo ""
    echo "=========================================="
    echo "  公文格式化工具 - .desktop 文件生成器"
    echo "  版本: $APP_VERSION"
    echo "=========================================="
    echo ""
    
    parse_args "$@"
    
    # 自动检测
    if [ -z "$APP_EXEC" ]; then
        detect_exec_path || {
            error "无法自动检测可执行路径，请手动指定 --exec 参数"
            exit 1
        }
    fi
    
    if [ -z "$APP_ICON" ]; then
        detect_icon_path
    fi
    
    # 确定输出目录
    determine_output_dir
    
    # 生成文件
    generate_desktop_file
    
    # 验证
    validate_desktop_file
    
    # 更新数据库
    update_desktop_database
    
    # 创建桌面快捷方式
    create_desktop_shortcut
    
    # 显示结果
    show_result
    
    success "所有操作完成！"
    exit 0
}

# 执行主函数
main "$@"
