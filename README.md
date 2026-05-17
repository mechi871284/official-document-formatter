# 公文格式化工具

## 项目概述

公文格式化工具是一款基于Python开发的自动化文档处理应用程序，严格遵循《党政机关公文格式》（GB/T 9704-2012）国家标准，能够将普通Word文档快速转换为符合规范的公文格式。工具支持多种文档格式，提供统一的API接口和友好的命令行操作方式，适用于政府机关、企事业单位的公文处理场景。

## 功能特性

### 核心功能
- **多格式原生支持**：原生支持.docx、.doc、.wps三种格式
- **智能识别**：自动识别公文标题、正文、各级标题等结构元素
- **格式标准化**：
  - 公文标题：方正小标宋简体，22磅，居中
  - 一级标题：黑体，16磅，首行缩进2字符
  - 二级标题：楷体_GB2312，16磅，加粗
  - 正文内容：仿宋_GB2312，16磅，首行缩进2字符
  - 落款/附件：仿宋_GB2312，16磅
- **页面设置**：自动设置A4纸张尺寸及标准页边距
- **页码处理**：自动添加标准格式页码
- **备份机制**：自动备份原文件，支持增量备份

### 高级功能
- **智能引擎选择**：优先使用WPS，备选Word
- **格式兼容性检测**：自动验证文档格式
- **错误恢复**：处理失败时自动从备份恢复
- **批量处理**：支持同时处理多个文档
- **拖拽操作**：拖拽文件到exe图标直接处理

## 安装步骤

### 环境要求
- **操作系统**：Windows 7 及以上版本 / Linux / macOS
- **Python版本**：Python 3.8 或更高版本
- **办公软件**：
  - Microsoft Word（处理.doc格式必需）
  - WPS Office（处理.wps格式可选）

### Windows 安装

1. 克隆或下载项目到本地目录
2. 打开命令行，进入项目根目录
3. 执行以下命令安装依赖：

```bash
pip install -r requirements.txt
```

### Linux 安装

#### 方式一：一键安装（推荐）

```bash
# 用户级安装（无需 root）
./scripts/install_linux.sh

# 系统级安装（需要 root）
sudo ./scripts/install_linux.sh --system
```

#### 方式二：手动安装

```bash
# 1. 安装依赖
pip3 install python-docx

# 2. 生成 .desktop 文件
./scripts/generate_desktop_file.sh

# 3. 或手动指定参数
./scripts/generate_desktop_file.sh \
  --name "公文格式化工具" \
  --exec "/opt/公文格式化工具/run.sh %f" \
  --icon "/opt/公文格式化工具/assets/icons/图标.png" \
  --system
```

#### 卸载

```bash
# 用户级卸载
./scripts/uninstall_linux.sh

# 系统级卸载
sudo ./scripts/uninstall_linux.sh --system
```

依赖包说明：
- `python-docx`：DOCX文档处理核心库
- `pywin32`：Windows COM接口，用于.doc和.wps格式转换
- `Pillow`：图像处理库，用于图标转换

## 使用方法

### 命令行方式

```bash
# 进入项目目录
cd 项目根目录

# 运行程序
python src/main.py

# 处理单个文件
python src/main.py 文档.docx

# 处理多个文件
python src/main.py 文档1.docx 文档2.doc 文档3.wps
```

### 拖拽方式

#### Windows
直接将文档文件拖拽到`公文格式化工具.exe`图标上，程序将自动处理。

#### Linux
安装 .desktop 文件后，支持以下拖放方式：
1. **应用图标拖放**：将文件拖拽到应用程序菜单中的图标上
2. **桌面快捷方式**：将文件拖拽到桌面上的快捷方式图标
3. **右键菜单**：右键点击文档 -> 打开方式 -> 公文格式化工具
4. **文件管理器**：在 Nautilus/Dolphin/Thunar 中直接拖放

### Python API方式

```python
from src.formatters.router import get_router

# 获取路由器实例
router = get_router()

# 格式化单个文件
backup_path = router.format_file(
    input_path='输入文件.docx',
    output_path='输出文件.docx',
    backup_path='备份文件-旧.docx'
)

# 验证文件格式
is_valid, message, handler = router.validate_file('文档.docx')
```

## 支持格式

| 格式 | 扩展名 | 处理方式 | 依赖要求 |
|------|--------|---------|---------|
| Word文档 | .docx | 原生处理 | python-docx |
| Word 97-2003 | .doc | COM接口原生处理 | pywin32 + WPS Office/Word |
| WPS文档 | .wps | COM接口原生处理 | pywin32 + WPS Office/Word |

### 字体规范

所有文本元素（中文、英文字母、数字）均使用对应位置的指定字体：

| 文本区域 | 指定字体 |
|---------|---------|
| 公文标题 | 方正小标宋简体 |
| 一级标题 | 黑体 |
| 二级标题 | 楷体_GB2312 |
| 正文/主送/落款/附件 | 仿宋_GB2312 |
| 页码 | 仿宋_GB2312 |

**注意**：已移除Times New Roman字体，统一使用指定字体。

## 项目结构

```
公文格式化工具/
├── assets/                       # 资源文件目录
│   └── icons/                    # 应用图标
│       ├── 公文格式化工具图标设计.ico
│       └── 公文格式化工具图标设计.png
├── config/                       # 配置文件
│   └── 公文格式化工具.spec       # PyInstaller打包配置
├── docs/                         # 文档目录
│   ├── Linux跨平台使用说明.md
│   └── 目录结构说明.md
├── scripts/                      # 工具脚本
│   ├── convert_icon.py           # 图标转换脚本
│   ├── generate_desktop_file.sh  # Linux .desktop 文件生成器
│   ├── install_linux.sh          # Linux 一键安装脚本
│   └── uninstall_linux.sh        # Linux 卸载脚本
├── src/                          # 源代码目录
│   ├── __init__.py
│   ├── main.py                   # 主程序入口（Windows）
│   ├── cross_platform_main.py     # 跨平台主程序入口
│   └── formatters/               # 格式处理器模块
│       ├── __init__.py
│       ├── base_handler.py       # 处理器基类
│       ├── docx_handler.py       # DOCX处理器
│       ├── doc_handler.py        # DOC处理器
│       ├── wps_handler.py        # WPS处理器
│       ├── router.py             # 格式路由器
│       ├── robustness.py         # 健壮性处理
│       └── cross_platform.py     # 跨平台支持
├── tests/                        # 测试目录
│   ├── __init__.py
│   ├── test_main.py              # 主程序测试
│   ├── test_formatters.py        # 格式处理器测试
│   ├── test_robustness.py        # 健壮性测试
│   └── linux_native_test.py      # Linux原生测试
├── .gitignore                    # Git忽略配置
├── README.md                     # 项目说明
├── requirements.txt              # 依赖列表
├── run.bat                       # Windows启动脚本
└── run.sh                        # Linux/macOS启动脚本
```

## 打包部署

### 打包为exe

```bash
# 进入项目目录
cd 项目根目录

# 执行打包
pyinstaller config/公文格式化工具.spec
```

打包完成后，可执行文件将生成在项目根目录。

### 分发说明

分发时只需提供以下文件：
- `公文格式化工具.exe`（主程序）
- `requirements.txt`（依赖说明）
- `README.md`（使用说明）

## 注意事项

### 字体要求
程序使用以下字体，请确保系统已安装：
- 方正小标宋简体（公文标题）
- 黑体（一级标题）
- 楷体_GB2312（二级标题）
- 仿宋_GB2312（正文、落款、附件）

如系统缺少相应字体，程序将使用默认字体替代。

### 格式限制
- .doc和.wps格式需要Microsoft Word或WPS Office支持
- 处理过程中会临时创建转换文件，处理完成后自动删除
- 建议处理前备份重要文档

### 性能说明
- 处理速度：约1-3秒/文档（取决于文档大小和复杂度）
- 内存占用：约50-100MB
- 支持批量处理，建议单次不超过50个文件

### 错误处理
- 文件不存在：提示错误并跳过
- 格式不支持：提示支持的格式列表
- 处理失败：自动恢复备份文件
- 依赖缺失：提示安装相应依赖包

## 开发指南

### 运行测试

```bash
# 运行所有测试
pytest tests/ -v

# 运行特定测试
pytest tests/test_formatters.py -v
```

### 扩展格式支持

1. 在`src/formatters/`目录下创建新的处理器类
2. 继承`FormatHandler`基类
3. 实现必需的抽象方法
4. 在`router.py`中注册新处理器

### 图标转换

```bash
python scripts/convert_icon.py
```

## 版本历史

- **v5.1.0**：优化doc和wps原生支持，修正字体设置
  - 优先使用WPS Office，备选Microsoft Word
  - 添加Word/WPS实例复用机制
  - 所有字符类型（中文、英文、数字）统一使用指定字体
  - 移除Times New Roman字体
  - 使用命名常量替代硬编码格式数字
- **v5.0.0**：新增.doc和.wps格式支持，统一API接口
- **v4.0.0**：全面代码质量优化，类型注解，错误处理增强
- **v3.7.3**：落款偏移量校正，首行缩进优化

## 许可证

本项目仅供内部使用。

## 联系方式

作者：茗记
