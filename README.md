# 公文格式化工具

## 项目概述

公文格式化工具是一款基于Python开发的自动化文档处理应用程序，严格遵循《党政机关公文格式》（GB/T 9704-2012）国家标准，能够将普通Word文档快速转换为符合规范的公文格式。工具支持多种文档格式，提供统一的API接口和友好的命令行操作方式，适用于政府机关、企事业单位的公文处理场景。

## 功能特性

### 核心功能
- **多格式支持**：原生支持.docx格式，通过转换支持.doc和.wps格式
- **智能识别**：自动识别公文标题、正文、一级标题、二级标题等结构元素
- **格式标准化**：
  - 公文标题：方正小标宋简体，22磅，居中
  - 一级标题：黑体，16磅，首行缩进2字符
  - 二级标题：楷体_GB2312，16磅，加粗
  - 正文内容：仿宋_GB2312，16磅，首行缩进2字符
- **页面设置**：自动设置A4纸张尺寸及标准页边距（上3.7cm，下3.5cm，左2.8cm，右2.6cm）
- **页码处理**：自动添加标准格式页码（— PAGE —）
- **落款处理**：智能识别发文单位和日期，自动计算对齐位置
- **附件处理**：自动识别附件块并规范缩进格式
- **表格处理**：统一表格内字体为仿宋_GB2312
- **备份机制**：自动备份原文件，支持增量备份避免覆盖

### 高级功能
- **格式兼容性检测**：自动验证文档格式，提示不支持的版本或异常格式
- **错误恢复**：处理失败时自动从备份恢复原文件
- **批量处理**：支持同时处理多个文档文件
- **拖拽操作**：支持拖拽文件到exe图标直接处理

## 安装步骤

### 环境要求
- **操作系统**：Windows 7 及以上版本
- **Python版本**：Python 3.8 或更高版本
- **办公软件**：
  - Microsoft Word（处理.doc格式必需）
  - WPS Office（处理.wps格式可选）

### 安装依赖

1. 克隆或下载项目到本地目录
2. 打开命令行，进入项目根目录
3. 执行以下命令安装依赖：

```bash
pip install -r requirements.txt
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

直接将文档文件拖拽到`公文格式化工具.exe`图标上，程序将自动处理。

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
| Word 97-2003 | .doc | COM转换 | pywin32 + Microsoft Word |
| WPS文档 | .wps | COM转换 | pywin32 + WPS Office/Word |

## 项目结构

```
公文格式化工具/
├── src/                          # 源代码目录
│   ├── __init__.py
│   ├── main.py                   # 主程序入口
│   └── formatters/               # 格式处理器模块
│       ├── __init__.py
│       ├── base_handler.py       # 处理器基类
│       ├── docx_handler.py       # DOCX处理器
│       ├── doc_handler.py        # DOC处理器
│       ├── wps_handler.py        # WPS处理器
│       ── router.py             # 格式路由器
├── tests/                        # 测试目录
│   ├── __init__.py
│   ├── test_main.py              # 主程序测试
│   └── test_formatters.py        # 格式处理器测试
── assets/                       # 资源文件目录
│   └── icons/                    # 图标文件
├── scripts/                      # 工具脚本
│   └── convert_icon.py           # 图标转换脚本
├── config/                       # 配置文件
│   └── 公文格式化工具.spec       # PyInstaller配置
├── docs/                         # 文档目录
│   └── 目录结构说明.md
├── .gitignore                    # Git忽略配置
├── README.md                     # 项目说明
── requirements.txt              # 依赖列表
└── 公文格式化工具.exe            # 可执行文件
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

- **v5.0.0**：新增.doc和.wps格式支持，统一API接口
- **v4.0.0**：全面代码质量优化，类型注解，错误处理增强
- **v3.7.3**：落款偏移量校正，首行缩进优化

## 许可证

本项目仅供内部使用。

## 联系方式

作者：茗记
