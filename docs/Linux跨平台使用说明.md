# 🐧 Linux/macOS 跨平台使用说明

## 功能矩阵

| 功能 | Windows | Linux | macOS |
|------|---------|-------|-------|
| .docx 格式处理 | ✅ 完整支持 | ✅ 完整支持 | ✅ 完整支持 |
| .doc 格式处理 | ✅ 原生支持 | ⚠️ 提示转换 | ⚠️ 提示转换 |
| .wps 格式处理 | ✅ 原生支持 | ⚠️ 提示转换 | ⚠️ 提示转换 |
| 自动备份 | ✅ | ✅ | ✅ |
| GB/T 9704-2012 标准 | ✅ | ✅ | ✅ |

> ⚠️ 提示：.doc/.wps 格式依赖 Windows COM 组件，在 Linux/macOS 下建议先在 Windows 转成 .docx

## 快速开始

### 方式一：使用启动脚本（推荐）

```bash
# 1. 进入项目目录
cd 公文格式化工具

# 2. 给脚本执行权限
chmod +x run.sh

# 3. 直接运行
./run.sh

# 4. 或者直接处理文档
./run.sh 文档.docx
./run.sh 文档1.docx 文档2.docx
```

### 方式二：直接调用 Python

```bash
# 进入项目根目录
cd 公文格式化工具

# 安装依赖
pip3 install python-docx

# 运行跨平台版本
python3 src/cross_platform_main.py 文档.docx
```

## 环境要求

| 软件 | 最低版本 |
|------|---------|
| Python | 3.8+ |
| python-docx | 0.8.11+ |

## 依赖安装

### Ubuntu/Debian

```bash
sudo apt update
sudo apt install python3 python3-pip
pip3 install python-docx
```

### RHEL/CentOS/Fedora

```bash
sudo dnf install python3 python3-pip
pip3 install python-docx
```

### macOS (Homebrew)

```bash
brew install python3
pip3 install python-docx
```

## 文件说明

```
公文格式化工具/
├── run.sh                      # Linux/macOS 启动脚本
├── src/
│   ├── cross_platform_main.py  # 跨平台主程序入口
│   ├── main.py                 # Windows 完整版主程序
│   └── formatters/
│       ├── robustness.py       # 健壮性增强模块
│       ├── base_handler.py     # 基类定义
│       ├── docx_handler.py     # DOCX处理器（跨平台）
│       ├── doc_handler.py      # DOC处理器（Windows-only）
│       ├── wps_handler.py      # WPS处理器（Windows-only）
│       └── router.py           # 格式路由
├── config/
├── assets/
├── tests/
├── docs/
└── README.md
```

## 功能对比

### 相同点（所有平台）

✅ 自动识别公文结构（标题、正文、落款）
✅ 标准化字体、字号、行间距设置
✅ 自动添加页码
✅ 文件自动备份（`-未修改`后缀）
✅ 命令行参数处理
✅ 多文件批量处理

### Windows 独有功能

⚡ `.doc` 格式原生处理（pywin32 + Word/WPS）
⚡ `.wps` 格式原生处理（pywin32 + WPS/Word）
⚡ EXE 可执行文件打包

## 格式转换建议

如果你在 Linux/macOS 上收到 `.doc/.wps` 不支持的提示：

```
方案一：在 Windows 系统转换
1. 在 Windows 打开 Word/WPS
2. 文件 -> 另存为 -> *.docx
3. 在 Linux/macOS 上处理 .docx

方案二：使用在线转换工具
1. 上传到 Google Docs / Microsoft 365
2. 导出为 .docx
3. 下载后处理
```

## 故障排除

### 问题1: ModuleNotFoundError

```bash
# 错误: No module named 'docx'
# 解决: 安装依赖
pip3 install python-docx
```

### 问题2: 权限不足

```bash
# 错误: Permission denied
# 解决: 添加执行权限
chmod +x run.sh
```

### 问题3: Python 版本太旧

```bash
# 检查版本
python3 --version  # 需要 3.8+

# Ubuntu 升级 Python
sudo add-apt-repository ppa:deadsnakes/ppa
sudo apt install python3.10
```

## 源代码说明

### cross_platform_main.py 核心逻辑

1. **系统检测**：使用 `platform.system()` 判断运行环境
2. **功能降级**：非 Windows 系统显示支持的格式
3. **错误提示**：针对不同平台给出友好的解决建议
4. **延迟导入**：避免 Windows-only 依赖在导入时报错

### router.py 自动适配

路由器在非 Windows 系统会自动返回只包含 `.docx` 的处理器：

```python
# 实际效果：
# Windows: [DocxFormatHandler(), DocFormatHandler(), WpsFormatHandler()]
# Linux:   [DocxFormatHandler()]
# macOS:   [DocxFormatHandler()]
```

这是因为 Doc/Wps Handler 的 `can_handle()` 会检测 `win32_available`，
非 Windows 下自动返回 `False`。

---

有问题欢迎提 Issue！🚀
