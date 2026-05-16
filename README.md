# 公文格式化工具

## 项目简介

公文格式化工具是一款用于自动格式化公文文档的Python应用程序，支持将普通Word文档转换为符合国家标准格式的公文。

## 功能特性

- 自动识别公文标题和正文
- 标准化字体、字号、行距设置
- 自动添加页码和页眉页脚
- 表格格式统一处理
- 落款位置自动计算
- 附件格式规范化
- 文件自动备份机制

## 快速开始

### 环境要求

- Python 3.8+
- Windows 7 及以上

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行程序

```bash
python src/main.py
```

### 打包为exe

```bash
pyinstaller config/公文格式化工具.spec
```

打包后的可执行文件位于 `dist/` 目录。

## 项目结构

详见 [目录结构说明.md](docs/目录结构说明.md)

## 开发指南

### 运行测试

```bash
pytest tests/ -v
```

### 图标转换

```bash
python scripts/convert_icon.py
```

## 许可证

本项目仅供内部使用。
