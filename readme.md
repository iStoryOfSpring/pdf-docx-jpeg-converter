# PDF·Word·JPG 互转工具

**武汉纺织大学管理学院媒体运营部**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)]()

一个基于 Python + Tkinter 的文档格式互转工具，支持 **PDF / Word / JPEG** 三种格式之间任意方向的无缝互转。

## ✨ 功能特性

- **6 条转换路径全覆盖**
  - PDF ↔ Word（docx）
  - PDF ↔ JPEG（支持多页逐页渲染）
  - Word ↔ JPEG
- **批量文件处理** — 一次添加多个文件，每行独立设置目标格式
- **导入文件夹** — 一键扫描整个目录中所有受支持格式的文件
- **.doc 旧格式兼容** — 通过 COM Word 桥接，自动将 `.doc` 转为 `.docx` 后进入转换管线
- **输出选项丰富**
  - 自定义输出目录，也可输出到源文件所在目录
  - JPEG 质量（1-100）滑块
  - 渲染 DPI（72-600）
  - PDF 页码范围（全部 / 指定范围）
  - 文件命名冲突策略：自动重命名 / 覆盖 / 跳过
- **友好的图形界面** — 基于 Tkinter，Treeview 文件列表，实时进度条，转换在后台线程运行不冻屏

## 📸 界面预览

```
┌──────────────────────────────────────────────────────────────┐
│  PDF·Word·JPG互转工具-武汉纺织大学管理学院媒体运营部           │
├──────────────────────────────────────────────────────────────┤
│  [+ 添加文件] [+ 导入文件夹] [— 移除选中] [清空列表]           │
│           [全部→PDF] [全部→Word] [全部→JPEG]  [仅选中的行]     │
├──────────────────────────────────────────────────────────────┤
│  # │ 文件名        │ 源格式 │ 目标格式 │ 大小   │ 状态        │
│  1 │ report.pdf    │ PDF    │ [Word ▼] │ 2.3MB  │ 完成 ✓      │
│  2 │ photo.jpg     │ JPEG   │ [PDF  ▼] │ 1.1MB  │ 就绪        │
│  3 │ contract.docx │ Word   │ [JPEG ▼] │ 450KB  │ 就绪        │
├──────────────────────────────────────────────────────────────┤
│  ┌─ 输出设置 ───────────────────────────────────────────┐    │
│  │ ☑ 输出到源文件所在目录                                 │    │
│  │ JPEG质量: [90] ██████░░  DPI: [200]                   │    │
│  │ PDF页面: ● 全部  ○ 指定范围: [1]—[10]                │    │
│  │ 文件冲突: ● 自动重命名  ○ 覆盖  ○ 跳过                │    │
│  └──────────────────────────────────────────────────────┘    │
├──────────────────────────────────────────────────────────────┤
│  ████████████░░░░░░░░  40%   进度: 2/5 个文件                 │
│  正在转换: report.pdf → report.docx (第 2/3 页)               │
├──────────────────────────────────────────────────────────────┤
│  [▶ 开始全部转换]  [取消]  [打开输出目录]    状态: 正在转换... │
└──────────────────────────────────────────────────────────────┘
```

## 🔧 转换矩阵

| # | 方向 | 实现方式 | 核心库 |
|---|------|----------|--------|
| 1 | PDF → Word | `pdf2docx.Converter` 提取文本/表格 | `pdf2docx` |
| 2 | PDF → JPEG | 逐页光栅化 `page.get_pixmap()` | `PyMuPDF` |
| 3 | Word → PDF | Windows COM（保真度最高） | `docx2pdf` |
| 4 | Word → JPEG | Word→临时PDF→JPEG 链式 | `docx2pdf` + `PyMuPDF` |
| 5 | JPEG → PDF | 创建页面并 `insert_image()` 嵌入 | `PyMuPDF` |
| 6 | JPEG → Word | `python-docx` 创建文档并 `add_picture()` | `python-docx` |
| * | `.doc` → `.docx` | COM Word/WPS 桥接，自动转临时 `.docx` | `pywin32` |

## 📋 环境要求

| 项目 | 要求 |
|------|------|
| 操作系统 | Windows（依赖 COM Word 组件） |
| Python | 3.10 及以上 |
| Microsoft Office 或 WPS Office | 已安装（Word→PDF、.doc 桥接需要） |

## 🚀 快速开始

### 1. 克隆仓库

```bash
git clone https://github.com/istoryofspring/pdf-word-jpg-converter.git
cd pdf-word-jpg-converter
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 启动程序

```bash
python main.py
```

## 📦 依赖清单

| 库 | 版本 | 用途 |
|----|------|------|
| `PyMuPDF` | ≥1.26.0 | PDF 读取、渲染、创建 |
| `python-docx` | ≥1.1.0 | Word 文档读写 |
| `pywin32` | ≥300 | COM Word/WPS 自动化（Word→PDF、.doc 桥接） |
| `pdf2docx` | ≥0.5.8 | PDF → Word |
| `Pillow` | ≥10.0.0 | 图片读写 |
| `tkinterdnd2` | ≥0.5.0 | 拖放文件支持 |

## 📖 使用说明

### 添加文件

- **添加文件**：点击 `+ 添加文件` → 弹出多选对话框（Ctrl+点击 多选）
- **拖拽文件**：直接从资源管理器拖入文件或文件夹到窗口
- **导入文件夹**：点击 `+ 导入文件夹` → 选择目录 → 自动扫描所有支持格式
- 支持格式：`.pdf` `.docx` `.doc` `.jpg` `.jpeg` `.png`

### 设置目标格式

- **逐行设置**：双击文件行的「目标格式」单元格 → 下拉选择 PDF / Word / JPEG
- **批量设置**：点击 `全部→PDF` / `全部→Word` / `全部→JPEG`
- 与源格式相同的选项自动过滤，避免无意义的同格式转换

### 输出选项

| 选项 | 说明 |
|------|------|
| JPEG 质量 | 1-100，默认 90（越高文件越大） |
| 渲染 DPI | 72-600，默认 200（PDF→JPEG 时控制图片清晰度） |
| PDF 页面范围 | `全部页面` 或指定起止页（仅 PDF→JPEG 生效） |
| 文件冲突策略 | `自动重命名`（末尾加 _1）/ `覆盖` / `跳过` |

### 转换状态

| 状态 | 含义 |
|------|------|
| 就绪 | 等待转换 |
| 转换中... | 正在转换 |
| 完成 ✓ | 转换成功 |
| 失败 ✗ | 转换失败，可查看汇总弹窗中的失败详情 |
| 文件缺失 | 文件已不存在，已自动跳过 |

## 📝 开源协议

本项目基于 [MIT License](https://opensource.org/licenses/MIT) 开源，可自由使用、修改和分发。

---

**作者**：[@istoryofspring](https://github.com/istoryofspring)
