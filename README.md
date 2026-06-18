# Lab Report Generator

**Claude Code 技能** — 基于 Word 模板 + 实验截图自动生成格式化实验报告（DOCX/PDF）。

从实验报告模板中提取章节结构（实验目的→实验原理→实验步骤→实验作业），填充答案和截图，生成格式规范的完整报告。

## 工作流程

```
阶段 0 ──→ 阶段 1 ──→ 阶段 2 ──→ 阶段 3
分析模板     分析格式     截图映射     生成报告
```

## 安装

### 依赖

```bash
pip install python-docx Pillow
```

### PDF 输出（可选）

仅当你需要生成 PDF 时才需要：

- **Windows**: 从 https://miktex.org/download 下载安装 MiKTeX
- **macOS**: `brew install texlive`
- **Linux**: `sudo apt install texlive-xetex`

安装后确保 `lualatex` 或 `xelatex` 可在终端中运行。

### 自动识图（可选）

仅当 AI 模型不具备原生读图能力时才需要：

1. 在阿里云百炼平台获取 API Key：https://bailian.console.aliyun.com/
2. 在 `skills/vision-skill/.env` 中配置 `DASHSCOPE_API_KEY=your-key`

## 用法

### 阶段 0：分析模板结构

```bash
python analyze_content.py --yes "实验报告模板.docx" 实验内容结构.json
```

### 阶段 1：分析模板格式

```bash
python analyze_template.py "实验报告模板.docx" 格式规范
```

### 阶段 2：截图分析（二选一）

**模式 A — AI 模型自身读图（推荐）**

```bash
python analyze_images.py --model-vision "截图目录/" 实验内容结构.json image_map.json
```
然后用 Read 工具逐张查看截图，编辑 image_map.json 填入路径。

**模式 B — 外部视觉 API（需配置 DASHSCOPE_API_KEY）**

```bash
python analyze_images.py "截图目录/" 实验内容结构.json image_map.json
```

### 阶段 3：生成报告

```bash
# 先出 MD 预览确认
python generate_report.py --content 实验内容结构.json --images image_map.json \
  --format-data 格式规范.json --template "模板.docx" --format md --output 预览

# 确认后生成最终 DOCX 或 PDF
python generate_report.py --content 实验内容结构.json --images image_map.json \
  --format-data 格式规范.json --latex-format 格式规范-LaTeX模板.tex \
  --template "模板.docx" --format docx --output 实验报告_最终版
```

## 文件说明

| 文件 | 用途 |
|------|------|
| `analyze_content.py` | 从 docx 模板提取章节结构和图片插入点 |
| `analyze_template.py` | 分析 docx 模板的字体/段落/页面格式 |
| `analyze_images.py` | 截图分类与映射（支持 AI 自读 + API 双模式） |
| `generate_report.py` | 最终报告生成（MD 预览 / DOCX / PDF） |
| `lib/docx_utils.py` | docx / JSON / LaTeX 公共工具函数 |
| `格式规范-LaTeX模板.tex` | PDF 输出的 LaTeX 格式模板 |
| `SKILL.md` | Claude Code 技能主文档 |

## 输出物

- `实验内容结构.json` — 章节结构 + 填充的答案文本
- `格式规范.json` — 字体/行距/页边距等格式配置
- `image_map.json` — 截图到插入点的路径映射
- `实验报告_最终版.docx` / `.pdf` — 生成好的实验报告

## 许可证

MIT
