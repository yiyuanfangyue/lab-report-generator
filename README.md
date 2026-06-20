# Lab Report Generator

**Claude Code 技能** — 基于 Word 模板 + 实验截图自动生成格式化实验报告（DOCX/PDF）。

从实验报告模板中提取章节结构（实验目的→实验原理→实验步骤→实验作业），填充答案和截图，生成格式规范的完整报告。

## 特性

- **双运行模式**：全自动模式（一键出报告） / 交互式模式（逐阶段确认）
- **智能封面检测**：自动判断模板是否有封面页（校徽/大字号/封面关键词）
- **双截图识别方式**：AI 模型自身读图 / 外部视觉 API（DashScope 千问VL）
- **图注自动生成**：根据图片数量和描述文字自动编号，无需硬编码
- **路径完整性校验**：生成前自动检查所有图片路径是否有效
- **输出格式**：MD 预览 / DOCX / PDF

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
2. 设置环境变量 `DASHSCOPE_API_KEY=your-key`

## 用法

### 全自动模式（Auto）

一次性跑完所有阶段，无中间确认，直接出最终报告：

```bash
# 阶段0：分析模板结构
python analyze_content.py --yes "实验报告模板.docx" 实验内容结构.json

# 阶段1：分析模板格式（跳过交互）
python analyze_template.py --yes "实验报告模板.docx" 格式规范
# → 检查 格式规范.json 中的 has_cover 字段，不对则用 Edit 工具修正

# 阶段2：截图映射
python analyze_images.py --model-vision "截图目录/" 实验内容结构.json image_map.json
# → 用 Read 工具查看截图，用 Edit 工具编辑 image_map.json 填入路径

# 阶段3：直接生成最终报告（跳过 MD 预览）
python generate_report.py --content 实验内容结构.json --images image_map.json \
  --format-data 格式规范.json --latex-format 格式规范-LaTeX模板.tex \
  --template "模板.docx" --format docx --output 实验报告_最终版
```

### 交互式模式（Interactive）

每阶段完成后展示结果，等待用户确认后进入下一阶段：

```bash
# 阶段0：分析模板结构
python analyze_content.py --yes "实验报告模板.docx" 实验内容结构.json

# 阶段1：分析模板格式（可交互调整字号）
python analyze_template.py "实验报告模板.docx" 格式规范

# 阶段2：截图分析（二选一）
python analyze_images.py --model-vision "截图目录/" 实验内容结构.json image_map.json  # AI自读
python analyze_images.py "截图目录/" 实验内容结构.json image_map.json               # 外部API

# 阶段3：先出 MD 预览，确认后再生成最终文件
python generate_report.py --content 实验内容结构.json --images image_map.json \
  --format-data 格式规范.json --template "模板.docx" --format md --output 预览
python generate_report.py --content 实验内容结构.json --images image_map.json \
  --format-data 格式规范.json --latex-format 格式规范-LaTeX模板.tex \
  --template "模板.docx" --format docx --output 实验报告_最终版
```

### 封面检测

`analyze_template.py` 会自动检测模板是否有封面页，判断依据：
- 首段是否含图片（如校徽）
- 是否有大字号居中文字（>= 13pt）
- 是否含有封面关键词（"大学/学院/实验报告/姓名/学号"等）
- 满足条件自动标记 `has_cover: true/false`

交互式模式中用户可翻转判定结果；全自动模式中 AI 检查 JSON 后用 Edit 工具修正。

## 文件说明

| 文件 | 用途 |
|------|------|
| `analyze_content.py` | 从 docx 模板提取章节结构和图片插入点 |
| `analyze_template.py` | 分析格式 + 封面检测（支持 `--yes` 跳过交互） |
| `analyze_images.py` | 截图映射（`--model-vision` AI 自读 / 普通 API 模式） |
| `generate_report.py` | 报告生成（MD 预览 / DOCX / PDF）+ 路径校验 |
| `lib/docx_utils.py` | docx / JSON / LaTeX 公共工具函数 |
| `格式规范-LaTeX模板.tex` | PDF 输出的 LaTeX 格式模板 |
| `SKILL.md` | Claude Code 技能主文档（含完整使用规范） |

## 输出物

| 文件 | 说明 |
|------|------|
| `实验内容结构.json` | 章节结构 + 填充的答案文本 |
| `格式规范.json` | 字体/行距/页边距等格式配置（含 `has_cover`） |
| `image_map.json` | 截图到插入点的路径映射 |
| `实验报告_最终版.docx/.pdf` | 生成好的实验报告 |

## 许可证

MIT
