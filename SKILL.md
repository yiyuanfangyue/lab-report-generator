---
name: lab-report-generator
description: Use when generating formal experiment/lab reports from Word templates, experiment content documents, and screenshots. The user provides a template .docx and screenshots.
---

# Lab Report Generator

## 核心原则

**本技能支持两种运行模式，在流程开始时让用户选择：**

- **全自动模式（Auto）** — 用户只需提供模板和内容，从阶段0到阶段3一键执行，无确认、无MD预览，直接出最终报告
- **交互式模式（Interactive）** — 每个阶段完成后展示结果，等用户确认后再进入下一阶段

**无论哪种模式，阶段 0 的输入文件始终是模板 .docx（定义章节目录）而不是实验指导书（提供填充内容）。两者角色不同，不能混淆。**

---

## 模式选择

**第一步，问用户：**"请选择模式：全自动（auto）还是交互式（interactive）？"

### 全自动模式

**全自动模式 = 一次性跑完阶段0→1→2→3，无中间确认直接出最终报告。**

流程：

1. **问用户提供以下信息（一次性收集）：**
   - 实验报告模板 .docx 路径
   - 截图目录路径
   - 输出格式（docx / pdf / both）
   - （可选）实验指导书或答案内容说明

2. **阶段 0：** 分析模板结构 → 填写作业题答案（由你根据专业知识或指导书内容填入 JSON）
   ```bash
   python "skills/lab-report-generator/analyze_content.py" --yes "<模板路径>" 实验内容结构.json
   ```

3. **阶段 1：** 分析模板格式（跳过交互确认）
   ```bash
   python "skills/lab-report-generator/analyze_template.py" --yes "<模板路径>" 格式规范
   ```
   **运行后读取 `格式规范.json` 检查 `has_cover` 字段**，告知用户封面检测结果。如果不对，手动编辑 `格式规范.json` 将 `has_cover` 改为 `true` 或 `false`，**禁止重跑脚本**。

4. **阶段 2：** 列出截图，用 Read 工具逐一识别后手动填入映射
   ```bash
   python "skills/lab-report-generator/analyze_images.py" --model-vision "<截图目录>" 实验内容结构.json image_map.json
   ```
   然后用 Read 工具逐张查看截图，编辑 image_map.json 填入路径。

5. **阶段 3：** 直接生成最终报告（跳过 MD 预览）
   ```bash
   python "skills/lab-report-generator/generate_report.py" \
     --content 实验内容结构.json --images image_map.json \
     --format-data 格式规范.json --latex-format 格式规范-LaTeX模板.tex \
     --template "<模板路径>" --format <docx|pdf|both> --output 实验报告_最终版
   ```

6. 告知用户文件路径

### 交互式模式

**交互式模式 = 每阶段展示结果 → 用户确认 → 进入下一阶段。用户说"继续"之前，绝不动手。**

**先生成 MD 预览给用户确认内容，再出最终文件。**

**禁止自行退出 skill 流程：一旦进入本 skill，必须按阶段顺序执行直到用户确认完成。不能中途切换到其他无关操作（如直接修改脚本、生成文件）。**

---

## 交互式模式工作流程

### 阶段 0：实验内容分析

**目标：从模板文件中提取报告的章节结构（实验目的→实验原理→实验步骤→实验作业等）**

1. 问用户："请上传你的实验报告模板 .docx"
2. 运行 `analyze_content.py` **在模板文件上**（不是实验指导书/内容文件）：
   ```bash
   python "skills/lab-report-generator/analyze_content.py" --yes "<模板路径>" 实验内容结构.json
   ```
3. 展示：章节标题、子节（heading2/heading3）、图片插入点列表
4. **不要问用户"选哪个实验"之类的问题**——模板结构就是报告结构，直接展示
5. 问用户："内容结构正确吗？"
6. 将实验指导书（如果有）的内容填充到对应章节中：
   - 在 `实验内容结构.json` 的对应章节 children 中添加 body 文本
   - 从指导书中提取实验目的、原理、步骤文字填入
   - 对于模板中原有的需要学生填写的空白项（如填空、问答题），**保留空白**，让用户后续手写
   - 对于模板中原有的作业题的题干，**保留题干文字**
   - 如需代写作业题答案，先问用户是否要帮忙写

**关键区分：**
- 模板文件（文件名含"模板"）→ 定义章节结构，作为 analyze_content.py 的输入
- 实验指导书（文件名含"指导"）→ 提供填充文本内容，手动提取并填入 JSON

### 阶段 1：模板格式分析

**自动检测模板是否有封面页。** 检测逻辑：
- 第一页是否含图片（如校徽）
- 是否有大字号居中文字（>= 13pt）
- 是否含有"大学/学院/实验报告/姓名/学号"等封面关键词
- 满足2条以上视为有封面

1. 问用户："请上传你的实验报告模板 .docx"
2. 运行：
   ```bash
   python "skills/lab-report-generator/analyze_template.py" "<路径>" 格式规范
   ```
3. 展示格式摘要（含封面判定结果）
4. 问用户："封面判定正确吗？"（如果自动判定不准，可让用户翻转）
5. 问用户："格式正确吗？"
6. 如需调字号/字体 → **直接编辑格式规范.json，禁止重跑 analyze_template.py**
   - `body_format.normal.font.size` = 12.0（小四）
   - `body_format.heading1.font.size` = 18.0（小二）

### 阶段 2：截图分析

**本阶段支持两种模式，根据 AI 模型是否具备原生读图能力选择：**

#### 模式 A：模型自身读图（推荐，支持原生视觉的模型如 GPT-4V、Claude Vision 等）

**流程：模型用自己的 Read 工具逐一查看截图，自主建立映射。**

1. 问用户："请提供截图目录路径"
2. 运行 `list_images` 模式，列出所有截图和插入点，输出空映射模板：
   ```bash
   python "skills/lab-report-generator/analyze_images.py" \
     --model-vision "<截图目录>" 实验内容结构.json image_map.json
   ```
3. **用 Read 工具逐一查看每张截图**（一次可看多张），确认每张截图的内容
4. 根据识图结果，手动编辑 `image_map.json`，将截图路径填入对应的插入点
   - 单图：`"img_01": "完整/路径/截图.png"`
   - 多图：`"img_03": ["路径1", "路径2", ...]`
5. 展示映射结果给用户
6. 问用户："映射正确吗？"
7. 如需调整，直接编辑 image_map.json

#### 模式 B：外部视觉 API（模型无原生读图能力时）

**流程：通过 vision.js（DashScope 千问VL）自动识别截图内容。**

1. 问用户："请提供截图目录路径"
2. 运行：
   ```bash
   python "skills/lab-report-generator/analyze_images.py" "<截图目录>" 实验内容结构.json image_map.json
   ```
3. 展示匹配结果
4. 问用户："映射正确吗？"
5. 如需调整，直接编辑 image_map.json

### 阶段 3：报告生成（关键步骤）

**必须先出 MD 预览，不能跳过！**

1. **先生成 MD 预览**：
   ```bash
   python "skills/lab-report-generator/generate_report.py" \
     --content 实验内容结构.json --images image_map.json \
     --format-data 格式规范.json \
     --template "模板.docx" \
     --format md --output 实验报告预览
   ```
2. **展示 MD 预览内容给用户，让用户确认**
3. **用户确认 MD 后，必须问用户要什么输出格式**：
   - 用 `input()` 或阻塞方式问："请选择输出格式：pdf、docx、both？"
   - **绝对不能**自作主张用 `--format both` 或任何默认值
   - **PDF 需要安装 MiKTeX**（含 lualatex/xelatex 编译器）；如果未安装，生成时会自动提示并提供下载链接
4. **根据用户选择生成最终文件**：
   ```bash
   # DOCX 需要 --template 和 --format-data
   # PDF 需要 --latex-format
   python "skills/lab-report-generator/generate_report.py" \
     --content 实验内容结构.json --images image_map.json \
     --format-data 格式规范.json \
     --latex-format 格式规范.json-LaTeX模板.tex \
     --template "模板.docx" \
     --format <用户选择> --output 输出文件名
   ```
5. 告知用户文件路径

### 关于 image_placeholder 的图注规则

**image_placeholder 的 `text` 字段会被用作图片的 caption（图注）。**

```python
# ✅ 正确：text = 简短图注，长说明文字单独用 body 段落
{"type": "body", "text": "①粘贴经修改过这两个参数的网元设备配置页的截图..."},  # 指令说明
{"type": "image_placeholder", "id": "xxx", "text": "网元设备配置页截图", "caption": "网元设备配置页截图"},  # 简短的图注

# ❌ 错误：text 放长段落 → 图注变成冗长的指令文字
{"type": "image_placeholder", "id": "xxx", "text": "①粘贴经修改过这两个参数的...", "caption": "..."},
```

**图注应该：**
- 简洁描述图片内容本身（如"网元设备配置页截图"）
- 不是粘贴要求（如"①粘贴经修改过这两个参数的网元设备配置页的截图"）
- 不是操作说明（如"截图要包括仿真平台左上角的学号信息"）

操作说明应放在图片前面的 body 段落中。

## 禁止项

以下禁止项区分模式：

### [交互式模式] 禁止跳过 MD 预览
阶段 3 必须先出 MD，用户确认后再出 DOCX/PDF。

**No exceptions:**
- "内容很简单不需要预览" → 不行
- "用户要的是最终文件" → 先出预览再出最终文件
- "之前已经确认过内容了" → 生成过程中的内容结构可能变化，必须再看一次
- "图片映射都确认过了直接生成吧" → 不行，MD 预览是内容确认，不是图片确认
- "上个用户刚确认过类似的内容" → 每次都不一样

### [交互式模式] 禁止问完不等人
问了用户问题（如"要什么格式？"）后，必须用 input() 或阻塞等待用户输入。不能自作主张执行。

**No exceptions:**
- "用户之前说过要 both" → 这次可能变，必须再问
- "用户没回就是默认" → 必须等到明确回答
- "先跑再问节约时间" → 跑错了浪费更多时间
- "both 是最全的" → 用户可能有不同需求

### [交互式模式] 禁止退出 skill 流程
一旦进入交互式模式，必须按阶段顺序执行，直到用户确认完成。

**No exceptions:**
- "这个 bug 修了，改一下脚本就行" → 先记录，走完流程再修
- "我去搜索一下这个问题的原因" → 在 skill 框架内进行
- "先出文件，用户等不及了" → 按阶段来，跳过=重做
- "我知道需求了直接生成" → 再急也要按流程

### [交互式模式] 禁止用 `--format both` 作为默认值
阶段 3 必须问用户要什么格式，`input()` 阻塞等待。

**No exceptions:**
- "用户之前要过 both" → 每次都要问
- "both 最安全" → 生成两个文件可能更慢

### [全模式通用] 禁止重跑 analyze_template.py
一旦用户确认格式后，任何时候都不要重跑此脚本。改格式直接改格式规范.json。

**No exceptions:**
- "只是更新一下" → 不行
- "上次跑的时候 has_image 是错的" → 手动改格式规范.json 中的 cover_paragraphs[0].has_image 即可
- "需要更新 XPath" → 改脚本，不是重跑

### [全模式通用] 禁止在阶段0用指导书当输入
`analyze_content.py` 的输入应该是**模板 .docx**（定义章节结构），而不是实验指导书（提供填充内容）。

**No exceptions:**
- "指导书内容更丰富" → 结构从模板来，内容手动填入
- "模板只有1个章节" → 模板结构就是报告结构，不要自己编造章节
- "我想把5个实验都放进去" → 如果模板只定义了1套结构，就按1套来

### [全模式通用] 禁止问实验选择问题
模板的章节结构就是报告的结构。不要在阶段0问用户"选哪个实验"。

**No exceptions:**
- "指导书有5个实验" → 模板有几次实验结构就生成几次
- "用户可能想合并" → 先展示模板结构，让用户自己说要不要改

### 禁止在 Python f-string 中写 LaTeX 命令
f-string 中 `\b`、`\t`、`\n` 等会被 Python 解释为转义字符。

**禁止这样写：**
```python
f'    format     = \\\\zihao{{{h1_zh}}}\\bfseries\\setstretch{{1.5}},'
# Python 输出: \\zihao{-2}\bfseries\setstretch{1.5}
# LaTeX 把 \\ 解释为换行，\b 被 f-string 吃掉 → zihao-2
```

**必须这样写：**
```python
'    format     = \\zihao{' + str(h1_zh) + r'}\bfseries\setstretch{1.5},'
# 用 r'' raw string + 拼接，保证反斜杠原样传给 LaTeX
```

### 禁止自行搜索文件系统
每一步都必须问用户路径。

### 禁止合并阶段步骤
完成一个阶段 → 展示 → 等待确认 → 下一个。

## 红牌自检清单

出现以下任一情况，立刻 STOP：

### 交互式模式红牌
- "格式已经确认过了，我直接改一下 JSON 就行，不用再展示" → STOP，展示修改后的结果
- "这个 bug 修了，重跑一下分析脚本更新数据" → STOP，重跑会覆盖用户手动改的值
- "内容很简单，直接出 DOCX 吧" → STOP，先出 MD 预览
- "我知道模板在哪里，不用问用户了" → STOP，必须问
- "用户之前说过格式正确，这次不展示了" → STOP，每次都展示
- "用户之前说要 both，我不用再问就直接执行" → STOP，每次都要问
- "用户已经确认了图片映射，直接生成最终文件吧" → STOP，先出 MD 预览
- "我先把图注改简短，不涉及内容结构问题不大" → STOP，MD 预览后才能生成
- "先 both 都出了，用户要哪个用哪个" → STOP，必须问用户

### 全模式通用红牌
- "反斜杠多加几层总能转义对" → STOP，用 r'' raw string 拼接
- "我把指导书跑一下 analyze_content 看看结构" → STOP，analyze_content 跑的是模板不是指导书
- "用户说全部实验都要，我合并一下" → STOP，模板有几套结构就生成几套
- "我跳出流程查一下这个问题" → STOP，在 skill 框架内解决问题
- "全自动模式太麻烦，我在交互式里跳过确认吧" → STOP，选了交互式就必须走完确认流程

## 已知边界情况

| 症状 | 原因 | 解决 |
|------|------|------|
| json.load 炸了（JSONDecodeError） | JSON 中有未转义的中文引号 " | 确保文本中的引号用 `“` `”` 或 “ ” |
| P0 校徽没出现 | `findall('w:drawing')` 少了 `.//` | XPath 要用 `.//` 查找所有后代 |
| lualatex 线程报错但 PDF 生成成功 | `capture_output=True, text=True` 在 GBK 环境有问题 | 去掉 text=True，用二进制 |
| PDF 只有 2 页 | 模板中有 `\end{document}` 提前终止 | generate_report.py 已自动截断模板中的 end{document} |
| 图片没出现在 PDF 里 | PNG 是 RGBA 模式 | generate_report.py 复制时自动转 RGB JPEG |
| 正文行距没变 | 行距值单位是半磅不是磅 | 固定 20 磅 = line=400（不是 20） |
| 正文行距固定值不生效 | 格式规范中 `line_rule` 是 `"exactly"` 但代码比较 `== 'exact'` | `generate_report.py:629` 把 `== 'exact'` 改为 `str(body_line_rule).startswith('exact')` |
| 封面字体变回默认 | 重跑了 analyze_template.py 覆盖了手动修改 | 严禁重跑！直接改格式规范.json |
| 标题显示 "zihao-2" 字样而非实际字号 | `\\\\zihao` 在 f-string 中转义成 `\\zihao`，LaTeX 把 `\\` 当换行 | 用 `'...' + str() + r'...'` 拼接，不用 f-string |
| `\bfseries` 变成了 `\x08fseries` | f-string 中 `\b` 被解释为 backspace 退格符 | 同上，用 r'' raw string |
| PDF 封面没有校徽 | 模板中硬编码 `logo.jpeg` 但没从 docx 提取 | generate_report.py 复制图片时自动从模板 media 提取 |
| 问了格式选择但没等回答就执行 | 没有用阻塞式 input 等待 | 用 input() 等待用户输入，不能假设默认值 |
| 图注显示长段落指令文字（如"①粘贴经修改过这两个参数的..."） | image_placeholder 的 `text` 字段被用作 caption，而 text 中放了长说明 | 把长说明文字移到前面的 body 段落中，`text` 只放简短图注（如"网元设备配置页截图"） |
| 阶段0后用户说"不对，你应该用模板" | 用实验指导书跑了 analyze_content.py 而不是模板文件 | 牢记：analyze_content.py 的输入是**模板 .docx**，不是指导书 |
| 用户说"你怎么跳出流程了" | 中途切换到脚本修改、文件搜索等无关操作 | 禁止退出 skill 流程，所有操作在 skill 框架内进行 |
