# AGENTS.md - AI Agent 完整工作指引

本文件指导 AI Agent 如何在本项目中正确工作，包括架构理解、操作技巧、常见陷阱和完整工作流。

**重要提醒**：在开始编写实验代码前，必须询问用户使用哪种编程语言（Python/Java/C++/C等），然后根据用户选择的语言生成对应的代码。

---

## 一、项目架构概览

### 1.1 核心思路

本项目**不是从零生成 Word**，而是**基于已有模板填充内容**。模板 `实验报告.docx` 包含完整的页面布局、表格边框、字体样式，脚本只负责往对应的单元格里塞文字和图片。

这是关键设计决策：Word 排版极其复杂（字体回退、行距计算、表格跨页），从零生成几乎不可能达到模板的排版质量。

### 1.2 文件关系

```
用户放置（根目录）          脚本自动创建           脚本生成
─────────────────      ──────────────      ──────────
实验报告.docx ──────→ (模板，只读)
实验指导书.docx ─────→ (解析内容)
                     src/exp1/  ...       ← 放入源代码
                     output/exp1/  ...    ← 放入截图
                                          report/*.docx
```

### 1.3 关键文件说明

| 文件                 | 角色   | AI 操作方式                                          |
| -------------------- | ------ | ---------------------------------------------------- |
| `实验报告.docx`      | 模板   | **只读**，用 `Document()` 打开后另存，绝不修改原文件 |
| `实验指导书.docx`    | 内容源 | **只读**，用 python-docx 解析段落文本                |
| `generate_report.py` | 主脚本 | AI 的主要工作对象，需要理解和维护                    |
| `src/exp{N}/`        | 源代码 | AI 生成实验代码放入此目录                            |
| `output/exp{N}/`     | 截图   | AI 运行代码后生成截图放入此目录                      |
| `report/`            | 输出   | 脚本自动生成，不要手动修改                           |

---

## 二、模板结构（最重要，必须理解）

### 2.1 模板格式自动检测

脚本支持**两种模板格式**，自动识别：

| 格式       | 特征                             | 示例               |
| ---------- | -------------------------------- | ------------------ |
| **多表格** | 每个实验独立表格（6-10行 × N个） | Java程序设计模板   |
| **单表格** | 一个大表格包含所有实验（24+行）  | 算法分析与设计模板 |

检测逻辑：扫描所有表格，找包含"实验报告"标题且有实验字段（题目/目的）的表格。

### 2.2 通用行类型检测

脚本通过**扫描单元格文本关键词**自动识别每行的作用，而非硬编码行号：

| 行类型   | 匹配关键词                         | 操作           |
| -------- | ---------------------------------- | -------------- |
| title    | "教学上机实验报告"                 | 跳过（不修改） |
| date     | "上机时间"、"上机日期"、"实验日期" | 填充日期       |
| topic    | "实验题目"                         | 从指导书填充   |
| purpose  | "实验目的"、"实验要求"             | 从指导书填充   |
| process  | "实验过程"、"实验步骤"             | 填入代码       |
| result   | "实验结果"                         | 插入截图       |
| analysis | "实验分析"、"实验小结"、"实验心得" | 生成分析       |
| score    | "实验成绩"、"成绩"、"评分"         | 跳过（不修改） |

这意味着**行顺序不同也能正确填充**。

### 2.3 标准结构（典型8行/实验）

模板文件包含 **2 个 Table**：

| Table   | 行数      | 内容              | 能否修改         |
| ------- | --------- | ----------------- | ---------------- |
| Table 1 | 3 行      | 封面页 + 评分标准 | **绝对不能修改** |
| Table 2 | **32 行** | 4 个实验报告区域  | 填充内容         |

### 2.4 典型行结构（每个实验 8 行）

无论多表格还是单表格，每个实验的标准布局：

| 行类型   | 内容                       | 操作方式                   |
| -------- | -------------------------- | -------------------------- |
| title    | "教学上机实验报告"（标题） | **不修改**                 |
| date     | 上机时间                   | 替换日期文本               |
| topic    | 实验题目                   | 智能检测（见下方说明）     |
| purpose  | 实验目的和要求             | 智能检测（见下方说明）     |
| process  | 实验过程                   | 清空内容段落后写入代码     |
| result   | 实验结果                   | 清空内容段落后插入截图     |
| analysis | 实验分析                   | 清空内容段落后写入分析文本 |
| score    | 实验成绩 + 日期            | **不修改**（教师填写）     |

**注意**：行顺序可以不同，脚本按关键词自动匹配。

**智能检测逻辑**（实验题目 +2、实验目的和要求 +3）：

1. **用户未填写** → 从指导书提取内容填充
2. **用户已填写，内容相似** → 保留用户填写的内容
3. **用户已填写，内容不相似** → 打印警告信息，保留用户填写的内容，提示用户检查

相似度判断基于关键词重叠率（阈值30%），如果用户填写的内容与指导书内容相差甚远，会在控制台输出警告，方便用户排查。

### 2.3 绝对不要做的事

- **不要复制 Table 2**。模板已经预留了所有实验的行，直接填充即可。
- **不要删除或新增行**。表格行数必须保持 32 行。
- **不要修改 +0 和 +7 行**。这是标题和评分区域。

---

## 三、单元格操作技巧

### 3.1 清空单元格但保留标题

每个内容单元格（+2 到 +6）的第一个段落是标题（如"实验目的和要求："），后面是空段落。

正确做法：

```python
def clear_cell_keep_first_label(cell):
    paragraphs = cell.paragraphs
    for p in paragraphs[1:]:  # 保留第1个段落（标题）
        p._element.getparent().remove(p._element)
```

然后用 `cell.add_paragraph()` 添加新内容。

### 3.2 添加带格式的段落

```python
def add_paragraph_to_cell(cell, text, fmt):
    p = cell.add_paragraph()
    p.paragraph_format.line_spacing = fmt["line_spacing"]
    p.paragraph_format.first_line_indent = Pt(indent)
    run = p.add_run(text)
    run.font.size = Pt(fmt["size"])
    run.bold = fmt["bold"]
    # 中文字体需要操作 XML
    rpr = run._element.get_or_add_rPr()
    rFonts = rpr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rpr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), fmt["font_ea"])
```

**关键陷阱**：`run.font.name = "宋体"` 只设置西文字体，中文字体必须通过 XML 操作 `w:rFonts` 的 `w:eastAsia` 属性。

### 3.3 日期替换

日期行（+1）有复杂格式，不能删除重建。正确做法是清空 runs 后追加：

```python
for run in p.runs:
    run.text = ""  # 清空原有 runs
run = p.add_run(f"上机时间   {exp_date} ")
set_run_font(run, size=10.5)
```

---

## 四、图片处理

### 4.1 自动缩放逻辑

图片插入前必须计算缩放尺寸，否则超长截图会撑破页面：

```python
from PIL import Image as PILImage

with PILImage.open(image_path) as img:
    orig_w, orig_h = img.size

aspect = orig_h / orig_w
scaled_height = max_width_cm * aspect  # 按宽度算高度

if scaled_height <= max_height_cm:
    final_width = max_width_cm    # 高度没超限
else:
    final_width = max_height_cm / aspect  # 反推宽度
```

默认参数：`max_width_cm=15`，`max_height_cm=20`。

### 4.2 插入图片

```python
p = cell.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run()
run.add_picture(str(image_path), width=Cm(final_width))
```

**不要加图片说明**（如"图：xxx"），用户要求只放截图。

---

## 五、实验代码生成规范

### 5.1 代码要求

AI 为每个实验生成的代码应满足（支持多种编程语言）：

1. **可独立运行**：代码可直接执行，无外部依赖
   - Python: `python exp{N}/xxx.py`
   - Java: `javac exp{N}/*.java && java -cp exp{N} Main`
   - C/C++: `g++ exp{N}/*.cpp -o exp{N}/main && ./exp{N}/main`
2. **包含多组测试用例**：至少 4 组，覆盖正常、边界、特殊情况
3. **有验证机制**：能用暴力法或其他方法验证算法正确性
4. **输出清晰**：有分隔线、测试用例编号、输入输出对照
5. **代码量控制**：60-80 行为宜，不超过 100 行

### 5.2 输出格式模板

如果是常规验证形的问题，就根据文档要求输出

如果是算法问题代码应输出清晰的测试结果，格式如下：

```
==================================================
XXX问题 - YYY算法求解
==================================================

测试用例 1: ...
输入: ...
输出: ...

测试用例 2: ...
输入: ...
输出: ...

==================================================
所有测试用例执行完毕
==================================================
```

**不同语言的输出示例**：

- Python: `print("=" * 50)`
- Java: `System.out.println("=".repeat(50));`
- C++: `std::cout << std::string(50, '=') << std::endl;`
- C: `printf("==================================================\n");`

### 5.3 多文件项目

如果实验需要多个文件（如 Java 的多类文件），全部放入 `src/exp{N}/` 目录。脚本会自动读取所有代码文件并在报告中用 `# === 相对路径 ===` 标注。

---

## 六、终端截图生成

### 6.1 生成方法

运行代码 → 捕获输出 → 用 Pillow 渲染为黑底白字图片：

```python
from PIL import Image, ImageDraw, ImageFont

# 字体：必须用支持中文的字体
font = ImageFont.truetype(r'C:\Windows\Fonts\msyh.ttc', 16)

# 关键：用 splitlines() 处理换行，不要用 split('\n')
lines = text.splitlines()

# 背景色 (12,12,12)，文字色 (204,204,204)
img = Image.new('RGB', (width, height), (12, 12, 12))
draw = ImageDraw.Draw(img)
for line in lines:
    draw.text((28, y), line, fill=(204, 204, 204), font=font)
    y += 26
```

### 6.2 字体选择

| 字体                | 中文 | 等宽 | 适合场景         |
| ------------------- | ---- | ---- | ---------------- |
| msyh.ttc (微软雅黑) | ✓    | ✗    | 截图生成（推荐） |
| Consolas            | ✗    | ✓    | Word 中代码块    |
| simhei.ttf (黑体)   | ✓    | ✗    | 备选             |

**绝对不能用 Consolas 生成截图**——它不支持中文，会显示为方块。

### 6.3 常见陷阱

- `text.split('\n')` 在 Windows 下可能不分行（因为 `\r\n`），必须用 `text.splitlines()`
- 读取文件时用 `'rb'` 模式 + `decode('utf-8')` 比 `open(file, 'r')` 更可靠
- 截图文件名必须是 `run_result.png`（脚本通过 `*.png` glob 查找）

---

## 七、指导书解析

### 7.1 解析方法

指导书是 `.docx` 文件，用 python-docx 读取段落文本，通过正则匹配提取结构：

```python
doc = Document("实验指导书.docx")
for para in doc.paragraphs:
    text = para.text.strip()
    # 匹配 "实验X：名称"
    exp_match = re.match(r"实验([一二三四五六七八九十\d]+)[：:](.+)", text)
    # 匹配 "一、实验目的"
    section_match = re.match(r"^[一二三四五六七八九十]+[、.](.+)", text)
```

### 7.2 中文数字转换

指导书用中文数字（"实验一"），脚本内部用阿拉伯数字（exp1）。映射表：

```python
EXP_NUM_MAP = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8}
```

### 7.3 指导书可能更新

用户可能替换指导书文件（增减实验数量）。脚本会自动识别，不需要硬编码实验数量。

---

## 八、文件查找机制

脚本**不硬编码路径**，而是自动搜索文件：

```python
def find_file(name, search_dirs=None):
    """在根目录和一级子目录中查找文件"""
    for d in [BASE_DIR] + subdirs:
        if (d / name).exists():
            return d / name
```

所以用户可以把 `实验报告.docx` 和 `实验指导书.docx` 放在根目录或任何一级子目录中。

源码和截图目录支持多种命名：`exp1`、`experiment1`、`实验1`、`实验一` 都能识别。

---

## 九、调试指南

### 9.1 验证生成的报告

```bash
# 查看文档结构
officecli view report/xxx.docx outline

# 检查表格行数和列数
officecli query report/xxx.docx "table"

# 检查图片是否正确插入
officecli query report/xxx.docx "picture"

# 查看某个单元格内容
officecli get report/xxx.docx "/body/tbl[2]/tr[5]/tc[1]" --depth 1

# 生成 HTML 预览
officecli view report/xxx.docx html -o preview.html
```

### 9.2 常见问题排查

| 现象                 | 原因                                 | 解决方法                            |
| -------------------- | ------------------------------------ | ----------------------------------- |
| 中文显示为方块       | 字体不支持中文                       | 改用 msyh.ttc 或 simhei.ttf         |
| 截图占满整页         | 未做缩放处理                         | 检查 `add_image_to_cell` 的缩放逻辑 |
| 代码格式错乱         | 用了 `split('\n')`                   | 改用 `splitlines()`                 |
| 表格行数不对         | 误复制了表格                         | 不要复制 Table 2，直接填充 32 行    |
| 实验分析全是占位文本 | `_generate_analysis()` 未改进        | 修改该函数或接入 AI API             |
| 找不到指导书         | 文件名不对                           | 文件名必须是 `实验指导书.docx`      |
| 单元格标题消失       | `clear_cell_keep_first_label` 删多了 | 确认保留 paragraphs[0]              |

### 9.3 单实验测试

```bash
# 只生成实验二，快速验证
python generate_report.py --exp 2 --date "2026  年 3  月 30  日"

# 查看输出
officecli view report/实验二_报告.docx text
```

---

## 十、扩展指南

### 10.1 新增实验

1. 在指导书中添加新实验（脚本会自动识别）
2. 模板的 Table 2 需要增加 8 行（手动在 Word 中复制一组 8 行）
3. 在 `src/exp{N}/` 放入代码
4. 在 `output/exp{N}/` 放入截图
5. 运行 `python generate_report.py`

### 10.2 改进实验分析

修改 `generate_report.py` 中的 `_generate_analysis()` 函数。当前返回模板文本，可以改为：

- 接入 OpenAI / 其他 LLM API，传入代码和实验名称生成分析
- 对代码做静态分析（计算行数、函数数、循环嵌套层数）
- 根据算法类型填充特定的复杂度分析

### 10.3 修改格式

所有格式常量集中在 `Fmt` 类中，修改后全局生效。不要在各处硬编码字体字号。

---

## 十一、完整工作流（AI 执行清单）

当用户说"帮我写实验报告"或类似请求时，按以下顺序执行：

**步骤0：判断是否需要询问语言**

检查用户消息中是否已明确指定编程语言：

| 用户消息示例             | 是否包含语言 | Agent 动作                                                 |
| ------------------------ | ------------ | ---------------------------------------------------------- |
| "帮我写实验报告"         | ❌ 未指定    | 询问："请问您使用哪种编程语言？（Python/Java/C++/C/其他）" |
| "帮我写 Java 实验报告"   | ✅ Java      | 跳过询问，直接使用 Java                                    |
| "用 Python 生成实验报告" | ✅ Python    | 跳过询问，直接使用 Python                                  |
| "帮我写 C++ 的算法实验"  | ✅ C++       | 跳过询问，直接使用 C++                                     |

**完整流程**：

```
0. [语言] 判断用户是否已指定语言
      - 未指定 → 询问用户
      - 已指定 → 直接使用该语言
1. [检测] 检测用户环境是否满足要求（见下方环境检测指南）
2. [识别] 自动识别指导书和模板文件（见下方文件识别指南）
3. [解析] 读取实验指导书 → 提取实验列表
4. [分析] 读取实验报告模板 → 对比用户已填写的实验题目和实验目的与指导书是否一致
5. [编码] 为每个实验编写指定语言的代码 → 放入 src/exp{N}/
6. [运行] 执行所有代码 → 确认无报错
7. [截图] 用 Pillow 生成黑底终端截图 → 放入 output/exp{N}/run_result.png
8. [生成] 运行 python generate_report.py --merge --date "..."
9. [验证] 用 officecli 检查报告结构和图片
10. [交付] 呈现 report/ 下的最终报告
```

**语言识别关键词**：

| 关键词                 | 识别为     |
| ---------------------- | ---------- |
| Java、java、JAVA       | Java       |
| Python、python、py     | Python     |
| C++、cpp、C plus plus  | C++        |
| C语言、纯C、C（非C++） | C          |
| JavaScript、js         | JavaScript |
| Go、golang             | Go         |
| Rust、rust             | Rust       |

---

## 十二、环境检测指南（步骤1详细说明）

### 12.1 检测流程

在用户告知编程语言后，Agent 必须立即检测环境：

```
用户回答语言 → 检测 Python 环境 → 检测对应编译环境 → 判断结果 → 执行动作
```

### 12.2 Python 环境检测（始终需要）

**检测命令**：

```bash
# 检测 Python 版本
python --version

# 如果 python 不行，尝试 python3
python3 --version

# 查看所有已安装的 Python 版本（Windows）
where python

# 查看 Python 安装路径
python -c "import sys; print(sys.executable)"
```

**常见情况处理**：

| 情况       | 现象                          | Agent 处理                               |
| ---------- | ----------------------------- | ---------------------------------------- |
| 正常       | `Python 3.10.x`               | 告诉用户"Python 环境正常"，继续下一步    |
| 版本过低   | `Python 3.9.x` 或更低         | 提示用户需要 3.10+，询问是否需要帮忙配置 |
| 多版本混乱 | `where python` 显示 8+ 个路径 | 询问用户使用哪个版本，或帮忙清理环境变量 |
| 未安装     | `python 不是内部命令`         | 提供下载链接，指导用户安装               |

**多版本混乱时的处理**：

```bash
# Agent 执行以下命令查看所有 Python 路径
where python

# 典型输出示例：
# C:\Python39\python.exe
# C:\Python310\python.exe
# C:\Users\xxx\AppData\Local\Programs\Python\Python311\python.exe
# C:\Users\xxx\AppData\Local\Microsoft\WindowsApps\python.exe
# ... 可能有 8-10 个

# Agent 应该：
# 1. 询问用户："检测到多个 Python 版本，请问您想使用哪个？"
# 2. 列出所有路径让用户选择
# 3. 帮用户配置正确的 Python 到 PATH 最前面
```

### 12.3 Java 环境检测（用户选择 Java 时）

**检测命令**：

```bash
# 检测 Java 运行时
java -version

# 检测 Java 编译器
javac -version

# 查看 JAVA_HOME 环境变量
echo %JAVA_HOME%

# 查看所有 Java 安装路径（Windows）
where java

# 查看注册表中的 Java 安装信息
reg query "HKLM\SOFTWARE\JavaSoft" /s 2>nul
```

**常见情况处理**：

| 情况             | 现象                         | Agent 处理                               |
| ---------------- | ---------------------------- | ---------------------------------------- |
| 正常             | `java version "17.x.x"`      | 告诉用户"Java 环境正常"，继续下一步      |
| 版本过低         | `java version "1.8.x"`       | 建议升级到 JDK 17，询问是否需要帮忙      |
| 有 java 没 javac | `java` 可用但 `javac` 不可用 | 用户装的是 JRE 不是 JDK，需要安装 JDK    |
| 路径复杂         | `where java` 显示多个路径    | 询问用户使用哪个版本，帮忙配置 JAVA_HOME |
| 未安装           | `java 不是内部命令`          | 提供下载链接，指导用户安装 JDK 17        |

**路径复杂时的处理**：

```bash
# Agent 执行
where java

# 典型输出：
# C:\Program Files\Java\jdk-17\bin\java.exe
# C:\Program Files\Java\jdk-11\bin\java.exe
# C:\Program Files\Java\jre1.8.0_361\bin\java.exe
# C:\Program Files\Eclipse Adoptium\jdk-17\bin\java.exe

# Agent 应该：
# 1. 分析哪些是 JDK（有 javac），哪些是 JRE（只有 java）
# 2. 询问用户："检测到多个 Java 版本，建议使用 JDK 17，是否需要帮忙配置？"
# 3. 如果用户同意，帮忙设置 JAVA_HOME 和 PATH
```

**帮忙配置 JAVA_HOME（用户同意时）**：

```bash
# 设置 JAVA_HOME（假设用户选择 JDK 17 路径）
setx JAVA_HOME "C:\Program Files\Java\jdk-17"

# 将 JDK 的 bin 目录添加到 PATH 最前面
# 注意：需要用户手动操作或使用 PowerShell
[Environment]::SetEnvironmentVariable("Path", "C:\Program Files\Java\jdk-17\bin;" + $env:Path, "User")
```

### 12.4 C/C++ 环境检测（用户选择 C/C++ 时）

**检测命令**：

```bash
# 检测 g++ 编译器
g++ --version

# 检测 gcc 编译器
gcc --version

# 查看所有编译器路径
where g++
where gcc
where cl

# 检测 MSVC（如果安装了 Visual Studio）
cl
```

**常见情况处理**：

| 情况             | 现象                                                             | Agent 处理                            |
| ---------------- | ---------------------------------------------------------------- | ------------------------------------- |
| 正常             | `g++ (x86_64-posix-seh-rev0, Built by MinGW-W64 project) 12.x.x` | 告诉用户"C++ 环境正常"，继续下一步    |
| 有 MSVC 无 MinGW | `cl` 可用但 `g++` 不可用                                         | 询问用户是否安装 MinGW，或使用 MSVC   |
| 路径复杂         | `where g++` 显示多个路径                                         | 询问用户使用哪个版本                  |
| 未安装           | `g++ 不是内部命令`                                               | 提供 MinGW-w64 下载链接，指导用户安装 |

### 12.5 环境配置指导（用户需要时）

**Python 安装指导**：

```
1. 访问 https://www.python.org/downloads/
2. 下载 Python 3.10 或更高版本
3. 安装时务必勾选 "Add Python to PATH"
4. 安装完成后重启终端
5. 验证：python --version
```

**JDK 17 安装指导**：

```
推荐下载地址（优先使用国内镜像）：
1. 华为镜像（推荐，国内最快）：https://repo.huaweicloud.com/java/jdk/
2. Eclipse Adoptium（免费）：https://adoptium.net/
3. Oracle JDK：https://www.oracle.com/java/technologies/downloads/

安装步骤：
1. 下载 JDK 17 Windows x64 安装包
2. 运行安装程序，记住安装路径
3. 配置环境变量：
   - 新建系统变量 JAVA_HOME = 安装路径（如 C:\Program Files\Java\jdk-17）
   - 编辑 Path 变量，添加 %JAVA_HOME%\bin
4. 重启终端，验证：java -version && javac -version
```

**MinGW-w64 安装指导（C/C++）**：

```
推荐下载地址（优先使用国内镜像）：
1. 华为镜像（推荐）：https://repo.huaweicloud.com/mingw/
2. MSYS2（推荐）：https://www.msys2.org/
3. 官网：https://www.mingw-w64.org/

MSYS2 安装步骤：
1. 下载并安装 MSYS2
2. 打开 MSYS2 UCRT64 终端
3. 执行：pacman -S mingw-w64-ucrt-x86_64-gcc
4. 将 C:\msys64\ucrt64\bin 添加到系统 PATH
5. 重启终端，验证：g++ --version
```

**pip 镜像源配置（重要）**：

```
安装 Python 依赖时，务必使用国内镜像源，否则可能非常慢甚至失败。

临时使用镜像源：
pip install python-docx Pillow -i https://pypi.tuna.tsinghua.edu.cn/simple

永久配置镜像源（推荐）：
pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple

常用国内镜像源：
1. 清华源（推荐）：https://pypi.tuna.tsinghua.edu.cn/simple
2. 阿里源：https://mirrors.aliyun.com/pypi/simple/
3. 腾讯源：https://mirrors.cloud.tencent.com/pypi/simple
4. 华为源：https://repo.huaweicloud.com/repository/pypi/simple

如果下载速度慢，Agent 应该：
1. 自动使用 -i 参数指定镜像源
2. 提示用户可以永久配置镜像源
3. 如果一个镜像源慢，尝试换另一个
```

### 12.6 Agent 检测后的标准回复模板

**环境正常时**：

```
✅ 环境检测通过！

- Python: 3.10.11 (C:\Python310\python.exe)
- Java: JDK 17.0.5 (C:\Program Files\Java\jdk-17)

环境满足要求，我继续为您生成实验报告...
```

**环境有问题时**：

```
⚠️ 环境检测发现问题：

1. Python: 检测到 8 个版本，当前默认为 Python 3.9（版本过低）
   - C:\Python39\python.exe
   - C:\Python310\python.exe
   - C:\Users\xxx\AppData\Local\Programs\Python\Python312\python.exe
   - ... 共 8 个

2. Java: 未检测到 JDK（仅有 JRE 1.8）

请问：
1. Python 您想使用哪个版本？（建议选择 3.10 或更高版本）
2. 是否需要我帮您配置 JDK 17 环境？
```

**Agent 安装依赖时的标准做法**：

当需要安装 Python 依赖时，Agent 应该：

1. **优先使用国内镜像源**，避免下载慢或失败
2. 如果一个镜像源失败，自动尝试其他镜像源

```bash
# 推荐使用清华源
pip install python-docx Pillow -i https://pypi.tuna.tsinghua.edu.cn/simple

# 如果清华源失败，尝试阿里源
pip install python-docx Pillow -i https://mirrors.aliyun.com/pypi/simple/

# 如果还是失败，尝试华为源
pip install python-docx Pillow -i https://repo.huaweicloud.com/repository/pypi/simple
```

**Agent 应该告诉用户**：

```
正在安装 Python 依赖（使用国内镜像源加速）...
✅ python-docx 安装成功
✅ Pillow 安装成功
```

**命令行参数说明**：

```bash
# 默认模式（使用根目录下的模板和指导书）
python generate_report.py --merge --date "2026年6月8日"

# 指定模板和指导书路径
python generate_report.py --template "path/to/模板.docx" --instruction "path/to/指导书.docx" --merge --date "2026年6月8日" --output-dir "output_folder"

# 仅生成指定实验
python generate_report.py --exp 2 --date "2026年6月8日"
```

| 参数            | 说明                   | 默认值         |
| --------------- | ---------------------- | -------------- |
| `--template`    | 模板文件路径           | 根目录自动查找 |
| `--instruction` | 指导书文件路径         | 根目录自动查找 |
| `--output-dir`  | 输出目录               | `report/`      |
| `--merge`       | 合并所有实验到一份报告 | 否（单独生成） |
| `--exp N`       | 仅生成第N个实验        | 全部           |
| `--date`        | 上机日期               | 当天日期       |

**语言选择说明**：

- 在开始编写代码前，必须询问用户使用哪种编程语言
- 支持的语言：Python、Java、C++、C、JavaScript、Go、Rust 等
- 根据用户选择的语言生成对应的代码文件
- 不同语言的文件扩展名：`.py`、`.java`、`.cpp`、`.c`、`.js`、`.go`、`.rs`

**内容分析说明**（步骤2）：

- 读取模板中的实验题目和实验目的字段
- 与指导书中的内容进行对比分析
- 如果内容相似或一致 → 保留用户填写的内容
- 如果内容相差甚远（如指导书写"人口增长"，模板写"跳远"） → 询问用户应该参考哪个

每一步都必须确认无错后再进入下一步。特别注意第3步和第4步，这是最容易出问题的环节。

---

## 十三、文件识别指南（步骤2详细说明）

### 13.1 识别流程

Agent 应该主动识别用户工作目录中的指导书和模板文件，而不是依赖固定文件名：

```
用户说"帮我写实验报告" → Agent 扫描工作目录 → 读取文件内容 → 识别指导书和模板
```

### 13.2 识别方法

**第一步：列出目录中的 .docx 文件**

```bash
# 使用 list_dir 工具查看目录结构
# 或使用 file_search 搜索 *.docx 文件
```

**第二步：读取文件内容进行识别**

对于每个 .docx 文件，读取前 20 个段落，通过关键词判断文件类型：

| 文件类型     | 识别关键词                                             | 示例                         |
| ------------ | ------------------------------------------------------ | ---------------------------- |
| 实验指导书   | "实验指导书"、"实验目的"、"实验内容"、"实验原理"       | 《算法分析与设计》实验指导书 |
| 实验报告模板 | "教学上机实验报告"、"实验题目"、"实验过程"、"实验成绩" | 河南理工大学教学上机实验报告 |

**第三步：确认识别结果**

Agent 应该告诉用户识别结果：

```
✅ 已识别文件：
- 指导书：实验指导书.docx（包含 4 个实验）
- 模板：实验报告.docx（单表格格式，32行）
```

### 13.3 常见文件命名变体

用户可能使用以下命名方式，Agent 都应该能识别：

| 文件类型     | 可能的命名                                                               |
| ------------ | ------------------------------------------------------------------------ |
| 实验指导书   | `实验指导书.docx`、`指导书.docx`、`算法实验指导书.docx`、`实验指导.docx` |
| 实验报告模板 | `实验报告.docx`、`报告模板.docx`、`算法实验报告.docx`、`模板.docx`       |

### 13.4 识别失败处理

如果 Agent 无法识别文件：

1. **列出找到的所有 .docx 文件**
2. **询问用户**："我找到了以下文件，请告诉我哪个是指导书，哪个是报告模板？"
3. **让用户确认**后继续

### 13.5 Agent 识别后的标准回复模板

**识别成功时**：

```
✅ 文件识别完成！

📁 找到的文件：
- 指导书：实验指导书.docx
  → 包含 4 个实验：递归算法、分治策略、动态规划、贪心算法
- 模板：实验报告.docx
  → 格式：单表格（32行，可容纳4个实验）

环境正常，我继续为您生成实验报告...
```

**识别失败时**：

```
⚠️ 文件识别遇到问题：

📁 找到的 .docx 文件：
1. 文档1.docx（无法识别类型）
2. 实验材料.docx（无法识别类型）

请告诉我：
1. 哪个文件是实验指导书？
2. 哪个文件是实验报告模板？
```
