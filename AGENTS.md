# AGENTS.md - AI Agent 工作指引

本文件指导 AI Agent 如何在本项目中正确工作。

**核心提醒**：在开始编写实验代码前，必须确认用户使用的编程语言。如果用户未指定，主动询问。

---

## 一、项目架构

本项目**基于已有 Word 模板填充内容**，不从零生成 Word。模板包含页面布局和表格边框，脚本只负责往对应单元格塞文字和图片。

```
用户放置                    脚本生成
─────────                ──────────
实验报告.docx (模板，只读)
实验指导书.docx (内容源，只读)
src/exp{N}/     ← 放代码
output/exp{N}/  ← 放截图   report/*.docx (输出)
```

**绝对不要做的事**：不要修改模板原文件、不要复制/删除/新增表格行、不要修改标题行和成绩行。

---

## 二、模板行类型（自动检测，无需硬编码行号）

脚本通过扫描单元格关键词自动识别每行用途，行顺序可以不同：

| 行类型   | 关键词                               | 操作       |
| -------- | ------------------------------------ | ---------- |
| title    | "教学上机实验报告"                   | **不修改** |
| date     | "上机时间"、"上机日期"               | 填日期     |
| topic    | "实验题目"                           | 指导书填充 |
| purpose  | "实验目的"、"实验要求"               | 指导书填充 |
| process  | "实验过程"、"实验步骤"               | 填入代码   |
| result   | "实验结果"                           | 插入截图   |
| analysis | "实验分析"、"实验小结"               | Agent写analysis.txt，脚本读取 |
| score    | "实验成绩"、"评分"                   | **不修改** |

模板格式自动检测：多表格（每个实验独立表格）和单表格（一个大表格），脚本自动识别。

---

## 三、python-docx 关键陷阱

**中文字体**：`run.font.name = "宋体"` 只设置西文。中文字体必须操作 XML：

```python
rpr = run._element.get_or_add_rPr()
rFonts = rpr.find(qn("w:rFonts"))
if rFonts is None:
    rFonts = OxmlElement("w:rFonts")
    rpr.insert(0, rFonts)
rFonts.set(qn("w:eastAsia"), "宋体")
```

**清空单元格**：保留第一个段落（标题），删除后续段落：

```python
for p in cell.paragraphs[1:]:
    p._element.getparent().remove(p._element)
```

**日期替换**：不能删除重建，必须清空 runs 后追加：

```python
for run in p.runs:
    run.text = ""
run = p.add_run(f"上机时间   {exp_date} ")
```

**图片插入不加说明文字**，只放截图。

---

## 四、代码生成规范

为每个实验生成的代码应满足：

1. **可独立运行**，无外部依赖
2. **至少 4 组测试用例**，覆盖正常、边界、特殊情况
3. **输出清晰**：有分隔线、用例编号、输入输出对照
4. **代码量**：60-80 行为宜，不超过 100 行

输出格式示例（算法类问题）：

```
==================================================
XXX问题 - YYY算法求解
==================================================

测试用例 1: ...
输入: ...
输出: ...

==================================================
所有测试用例执行完毕
==================================================
```

各语言分隔线：Python `print("=" * 50)` / Java `System.out.println("=".repeat(50))` / C++ `std::cout << std::string(50, '=') << std::endl`

多文件项目（如 Java）全部放入 `src/exp{N}/`，脚本自动读取并用 `# === 相对路径 ===` 标注。

---

## 五、终端截图生成

使用项目内置的 `capture_output.py`，**不要自己写 Pillow 渲染代码**：

```bash
# 运行命令并截图
python capture_output.py --cmd "java -cp src/exp1 Main" --output output/exp1/run_result.png

# 从文本文件生成截图
python capture_output.py --file output/exp1/output.txt --output output/exp1/run_result.png

# 直接传入文本生成截图
python capture_output.py --text "输出内容..." --output output/exp1/run_result.png
```

截图效果：黑底（#0C0C0C）白字（#CCCCCC），微软雅黑字体，支持中文。

---

## 六、完整工作流（AI 执行清单）

当用户说"帮我写实验报告"或类似请求时，按以下顺序执行：

### 步骤 0：确认编程语言

检查用户消息中是否已指定语言。未指定则询问："请问您使用哪种编程语言？（Python/Java/C++/C/其他）"

### 步骤 1：检测环境

运行环境检测脚本，或手动检查：

```bash
python check_environment.py         # 全面检测（Python/Java/C++/磁盘）
python --version                    # 快速检查 Python
java -version && javac -version     # 快速检查 Java
```

环境有问题时，帮用户配置或指导安装。安装 pip 依赖时使用国内镜像源：
```bash
pip install python-docx Pillow -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 步骤 2：识别文件

脚本会自动查找模板和指导书（不依赖固定文件名）。Agent 也可以用 `--dry-run` 预览：

```bash
python generate_report.py --dry-run --merge --date "2026年6月8日"
```

如果识别失败，列出所有 .docx 文件让用户确认哪个是模板、哪个是指导书。

### 步骤 3：解析指导书

`generate_report.py` 会自动解析指导书提取实验列表。Agent 通过 `--dry-run` 输出确认解析结果。

### 步骤 4：生成代码

为每个实验编写代码，放入 `src/exp{N}/`。确认代码可运行、无报错。

### 步骤 5：运行、截图、写分析

运行每个实验的代码，用 `capture_output.py` 生成截图，并为每个实验写一份实验分析：

```bash
# 对每个实验执行
python capture_output.py --cmd "<运行命令>" --output output/exp{N}/run_result.png
```

然后为每个实验写一份 `output/exp{N}/analysis.txt`，内容是你作为 AI 对这个实验的真实分析。

**分析格式（3段，共150-200字）：**

```
一、实验内容分析
（做了什么、用了什么技术、结果如何——2~3句话概括）

二、关键知识点
（本实验涉及的核心概念——如继承、封装、递归等，结合代码点拨——3~4句话）

三、实验心得
（调试中遇到的具体问题或收获——1~2句话，不要套话）
```

**写作要求：**
- 必须基于实际代码和运行结果，不要写模板套话
- 根据课程类型调整侧重点：
  - **程序设计/OOP课程**：侧重类设计、封装、继承、多态、接口等面向对象概念
  - **算法课程**：侧重算法正确性、复杂度分析、边界条件处理
- 简洁为主，每段不超过4句话

脚本会优先读取 `analysis.txt` 填入报告。如果没有这个文件，会回退到脚本内置的模板分析。

### 步骤 6：生成报告

```bash
python generate_report.py --merge --date "2026年6月8日"
```

支持的日期格式：`"2026年6月8日"`、`"2026-06-08"`、`"2026/06/08"` 均可。

### 步骤 7：验证并交付

检查 `report/` 下生成的报告文件。向用户报告：

- 生成了哪些文件
- 每个实验包含什么内容（代码、截图数量）
- 如有实验生成失败，说明原因

---

## 七、关键命令行参数速查

```bash
python generate_report.py --merge --date "2026年6月8日"    # 合并所有实验
python generate_report.py --exp 2 --date "2026年6月8日"    # 仅生成实验二
python generate_report.py --dry-run --merge                 # 预览模式
python generate_report.py --verbose --merge                 # 详细日志
python generate_report.py --template "x.docx" --instruction "y.docx"  # 指定文件
```

更多参数说明见 `python generate_report.py --help` 或 readme.md。
