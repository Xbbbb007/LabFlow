# 大学生实验报告自动化生成

自动生成《算法分析与设计》课程实验报告。

基于 Word 模板填充，封面页和评分页保持原样，脚本负责生成各实验内容页。

## 环境要求

### 运行依赖（生成报告必需）

| 工具        | 版本     | 用途                                          |
| ----------- | -------- | --------------------------------------------- |
| Python      | >= 3.10  | 运行环境                                      |
| python-docx | >= 1.2.0 | Word 文件操作（模板读取、表格填充、图片插入） |
| Pillow      | >= 10.0  | 读取图片原始尺寸，用于自动缩放计算            |

安装依赖：

```bash
pip install python-docx Pillow
```

### 字体依赖（Windows 系统自带）

报告模板和截图生成依赖以下字体，Windows 系统已预装：

| 字体                       | 用途                 |
| -------------------------- | -------------------- |
| 宋体 (SimSun)              | 报告正文段落         |
| 黑体 (SimHei)              | 报告标题和小节标题   |
| Times New Roman            | 报告西文字体         |
| Consolas                   | 代码块、终端截图     |
| 微软雅黑 (Microsoft YaHei) | 截图生成时的中文渲染 |

如在非 Windows 环境运行，需手动安装上述字体。

### 开发调试工具（可选）

| 工具                                | 用途                                         |
| ----------------------------------- | -------------------------------------------- |
| [officecli](https://d.officecli.ai) | 检查 docx 结构、验证输出质量、生成 HTML 预览 |
| Word / WPS                          | 打开生成的报告确认排版效果                   |

officecli 安装：

```bash
# Windows
irm https://d.officecli.ai/install.ps1 | iex

# macOS / Linux
curl -fsSL https://d.officecli.ai/install.sh | bash
```

## 编程语言环境配置

本工具支持多种编程语言，以下是各语言的环境配置指南：

### Python 环境

**检查是否已安装**：

```bash
python --version
# 或
python3 --version
```

**安装方法**：

1. **Windows**：从 [Python官网](https://www.python.org/downloads/) 下载安装包，安装时勾选"Add Python to PATH"
2. **macOS**：使用 Homebrew：`brew install python`
3. **Linux**：使用包管理器：`sudo apt install python3` (Ubuntu/Debian) 或 `sudo yum install python3` (CentOS/RHEL)

**验证安装**：

```bash
python --version  # 应显示 Python 3.10+
pip --version     # 应显示 pip 版本
```

### Java 环境

**检查是否已安装**：

```bash
java -version
javac -version
```

**安装方法**：

1. **Windows**：从 [Oracle官网](https://www.oracle.com/java/technologies/downloads/) 或 [Adoptium](https://adoptium.net/) 下载 JDK
2. **macOS**：使用 Homebrew：`brew install openjdk`
3. **Linux**：使用包管理器：`sudo apt install default-jdk` (Ubuntu/Debian) 或 `sudo yum install java-11-openjdk-devel` (CentOS/RHEL)

**环境变量配置**：

- **JAVA_HOME**：指向 JDK 安装目录
- **PATH**：添加 `%JAVA_HOME%\bin` (Windows) 或 `$JAVA_HOME/bin` (Linux/macOS)

**验证安装**：

```bash
java -version   # 应显示 Java 版本
javac -version  # 应显示编译器版本
```

### C/C++ 环境

**检查是否已安装**：

```bash
gcc --version
g++ --version
```

**安装方法**：

1. **Windows**：
   - 安装 [MinGW-w64](https://www.mingw-w64.org/) 或 [MSYS2](https://www.msys2.org/)
   - 或安装 [Visual Studio](https://visualstudio.microsoft.com/) 并选择"使用C++的桌面开发"工作负载
2. **macOS**：安装 Xcode Command Line Tools：`xcode-select --install`
3. **Linux**：使用包管理器：`sudo apt install build-essential` (Ubuntu/Debian) 或 `sudo yum groupinstall "Development Tools"` (CentOS/RHEL)

**验证安装**：

```bash
gcc --version  # 应显示 GCC 版本
g++ --version  # 应显示 G++ 版本
```

### 环境检测脚本

我们提供了一个环境检测脚本，可以自动检查你的环境是否配置正确：

```bash
python check_environment.py
```

该脚本会检查：

- Python 版本和依赖包
- Java 环境 (JDK)
- C/C++ 编译器 (GCC/G++)
- 必要的系统工具
- 磁盘空间

**示例输出**：

```
============================================================
算法分析与设计 - 实验报告环境检测工具
============================================================

🐍 检查 Python 环境...
   ✅ Python 版本: 3.10.12
   ✅ 依赖包已安装

☕ 检查 Java 环境...
   ✅ Java 版本: openjdk version "17.0.8" 2023-07-18
   ✅ Java 编译器: javac 17.0.8

⚙️  检查 C/C++ 环境...
   ✅ G++ 编译器: g++ (Ubuntu 11.4.0-1ubuntu1~22.04) 11.4.0
   ✅ GCC 编译器: gcc (Ubuntu 11.4.0-1ubuntu1~22.04) 11.4.0

🔧 检查系统工具...
   ✅ git: 版本控制工具
   ✅ pip: Python 包管理器

💾 检查磁盘空间...
   ✅ 磁盘空间充足: 12345.6 MB 可用

============================================================
检查结果汇总:
============================================================
✅ 通过 Python 环境
✅ 通过 Java 环境
✅ 通过 C/C++ 环境
✅ 通过 系统工具
✅ 通过 磁盘空间

总计: 5/5 项检查通过

🎉 恭喜！你的环境配置完全正确，可以开始使用实验报告生成工具了！
```

### 常见问题解决

**问题1：Python 找不到**

```
'python' 不是内部或外部命令
```

**解决方案**：

1. 重新安装 Python，确保勾选"Add Python to PATH"
2. 手动添加 Python 到系统 PATH
3. 使用 `python3` 代替 `python`

**问题2：Java 编译错误**

```
'javac' 不是内部或外部命令
```

**解决方案**：

1. 确保安装了 JDK 而不仅仅是 JRE
2. 配置 JAVA_HOME 环境变量
3. 将 `%JAVA_HOME%\bin` 添加到 PATH

**问题3：GCC/G++ 找不到**

```
'g++' 不是内部或外部命令
```

**解决方案**：

1. **Windows**：安装 MinGW-w64 并添加到 PATH
2. **Linux**：安装 build-essential 包
3. **macOS**：安装 Xcode Command Line Tools

**问题4：权限问题**

```
Permission denied
```

**解决方案**：

1. **Windows**：以管理员身份运行命令提示符
2. **Linux/macOS**：使用 `sudo` 或检查文件权限

## 目录结构

```
算法设计与分析/
├── templates/              # 参考模板（仅供参考）
│   ├── 实验报告模板.docx    # 报告模板示例
│   └── 实验指导书模板.docx  # 指导书示例
├── src/                    # 实验源代码
│   ├── exp1/               # 实验一代码
│   │   └── *.py
│   ├── exp2/               # 实验二代码
│   └── ...
├── output/                 # 运行结果
│   ├── exp1/               # 实验一结果
│   │   ├── *.png           # 运行截图（必须）
│   │   └── output.txt      # 文本输出（可选）
│   └── ...
├── report/                 # 生成的报告（输出目录）
├── 实验报告.docx            # 报告模板（用户自行上传）
├── 实验指导书.docx          # 实验指导书（用户自行上传）
├── generate_report.py      # 自动生成脚本
├── check_environment.py    # 环境检测脚本
├── readme.md
└── AGENTS.md
```

### ⚠️ 重要说明

**用户需要自行上传两个文件到主目录**：

1. **实验报告模板**（命名：`实验报告.docx` 或其他 .docx 文件）
   - 这是实验报告模板
   - 脚本会自动识别，不一定要用固定文件名

2. **实验指导书**（命名：`实验指导书.docx` 或其他 .docx 文件）
   - 这是你们课程的实验指导书
   - 脚本会自动识别，不一定要用固定文件名

**templates 文件夹中的文件仅供参考**，展示模板的格式和结构，请勿直接使用。

### 文件识别说明

脚本支持智能识别文件，**不需要固定文件名**：

| 文件类型     | 可能的命名                                                                                  |
| ------------ | ------------------------------------------------------------------------------------------- |
| 实验报告模板 | `实验报告.docx`、`报告模板.docx`、`算法实验报告.docx`、或任何包含"教学上机实验报告"的 .docx |
| 实验指导书   | `实验指导书.docx`、`指导书.docx`、`算法实验指导书.docx`、或任何包含"实验指导书"的 .docx     |

脚本会读取文件内容来自动识别，所以你可以用任意文件名。

### 目录命名规则

实验目录支持多种命名方式，脚本会自动识别：

- `exp1` / `exp2` / `exp3`（推荐）
- `experiment1` / `实验1` / `实验一`

多个代码文件时（如 Java 项目），脚本会读取目录下所有代码文件并在报告中显示相对路径。

## 使用方法

**语言选择说明**：

- 在开始编写实验代码前，需要确认使用哪种编程语言
- 支持的语言：Python、Java、C++、C、JavaScript、Go、Rust 等
- 不同语言的代码文件应放在对应的 `src/exp{N}/` 目录中

```bash
# 生成所有实验报告（默认模式）
python generate_report.py

# 仅生成指定实验
python generate_report.py --exp 1

# 合并所有实验到一份报告
python generate_report.py --merge

# 指定上机日期
python generate_report.py --merge --date "2026年6月8日"

# 使用其他模板和指导书
python generate_report.py --template "path/to/模板.docx" --instruction "path/to/指导书.docx" --merge --date "2026年6月8日" --output-dir "output_folder"

# 查看目录结构指南
python generate_report.py --guide
```

### 命令行参数

| 参数            | 说明                   | 默认值         |
| --------------- | ---------------------- | -------------- |
| `--template`    | 模板文件路径           | 根目录自动查找 |
| `--instruction` | 指导书文件路径         | 根目录自动查找 |
| `--output-dir`  | 输出目录               | `report/`      |
| `--merge`       | 合并所有实验到一份报告 | 否（单独生成） |
| `--exp N`       | 仅生成第N个实验        | 全部           |
| `--date`        | 上机日期               | 当天日期       |

### 模板格式自动检测

脚本支持两种模板格式，**自动识别**：

| 格式       | 特征                             | 示例               |
| ---------- | -------------------------------- | ------------------ |
| **多表格** | 每个实验独立表格（6-10行 × N个） | Java程序设计模板   |
| **单表格** | 一个大表格包含所有实验（24+行）  | 算法分析与设计模板 |

行顺序不同也能正确填充，脚本通过关键词自动匹配每行的用途。

## 报告结构

每个实验在报告中对应一个表格页面，包含以下区域：

| 区域           | 内容来源     | 说明                              |
| -------------- | ------------ | --------------------------------- |
| 实验题目       | 指导书       | 仅显示实验名称                    |
| 实验目的和要求 | 指导书       | 自动提取，不改变原意              |
| 实验过程       | src/ 目录    | 直接贴完整代码，多文件带相对路径  |
| 实验结果       | output/ 目录 | 仅插入截图，自动缩放适配页面      |
| 实验分析       | 自动生成     | 正确性、时间/空间复杂度、实验心得 |
| 实验成绩       | 留空         | 教师填写                          |

## 截图要求

- 格式：PNG（优先）或 JPG
- 背景：黑色终端背景
- 图片会自动缩放到页面宽度（15cm），高度超限（20cm）时自动等比缩小
