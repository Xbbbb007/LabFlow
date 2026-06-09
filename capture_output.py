#!/usr/bin/env python3
"""
终端截图生成工具

运行命令或将已有文本渲染为黑底白字的终端风格截图。

用法:
    # 运行命令并截图
    python capture_output.py --cmd "python src/exp1/main.py" --output output/exp1/run_result.png

    # 从文本文件生成截图
    python capture_output.py --file output/exp1/output.txt --output output/exp1/run_result.png

    # 直接传入文本生成截图
    python capture_output.py --text "Hello World" --output output/exp1/run_result.png

依赖:
    pip install Pillow
"""

import argparse
import subprocess
import sys
import textwrap
from pathlib import Path

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("[错误] 缺少 Pillow，请运行: pip install Pillow")
    sys.exit(1)


# ============================================================
# 字体选择
# ============================================================

def find_font(size: int = 16) -> ImageFont.FreeTypeFont:
    """
    查找支持中文的等宽字体。
    按优先级依次尝试多个字体路径。
    """
    # 候选字体列表（按优先级排序）
    candidates = []

    if sys.platform == "win32":
        candidates = [
            r"C:\Windows\Fonts\msyh.ttc",       # 微软雅黑（推荐）
            r"C:\Windows\Fonts\msyhbd.ttc",      # 微软雅黑 Bold
            r"C:\Windows\Fonts\simhei.ttf",       # 黑体
            r"C:\Windows\Fonts\consola.ttf",      # Consolas（不支持中文，兜底）
        ]
    elif sys.platform == "darwin":
        candidates = [
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Light.ttc",
            "/Library/Fonts/Arial Unicode.ttf",
        ]
    else:  # Linux
        candidates = [
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
            "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
        ]

    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except (OSError, IOError):
            continue

    # 最后兜底：使用默认字体
    print("  ⚠️  未找到合适的中文字体，使用默认字体（中文可能显示为方块）")
    return ImageFont.load_default()


# ============================================================
# 文本获取
# ============================================================

def run_command(cmd: str, cwd: str = None, timeout: int = 30) -> str:
    """运行命令并捕获输出"""
    print(f"  执行命令: {cmd}")
    try:
        result = subprocess.run(
            cmd,
            shell=True,
            capture_output=True,
            text=True,
            cwd=cwd,
            timeout=timeout,
            encoding="utf-8",
            errors="replace",
        )
        output = ""
        if result.stdout:
            output += result.stdout
        if result.stderr:
            output += result.stderr
        if not output.strip():
            output = f"(命令执行完毕，退出码: {result.returncode})"
        return output
    except subprocess.TimeoutExpired:
        return f"(命令执行超时: {timeout}秒)"
    except Exception as e:
        return f"(命令执行失败: {e})"


def read_file(filepath: str) -> str:
    """读取文本文件"""
    path = Path(filepath)
    if not path.exists():
        return f"(文件不存在: {filepath})"
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="gbk", errors="replace")


# ============================================================
# 截图渲染
# ============================================================

def render_terminal_screenshot(
    text: str,
    output_path: str,
    font_size: int = 16,
    padding_x: int = 28,
    padding_y: int = 20,
    line_height: int = None,
    max_width: int = 900,
    bg_color: tuple = (12, 12, 12),
    text_color: tuple = (204, 204, 204),
):
    """
    将文本渲染为黑底白字的终端风格截图。

    参数:
        text: 要渲染的文本内容
        output_path: 输出图片路径
        font_size: 字体大小
        padding_x: 左右内边距
        padding_y: 上下内边距
        line_height: 行高（None 则自动计算为 font_size * 1.6）
        max_width: 最大宽度
        bg_color: 背景色
        text_color: 文字色
    """
    font = find_font(font_size)

    if line_height is None:
        line_height = int(font_size * 1.6)

    # 使用 splitlines() 处理换行（兼容 Windows \r\n）
    lines = text.splitlines()

    # 自动换行：如果某行过长，按 max_width 拆分
    wrapped_lines = []
    # 估算每行最大字符数（中文约 font_size 宽，英文约 font_size/2 宽）
    chars_per_line = max(40, int((max_width - 2 * padding_x) / (font_size * 0.6)))

    for line in lines:
        if len(line) > chars_per_line:
            wrapped_lines.extend(textwrap.wrap(line, width=chars_per_line) or [""])
        else:
            wrapped_lines.append(line)

    # 计算图片尺寸
    img_width = max_width
    img_height = padding_y * 2 + len(wrapped_lines) * line_height + 10

    # 创建图片
    img = Image.new("RGB", (img_width, img_height), bg_color)
    draw = ImageDraw.Draw(img)

    # 逐行绘制文本
    y = padding_y
    for line in wrapped_lines:
        draw.text((padding_x, y), line, fill=text_color, font=font)
        y += line_height

    # 保存
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    img.save(str(output))
    print(f"  截图已保存: {output_path} ({img_width}x{img_height})")


# ============================================================
# 主入口
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="终端截图生成工具 - 运行命令或渲染文本为终端风格截图"
    )

    # 文本来源（三选一）
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument(
        "--cmd", type=str,
        help="要执行的命令（如 'python src/exp1/main.py'）"
    )
    source.add_argument(
        "--file", type=str,
        help="读取文本的文件路径"
    )
    source.add_argument(
        "--text", type=str,
        help="直接传入的文本内容"
    )

    # 输出路径
    parser.add_argument(
        "--output", "-o", type=str, required=True,
        help="截图输出路径（如 output/exp1/run_result.png）"
    )

    # 可选参数
    parser.add_argument(
        "--cwd", type=str, default=None,
        help="命令执行的工作目录"
    )
    parser.add_argument(
        "--timeout", type=int, default=30,
        help="命令执行超时时间（秒，默认30）"
    )
    parser.add_argument(
        "--font-size", type=int, default=16,
        help="字体大小（默认16）"
    )
    parser.add_argument(
        "--max-width", type=int, default=900,
        help="截图最大宽度（默认900）"
    )

    args = parser.parse_args()

    # 获取文本
    if args.cmd:
        text = run_command(args.cmd, cwd=args.cwd, timeout=args.timeout)
    elif args.file:
        text = read_file(args.file)
    else:
        text = args.text

    if not text.strip():
        print("  ⚠️  文本为空，跳过截图生成")
        return

    # 渲染截图
    render_terminal_screenshot(
        text=text,
        output_path=args.output,
        font_size=args.font_size,
        max_width=args.max_width,
    )


if __name__ == "__main__":
    main()
