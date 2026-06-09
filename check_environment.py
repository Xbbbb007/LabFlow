#!/usr/bin/env python3
"""
环境检测脚本

检查用户环境是否满足实验报告生成的要求。
支持检查 Python、Java、C/C++ 环境。
"""

import sys
import subprocess
import shutil
import platform
from pathlib import Path


def check_python():
    """检查 Python 环境"""
    print("🐍 检查 Python 环境...")
    
    # 检查 Python 版本
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 10):
        print(f"   ❌ Python 版本过低: {version.major}.{version.minor}.{version.micro}")
        print("      需要 Python 3.10 或更高版本")
        return False
    
    print(f"   ✅ Python 版本: {version.major}.{version.minor}.{version.micro}")
    
    # 检查依赖包
    required_packages = ['docx', 'PIL']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"   ❌ 缺少依赖包: {', '.join(missing_packages)}")
        print("      请运行: pip install python-docx Pillow")
        return False
    
    print("   ✅ 依赖包已安装")
    return True


def check_java():
    """检查 Java 环境"""
    print("☕ 检查 Java 环境...")
    
    # 检查 java 命令
    if not shutil.which('java'):
        print("   ❌ 未找到 Java 运行时环境 (JRE)")
        print("      请安装 JDK: https://adoptium.net/")
        return False
    
    # 检查 javac 命令
    if not shutil.which('javac'):
        print("   ❌ 未找到 Java 编译器 (JDK)")
        print("      请安装 JDK: https://adoptium.net/")
        return False
    
    try:
        # 获取 Java 版本
        result = subprocess.run(['java', '-version'], capture_output=True, text=True)
        version_line = result.stderr.split('\n')[0] if result.stderr else result.stdout.split('\n')[0]
        print(f"   ✅ Java 版本: {version_line}")
        
        # 获取 javac 版本
        result = subprocess.run(['javac', '-version'], capture_output=True, text=True)
        version_line = result.stdout.split('\n')[0] if result.stdout else result.stderr.split('\n')[0]
        print(f"   ✅ Java 编译器: {version_line}")
        
        return True
    except Exception as e:
        print(f"   ❌ Java 环境检查失败: {e}")
        return False


def check_cpp():
    """检查 C/C++ 环境"""
    print("⚙️  检查 C/C++ 环境...")
    
    system = platform.system().lower()
    
    # 检查 g++ (C++)
    if shutil.which('g++'):
        try:
            result = subprocess.run(['g++', '--version'], capture_output=True, text=True)
            version_line = result.stdout.split('\n')[0] if result.stdout else "未知版本"
            print(f"   ✅ G++ 编译器: {version_line}")
        except Exception:
            print("   ⚠️  G++ 命令存在但无法获取版本信息")
    else:
        print("   ❌ 未找到 G++ 编译器")
        if system == "windows":
            print("      请安装 MinGW-w64: https://www.mingw-w64.org/")
            print("      或安装 Visual Studio: https://visualstudio.microsoft.com/")
        elif system == "linux":
            print("      请运行: sudo apt install build-essential (Ubuntu/Debian)")
            print("      或运行: sudo yum groupinstall 'Development Tools' (CentOS/RHEL)")
        elif system == "darwin":
            print("      请运行: xcode-select --install")
        return False
    
    # 检查 gcc (C)
    if shutil.which('gcc'):
        try:
            result = subprocess.run(['gcc', '--version'], capture_output=True, text=True)
            version_line = result.stdout.split('\n')[0] if result.stdout else "未知版本"
            print(f"   ✅ GCC 编译器: {version_line}")
        except Exception:
            print("   ⚠️  GCC 命令存在但无法获取版本信息")
    else:
        print("   ⚠️  未找到 GCC 编译器 (可选，G++ 已支持 C 语言)")
    
    return True


def check_system_tools():
    """检查系统工具"""
    print("🔧 检查系统工具...")
    
    tools = {
        'git': '版本控制工具',
        'pip': 'Python 包管理器',
    }
    
    all_ok = True
    for tool, description in tools.items():
        if shutil.which(tool):
            print(f"   ✅ {tool}: {description}")
        else:
            print(f"   ⚠️  {tool}: 未找到 ({description})")
            if tool == 'pip':
                all_ok = False
    
    return all_ok


def check_disk_space():
    """检查磁盘空间"""
    print("💾 检查磁盘空间...")
    
    try:
        # 获取当前目录所在磁盘的可用空间
        current_dir = Path.cwd()
        if platform.system().lower() == "windows":
            import ctypes
            free_bytes = ctypes.c_ulonglong(0)
            ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                ctypes.c_wchar_p(str(current_dir)),
                None, None, ctypes.pointer(free_bytes)
            )
            free_mb = free_bytes.value / (1024 * 1024)
        else:
            import os
            stat = os.statvfs(str(current_dir))
            free_mb = (stat.f_bavail * stat.f_frsize) / (1024 * 1024)
        
        if free_mb < 100:  # 少于 100MB
            print(f"   ⚠️  磁盘空间不足: {free_mb:.1f} MB 可用")
            print("      建议至少 100MB 可用空间")
            return False
        else:
            print(f"   ✅ 磁盘空间充足: {free_mb:.1f} MB 可用")
            return True
    except Exception as e:
        print(f"   ⚠️  无法检查磁盘空间: {e}")
        return True


def main():
    """主函数"""
    print("=" * 60)
    print("算法分析与设计 - 实验报告环境检测工具")
    print("=" * 60)
    print()
    
    checks = [
        ("Python 环境", check_python),
        ("Java 环境", check_java),
        ("C/C++ 环境", check_cpp),
        ("系统工具", check_system_tools),
        ("磁盘空间", check_disk_space),
    ]
    
    results = []
    for name, check_func in checks:
        try:
            result = check_func()
            results.append((name, result))
        except Exception as e:
            print(f"   ❌ 检查 {name} 时出错: {e}")
            results.append((name, False))
        print()
    
    # 统计结果
    print("=" * 60)
    print("检查结果汇总:")
    print("=" * 60)
    
    passed = 0
    total = len(results)
    
    for name, result in results:
        status = "✅ 通过" if result else "❌ 失败"
        print(f"{status} {name}")
        if result:
            passed += 1
    
    print()
    print(f"总计: {passed}/{total} 项检查通过")
    
    if passed == total:
        print("\n🎉 恭喜！你的环境配置完全正确，可以开始使用实验报告生成工具了！")
        print("\n下一步:")
        print("1. 将实验报告模板放在根目录")
        print("2. 将实验指导书放在根目录")
        print("3. 在 src/exp1/ 中放入实验代码")
        print("4. 运行: python generate_report.py --merge")
    else:
        print("\n⚠️  部分环境检查未通过，请根据上述提示进行修复。")
        print("\n需要帮助？请查看:")
        print("- readme.md 中的环境配置指南")
        print("- 或访问项目 GitHub 页面提交 Issue")
    
    print()
    return passed == total


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)