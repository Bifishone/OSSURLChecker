import sys
import re
from colorama import init, Fore, Back, Style

# 初始化colorama，支持Windows系统
init(autoreset=True)


def remove_ansi_codes(s):
    """移除字符串中的ANSI颜色代码，用于准确计算显示长度"""
    ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
    return ansi_escape.sub('', s)


def print_ascii_art():
    # 定义ASCII艺术内容，使用原始字符串标记r解决转义问题
    ascii_art = [
        r"___  ____ ____  _   _ ____  _     ____ _               _",
        r" / _ \/ ___/ ___|| | | |  _ \| |   / ___| |__   ___  ___| | _____ _ __",
        r"| | | \___ \___ \| | | | |_) | |  | |   | '_ \ / _ \/ __| |/ / _ \ '__|",
        r"| |_| |___) |__) | |_| |  _ <| |__| |___| | | |  __/ (__|   <  __/ |",
        r" \___/|____/____/ \___/|_| \_\_____\____|_| |_|\___|\___|_|\_\___|_|"
    ]

    # 打印彩色ASCII艺术
    print("\n" + " " * 10, end="")  # 增加左侧缩进
    print(Fore.CYAN + Style.BRIGHT + ascii_art[0])

    for line in ascii_art[1:]:
        print(" " * 10, end="")  # 保持一致的缩进
        # 为不同行设置不同的颜色，创造渐变效果
        if line == ascii_art[1]:
            print(Fore.GREEN + Style.BRIGHT + line)
        elif line == ascii_art[2]:
            print(Fore.YELLOW + Style.BRIGHT + line)
        elif line == ascii_art[3]:
            print(Fore.MAGENTA + Style.BRIGHT + line)
        else:
            print(Fore.RED + Style.BRIGHT + line)

    # 计算ASCII艺术中最长行的长度，用于分隔线对齐
    max_ascii_length = max(len(line) for line in ascii_art)
    # 打印分隔线
    print("\n" + "=" * (max_ascii_length + 20))

    # 打印作者信息
    author_info = f"Author: {Fore.CYAN}Bifish"
    github_info = f"GitHub: {Fore.GREEN}https://github.com/Bifishone"

    # 计算实际显示长度（排除ANSI颜色代码）
    author_visible_len = len(remove_ansi_codes(author_info))
    github_visible_len = len(remove_ansi_codes(github_info))
    # 计算居中对齐的缩进（基于最长可见长度）
    max_visible_length = max(author_visible_len, github_visible_len, max_ascii_length)
    author_indent = " " * ((max_visible_length - author_visible_len) // 2 + 10)
    github_indent = " " * ((max_visible_length - github_visible_len) // 2 + 10)

    print(author_indent + author_info)
    print(github_indent + github_info)

    # 底部装饰线
    print("\n" + "=" * (max_ascii_length + 20) + "\n")


if __name__ == "__main__":
    try:
        print_ascii_art()
    except Exception as e:
        print(f"发生错误: {e}")
        sys.exit(1)
