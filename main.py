"""
PDF·Word·JPG 互转工具 — 入口
武汉纺织大学管理学院媒体运营部

用法: python main.py
"""

import sys
import os

# 确保脚本所在目录在 path 中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if __name__ == "__main__":
    from converter_gui import ConverterApp
    app = ConverterApp()
    app.mainloop()
