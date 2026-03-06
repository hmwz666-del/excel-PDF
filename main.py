# -*- coding: utf-8 -*-
"""
Excel 转 PDF 批量转换工具

主程序入口。启动 GUI 界面，由用户选择文件夹并开始转换。

使用方法:
    python main.py

打包为 EXE:
    pip install pyinstaller pywin32
    pyinstaller -F -w --name "Excel转PDF工具" main.py
"""

import sys
import os
import multiprocessing


def main():
    """程序入口"""
    # Windows 多进程打包支持（PyInstaller 必须）
    multiprocessing.freeze_support()

    import tkinter as tk
    from gui import ExcelToPdfApp

    root = tk.Tk()
    app = ExcelToPdfApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
