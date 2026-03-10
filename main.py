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
    from tkinter import messagebox
    from gui import ExcelToPdfApp

    root = tk.Tk()
    app = ExcelToPdfApp(root)

    # 启动时检查 pypdf 是否可用（用于空白页删除）
    try:
        import pypdf
        pypdf_ok = True
    except ImportError:
        pypdf_ok = False

    if not pypdf_ok:
        root.after(500, lambda: messagebox.showwarning(
            "功能受限",
            "pypdf 模块未找到，空白页自动删除功能不可用。\n\n"
            "请使用最新版本的「一键打包EXE.bat」重新打包。"
        ))

    root.mainloop()


if __name__ == "__main__":
    main()
