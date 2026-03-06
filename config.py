# -*- coding: utf-8 -*-
"""
Excel 转 PDF 批量转换工具 - 配置模块
"""

import os
import sys

# ============================================================
# 版本信息
# ============================================================
APP_NAME = "Excel 转 PDF 批量转换工具"
APP_VERSION = "1.0.0"

# ============================================================
# 文件格式支持
# ============================================================
SUPPORTED_EXTENSIONS = {'.xls', '.xlsx', '.xlsm'}

# ============================================================
# 并发配置
# ============================================================
DEFAULT_WORKERS = 4      # 默认工作进程数
MIN_WORKERS = 1          # 最少进程数
MAX_WORKERS = 8          # 最多进程数

# ============================================================
# 转换配置
# ============================================================
RETRY_COUNT = 1          # 失败重试次数
EXCEL_VISIBLE = False    # Excel 是否可见（调试用）
EXCEL_ALERTS = False     # 是否显示 Excel 弹窗提示

# PDF 导出格式常量 (Excel COM xlTypePDF = 0)
XL_TYPE_PDF = 0

# ============================================================
# 日志配置
# ============================================================
LOG_FILENAME = "conversion_log_{timestamp}.txt"
LOG_DATE_FORMAT = "%Y%m%d_%H%M%S"
LOG_ENCODING = "utf-8"

# ============================================================
# GUI 配置
# ============================================================
WINDOW_WIDTH = 750
WINDOW_HEIGHT = 580
WINDOW_MIN_WIDTH = 650
WINDOW_MIN_HEIGHT = 500

# 颜色主题
THEME = {
    "bg": "#f0f2f5",
    "card_bg": "#ffffff",
    "primary": "#1890ff",
    "primary_hover": "#40a9ff",
    "success": "#52c41a",
    "warning": "#faad14",
    "error": "#ff4d4f",
    "text": "#333333",
    "text_secondary": "#666666",
    "border": "#d9d9d9",
    "log_bg": "#1e1e1e",
    "log_text": "#d4d4d4",
}

# ============================================================
# 工具函数
# ============================================================

def get_resource_path(relative_path):
    """获取资源文件的绝对路径（兼容 PyInstaller 打包后的环境）"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包后的临时目录
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def get_default_output_dir(input_dir):
    """获取默认输出目录：输入目录下的 PDF_Output 子文件夹"""
    return os.path.join(input_dir, "PDF_Output")
