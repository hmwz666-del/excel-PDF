# -*- coding: utf-8 -*-
"""
Excel 转 PDF 批量转换工具 - 核心转换引擎

使用 win32com COM 自动化调用 Microsoft Excel 进行高保真转换。
"""

import os
import logging
import traceback

try:
    import win32com.client
    import pythoncom
    from pywintypes import com_error
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

from config import XL_TYPE_PDF, EXCEL_VISIBLE, EXCEL_ALERTS, RETRY_COUNT

logger = logging.getLogger(__name__)


class ConversionResult:
    """单个文件的转换结果"""
    SUCCESS = "success"
    FAILED = "failed"
    SKIPPED = "skipped"

    def __init__(self, filepath, status, message="", output_path=""):
        self.filepath = filepath
        self.filename = os.path.basename(filepath)
        self.status = status
        self.message = message
        self.output_path = output_path

    def __repr__(self):
        return f"ConversionResult({self.filename}, {self.status}, {self.message})"


class ExcelConverter:
    """
    Excel 转 PDF 转换器

    管理一个 Excel COM 实例，负责文件的打开、打印区域设置和 PDF 导出。
    设计为在单独的进程中运行（COM 对象不能跨线程共享）。
    """

    def __init__(self):
        self.excel_app = None
        self._initialized = False

    def initialize(self):
        """初始化 COM 环境和 Excel 实例"""
        if not HAS_WIN32COM:
            raise RuntimeError(
                "未找到 win32com 模块。请确保在 Windows 系统上运行，"
                "并已安装 pywin32: pip install pywin32"
            )

        # 初始化当前线程的 COM 环境
        pythoncom.CoInitialize()

        try:
            self.excel_app = win32com.client.Dispatch("Excel.Application")
            self.excel_app.Visible = EXCEL_VISIBLE
            self.excel_app.DisplayAlerts = EXCEL_ALERTS
            # 禁用屏幕刷新以提升性能
            self.excel_app.ScreenUpdating = False
            self._initialized = True
            logger.info("Excel COM 实例初始化成功")
        except Exception as e:
            logger.error(f"无法启动 Excel: {e}")
            self.cleanup()
            raise RuntimeError(f"无法启动 Microsoft Excel: {e}")

    def cleanup(self):
        """清理 Excel COM 实例和 COM 环境"""
        if self.excel_app is not None:
            try:
                self.excel_app.Quit()
            except Exception:
                pass
            self.excel_app = None

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

        self._initialized = False
        logger.info("Excel COM 实例已清理")

    def convert_file(self, excel_path, output_dir):
        """
        将单个 Excel 文件转换为 PDF

        Args:
            excel_path: Excel 文件的绝对路径
            output_dir: PDF 输出目录的绝对路径

        Returns:
            ConversionResult 对象
        """
        if not self._initialized:
            return ConversionResult(
                excel_path, ConversionResult.FAILED,
                "转换器未初始化"
            )

        # 构建输出路径
        base_name = os.path.splitext(os.path.basename(excel_path))[0]
        pdf_path = os.path.join(output_dir, f"{base_name}.pdf")

        # 处理同名文件
        if os.path.exists(pdf_path):
            counter = 1
            while os.path.exists(pdf_path):
                pdf_path = os.path.join(output_dir, f"{base_name}_{counter}.pdf")
                counter += 1

        workbook = None
        last_error = None

        for attempt in range(RETRY_COUNT + 1):
            try:
                # 以只读方式打开工作簿
                workbook = self.excel_app.Workbooks.Open(
                    os.path.abspath(excel_path),
                    ReadOnly=True,
                    UpdateLinks=0,       # 不更新外部链接
                    IgnoreReadOnlyRecommended=True,
                    CorruptLoad=1,       # 尝试修复打开
                )

                # 处理每个有内容的工作表的打印区域
                self._optimize_print_area(workbook)

                # 导出为 PDF
                workbook.ExportAsFixedFormat(
                    Type=XL_TYPE_PDF,
                    Filename=os.path.abspath(pdf_path),
                    Quality=0,           # 标准质量 (xlQualityStandard)
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False,
                )

                logger.info(f"✅ 转换成功: {os.path.basename(excel_path)}")
                return ConversionResult(
                    excel_path, ConversionResult.SUCCESS,
                    "转换成功", pdf_path
                )

            except com_error as e:
                last_error = str(e)
                error_code = getattr(e, 'hresult', None)

                # 密码保护的文件
                if error_code and (error_code == -2147352567 or "password" in str(e).lower()):
                    logger.warning(f"⏭️ 跳过密码保护文件: {os.path.basename(excel_path)}")
                    return ConversionResult(
                        excel_path, ConversionResult.SKIPPED,
                        "文件有密码保护"
                    )

                if attempt < RETRY_COUNT:
                    logger.info(f"🔄 重试 ({attempt + 1}/{RETRY_COUNT}): {os.path.basename(excel_path)}")
                    continue

            except Exception as e:
                last_error = str(e)
                logger.error(f"❌ 转换失败: {os.path.basename(excel_path)} - {e}")
                if attempt < RETRY_COUNT:
                    continue

            finally:
                # 确保工作簿被关闭
                if workbook is not None:
                    try:
                        workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
                    workbook = None

        logger.error(f"❌ 最终失败: {os.path.basename(excel_path)} - {last_error}")
        return ConversionResult(
            excel_path, ConversionResult.FAILED,
            f"转换失败: {last_error}"
        )

    def _optimize_print_area(self, workbook):
        """
        优化每个工作表的打印区域，避免空白页和多余留白

        使用 Cells.Find 精准查找实际有数据内容的最后行/列，
        排除仅有边框/格式但无数据的空行空列。
        """
        for sheet in workbook.Worksheets:
            try:
                # ===== 第一步：查找实际有数据的范围 =====
                # 使用 Find 方法查找最后一个有数据的单元格
                # 这比 UsedRange 更精准，不会把仅有边框格式的空行算进去

                # 查找最后有数据的行 (按行搜索, 从末尾往前找)
                last_row_cell = sheet.Cells.Find(
                    What="*",
                    SearchOrder=1,       # xlByRows
                    SearchDirection=2,   # xlPrevious
                    LookIn=-4163,        # xlValues (按值查找,忽略格式)
                )

                if last_row_cell is None:
                    # 整个工作表没有任何数据内容
                    if workbook.Worksheets.Count > 1:
                        sheet.Visible = False
                    continue

                last_row = last_row_cell.Row

                # 查找最后有数据的列 (按列搜索, 从末尾往前找)
                last_col_cell = sheet.Cells.Find(
                    What="*",
                    SearchOrder=2,       # xlByColumns
                    SearchDirection=2,   # xlPrevious
                    LookIn=-4163,        # xlValues
                )
                last_col = last_col_cell.Column if last_col_cell else 1

                # 查找第一个有数据的行 (从开头往后找)
                first_row_cell = sheet.Cells.Find(
                    What="*",
                    SearchOrder=1,       # xlByRows
                    SearchDirection=1,   # xlNext
                    LookIn=-4163,        # xlValues
                )
                first_row = first_row_cell.Row if first_row_cell else 1

                # 查找第一个有数据的列
                first_col_cell = sheet.Cells.Find(
                    What="*",
                    SearchOrder=2,       # xlByColumns
                    SearchDirection=1,   # xlNext
                    LookIn=-4163,        # xlValues
                )
                first_col = first_col_cell.Column if first_col_cell else 1

                # 从第1行开始(保留表头),列从第1列开始
                start_row = min(first_row, 1)
                start_col = min(first_col, 1)

                # ===== 第二步：设置精准打印区域 =====
                print_range = sheet.Range(
                    sheet.Cells(start_row, start_col),
                    sheet.Cells(last_row, last_col)
                )
                sheet.PageSetup.PrintArea = print_range.Address

                logger.debug(
                    f"工作表 '{sheet.Name}': 打印区域 "
                    f"Row {start_row}-{last_row}, Col {start_col}-{last_col}"
                )

                # ===== 第三步：优化页面布局 =====
                page_setup = sheet.PageSetup

                # 自动判断横向/纵向
                if last_col > 8:
                    page_setup.Orientation = 2  # xlLandscape (横向)
                else:
                    page_setup.Orientation = 1  # xlPortrait (纵向)

                # 自动缩放适配页面宽度 (关键: 内容填满页宽)
                page_setup.Zoom = False
                page_setup.FitToPagesWide = 1    # 宽度缩放到1页
                page_setup.FitToPagesTall = False  # 高度不限制,自动分页

                # 最小化页边距 (单位: 磅, 1英寸=72磅)
                page_setup.LeftMargin = 7.2     # 约 0.1 英寸
                page_setup.RightMargin = 7.2    # 约 0.1 英寸
                page_setup.TopMargin = 14.4     # 约 0.2 英寸
                page_setup.BottomMargin = 14.4  # 约 0.2 英寸
                page_setup.HeaderMargin = 0
                page_setup.FooterMargin = 0

                # 水平居中打印
                page_setup.CenterHorizontally = True

            except Exception as e:
                # 某个工作表设置失败不影响整体转换
                logger.debug(f"设置工作表 '{sheet.Name}' 打印区域时出错: {e}")
                continue

    def __enter__(self):
        self.initialize()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.cleanup()
        return False
