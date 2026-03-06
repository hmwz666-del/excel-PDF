# -*- coding: utf-8 -*-
"""
Excel 转 PDF 批量转换工具 - 核心转换引擎

使用 win32com COM 自动化调用 Microsoft Excel 进行高保真转换。
"""

import os
import re
import logging
import traceback

try:
    import win32com.client
    import pythoncom
    from pywintypes import com_error
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

try:
    from pypdf import PdfReader, PdfWriter
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

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

    # Excel COM 常量
    XL_UP = -4162       # xlUp
    XL_TO_LEFT = -4159  # xlToLeft

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
            # DispatchEx 强制创建独立的 Excel 进程（不会复用已有的 Excel）
            # 这样用户可以在转换期间正常使用自己的 Excel
            self.excel_app = win32com.client.DispatchEx("Excel.Application")
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

    def convert_file(self, excel_path, output_dir, input_dir=None):
        """
        将单个 Excel 文件转换为 PDF

        Args:
            excel_path: Excel 文件的绝对路径
            output_dir: PDF 输出目录的绝对路径
            input_dir: 输入根目录（用于保留子目录结构）

        Returns:
            ConversionResult 对象
        """
        if not self._initialized:
            return ConversionResult(
                excel_path, ConversionResult.FAILED,
                "转换器未初始化"
            )

        # 构建输出路径（保留子目录结构，避免不同文件夹的同名文件冲突）
        base_name = os.path.splitext(os.path.basename(excel_path))[0]

        # 如果有 input_dir，保留相对子目录结构
        if input_dir:
            rel_dir = os.path.relpath(os.path.dirname(excel_path), input_dir)
            target_dir = os.path.join(output_dir, rel_dir) if rel_dir != '.' else output_dir
            os.makedirs(target_dir, exist_ok=True)
        else:
            target_dir = output_dir

        pdf_path = os.path.join(target_dir, f"{base_name}.pdf")

        # 处理同名文件：追加序号（兜底机制）
        renamed = False
        original_pdf_name = f"{base_name}.pdf"
        if os.path.exists(pdf_path):
            counter = 1
            while os.path.exists(pdf_path):
                pdf_path = os.path.join(target_dir, f"{base_name}_{counter}.pdf")
                counter += 1
            renamed = True
            logger.warning(
                f"⚠️ 同名文件: {original_pdf_name} 已存在，"
                f"已重命名为 {os.path.basename(pdf_path)}"
            )

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

                # 优化打印区域和页面布局
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

                # 后处理：移除空白页
                removed = self._remove_blank_pages(pdf_path)

                if renamed:
                    msg = f"转换成功 (同名文件 {original_pdf_name} 已重命名为 {os.path.basename(pdf_path)})"
                    logger.info(f"✅ 转换成功: {os.path.basename(excel_path)} → {os.path.basename(pdf_path)}")
                else:
                    msg = "转换成功"
                    logger.info(f"✅ 转换成功: {os.path.basename(excel_path)}")

                if removed > 0:
                    msg += f" (已删除 {removed} 个空白页)"
                    logger.info(f"   🗑️ 已删除 {removed} 个空白页")

                return ConversionResult(
                    excel_path, ConversionResult.SUCCESS,
                    msg, pdf_path
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

    # ==================== 打印区域优化 ====================

    def _get_data_last_row(self, sheet):
        """
        查找工作表中最后有数据的行号

        使用 Cells(Rows.Count, col).End(xlUp) —— 这是 Excel VBA 标准做法：
        从每列的最底行向上搜索，找到第一个有值的单元格。
        此方法只看单元格值，完全忽略格式/边框，且不受 Find 的 After 参数 Bug 影响。

        Returns:
            最后有数据的行号；无数据返回 0
        """
        used_range = sheet.UsedRange
        if used_range is None:
            return 0

        max_row = 0
        start_col = used_range.Column
        num_cols = min(used_range.Columns.Count, 100)

        for i in range(num_cols):
            col = start_col + i
            try:
                last_cell = sheet.Cells(sheet.Rows.Count, col).End(self.XL_UP)
                if last_cell.Value is not None and last_cell.Row > max_row:
                    max_row = last_cell.Row
            except Exception:
                continue

        return max_row

    def _get_data_last_col(self, sheet):
        """
        查找工作表中最后有数据的列号

        使用 Cells(row, Columns.Count).End(xlToLeft) 方法。

        Returns:
            最后有数据的列号；无数据返回 0
        """
        used_range = sheet.UsedRange
        if used_range is None:
            return 0

        max_col = 0
        start_row = used_range.Row
        # 不需要检查所有行，抽样检查前 200 行足够
        num_rows = min(used_range.Rows.Count, 200)

        for i in range(num_rows):
            row = start_row + i
            try:
                last_cell = sheet.Cells(row, sheet.Columns.Count).End(self.XL_TO_LEFT)
                if last_cell.Value is not None and last_cell.Column > max_col:
                    max_col = last_cell.Column
            except Exception:
                continue

        return max_col

    def _optimize_print_area(self, workbook):
        """
        优化每个工作表的打印区域和页面布局

        核心逻辑：
        1. 用 End(xlUp)/End(xlToLeft) 精确定位最后有数据的行/列
        2. 设置 PrintArea 只包含有数据的区域（排除空行空列）
        3. 优化页面边距、方向、缩放
        """
        for sheet in workbook.Worksheets:
            try:
                used_range = sheet.UsedRange
                if used_range is None or used_range.Count == 0:
                    if workbook.Worksheets.Count > 1:
                        sheet.Visible = False
                    continue

                # ===== 精确定位数据范围 =====
                last_row = self._get_data_last_row(sheet)
                last_col = self._get_data_last_col(sheet)

                if last_row == 0 or last_col == 0:
                    # 工作表没有任何数据值，隐藏
                    if workbook.Worksheets.Count > 1:
                        sheet.Visible = False
                    continue

                # ===== 设置精确的打印区域 =====
                # 从第1行第1列开始，到最后有数据的行/列结束
                print_range = sheet.Range(
                    sheet.Cells(1, 1),
                    sheet.Cells(last_row, last_col)
                )
                sheet.PageSetup.PrintArea = print_range.Address

                logger.debug(
                    f"工作表 '{sheet.Name}': "
                    f"UsedRange={used_range.Rows.Count}行x{used_range.Columns.Count}列, "
                    f"实际数据={last_row}行x{last_col}列"
                )

                # ===== 优化页面布局 =====
                page_setup = sheet.PageSetup

                if last_col > 8:
                    page_setup.Orientation = 2  # xlLandscape
                else:
                    page_setup.Orientation = 1  # xlPortrait

                page_setup.Zoom = False
                page_setup.FitToPagesWide = 1
                page_setup.FitToPagesTall = False

                page_setup.LeftMargin = 7.2     # ~0.1 英寸
                page_setup.RightMargin = 7.2
                page_setup.TopMargin = 14.4     # ~0.2 英寸
                page_setup.BottomMargin = 14.4
                page_setup.HeaderMargin = 0
                page_setup.FooterMargin = 0

                page_setup.CenterHorizontally = True

            except Exception as e:
                logger.debug(f"设置工作表 '{sheet.Name}' 打印区域时出错: {e}")
                continue

    # ==================== PDF 空白页移除 ====================

    def _remove_blank_pages(self, pdf_path):
        """
        移除 PDF 中的空白页（双重检测）

        检测逻辑：
        - 提取文本，检查是否有可见字符（字母/数字/中文）
        - 检查页面是否包含图片
        - 没有文字 AND 没有图片 → 删除

        Returns:
            被删除的空白页数量
        """
        if not HAS_PYPDF:
            return 0

        try:
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)

            if total_pages <= 1:
                return 0

            writer = PdfWriter()
            removed = 0

            for i, page in enumerate(reader.pages):
                # 检测1: 是否有可见文字
                text = page.extract_text() or ""
                has_text = bool(re.search(r'[\w\u4e00-\u9fff]', text))

                # 检测2: 是否有图片
                has_images = False
                try:
                    if '/Resources' in page:
                        resources = page['/Resources']
                        if '/XObject' in resources:
                            xobjects = resources['/XObject']
                            if xobjects and len(xobjects) > 0:
                                has_images = True
                except Exception:
                    pass

                # 有文字 OR 有图片 → 保留
                if has_text or has_images:
                    writer.add_page(page)
                else:
                    removed += 1
                    logger.debug(f"  删除空白页: 第 {i+1} 页")

            if removed > 0 and len(writer.pages) > 0:
                with open(pdf_path, "wb") as f:
                    writer.write(f)

            return removed

        except Exception as e:
            logger.debug(f"移除空白页时出错: {e}")
            return 0

    # ==================== 上下文管理器 ====================

    def __enter__(self):
        self.initialize()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.cleanup()
        return False
