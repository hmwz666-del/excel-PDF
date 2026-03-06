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

    管理一个 Excel COM 实例，负责文件的打开、边距优化和 PDF 导出。
    设计为在单独的进程中运行（COM 对象不能跨线程共享）。

    策略：最小干预
    - 不修改 PrintArea（避免内容丢失/断页）
    - 不修改 Orientation（尊重原文件设置）
    - 不修改 Zoom/FitToPages（尊重原文件设置）
    - 只设置小边距（安全优化）
    - 导出后用 pypdf 删除末尾空白页
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

        pythoncom.CoInitialize()

        try:
            # DispatchEx 创建独立 Excel 进程，不影响用户正在使用的 Excel
            self.excel_app = win32com.client.DispatchEx("Excel.Application")
            self.excel_app.Visible = EXCEL_VISIBLE
            self.excel_app.DisplayAlerts = EXCEL_ALERTS
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

        # 构建输出路径（保留子目录结构）
        base_name = os.path.splitext(os.path.basename(excel_path))[0]

        if input_dir:
            rel_dir = os.path.relpath(os.path.dirname(excel_path), input_dir)
            target_dir = os.path.join(output_dir, rel_dir) if rel_dir != '.' else output_dir
            os.makedirs(target_dir, exist_ok=True)
        else:
            target_dir = output_dir

        pdf_path = os.path.join(target_dir, f"{base_name}.pdf")

        # 处理同名文件
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
                workbook = self.excel_app.Workbooks.Open(
                    os.path.abspath(excel_path),
                    ReadOnly=True,
                    UpdateLinks=0,
                    IgnoreReadOnlyRecommended=True,
                    CorruptLoad=1,
                )

                # 只设置边距（最小干预）
                self._set_margins(workbook)

                # 导出为 PDF
                workbook.ExportAsFixedFormat(
                    Type=XL_TYPE_PDF,
                    Filename=os.path.abspath(pdf_path),
                    Quality=0,
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False,
                )

                # 后处理：删除末尾空白页
                removed = self._remove_last_blank_page(pdf_path)

                if renamed:
                    msg = f"转换成功 (同名文件 {original_pdf_name} 已重命名为 {os.path.basename(pdf_path)})"
                    logger.info(f"✅ 转换成功: {os.path.basename(excel_path)} → {os.path.basename(pdf_path)}")
                else:
                    msg = "转换成功"
                    logger.info(f"✅ 转换成功: {os.path.basename(excel_path)}")

                if removed:
                    msg += " (已删除末尾空白页)"
                    logger.info(f"   🗑️ 已删除末尾空白页")

                return ConversionResult(
                    excel_path, ConversionResult.SUCCESS,
                    msg, pdf_path
                )

            except com_error as e:
                last_error = str(e)
                error_code = getattr(e, 'hresult', None)

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

    def _set_margins(self, workbook):
        """
        只设置页边距，不修改任何其他页面设置

        这是唯一安全的优化：缩小边距让内容有更多空间。
        不动 PrintArea / Orientation / Zoom / FitToPages。
        """
        for sheet in workbook.Worksheets:
            try:
                page_setup = sheet.PageSetup
                # 最小化页边距 (单位: 磅, 1英寸=72磅)
                page_setup.LeftMargin = 7.2     # ~0.1 英寸
                page_setup.RightMargin = 7.2
                page_setup.TopMargin = 14.4     # ~0.2 英寸
                page_setup.BottomMargin = 14.4
                page_setup.HeaderMargin = 0
                page_setup.FooterMargin = 0
            except Exception as e:
                logger.debug(f"设置工作表 '{sheet.Name}' 边距时出错: {e}")
                continue

    def _remove_last_blank_page(self, pdf_path):
        """
        检查并删除 PDF 最后一页（如果是空白页）

        只检查最后一页，不动其他任何页面。
        判定空白：没有可见文字 AND 没有图片。

        Returns:
            True 如果删除了空白页，False 如果没有
        """
        if not HAS_PYPDF:
            return False

        try:
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)

            if total_pages <= 1:
                return False

            # 只检查最后一页
            last_page = reader.pages[-1]

            # 检查是否有可见文字
            text = last_page.extract_text() or ""
            has_text = bool(re.search(r'[\w\u4e00-\u9fff]', text))

            # 检查是否有图片
            has_images = False
            try:
                if '/Resources' in last_page:
                    resources = last_page['/Resources']
                    if '/XObject' in resources:
                        xobjects = resources['/XObject']
                        if xobjects and len(xobjects) > 0:
                            has_images = True
            except Exception:
                pass

            # 有文字 OR 有图片 → 不是空白页，保留
            if has_text or has_images:
                return False

            # 确认是空白页，删除最后一页
            writer = PdfWriter()
            for i in range(total_pages - 1):
                writer.add_page(reader.pages[i])

            with open(pdf_path, "wb") as f:
                writer.write(f)

            return True

        except Exception as e:
            logger.debug(f"检查末尾空白页时出错: {e}")
            return False

    def __enter__(self):
        self.initialize()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.cleanup()
        return False
