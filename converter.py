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

    def _remove_blank_pages(self, pdf_path):
        """
        移除 PDF 中的空白页

        检测每页的文本内容，如果文本量非常少（< 10个字符）
        则判定为空白页并删除。

        Args:
            pdf_path: PDF 文件路径

        Returns:
            被删除的空白页数量
        """
        if not HAS_PYPDF:
            return 0

        try:
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)

            if total_pages <= 1:
                # 只有1页，不删除（避免生成空文件）
                return 0

            writer = PdfWriter()
            removed = 0

            for i, page in enumerate(reader.pages):
                text = page.extract_text() or ""
                # 只删除完全没有任何内容的页面
                clean_text = text.strip()

                if len(clean_text) == 0:
                    # 空白页，跳过
                    removed += 1
                    logger.debug(f"  删除空白页: 第 {i+1} 页")
                else:
                    writer.add_page(page)

            if removed > 0 and len(writer.pages) > 0:
                # 重新写入 PDF（覆盖原文件）
                with open(pdf_path, "wb") as f:
                    writer.write(f)

            return removed

        except Exception as e:
            logger.debug(f"移除空白页时出错: {e}")
            return 0

    def _optimize_print_area(self, workbook):
        """
        优化页面布局设置

        策略：不主动设置 PrintArea（避免数据丢失），
        只调整页边距、方向和缩放比例，让 Excel 自己决定打印内容。
        """
        for sheet in workbook.Worksheets:
            try:
                used_range = sheet.UsedRange
                if used_range is None or used_range.Count == 0:
                    # 空工作表，隐藏以避免空白页
                    if workbook.Worksheets.Count > 1:
                        sheet.Visible = False
                    continue

                # 清除任何已有的打印区域限制，让 Excel 导出全部内容
                sheet.PageSetup.PrintArea = ""

                # ===== 页面布局优化 =====
                page_setup = sheet.PageSetup

                # 自动判断横向/纵向（列多用横向）
                cols = used_range.Columns.Count
                if cols > 8:
                    page_setup.Orientation = 2  # xlLandscape (横向)
                else:
                    page_setup.Orientation = 1  # xlPortrait (纵向)

                # 自动缩放适配页面宽度
                page_setup.Zoom = False
                page_setup.FitToPagesWide = 1    # 宽度缩放到1页
                page_setup.FitToPagesTall = False  # 高度不限制

                # 最小化页边距 (单位: 磅, 1英寸=72磅)
                page_setup.LeftMargin = 7.2     # 约 0.1 英寸
                page_setup.RightMargin = 7.2
                page_setup.TopMargin = 14.4     # 约 0.2 英寸
                page_setup.BottomMargin = 14.4
                page_setup.HeaderMargin = 0
                page_setup.FooterMargin = 0

                # 水平居中打印
                page_setup.CenterHorizontally = True

            except Exception as e:
                logger.debug(f"设置工作表 '{sheet.Name}' 页面布局时出错: {e}")
                continue

    def __enter__(self):
        self.initialize()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.cleanup()
        return False
