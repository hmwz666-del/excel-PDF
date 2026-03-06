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

                # 隐藏连续空白行（避免空边框行跨页）
                self._hide_empty_rows(workbook)

                # 调整分页符，防止行跨页切割
                self._adjust_page_breaks(workbook)

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

    def _remove_trailing_page_breaks(self, sheet):
        """
        只删除最后有可见数据的行之后的手动分页符

        有些 Excel 文件（如报关单）在数据区域内有有意的手动分页符，
        不能用 ResetAllPageBreaks() 全部删除。
        只删除数据结束后的多余分页符（这些才是产生空白页的原因）。
        """
        try:
            # 找最后一行有可见内容的行号
            used = sheet.UsedRange
            if used is None:
                return

            last_visible_row = 0
            total_rows = used.Rows.Count
            start_row = used.Row
            total_cols = min(used.Columns.Count, 30)
            start_col = used.Column

            # 从最后一行往上找第一行有可见内容的行
            for r in range(start_row + total_rows - 1, start_row - 1, -1):
                has_visible = False
                for c in range(start_col, start_col + total_cols):
                    try:
                        val = sheet.Cells(r, c).Value
                        if val is not None and str(val).strip() != '':
                            has_visible = True
                            break
                    except Exception:
                        continue
                if has_visible:
                    last_visible_row = r
                    break

            if last_visible_row == 0:
                return

            # 从后往前删除位于 last_visible_row 之后的手动分页符
            breaks = sheet.HPageBreaks
            for i in range(breaks.Count, 0, -1):
                try:
                    pb = breaks(i)
                    if pb.Location.Row > last_visible_row:
                        pb.Delete()
                except Exception:
                    continue

        except Exception as e:
            logger.debug(f"删除尾部分页符时出错: {e}")

    def _set_margins(self, workbook):
        """
        设置小边距 + 删除尾部多余分页符

        - 删除数据末尾之后的手动分页符（避免空白页）
        - 保留数据区域内的分页符（如报关单的有意分页）
        - 缩小边距：让内容有更多空间
        """
        for sheet in workbook.Worksheets:
            try:
                # 只删除数据区域之后的多余分页符（保留有意的分页符）
                self._remove_trailing_page_breaks(sheet)

                # 设置小边距
                page_setup = sheet.PageSetup
                page_setup.LeftMargin = 7.2     # ~0.1 英寸
                page_setup.RightMargin = 7.2
                page_setup.TopMargin = 14.4     # ~0.2 英寸
                page_setup.BottomMargin = 14.4
                page_setup.HeaderMargin = 0
                page_setup.FooterMargin = 0

                # 只在没有手动分页符的工作表上使用宽度适配
                # 有手动分页符的文档（如报关单）需要保持原始布局
                has_manual_breaks = False
                try:
                    for pb in sheet.HPageBreaks:
                        if pb.Type == 1:  # xlPageBreakManual = -4135, but Type=1 for manual
                            has_manual_breaks = True
                            break
                except Exception:
                    pass

                if not has_manual_breaks:
                    page_setup.Zoom = False
                    page_setup.FitToPagesWide = 1    # 宽度缩放到1页
                    page_setup.FitToPagesTall = False  # 高度不限制
            except Exception as e:
                logger.debug(f"设置工作表 '{sheet.Name}' 边距时出错: {e}")
                continue

    def _hide_empty_rows(self, workbook):
        """
        隐藏连续的空白行（只有空字符串和边框，无可见内容）

        很多 Excel 模板在数据区和汇总区之间有大量空行，
        这些行的单元格值是空字符串 '' 而非 None，
        导致 Excel 认为它们有数据并打印出来。

        规则：连续 3 行以上的"全空行"（所有值都是 None 或空字符串），
        整批隐藏。保留单独的空行（可能是有意的间距）。
        """
        for sheet in workbook.Worksheets:
            try:
                used_range = sheet.UsedRange
                if used_range is None:
                    continue

                total_rows = used_range.Rows.Count
                start_row = used_range.Row
                total_cols = min(used_range.Columns.Count, 30)
                start_col = used_range.Column

                # 扫描所有行，标记哪些是"视觉空行"
                empty_runs = []   # [(start, end), ...]
                current_run_start = None

                for r in range(start_row, start_row + total_rows):
                    is_visually_empty = True
                    for c in range(start_col, start_col + total_cols):
                        try:
                            val = sheet.Cells(r, c).Value
                            if val is not None and str(val).strip() != '':
                                is_visually_empty = False
                                break
                        except Exception:
                            continue

                    if is_visually_empty:
                        if current_run_start is None:
                            current_run_start = r
                    else:
                        if current_run_start is not None:
                            run_length = r - current_run_start
                            if run_length >= 3:
                                empty_runs.append((current_run_start, r - 1))
                            current_run_start = None

                # 处理末尾的连续空行
                if current_run_start is not None:
                    run_length = (start_row + total_rows) - current_run_start
                    if run_length >= 3:
                        empty_runs.append((current_run_start, start_row + total_rows - 1))

                # 隐藏连续空行
                for run_start, run_end in empty_runs:
                    try:
                        hide_range = sheet.Range(
                            sheet.Rows(run_start),
                            sheet.Rows(run_end)
                        )
                        hide_range.Hidden = True
                        logger.debug(
                            f"工作表 '{sheet.Name}': "
                            f"隐藏空行 {run_start}-{run_end} ({run_end - run_start + 1} 行)"
                        )
                    except Exception:
                        continue

            except Exception as e:
                logger.debug(f"检测工作表 '{sheet.Name}' 空行时出错: {e}")
                continue

    def _adjust_page_breaks(self, workbook):
        """
        调整自动分页符位置，防止行被跨页切割

        检查每个自动分页符：
        - 如果分页符落在合并单元格区域中间 → 上移到合并区域开始行
        - 如果分页符落在非空行上 → 上移到该行之前
        这样确保每一行完整地出现在同一页上。
        """
        for sheet in workbook.Worksheets:
            try:
                breaks = sheet.HPageBreaks
                if breaks.Count == 0:
                    continue

                used_cols = min(sheet.UsedRange.Columns.Count, 20)
                start_col = sheet.UsedRange.Column

                # 从后往前调整（避免修改后影响后续索引）
                for i in range(breaks.Count, 0, -1):
                    try:
                        pb = breaks(i)
                        break_row = pb.Location.Row

                        # 检查分页符所在行的前几列是否有合并单元格
                        new_break_row = break_row
                        for col_idx in range(start_col, start_col + used_cols):
                            try:
                                cell = sheet.Cells(break_row, col_idx)
                                if cell.MergeCells:
                                    merge_start = cell.MergeArea.Row
                                    if merge_start < break_row:
                                        # 分页符在合并区域中间，需要上移
                                        new_break_row = min(new_break_row, merge_start)
                            except Exception:
                                continue

                        if new_break_row < break_row:
                            sheet.HPageBreaks.Add(sheet.Cells(new_break_row, 1))
                            logger.debug(
                                f"工作表 '{sheet.Name}': "
                                f"分页符从第 {break_row} 行上移到第 {new_break_row} 行"
                            )
                    except Exception:
                        continue

            except Exception as e:
                logger.debug(f"调整工作表 '{sheet.Name}' 分页符时出错: {e}")
                continue

    def _remove_last_blank_page(self, pdf_path):
        """
        从 PDF 末尾循环删除所有连续的空白页

        从最后一页往前检查，遇到有内容的页就停止。
        判定空白：没有可见文字 AND 没有实际图片（忽略表单/字体等 XObject）。

        Returns:
            True 如果删除了空白页，False 如果没有
        """
        if not HAS_PYPDF:
            logger.warning("⚠️ pypdf 未安装，无法自动删除空白页。请运行: pip install pypdf")
            return False

        try:
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)

            if total_pages <= 1:
                return False

            # 从末尾往前检查，找到最后一个有内容的页
            last_content_page = total_pages - 1

            while last_content_page >= 1:  # 至少保留第1页
                page = reader.pages[last_content_page]

                # 检查是否有可见文字
                text = page.extract_text() or ""
                has_text = bool(re.search(r'[\w\u4e00-\u9fff]', text))

                if has_text:
                    break  # 有文字，停止

                # 检查是否有实际图片（不是表单/字体等 XObject）
                has_real_image = False
                try:
                    if '/Resources' in page:
                        resources = page['/Resources']
                        if '/XObject' in resources:
                            xobjects = resources['/XObject']
                            if xobjects:
                                for key in xobjects:
                                    try:
                                        xobj = xobjects[key]
                                        obj = xobj.get_object() if hasattr(xobj, 'get_object') else xobj
                                        # 只有 Subtype 为 /Image 的才是真正的图片
                                        subtype = obj.get('/Subtype', '')
                                        if subtype == '/Image':
                                            has_real_image = True
                                            break
                                    except Exception:
                                        continue
                except Exception:
                    pass

                if has_real_image:
                    break  # 有图片，停止

                # 这一页既没有文字也没有图片 → 空白页，继续往前检查
                logger.debug(f"  检测到空白页: 第 {last_content_page + 1} 页")
                last_content_page -= 1

            # 计算要删除的页数
            pages_to_keep = last_content_page + 1
            removed_count = total_pages - pages_to_keep

            if removed_count <= 0:
                return False

            # 重建 PDF，只保留有内容的页
            writer = PdfWriter()
            for i in range(pages_to_keep):
                writer.add_page(reader.pages[i])

            with open(pdf_path, "wb") as f:
                writer.write(f)

            logger.info(f"  🗑️ 已删除末尾 {removed_count} 个空白页")
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
