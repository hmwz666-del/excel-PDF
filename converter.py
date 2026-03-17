# -*- coding: utf-8 -*-
"""
Excel 转 PDF 批量转换工具 - 核心转换引擎

使用 win32com COM 自动化调用 Microsoft Excel 进行高保真转换。
"""

import os
import re
import logging
import traceback
import tempfile
import shutil
import uuid

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
        self._original_printer = None

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

            # 统一打印机：避免不同电脑默认打印机不同导致分页差异
            self._set_unified_printer()

            self._initialized = True
            logger.info("Excel COM 实例初始化成功")
        except Exception as e:
            logger.error(f"无法启动 Excel: {e}")
            self.cleanup()
            raise RuntimeError(f"无法启动 Microsoft Excel: {e}")

    def cleanup(self):
        """清理 Excel COM 实例和 COM 环境"""
        if self.excel_app is not None:
            # 恢复原打印机
            self._restore_printer()
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

    def _set_unified_printer(self):
        """
        统一使用 'Microsoft Print to PDF' 打印机

        不同电脑的默认打印机不同，Excel COM 导出 PDF 时
        依赖打印机驱动计算页面布局（纸张大小、可打印区域等），
        同一个 Excel 在不同打印机下可能分页不同。

        统一使用系统自带的 PDF 打印机，确保一致的输出。
        """
        try:
            self._original_printer = self.excel_app.ActivePrinter
            logger.debug(f"原始打印机: {self._original_printer}")

            # Windows 自带的 PDF 打印机（Win10/11 都有）
            pdf_printer = "Microsoft Print to PDF"
            self.excel_app.ActivePrinter = pdf_printer
            logger.info(f"已切换打印机: {pdf_printer}")
        except Exception as e:
            # 如果切换失败（如打印机不存在），不影响正常转换
            logger.debug(f"切换打印机失败（使用默认打印机）: {e}")
            self._original_printer = None

    def _restore_printer(self):
        """恢复原始打印机设置"""
        if self._original_printer and self.excel_app:
            try:
                self.excel_app.ActivePrinter = self._original_printer
                logger.debug(f"已恢复打印机: {self._original_printer}")
            except Exception:
                pass

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

                # 预处理工作簿（根据文件类型自动选择策略）
                self._prepare_workbook(workbook)

                # 回退：不能使用 .tmp 因为 ExportAsFixedFormat 强制会自动加上 .pdf
                # 这会导致生成的是 .tmp.pdf，从而触发加密并让 python 找不到 .tmp 文件报错。
                temp_pdf_path = os.path.join(
                    tempfile.gettempdir(),
                    f"excel_to_pdf_temp_{uuid.uuid4().hex}.pdf"
                )

                # 第一步：先导出到系统临时目录（不被加密）
                workbook.ExportAsFixedFormat(
                    Type=XL_TYPE_PDF,
                    Filename=os.path.abspath(temp_pdf_path),
                    Quality=0,
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False,
                )

                # 第二步：在临时目录（未加密状态）下删除空白页
                removed = self._remove_last_blank_page(temp_pdf_path)

                # 第三步：将干净的 PDF 复制到最终目录（复制过程中会被企业软件自动加密）
                shutil.copy2(temp_pdf_path, os.path.abspath(pdf_path))

                # 第四步：安全销毁（阅后即焚）临时未加密文件
                try:
                    os.remove(temp_pdf_path)
                    logger.debug(f"已清理临时文件: {temp_pdf_path}")
                except Exception as e:
                    logger.warning(f"无法清理临时文件 {temp_pdf_path}: {e}")

                if renamed:
                    msg = f"转换成功 (同名文件 {original_pdf_name} 已重命名为 {os.path.basename(pdf_path)})"
                    logger.info(f"✅ 转换成功: {os.path.basename(excel_path)} → {os.path.basename(pdf_path)}")
                else:
                    msg = "转换成功"
                    logger.info(f"✅ 转换成功: {os.path.basename(excel_path)}")

                if removed:
                    msg += " (已删除尾部空白页)"
                    logger.info(f"   🗑️ 已删除尾部空白页")

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

    def _prepare_workbook(self, workbook):
        """
        预处理工作簿，根据文件特征选择不同策略

        策略1（有手动分页符，如报关单）：
          - 完全不修改页面设置（边距、缩放、方向等）
          - 只隐藏连续空行
          - 只删除数据末尾之后的多余分页符

        策略2（无手动分页符，如箱单/发票）：
          - 缩小边距
          - 宽度适配 FitToPagesWide
          - 隐藏连续空行
          - 调整分页符避免行跨页

        最后对所有工作表设置 PrintArea 到实际数据范围，
        从源头杜绝空白页（兼容企业透明加密环境）。
        """
        for sheet in workbook.Worksheets:
            try:
                has_manual_breaks = self._has_manual_page_breaks(sheet)

                if has_manual_breaks:
                    logger.debug(
                        f"工作表 '{sheet.Name}': 检测到手动分页符，保持原始布局"
                    )
                    # 只删除数据末尾后的多余分页符
                    self._remove_trailing_page_breaks(sheet)
                else:
                    logger.debug(
                        f"工作表 '{sheet.Name}': 无手动分页符，应用优化"
                    )
                    # 设置小边距
                    page_setup = sheet.PageSetup
                    page_setup.LeftMargin = 7.2
                    page_setup.RightMargin = 7.2
                    page_setup.TopMargin = 14.4
                    page_setup.BottomMargin = 14.4
                    page_setup.HeaderMargin = 0
                    page_setup.FooterMargin = 0
                    # 宽度适配
                    page_setup.Zoom = False
                    page_setup.FitToPagesWide = 1
                    page_setup.FitToPagesTall = False

            except Exception as e:
                logger.debug(f"预处理工作表 '{sheet.Name}' 时出错: {e}")
                continue

        # 隐藏连续空行（适用于所有文件）
        self._hide_empty_rows(workbook)

        # 设置 PrintArea 到实际数据范围（在所有预处理之后执行）
        # 这是空白页消除的核心手段，完全在 Excel COM 层面操作，
        # 不依赖 PDF 后处理，兼容企业透明加密环境
        for sheet in workbook.Worksheets:
            try:
                self._set_print_area_to_data(sheet)
            except Exception as e:
                logger.debug(f"设置 PrintArea 时出错 '{sheet.Name}': {e}")
                continue

    def _has_manual_page_breaks(self, sheet):
        """检查工作表是否有手动分页符"""
        try:
            return sheet.HPageBreaks.Count > 0
        except Exception:
            return False

    def _remove_trailing_page_breaks(self, sheet):
        """
        只删除最后有可见数据的行之后的手动分页符

        有些 Excel 文件（如报关单）在数据区域内有有意的手动分页符，
        不能用 ResetAllPageBreaks() 全部删除。
        只删除数据结束后的多余分页符（这些才是产生空白页的原因）。
        """
        try:
            used = sheet.UsedRange
            if used is None:
                return

            start_row = used.Row
            total_rows = used.Rows.Count

            # 一次性读取所有值（单次 COM 调用）
            all_values = used.Value
            if all_values is None:
                return

            # 单行时 COM 返回一维 tuple，统一转为二维
            if total_rows == 1:
                all_values = (all_values,)

            # 从最后一行往上找第一行有可见内容的行
            last_visible_row = 0
            for row_idx in range(len(all_values) - 1, -1, -1):
                row_data = all_values[row_idx]
                if row_data is None:
                    continue
                # 单列时 row_data 不是 tuple
                if not isinstance(row_data, tuple):
                    row_data = (row_data,)
                for val in row_data:
                    if val is not None and str(val).strip() != '':
                        last_visible_row = start_row + row_idx
                        break
                if last_visible_row > 0:
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

                # 一次性读取所有值（单次 COM 调用，替代数千次逐格调用）
                all_values = used_range.Value
                if all_values is None:
                    continue

                # 单行时 COM 返回一维 tuple，统一转为二维
                if total_rows == 1:
                    all_values = (all_values,)

                # 在 Python 内存中判断每行是否"视觉空行"
                empty_runs = []
                current_run_start = None

                for row_idx, row_data in enumerate(all_values):
                    is_visually_empty = True

                    if row_data is not None:
                        # 单列时 row_data 不是 tuple
                        cells = row_data if isinstance(row_data, tuple) else (row_data,)
                        for val in cells:
                            if val is not None and str(val).strip() != '':
                                is_visually_empty = False
                                break

                    actual_row = start_row + row_idx

                    if is_visually_empty:
                        if current_run_start is None:
                            current_run_start = actual_row
                    else:
                        if current_run_start is not None:
                            run_length = actual_row - current_run_start
                            if run_length >= 3:
                                empty_runs.append((current_run_start, actual_row - 1))
                            current_run_start = None

                # 处理末尾的连续空行
                if current_run_start is not None:
                    last_row = start_row + total_rows - 1
                    run_length = last_row - current_run_start + 1
                    if run_length >= 3:
                        empty_runs.append((current_run_start, last_row))

                # 隐藏连续空行（仅需少量 COM 调用）
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

    def _set_print_area_to_data(self, sheet):
        """
        将工作表的 PrintArea 精确设置到实际有数据的范围

        使用 Excel 的 Cells.Find 方法快速定位最后有数据的行和列，
        然后设置 PrintArea 限制打印范围，从源头杜绝空白页。

        规则：
        - 如果已有用户设置的 PrintArea → 不覆盖（保证数据完整性）
        - 如果没有 PrintArea → 设置为 $A$1:${最后列}${最后行}
        - 找不到数据时 → 不设置（避免误操作）
        """
        try:
            # 检查是否已有 PrintArea（用户手动设置的，不覆盖）
            current_print_area = sheet.PageSetup.PrintArea
            if current_print_area and str(current_print_area).strip():
                logger.debug(
                    f"工作表 '{sheet.Name}': 已有 PrintArea="
                    f"{current_print_area}，保持不变"
                )
                return

            # Excel COM 常量
            xlByRows = 1       # 按行搜索
            xlByColumns = 2    # 按列搜索
            xlPrevious = 2     # 从后往前搜索
            xlValues = -4163   # 搜索值（非公式）
            xlPart = 2         # 部分匹配

            # 用 Find 快速定位最后有数据的行（从 A1 反向搜索，绕到最后一个单元格）
            last_row_cell = sheet.Cells.Find(
                What="*",
                After=sheet.Cells(1, 1),
                LookIn=xlValues,
                LookAt=xlPart,
                SearchOrder=xlByRows,
                SearchDirection=xlPrevious,
            )

            # 用 Find 快速定位最后有数据的列
            last_col_cell = sheet.Cells.Find(
                What="*",
                After=sheet.Cells(1, 1),
                LookIn=xlValues,
                LookAt=xlPart,
                SearchOrder=xlByColumns,
                SearchDirection=xlPrevious,
            )

            last_data_row = last_row_cell.Row if last_row_cell else 1
            last_data_col = last_col_cell.Column if last_col_cell else 1

            # --- 核心第一步：探测纯数据本应该占用的完美页数 ---
            data_col_letter = self._col_num_to_letter(last_data_col)
            sheet.PageSetup.PrintArea = f"$A$1:${data_col_letter}${last_data_row}"
            # 读取当前水平分页数。页数 = HPageBreaks 数量 + 1
            data_tall_pages = sheet.HPageBreaks.Count + 1

            # 遍历所有浮动图片/形状（如盖章），防止它们被 PrintArea 裁剪
            # Forms, Pictures, OLEObjects 等都包含在 Shapes 集合中
            max_shape_row = 1
            max_shape_col = 1
            
            try:
                for shape in sheet.Shapes:
                    try:
                        # 获取形状的右下角所在单元格
                        br_cell = shape.BottomRightCell
                        if br_cell:
                            max_shape_row = max(max_shape_row, br_cell.Row)
                            max_shape_col = max(max_shape_col, br_cell.Column)
                    except Exception:
                        continue
            except Exception as e:
                logger.debug(f"检查 Shapes 时出错 '{sheet.Name}': {e}")

            # 最终边界取 数据边界 和 形状边界 的最大值
            final_row = max(last_data_row, max_shape_row)
            final_col = max(last_data_col, max_shape_col)

            if final_row == 1 and final_col == 1 and last_row_cell is None:
                logger.debug(f"工作表 '{sheet.Name}': 既无数据也无图片，跳过 PrintArea 设置")
                return

            # --- 核心第二步：设置包含 Shapes 的最终 PrintArea ---
            # 转换列号为字母
            col_letter = self._col_num_to_letter(final_col)

            # 设置 PrintArea
            print_area = f"$A$1:${col_letter}${final_row}"
            sheet.PageSetup.PrintArea = print_area

            logger.info(
                f"  📐 工作表 '{sheet.Name}': "
                f"PrintArea → {print_area}"
            )

            # --- 核心第三步：消除盖章透明底边造成的虚无溢出空白页 ---
            # 如果加入 Shapes 后，导致数据的右边或下边延展了（通常是盖章的透明 Padding 造成的）
            if final_row > last_data_row or final_col > last_data_col:
                logger.info(
                    f"  🔄 发现形状(盖章)边界超出纯数据区 (行: {last_data_row}->{final_row}, 列: {last_data_col}->{final_col})。 "
                    f"启用防溢出收缩，强制束缚在 {data_tall_pages} 页内以消灭空白页。"
                )
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = data_tall_pages

        except Exception as e:
            # 设置失败不影响正常转换
            logger.debug(f"设置 PrintArea 失败 '{sheet.Name}': {e}")

    @staticmethod
    def _col_num_to_letter(col_num):
        """列号转 Excel 列字母 (1→A, 26→Z, 27→AA, 702→ZZ, 703→AAA)"""
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def _remove_last_blank_page(self, pdf_path):
        """
        从 PDF 末尾循环删除所有连续的空白页

        从最后一页往前检查，遇到有内容的页就停止。
        判定空白：没有可见文字 AND 没有实际图片（忽略表单/字体等 XObject）。

        支持处理被系统自动加密的 PDF（如企业电脑的透明加密软件）。

        Returns:
            True 如果删除了空白页，False 如果没有
        """
        if not HAS_PYPDF:
            logger.warning("⚠️ pypdf 未安装，无法自动删除空白页。请运行: pip install pypdf")
            return False

        try:
            reader = PdfReader(pdf_path)

            # 处理加密的 PDF（企业电脑可能自动加密输出文件）
            if reader.is_encrypted:
                logger.info(f"  🔒 检测到 PDF 已加密，尝试解密...")
                try:
                    # 尝试用空密码解密（大部分自动加密软件使用空密码或 owner 密码）
                    decrypt_result = reader.decrypt("")
                    if decrypt_result == 0:
                        # 空密码解密失败，再尝试不提供密码直接读取
                        logger.warning(
                            f"  ⚠️ PDF 已加密且无法用空密码解密，跳过空白页处理。"
                            f"这可能是企业加密软件导致的，不影响 PDF 内容正确性。"
                        )
                        return False
                    logger.info(f"  🔓 PDF 解密成功，继续处理空白页")
                except Exception as e:
                    logger.warning(
                        f"  ⚠️ PDF 解密失败: {e}，跳过空白页处理。"
                        f"这可能是企业加密软件导致的，不影响 PDF 内容正确性。"
                    )
                    return False

            total_pages = len(reader.pages)

            if total_pages <= 1:
                return False

            # 从末尾往前检查，找到最后一个有内容的页
            last_content_page = total_pages - 1

            while last_content_page >= 1:  # 至少保留第1页
                page = reader.pages[last_content_page]

                # 检查是否有可见文字
                try:
                    text = page.extract_text() or ""
                except Exception:
                    # 加密或损坏的页面可能无法提取文字，视为有内容（保险起见不删）
                    logger.debug(f"  第 {last_content_page + 1} 页无法提取文字，跳过")
                    break

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
            # 捕获所有异常，包括加密导致的读取失败
            error_msg = str(e).lower()
            if 'encrypt' in error_msg or 'password' in error_msg or 'decrypt' in error_msg:
                logger.warning(
                    f"  ⚠️ PDF 文件已加密，无法处理空白页: {e}。"
                    f"这可能是企业电脑自动加密导致的。"
                )
            else:
                logger.debug(f"检查末尾空白页时出错: {e}")
            return False

    def __enter__(self):
        self.initialize()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.cleanup()
        return False
