# -*- coding: utf-8 -*-
"""
Excel 转 PDF 批量转换工具 - 多进程工作模块

管理多个 Excel 转换进程，实现并发文件转换。
每个工作进程独立管理一个 Excel COM 实例。
"""

import os
import logging
import multiprocessing
from multiprocessing import Process, Queue
import time

from config import SUPPORTED_EXTENSIONS
from converter import ExcelConverter, ConversionResult

logger = logging.getLogger(__name__)


def worker_process(task_queue, result_queue, worker_id):
    """
    工作进程的主函数

    从任务队列获取文件路径，使用独立的 Excel COM 实例进行转换，
    将结果放入结果队列。

    Args:
        task_queue: 输入任务队列 (excel_path, output_dir)
        result_queue: 输出结果队列 (ConversionResult)
        worker_id: 工作进程编号
    """
    converter = ExcelConverter()

    try:
        converter.initialize()
        logger.info(f"工作进程 {worker_id} 启动成功")

        while True:
            try:
                # 获取任务，超时 2 秒
                task = task_queue.get(timeout=2)
            except Exception:
                # 队列为空或超时，检查是否有终止信号
                if task_queue.empty():
                    break
                continue

            # 终止信号
            if task is None:
                break

            excel_path, output_dir, input_dir = task

            try:
                result = converter.convert_file(excel_path, output_dir, input_dir)
                result_queue.put(result)
            except Exception as e:
                result = ConversionResult(
                    excel_path, ConversionResult.FAILED,
                    f"进程内部错误: {str(e)}"
                )
                result_queue.put(result)

    except Exception as e:
        logger.error(f"工作进程 {worker_id} 初始化失败: {e}")
        # 把当前队列中所有任务标记为失败
        while not task_queue.empty():
            try:
                task = task_queue.get_nowait()
                if task is not None:
                    excel_path, output_dir, input_dir = task
                    result_queue.put(ConversionResult(
                        excel_path, ConversionResult.FAILED,
                        f"工作进程启动失败: {str(e)}"
                    ))
            except Exception:
                break

    finally:
        converter.cleanup()
        logger.info(f"工作进程 {worker_id} 已退出")


def scan_excel_files(input_dir):
    """
    扫描目录下所有支持的 Excel 文件

    Args:
        input_dir: 输入目录路径

    Returns:
        文件绝对路径列表
    """
    files = []
    for root, dirs, filenames in os.walk(input_dir):
        # 跳过输出目录和隐藏目录
        dirs[:] = [d for d in dirs if not d.startswith('.') and d != 'PDF_Output']

        for filename in filenames:
            # 跳过临时文件（以 ~ 开头）
            if filename.startswith('~'):
                continue

            ext = os.path.splitext(filename)[1].lower()
            if ext in SUPPORTED_EXTENSIONS:
                files.append(os.path.join(root, filename))

    # 按文件名排序，方便追踪
    files.sort()
    return files


class ConversionManager:
    """
    转换管理器

    负责协调多个工作进程，管理任务分发和结果收集。
    """

    def __init__(self, num_workers=4, progress_callback=None, log_callback=None):
        """
        Args:
            num_workers: 工作进程数量
            progress_callback: 进度回调 (current, total, result)
            log_callback: 日志回调 (message)
        """
        self.num_workers = num_workers
        self.progress_callback = progress_callback
        self.log_callback = log_callback
        self._workers = []
        self._is_running = False
        self._should_stop = False

    def start_conversion(self, input_dir, output_dir):
        """
        启动批量转换

        Args:
            input_dir: Excel 文件目录
            output_dir: PDF 输出目录

        Returns:
            (success_count, failed_count, skipped_count, results_list)
        """
        self._is_running = True
        self._should_stop = False

        # 扫描文件
        self._log(f"📂 扫描目录: {input_dir}")
        files = scan_excel_files(input_dir)
        total = len(files)

        if total == 0:
            self._log("⚠️ 未找到任何 Excel 文件")
            self._is_running = False
            return 0, 0, 0, []

        self._log(f"📋 发现 {total} 个 Excel 文件")

        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)

        # 创建任务队列和结果队列
        task_queue = Queue()
        result_queue = Queue()

        # 填充任务队列
        for filepath in files:
            task_queue.put((filepath, output_dir, input_dir))

        # 添加终止信号（每个 worker 一个）
        actual_workers = min(self.num_workers, total)
        for _ in range(actual_workers):
            task_queue.put(None)

        # 启动工作进程
        self._log(f"🚀 启动 {actual_workers} 个工作进程...")
        self._workers = []
        for i in range(actual_workers):
            p = Process(
                target=worker_process,
                args=(task_queue, result_queue, i + 1),
                daemon=True
            )
            p.start()
            self._workers.append(p)

        # 收集结果
        results = []
        success_count = 0
        failed_count = 0
        skipped_count = 0
        completed = 0

        start_time = time.time()

        while completed < total:
            if self._should_stop:
                self._log("⏹️ 用户请求停止转换")
                break

            try:
                result = result_queue.get(timeout=30)  # 单个文件最多等 30 秒
                completed += 1
                results.append(result)

                if result.status == ConversionResult.SUCCESS:
                    success_count += 1
                    self._log(f"✅ [{completed}/{total}] {result.filename}")
                elif result.status == ConversionResult.SKIPPED:
                    skipped_count += 1
                    self._log(f"⏭️ [{completed}/{total}] {result.filename} - {result.message}")
                else:
                    failed_count += 1
                    self._log(f"❌ [{completed}/{total}] {result.filename} - {result.message}")

                if self.progress_callback:
                    self.progress_callback(completed, total, result)

            except Exception:
                # 超时 - 检查进程是否还活着
                alive = any(p.is_alive() for p in self._workers)
                if not alive:
                    self._log("⚠️ 所有工作进程已退出")
                    break

        elapsed = time.time() - start_time

        # 等待所有进程结束
        for p in self._workers:
            p.join(timeout=10)
            if p.is_alive():
                p.terminate()

        self._workers = []
        self._is_running = False

        # 输出统计
        self._log(f"\n{'='*50}")
        self._log(f"📊 转换完成统计:")
        self._log(f"   总计: {total} 个文件")
        self._log(f"   ✅ 成功: {success_count}")
        self._log(f"   ❌ 失败: {failed_count}")
        self._log(f"   ⏭️ 跳过: {skipped_count}")
        self._log(f"   ⏱️ 耗时: {elapsed:.1f} 秒 ({elapsed/60:.1f} 分钟)")
        self._log(f"{'='*50}")

        return success_count, failed_count, skipped_count, results

    def stop(self):
        """请求停止转换"""
        self._should_stop = True
        self._log("🛑 正在停止转换...")

    @property
    def is_running(self):
        return self._is_running

    def _log(self, message):
        """输出日志"""
        if self.log_callback:
            self.log_callback(message)
        logger.info(message)
