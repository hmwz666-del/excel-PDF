# -*- coding: utf-8 -*-
"""
Excel 转 PDF 批量转换工具 - GUI 界面模块

基于 tkinter 的图形用户界面，提供文件夹选择、并发设置、
实时进度显示和日志输出功能。
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import datetime
import logging

from config import (
    APP_NAME, APP_VERSION,
    DEFAULT_WORKERS, MIN_WORKERS, MAX_WORKERS,
    WINDOW_WIDTH, WINDOW_HEIGHT,
    WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT,
    THEME, LOG_FILENAME, LOG_DATE_FORMAT, LOG_ENCODING,
)
from worker import ConversionManager

logger = logging.getLogger(__name__)


class ExcelToPdfApp:
    """Excel 转 PDF 图形界面应用"""

    def __init__(self, root):
        self.root = root
        self.manager = None
        self._conversion_thread = None
        self._log_file = None
        self._close_after_stop = False

        # 界面变量
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.worker_count = tk.IntVar(value=DEFAULT_WORKERS)
        self.progress_var = tk.DoubleVar(value=0)
        self.status_text = tk.StringVar(value="就绪")

        self._setup_window()
        self._create_widgets()
        self._setup_logging()

    def _setup_window(self):
        """配置主窗口"""
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
        self.root.configure(bg=THEME["bg"])

        # 窗口居中
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - WINDOW_WIDTH) // 2
        y = (self.root.winfo_screenheight() - WINDOW_HEIGHT) // 2
        self.root.geometry(f"+{x}+{y}")

        # 关闭窗口时的处理
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _create_widgets(self):
        """创建界面组件"""
        # 主容器
        main_frame = tk.Frame(self.root, bg=THEME["bg"], padx=15, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ========== 标题区 ==========
        title_frame = tk.Frame(main_frame, bg=THEME["bg"])
        title_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(
            title_frame, text="📊 " + APP_NAME,
            font=("Microsoft YaHei UI", 16, "bold"),
            bg=THEME["bg"], fg=THEME["text"]
        ).pack(side=tk.LEFT)

        tk.Label(
            title_frame, text=f"v{APP_VERSION}",
            font=("Microsoft YaHei UI", 9),
            bg=THEME["bg"], fg=THEME["text_secondary"]
        ).pack(side=tk.LEFT, padx=(8, 0), pady=(6, 0))

        # ========== 文件夹选择区 ==========
        folder_card = self._create_card(main_frame, "📁 文件夹设置")

        # 输入目录
        input_row = tk.Frame(folder_card, bg=THEME["card_bg"])
        input_row.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            input_row, text="Excel 目录:",
            font=("Microsoft YaHei UI", 9),
            bg=THEME["card_bg"], fg=THEME["text"], width=10, anchor="w"
        ).pack(side=tk.LEFT)

        input_entry = tk.Entry(
            input_row, textvariable=self.input_dir,
            font=("Microsoft YaHei UI", 9),
            relief=tk.SOLID, bd=1
        )
        input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        self._create_button(
            input_row, "浏览...", self._browse_input,
            width=8, style="default"
        ).pack(side=tk.RIGHT)

        # 输出目录
        output_row = tk.Frame(folder_card, bg=THEME["card_bg"])
        output_row.pack(fill=tk.X)

        tk.Label(
            output_row, text="输出目录:",
            font=("Microsoft YaHei UI", 9),
            bg=THEME["card_bg"], fg=THEME["text"], width=10, anchor="w"
        ).pack(side=tk.LEFT)

        output_entry = tk.Entry(
            output_row, textvariable=self.output_dir,
            font=("Microsoft YaHei UI", 9),
            relief=tk.SOLID, bd=1
        )
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        self._create_button(
            output_row, "浏览...", self._browse_output,
            width=8, style="default"
        ).pack(side=tk.RIGHT)

        # ========== 设置区 ==========
        settings_card = self._create_card(main_frame, "⚙️ 转换设置")

        settings_row = tk.Frame(settings_card, bg=THEME["card_bg"])
        settings_row.pack(fill=tk.X)

        tk.Label(
            settings_row, text="并发进程数:",
            font=("Microsoft YaHei UI", 9),
            bg=THEME["card_bg"], fg=THEME["text"]
        ).pack(side=tk.LEFT)

        worker_spin = tk.Spinbox(
            settings_row,
            from_=MIN_WORKERS, to=MAX_WORKERS,
            textvariable=self.worker_count,
            font=("Microsoft YaHei UI", 9),
            width=5, relief=tk.SOLID, bd=1
        )
        worker_spin.pack(side=tk.LEFT, padx=(8, 0))

        tk.Label(
            settings_row,
            text=f"（推荐 4，范围 {MIN_WORKERS}-{MAX_WORKERS}）",
            font=("Microsoft YaHei UI", 8),
            bg=THEME["card_bg"], fg=THEME["text_secondary"]
        ).pack(side=tk.LEFT, padx=(8, 0))

        # ========== 操作按钮区 ==========
        btn_frame = tk.Frame(main_frame, bg=THEME["bg"])
        btn_frame.pack(fill=tk.X, pady=10)

        self.start_btn = self._create_button(
            btn_frame, "🚀 开始转换", self._start_conversion,
            width=20, style="primary"
        )
        self.start_btn.pack(side=tk.LEFT)

        self.stop_btn = self._create_button(
            btn_frame, "⏹️ 停止", self._stop_conversion,
            width=12, style="danger"
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(10, 0))
        self.stop_btn.configure(state=tk.DISABLED)

        self.open_dir_btn = self._create_button(
            btn_frame, "📂 打开输出目录", self._open_output_dir,
            width=15, style="default"
        )
        self.open_dir_btn.pack(side=tk.RIGHT)

        # ========== 进度条区 ==========
        progress_frame = tk.Frame(main_frame, bg=THEME["bg"])
        progress_frame.pack(fill=tk.X, pady=(0, 5))

        style = ttk.Style()
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=THEME["border"],
            background=THEME["primary"],
            thickness=20
        )

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill=tk.X, side=tk.LEFT, expand=True)

        self.progress_label = tk.Label(
            progress_frame, text="0%",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg=THEME["bg"], fg=THEME["primary"], width=6
        )
        self.progress_label.pack(side=tk.RIGHT, padx=(8, 0))

        # ========== 日志区 ==========
        log_frame = tk.LabelFrame(
            main_frame, text=" 📋 运行日志 ",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg=THEME["bg"], fg=THEME["text"],
            relief=tk.GROOVE, bd=1
        )
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        # 日志文本框 + 滚动条
        log_container = tk.Frame(log_frame, bg=THEME["log_bg"])
        log_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        scrollbar = tk.Scrollbar(log_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(
            log_container,
            font=("Consolas", 9),
            bg=THEME["log_bg"], fg=THEME["log_text"],
            relief=tk.FLAT, wrap=tk.WORD,
            yscrollcommand=scrollbar.set,
            state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)

        # 设置日志标签颜色
        self.log_text.tag_configure("success", foreground="#52c41a")
        self.log_text.tag_configure("error", foreground="#ff4d4f")
        self.log_text.tag_configure("warning", foreground="#faad14")
        self.log_text.tag_configure("info", foreground="#1890ff")
        self.log_text.tag_configure("normal", foreground="#d4d4d4")

        # ========== 状态栏 ==========
        status_frame = tk.Frame(main_frame, bg=THEME["border"], height=1)
        status_frame.pack(fill=tk.X, pady=(5, 0))

        self.status_label = tk.Label(
            main_frame, textvariable=self.status_text,
            font=("Microsoft YaHei UI", 8),
            bg=THEME["bg"], fg=THEME["text_secondary"],
            anchor="w"
        )
        self.status_label.pack(fill=tk.X)

    def _create_card(self, parent, title):
        """创建卡片样式的容器"""
        frame = tk.LabelFrame(
            parent, text=f" {title} ",
            font=("Microsoft YaHei UI", 9, "bold"),
            bg=THEME["card_bg"], fg=THEME["text"],
            relief=tk.GROOVE, bd=1,
            padx=12, pady=8
        )
        frame.pack(fill=tk.X, pady=(0, 8))
        return frame

    def _create_button(self, parent, text, command, width=10, style="default"):
        """创建统一风格的按钮"""
        colors = {
            "primary": (THEME["primary"], "white"),
            "danger": (THEME["error"], "white"),
            "default": (THEME["card_bg"], THEME["text"]),
        }
        bg, fg = colors.get(style, colors["default"])

        btn = tk.Button(
            parent, text=text, command=command,
            font=("Microsoft YaHei UI", 9),
            bg=bg, fg=fg, width=width,
            relief=tk.RAISED, bd=1,
            cursor="hand2",
            activebackground=bg, activeforeground=fg,
        )
        return btn

    def _setup_logging(self):
        """配置文件日志"""
        timestamp = datetime.datetime.now().strftime(LOG_DATE_FORMAT)
        log_filename = LOG_FILENAME.format(timestamp=timestamp)

        # 获取程序所在目录
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))

        log_path = os.path.join(app_dir, log_filename)

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s [%(levelname)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S',
            handlers=[
                logging.FileHandler(log_path, encoding=LOG_ENCODING),
            ]
        )
        self._log_file_path = log_path
        self._append_log(f"日志文件: {log_path}", "info")

    # ========== 事件处理 ==========

    def _browse_input(self):
        """选择输入文件夹"""
        folder = filedialog.askdirectory(title="选择 Excel 文件所在的文件夹")
        if folder:
            self.input_dir.set(folder)
            # 如果输出目录为空，自动设置默认输出目录
            if not self.output_dir.get():
                self.output_dir.set(os.path.join(folder, "PDF_Output"))

    def _browse_output(self):
        """选择输出文件夹"""
        folder = filedialog.askdirectory(title="选择 PDF 输出文件夹")
        if folder:
            self.output_dir.set(folder)

    def _start_conversion(self):
        """开始转换"""
        input_dir = self.input_dir.get().strip()
        output_dir = self.output_dir.get().strip()

        # 参数校验
        if not input_dir:
            messagebox.showwarning("提示", "请先选择 Excel 文件所在的文件夹")
            return

        if not os.path.isdir(input_dir):
            messagebox.showerror("错误", f"输入目录不存在:\n{input_dir}")
            return

        if not output_dir:
            output_dir = os.path.join(input_dir, "PDF_Output")
            self.output_dir.set(output_dir)

        # 确认开始
        if not messagebox.askyesno(
            "确认",
            f"即将开始转换:\n\n"
            f"📂 输入: {input_dir}\n"
            f"📂 输出: {output_dir}\n"
            f"👷 进程数: {self.worker_count.get()}\n\n"
            f"确定开始？"
        ):
            return

        # 更新 UI 状态
        self._close_after_stop = False
        self.start_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        self.progress_var.set(0)
        self.progress_label.configure(text="0%")
        self.status_text.set("转换中...")

        # 清空日志
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)

        # 在后台线程中运行转换
        self._conversion_thread = threading.Thread(
            target=self._run_conversion,
            args=(input_dir, output_dir),
            daemon=True
        )
        self._conversion_thread.start()

    def _run_conversion(self, input_dir, output_dir):
        """在后台线程中执行转换（不阻塞 GUI）"""
        try:
            self.manager = ConversionManager(
                num_workers=self.worker_count.get(),
                progress_callback=self._on_progress,
                log_callback=lambda msg: self.root.after(0, self._append_log, msg),
            )

            success, failed, skipped, results = self.manager.start_conversion(
                input_dir, output_dir
            )

            # 转换完成，更新 UI
            stopped = self.manager.was_stopped if self.manager else False
            self.root.after(0, self._on_conversion_complete, success, failed, skipped, stopped)

        except Exception as e:
            self.root.after(
                0, self._append_log,
                f"❌ 转换过程异常: {str(e)}", "error"
            )
            self.root.after(0, self._reset_ui)

    def _on_progress(self, current, total, result):
        """进度回调（在工作线程中调用）"""
        percent = (current / total) * 100 if total > 0 else 0
        self.root.after(0, self._update_progress, percent, current, total)

    def _update_progress(self, percent, current, total):
        """更新进度条（在主线程中执行）"""
        self.progress_var.set(percent)
        self.progress_label.configure(text=f"{percent:.0f}%")
        self.status_text.set(f"转换中... {current}/{total}")

    def _on_conversion_complete(self, success, failed, skipped, stopped=False):
        """转换完成后的回调"""
        self._reset_ui()

        total = success + failed + skipped
        self.progress_var.set(100)
        self.progress_label.configure(text="100%")
        self.status_text.set("已停止" if stopped else "转换完成!")

        if self._close_after_stop:
            self.root.destroy()
            return

        # 弹出结果统计
        if stopped:
            title = "转换已停止"
            body = (
                f"📊 已停止本次转换\n\n"
                f"总计: {total} 个文件\n"
                f"✅ 成功: {success}\n"
                f"❌ 失败: {failed}\n"
                f"⏭️ 未处理/跳过: {skipped}\n\n"
                f"📂 输出目录: {self.output_dir.get()}"
            )
        else:
            title = "转换完成"
            body = (
                f"📊 转换结果统计\n\n"
                f"总计: {total} 个文件\n"
                f"✅ 成功: {success}\n"
                f"❌ 失败: {failed}\n"
                f"⏭️ 跳过: {skipped}\n\n"
                f"📂 输出目录: {self.output_dir.get()}"
            )

        messagebox.showinfo(title, body)

    def _stop_conversion(self):
        """停止转换"""
        if self.manager and self.manager.is_running:
            self.manager.stop()
            self.stop_btn.configure(state=tk.DISABLED)
            self.status_text.set("正在停止，等待当前文件完成...")
            self._append_log("🛑 正在停止转换，请稍候...", "warning")

    def _open_output_dir(self):
        """打开输出目录"""
        output_dir = self.output_dir.get().strip()
        if output_dir and os.path.isdir(output_dir):
            os.startfile(output_dir)
        elif output_dir:
            messagebox.showinfo("提示", f"输出目录尚不存在:\n{output_dir}")
        else:
            messagebox.showinfo("提示", "请先设置输出目录")

    def _reset_ui(self):
        """重置 UI 到初始状态"""
        self.start_btn.configure(state=tk.NORMAL)
        self.stop_btn.configure(state=tk.DISABLED)

    def _append_log(self, message, tag="normal"):
        """追加日志消息"""
        # 自动检测颜色标签
        if "✅" in message or "成功" in message:
            tag = "success"
        elif "❌" in message or "失败" in message:
            tag = "error"
        elif "⏭️" in message or "跳过" in message or "⚠️" in message:
            tag = "warning"
        elif "🚀" in message or "📂" in message or "📋" in message or "📊" in message:
            tag = "info"

        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted = f"[{timestamp}] {message}\n"

        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, formatted, tag)
        self.log_text.see(tk.END)  # 自动滚动到底部
        self.log_text.configure(state=tk.DISABLED)

    def _on_closing(self):
        """窗口关闭事件"""
        if self.manager and self.manager.is_running:
            if messagebox.askyesno("确认", "转换正在进行中，确定要退出吗？"):
                self._close_after_stop = True
                self.manager.stop()
                self.status_text.set("正在停止，准备退出...")
                self.root.after(300, self._wait_for_shutdown)
            return
        self.root.destroy()

    def _wait_for_shutdown(self):
        """等待后台转换线程和工作进程自然退出后再关闭窗口。"""
        manager_running = self.manager and self.manager.is_running
        thread_running = self._conversion_thread and self._conversion_thread.is_alive()

        if manager_running or thread_running:
            self.root.after(300, self._wait_for_shutdown)
            return

        self.root.destroy()
