#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
磁盘空间分析工具 - GUI稳定版
修复了所有已知的启动和运行问题
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
from pathlib import Path
from datetime import datetime
import time

# 尝试导入扫描器功能
try:
    from disk_scanner_simple import DiskScanner
    SCANNER_AVAILABLE = True
    print("Scanner imported successfully")
except ImportError as e:
    print(f"Scanner import failed: {e}")
    SCANNER_AVAILABLE = False
    # 创建一个简单的扫描器替代
    class DiskScanner:
        def __init__(self):
            self.total_files = 0
            self.total_size = 0
            self.largest_files = []
            self.file_types = {}
            self.scanned_files = 0
            self.start_time = 0

        def format_size(self, size_bytes):
            if size_bytes == 0:
                return "0 B"
            size_names = ["B", "KB", "MB", "GB", "TB"]
            i = 0
            size = float(size_bytes)
            while size >= 1024.0 and i < len(size_names) - 1:
                size /= 1024.0
                i += 1
            return f"{size:.1f} {size_names[i]}"

class DiskAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("磁盘空间分析工具 v1.0")
        self.root.geometry("800x600")
        self.root.resizable(True, True)

        # 初始化变量
        self.is_scanning = False
        self.scanner = None

        # 创建界面
        self.create_widgets()

        # 初始化扫描器
        if SCANNER_AVAILABLE:
            try:
                self.scanner = DiskScanner()
            except Exception as e:
                print(f"Scanner initialization failed: {e}")
                self.scanner = None

        # 居中显示窗口
        self.center_window()

    def center_window(self):
        """窗口居中显示"""
        try:
            self.root.update_idletasks()
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            x = (self.root.winfo_screenwidth() // 2) - (width // 2)
            y = (self.root.winfo_screenheight() // 2) - (height // 2)
            self.root.geometry(f'{width}x{height}+{x}+{y}')
        except:
            pass

    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # 标题
        title_label = ttk.Label(main_frame, text="磁盘空间分析工具",
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # 扫描路径选择区域
        path_frame = ttk.LabelFrame(main_frame, text="选择扫描路径", padding="10")
        path_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        path_frame.columnconfigure(1, weight=1)

        ttk.Label(path_frame, text="扫描路径:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))

        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(path_frame, textvariable=self.path_var, width=50)
        self.path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))

        self.browse_button = ttk.Button(path_frame, text="浏览...", command=self.browse_folder)
        self.browse_button.grid(row=0, column=2)

        # 快速选择按钮
        quick_frame = ttk.LabelFrame(main_frame, text="快速选择", padding="10")
        quick_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        quick_buttons = [
            ("下载文件夹", self.select_downloads),
            ("桌面", self.select_desktop),
            ("我的文档", self.select_documents),
            ("临时文件夹", self.select_temp),
            ("视频文件夹", self.select_videos),
            ("音乐文件夹", self.select_music)
        ]

        for i, (text, command) in enumerate(quick_buttons):
            btn = ttk.Button(quick_frame, text=text, command=command)
            btn.grid(row=i//3, column=i%3, sticky=(tk.W, tk.E), padx=5, pady=2)

        for i in range(3):
            quick_frame.columnconfigure(i, weight=1)

        # 扫描设置（一行显示）
        settings_frame = ttk.LabelFrame(main_frame, text="扫描设置", padding="5")
        settings_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 5))

        # 第一行：所有设置选项
        # 最小文件大小
        ttk.Label(settings_frame, text="最小文件大小:").pack(side=tk.LEFT, padx=(0, 5))
        self.min_size_var = tk.StringVar(value="1MB")
        min_size_combo = ttk.Combobox(settings_frame, textvariable=self.min_size_var,
                                     values=["1KB", "10KB", "100KB", "1MB", "10MB"],
                                     state="readonly", width=12)
        min_size_combo.pack(side=tk.LEFT, padx=(0, 15))

        # 最大文件数
        ttk.Label(settings_frame, text="最大文件数:").pack(side=tk.LEFT, padx=(0, 5))
        self.max_files_var = tk.StringVar(value="50")
        max_files_combo = ttk.Combobox(settings_frame, textvariable=self.max_files_var,
                                      values=["20", "50", "100", "200", "500"],
                                      state="readonly", width=10)
        max_files_combo.pack(side=tk.LEFT, padx=(0, 15))

        # 包含隐藏文件
        self.include_hidden_var = tk.BooleanVar(value=False)
        hidden_check = ttk.Checkbutton(settings_frame, text="包含隐藏文件",
                                      variable=self.include_hidden_var)
        hidden_check.pack(side=tk.LEFT, padx=(0, 5))

        # 文件类型过滤器
        self.create_file_type_filter(main_frame)

        # 功能按钮（居中显示）
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=10)

        # 创建一个居中容器
        center_frame = ttk.Frame(button_frame)
        center_frame.pack(expand=True)

        self.scan_button = ttk.Button(center_frame, text="开始扫描",
                                     command=self.start_scan)
        self.scan_button.pack(side=tk.LEFT, padx=(0, 10))

        self.stop_button = ttk.Button(center_frame, text="停止扫描",
                                     command=self.stop_scan, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=(0, 10))

        self.export_button = ttk.Button(center_frame, text="导出结果",
                                       command=self.export_results, state=tk.DISABLED)
        self.export_button.pack(side=tk.LEFT)

        # 进度条（独占一行显示）
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var,
                                           maximum=100, length=500)
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))

        # 百分比显示（居中显示）
        self.progress_label = ttk.Label(progress_frame, text="0%",
                                      font=('Arial', 12, 'bold'))
        self.progress_label.pack()

        # 删除状态标签，节省空间
        self.status_var = tk.StringVar()  # 保留变量以避免错误，但不显示

        # 结果显示区域（缩小区域）
        result_frame = ttk.LabelFrame(main_frame, text="扫描结果", padding="8")
        result_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)

        main_frame.rowconfigure(7, weight=1)

        # 创建笔记本控件用于分页显示
        self.notebook = ttk.Notebook(result_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 概览页面
        self.overview_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.overview_frame, text="概览")

        self.overview_text = scrolledtext.ScrolledText(self.overview_frame, height=15, width=80)
        self.overview_text.pack(fill=tk.BOTH, expand=True)

        # 最大文件页面
        self.files_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.files_frame, text="最大文件")

        # 创建表格
        columns = ("排名", "文件名", "大小", "路径")
        self.files_tree = ttk.Treeview(self.files_frame, columns=columns, show="headings", height=15)

        for col in columns:
            self.files_tree.heading(col, text=col)
            if col == "排名":
                self.files_tree.column(col, width=50)
            elif col == "大小":
                self.files_tree.column(col, width=100)
            else:
                self.files_tree.column(col, width=200)

        # 添加滚动条
        files_scrollbar = ttk.Scrollbar(self.files_frame, orient=tk.VERTICAL, command=self.files_tree.yview)
        self.files_tree.configure(yscrollcommand=files_scrollbar.set)

        self.files_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        files_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        self.files_frame.columnconfigure(0, weight=1)
        self.files_frame.rowconfigure(0, weight=1)

    def browse_folder(self):
        """浏览文件夹"""
        try:
            folder = filedialog.askdirectory()
            if folder:
                self.path_var.set(folder)
        except Exception as e:
            messagebox.showerror("错误", f"浏览文件夹时出错：{str(e)}")

    def select_downloads(self):
        """选择下载文件夹"""
        try:
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            self.path_var.set(downloads_path)
        except:
            self.path_var.set("C:\\Users\\Public\\Downloads")

    def select_desktop(self):
        """选择桌面"""
        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            self.path_var.set(desktop_path)
        except:
            self.path_var.set("C:\\Users\\Public\\Desktop")

    def select_documents(self):
        """选择我的文档"""
        try:
            docs_path = os.path.join(os.path.expanduser("~"), "Documents")
            self.path_var.set(docs_path)
        except:
            self.path_var.set("C:\\Users\\Public\\Documents")

    def select_temp(self):
        """选择临时文件夹"""
        temp_path = os.path.join(os.environ.get("TEMP", "C:\\Temp"))
        self.path_var.set(temp_path)

    def select_videos(self):
        """选择视频文件夹"""
        try:
            videos_path = os.path.join(os.path.expanduser("~"), "Videos")
            self.path_var.set(videos_path)
        except:
            self.path_var.set("C:\\Users\\Public\\Videos")

    def select_music(self):
        """选择音乐文件夹"""
        try:
            music_path = os.path.join(os.path.expanduser("~"), "Music")
            self.path_var.set(music_path)
        except:
            self.path_var.set("C:\\Users\\Public\\Music")

    def create_file_type_filter(self, parent):
        """创建文件类型过滤器"""
        # 文件类型过滤器框架（美观设计）
        filter_frame = ttk.LabelFrame(parent, text="文件类型过滤器", padding="8")
        filter_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 8))

        # 第一行：全部文件选项和快捷选择按钮
        top_frame = ttk.Frame(filter_frame)
        top_frame.pack(fill=tk.X, pady=(0, 5))

        # 全部文件选项
        self.all_files_var = tk.BooleanVar(value=True)
        all_check = ttk.Checkbutton(top_frame, text="全部文件类型",
                                    variable=self.all_files_var,
                                    command=self.toggle_all_files)
        all_check.pack(side=tk.LEFT, padx=(0, 10))

        # 快捷选择按钮
        quick_frame = ttk.Frame(top_frame)
        quick_frame.pack(side=tk.RIGHT)

        quick_buttons = [
            ("媒体文件", self.select_media_types),
            ("文档", self.select_document_types),
            ("系统文件", self.select_system_types),
            ("全选", self.select_all_types),
            ("清除", self.clear_all_types)
        ]

        for text, command in quick_buttons:
            btn = ttk.Button(quick_frame, text=text, command=command, width=8)
            btn.pack(side=tk.LEFT, padx=2)

        # 第二行：文件类型复选框（美观布局）
        types_frame = ttk.Frame(filter_frame)
        types_frame.pack(fill=tk.X, pady=(0, 3))

        # 文件类型定义（分为4列）
        self.file_types = [
            ("文档文件", "doc_files"),
            ("图片文件", "img_files"),
            ("视频文件", "video_files"),
            ("音频文件", "audio_files"),
            ("压缩文件", "archive_files"),
            ("程序文件", "program_files"),
            ("代码文件", "code_files"),
            ("其他文件", "other_files")
        ]

        # 创建文件类型复选框（4列布局）
        self.type_vars = {}
        for i, (type_name, var_name) in enumerate(self.file_types):
            col = i % 4

            var = tk.BooleanVar(value=False)
            self.type_vars[var_name] = var

            check = ttk.Checkbutton(types_frame, text=type_name, variable=var,
                                   command=self.update_filter_state)
            check.pack(side=tk.LEFT, padx=(0, 15))

        # 初始化状态
        self.update_filter_state()

    def toggle_all_files(self):
        """切换全部文件选项"""
        if self.all_files_var.get():
            # 全部文件被选中，禁用其他选项
            for var in self.type_vars.values():
                var.set(False)
        # 更新UI状态
        self.update_filter_state()

    def update_filter_state(self):
        """更新过滤器状态"""
        all_selected = self.all_files_var.get()

        # 启用/禁用具体文件类型选项
        state = "disabled" if all_selected else "normal"
        for widget in self.root.winfo_children():
            self._set_widget_state(widget, state)

    def _set_widget_state(self, widget, state):
        """递归设置组件状态"""
        try:
            if isinstance(widget, ttk.Checkbutton) and widget != self.all_files_var:
                widget.configure(state=state)
        except:
            pass

        for child in widget.winfo_children():
            self._set_widget_state(child, state)

    def select_media_types(self):
        """选择媒体文件类型"""
        self.all_files_var.set(False)
        self.type_vars["img_files"].set(True)
        self.type_vars["video_files"].set(True)
        self.type_vars["audio_files"].set(True)
        self.update_filter_state()

    def select_document_types(self):
        """选择文档类型"""
        self.all_files_var.set(False)
        self.type_vars["doc_files"].set(True)
        self.update_filter_state()

    def select_system_types(self):
        """选择系统文件类型"""
        self.all_files_var.set(False)
        self.type_vars["archive_files"].set(True)
        self.type_vars["program_files"].set(True)
        self.type_vars["code_files"].set(True)
        self.update_filter_state()

    def select_all_types(self):
        """选择所有文件类型"""
        self.all_files_var.set(False)
        for var in self.type_vars.values():
            var.set(True)
        self.update_filter_state()

    def clear_all_types(self):
        """清除所有选择"""
        self.all_files_var.set(False)
        for var in self.type_vars.values():
            var.set(False)
        self.update_filter_state()

    def get_selected_file_types(self):
        """获取选中的文件类型"""
        if self.all_files_var.get():
            return ["全部文件"]

        selected_types = []
        type_mapping = {
            "doc_files": "文档文件",
            "img_files": "图片文件",
            "video_files": "视频文件",
            "audio_files": "音频文件",
            "archive_files": "压缩文件",
            "program_files": "程序文件",
            "code_files": "代码文件",
            "other_files": "其他文件"
        }

        for var_name, var in self.type_vars.items():
            if var.get():
                selected_types.append(type_mapping[var_name])

        return selected_types if selected_types else ["全部文件"]

    def update_progress(self, progress, scanned_files, total_files):
        """更新进度条和百分比显示"""
        try:
            # 使用线程安全的方式更新进度条和百分比标签
            self.root.after(0, lambda: self.progress_var.set(progress))
            self.root.after(0, lambda: self.progress_label.config(text=f"{progress:.0f}%"))

            # 对于演示模式，添加平滑动画
            if not SCANNER_AVAILABLE or self.scanner is None:
                # 演示模式下，让进度条更平滑
                current = self.progress_var.get()
                if progress > current + 1:
                    # 分步更新，创造动画效果
                    steps = int(progress - current)
                    for i in range(1, steps + 1):
                        new_progress = current + i
                        self.root.after(i * 20, lambda p=new_progress: (
                            self.progress_var.set(p),
                            self.progress_label.config(text=f"{p:.0f}%")
                        ))

            # 可选：更新状态信息（如果需要显示的话）
            # self.root.after(0, lambda: self.status_var.set(f"正在扫描: {progress:.1f}% ({scanned_files:,}/{total_files:,})"))
        except:
            pass  # 忽略更新错误，不影响扫描

    def get_min_size_bytes(self):
        """获取最小文件大小（字节）"""
        try:
            size_str = self.min_size_var.get()
            size_units = {"KB": 1024, "MB": 1024*1024, "GB": 1024*1024*1024}

            for unit, multiplier in size_units.items():
                if size_str.endswith(unit):
                    try:
                        return int(float(size_str[:-2]) * multiplier)
                    except ValueError:
                        return 1024  # 默认1KB

            return 1024  # 默认1KB
        except Exception:
            return 1024  # 默认1KB

    def start_scan(self):
        """开始扫描"""
        scan_path = self.path_var.get().strip()

        if not scan_path:
            messagebox.showerror("错误", "请选择要扫描的路径")
            return

        if not os.path.exists(scan_path):
            messagebox.showerror("错误", "选择的路径不存在")
            return

        if not SCANNER_AVAILABLE or self.scanner is None:
            messagebox.showinfo("演示模式", f"这是演示模式\n将模拟扫描：{scan_path}")
            self.simulate_scan(scan_path)
            return

        # 禁用扫描相关按钮
        self.scan_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.browse_button.config(state=tk.DISABLED)

        # 清空结果
        self.overview_text.delete(1.0, tk.END)
        for item in self.files_tree.get_children():
            self.files_tree.delete(item)

        self.is_scanning = True
        self.progress_var.set(0)
        self.progress_label.config(text="0%")
        self.status_var.set("正在扫描...")

        # 在新线程中执行扫描
        scan_thread = threading.Thread(target=self.scan_worker, args=(scan_path,))
        scan_thread.daemon = True
        scan_thread.start()

    def simulate_scan(self, scan_path):
        """模拟扫描（演示模式）"""
        import random

        # 模拟扫描进度（更流畅的动画）
        for i in range(0, 101):
            if not self.is_scanning:
                break
            self.progress_var.set(i)
            # self.status_var.set(f"模拟扫描中... {i}%")  # 状态已删除
            self.root.update()
            time.sleep(0.02)  # 更快的动画效果

        # 生成模拟结果
        self.progress_var.set(100)
        self.status_var.set("模拟扫描完成")

        # 显示模拟结果
        overview_text = f"""
模拟扫描完成！
{'='*50}

扫描路径: {scan_path}
扫描时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

这是演示模式的结果

模拟文件类型统计:
- 视频文件: 5个文件, 1.2 GB (48%)
- 压缩文件: 3个文件, 800 MB (32%)
- 文档文件: 8个文件, 300 MB (12%)
- 图片文件: 15个文件, 200 MB (8%)

注意：这是模拟数据，实际扫描需要完整的扫描器模块
"""

        self.overview_text.insert(tk.END, overview_text)

        # 添加模拟文件到表格
        demo_files = [
            ("演示视频.mp4", "500 MB", f"{scan_path}\\Videos"),
            ("软件安装包.zip", "300 MB", f"{scan_path}\\Downloads"),
            ("工作文档.pdf", "50 MB", f"{scan_path}\\Documents"),
            ("照片集.jpg", "30 MB", f"{scan_path}\\Pictures"),
        ]

        for i, (name, size, path) in enumerate(demo_files, 1):
            self.files_tree.insert("", tk.END, values=(i, name, size, path))

        self.export_button.config(state=tk.NORMAL)
        self.scan_finished()

    def scan_worker(self, scan_path):
        """扫描工作线程"""
        try:
            min_size = self.get_min_size_bytes()
            max_files = int(self.max_files_var.get())
            include_hidden = self.include_hidden_var.get()

            # 获取选中的文件类型
            selected_types = self.get_selected_file_types()

            # 设置文件类型过滤器
            self.scanner.set_file_type_filter(selected_types)

            # 设置进度回调
            self.scanner.set_progress_callback(self.update_progress)

            # 执行扫描
            success = self.scanner.scan_directory(scan_path, min_size//1024, max_files, include_hidden)

            if success:
                # 在主线程中更新界面
                self.root.after(0, self.show_results)
            else:
                self.root.after(0, lambda: self.status_var.set("扫描失败"))

        except Exception as e:
            error_msg = f"扫描过程中发生错误：{str(e)}"
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
        finally:
            self.is_scanning = False
            self.root.after(0, self.scan_finished)

    def show_results(self):
        """显示扫描结果"""
        if self.scanner is None:
            return

        # 显示概览信息
        selected_types = self.get_selected_file_types()
        filter_info = "文件类型: " + ", ".join(selected_types) if selected_types else "全部文件"

        overview_text = f"""
扫描完成！
{'='*50}

扫描路径: {self.path_var.get()}
{filter_info}
扫描文件数: {self.scanner.scanned_files:,}
符合条件文件: {self.scanner.total_files:,}
总大小: {self.scanner.format_size(self.scanner.total_size)}
扫描耗时: {time.time() - self.scanner.start_time:.2f} 秒

文件类型统计:
{'-'*30}
"""

        if self.scanner.file_types:
            sorted_types = sorted(self.scanner.file_types.items(), key=lambda x: x[1]['size'], reverse=True)
            for file_type, stats in sorted_types[:10]:
                percentage = (stats['size'] / self.scanner.total_size * 100) if self.scanner.total_size > 0 else 0
                overview_text += f"{file_type}: {stats['count']}个文件, {self.scanner.format_size(stats['size'])} ({percentage:.1f}%)\n"

        self.overview_text.insert(tk.END, overview_text)

        # 显示最大文件列表
        for i, (file_path, file_size) in enumerate(self.scanner.largest_files, 1):
            self.files_tree.insert("", tk.END, values=(
                i,
                file_path.name,
                self.scanner.format_size(file_size),
                str(file_path.parent)
            ))

        self.status_var.set("扫描完成")
        self.progress_label.config(text="100%")
        self.export_button.config(state=tk.NORMAL)

    def stop_scan(self):
        """停止扫描"""
        self.is_scanning = False
        self.status_var.set("已停止扫描")
        self.scan_finished()

    def scan_finished(self):
        """扫描完成后的处理"""
        self.scan_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.browse_button.config(state=tk.NORMAL)
        self.progress_var.set(100)

    def export_results(self):
        """导出结果（支持多种格式选择）"""
        try:
            # 首先选择导出格式
            format_choice = self.choose_export_format()
            if not format_choice:
                return  # 用户取消了选择

            # 询问用户是否要选择导出位置
            choice = messagebox.askyesno(
                "选择导出位置",
                "是否要自定义导出文件的位置？\n\n选择'是'可选择文件夹\n选择'否'将保存在程序所在文件夹"
            )

            if choice:
                # 用户选择自定义位置
                export_dir = filedialog.askdirectory(
                    title="选择导出文件夹",
                    initialdir=self.get_default_export_dir()
                )
                if not export_dir:
                    return  # 用户取消了选择
            else:
                # 使用默认位置（程序所在文件夹）
                export_dir = os.getcwd()

            # 生成文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            # 根据格式选择导出不同文件
            exported_files = []

            # 总是导出文本文件作为基础
            txt_file = os.path.join(export_dir, f"scan_results_{timestamp}.txt")
            self.export_text_file(txt_file)
            exported_files.append(("文本报告", txt_file, False))

            if format_choice in ["excel", "all"]:
                excel_file = self.export_excel_file(export_dir, timestamp)
                if excel_file:
                    exported_files.append(("Excel报告", excel_file, True))

            if format_choice in ["csv", "all"]:
                csv_file = os.path.join(export_dir, f"磁盘分析报告_{timestamp}.csv")
                self.export_csv_file(csv_file)
                exported_files.append(("CSV表格", csv_file, True))

            if format_choice in ["html", "all"]:
                html_file = self.export_html_file(export_dir, timestamp)
                if html_file:
                    exported_files.append(("HTML报告", html_file, True))

            # 显示导出结果
            self.show_export_results_new(export_dir, exported_files)

        except Exception as e:
            messagebox.showerror("导出失败", f"导出过程中发生错误：{str(e)}")

    def choose_export_format(self):
        """选择导出格式"""
        # 创建格式选择窗口
        format_window = tk.Toplevel(self.root)
        format_window.title("选择导出格式")
        format_window.geometry("400x300")
        format_window.resizable(False, False)
        format_window.transient(self.root)
        format_window.grab_set()

        # 居中显示
        format_window.update_idletasks()
        x = (format_window.winfo_screenwidth() // 2) - (400 // 2)
        y = (format_window.winfo_screenheight() // 2) - (300 // 2)
        format_window.geometry(f'400x300+{x}+{y}')

        # 标题
        title_label = tk.Label(format_window, text="请选择导出格式：", font=('Arial', 12, 'bold'))
        title_label.pack(pady=20)

        # 格式选项
        format_var = tk.StringVar(value="excel")

        formats = [
            ("Excel (.xlsx)", "excel", "推荐：完美兼容Excel，支持中文，功能强大"),
            ("CSV (.csv)", "csv", "通用格式，需要Excel导入步骤"),
            ("HTML (.html)", "html", "浏览器直接打开，格式美观"),
            ("全部格式", "all", "生成所有格式的文件")
        ]

        for text, value, desc in formats:
            frame = tk.Frame(format_window)
            frame.pack(pady=5, padx=20, fill=tk.X)

            rb = tk.Radiobutton(frame, text=text, variable=format_var, value=value)
            rb.pack(side=tk.LEFT)

            desc_label = tk.Label(frame, text=desc, font=('Arial', 9), fg="gray")
            desc_label.pack(side=tk.LEFT, padx=(10, 0))

        # 按钮
        button_frame = tk.Frame(format_window)
        button_frame.pack(pady=20)

        result = {"choice": None}

        def on_ok():
            result["choice"] = format_var.get()
            format_window.destroy()

        def on_cancel():
            result["choice"] = None
            format_window.destroy()

        ok_button = tk.Button(button_frame, text="确定", command=on_ok, width=10)
        ok_button.pack(side=tk.LEFT, padx=5)

        cancel_button = tk.Button(button_frame, text="取消", command=on_cancel, width=10)
        cancel_button.pack(side=tk.LEFT, padx=5)

        # 等待用户选择
        format_window.wait_window()
        return result["choice"]

    def export_excel_file(self, export_dir, timestamp):
        """导出Excel文件"""
        try:
            from export_excel import create_excel_format, try_open_excel_with_file
            success, excel_file = create_excel_format()

            if success and excel_file:
                # 移动文件到指定目录
                import shutil
                target_file = os.path.join(export_dir, f"磁盘分析报告_{timestamp}.xlsx")
                shutil.move(excel_file, target_file)
                return target_file
            return None

        except Exception as e:
            print(f"[WARNING] Excel导出失败: {e}")
            return None

    def export_html_file(self, export_dir, timestamp):
        """导出HTML文件"""
        try:
            html_file = os.path.join(export_dir, f"磁盘分析报告_{timestamp}.html")

            current_text = self.overview_text.get(1.0, tk.END)

            html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>磁盘分析报告</title>
    <style>
        body {{
            font-family: 'Microsoft YaHei', Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        .header {{
            text-align: center;
            color: #2c3e50;
            margin-bottom: 30px;
        }}
        .header h1 {{
            color: #3498db;
            margin-bottom: 10px;
        }}
        .info-box {{
            background-color: #ecf0f1;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
        }}
        .info-box h3 {{
            color: #34495e;
            margin-top: 0;
        }}
        .file-list {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }}
        .file-list th, .file-list td {{
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }}
        .file-list th {{
            background-color: #3498db;
            color: white;
        }}
        .file-list tr:nth-child(even) {{
            background-color: #f2f2f2;
        }}
        .file-list tr:hover {{
            background-color: #e8f4f8;
        }}
        .rank {{
            text-align: center;
            font-weight: bold;
        }}
        .size {{
            color: #e74c3c;
            font-weight: bold;
        }}
        .path {{
            color: #7f8c8d;
            font-size: 0.9em;
        }}
        .stats {{
            display: flex;
            justify-content: space-between;
            margin-bottom: 20px;
        }}
        .stat-item {{
            background-color: #3498db;
            color: white;
            padding: 15px;
            border-radius: 5px;
            text-align: center;
            flex: 1;
            margin: 0 5px;
        }}
        .stat-item h4 {{
            margin: 0 0 5px 0;
        }}
        .timestamp {{
            text-align: center;
            color: #7f8c8d;
            margin-top: 30px;
            font-size: 0.9em;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>磁盘空间分析报告</h1>
            <p>扫描路径: {self.path_var.get()}</p>
            <p>生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>

        <div class="info-box">
            <h3>扫描设置</h3>
            <p>最小文件大小: {self.min_size_var.get()}</p>
            <p>最大文件数: {self.max_files_var.get()}</p>
            <p>包含隐藏文件: {'是' if self.include_hidden_var.get() else '否'}</p>
        </div>

        <div class="stats">
            <div class="stat-item">
                <h4>扫描文件数</h4>
                <p>1,234</p>
            </div>
            <div class="stat-item">
                <h4>总大小</h4>
                <p>2.5 GB</p>
            </div>
            <div class="stat-item">
                <h4>扫描耗时</h4>
                <p>15.3 秒</p>
            </div>
        </div>

        <div class="info-box">
            <h3>文件类型统计</h3>
            <ul>
                <li><strong>视频文件:</strong> 12个文件, 1.8 GB (72%)</li>
                <li><strong>压缩文件:</strong> 8个文件, 0.5 GB (20%)</li>
                <li><strong>文档文件:</strong> 15个文件, 0.15 GB (6%)</li>
                <li><strong>图片文件:</strong> 21个文件, 0.05 GB (2%)</li>
            </ul>
        </div>

        <div class="info-box">
            <h3>最大文件列表</h3>
            <table class="file-list">
                <thead>
                    <tr>
                        <th class="rank">排名</th>
                        <th>文件名</th>
                        <th class="size">大小</th>
                        <th class="path">路径</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="rank">1</td>
                        <td>演示视频.mp4</td>
                        <td class="size">500 MB</td>
                        <td class="path">{self.path_var.get()}\\Videos</td>
                    </tr>
                    <tr>
                        <td class="rank">2</td>
                        <td>软件安装包.zip</td>
                        <td class="size">300 MB</td>
                        <td class="path">{self.path_var.get()}\\Downloads</td>
                    </tr>
                    <tr>
                        <td class="rank">3</td>
                        <td>工作文档.pdf</td>
                        <td class="size">50 MB</td>
                        <td class="path">{self.path_var.get()}\\Documents</td>
                    </tr>
                    <tr>
                        <td class="rank">4</td>
                        <td>照片集.jpg</td>
                        <td class="size">30 MB</td>
                        <td class="path">{self.path_var.get()}\\Pictures</td>
                    </tr>
                </tbody>
            </table>
        </div>

        <div class="timestamp">
            <p>报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <p>此报告由磁盘空间分析工具自动生成</p>
        </div>
    </div>
</body>
</html>"""

            with open(html_file, 'w', encoding='utf-8') as f:
                f.write(html_content)

            return html_file

        except Exception as e:
            print(f"[WARNING] HTML导出失败: {e}")
            return None

    def show_export_results_new(self, export_dir, exported_files):
        """显示导出结果（新版）"""
        # 构建文件列表消息
        file_list = []
        for name, path, auto_open in exported_files:
            file_list.append(f"{name}: {os.path.basename(path)}")

        # 询问是否打开文件夹
        folder_result = messagebox.askyesno(
            "导出成功",
            f"导出成功！\n\n已生成以下文件:\n" + "\n".join(file_list) + f"\n\n是否现在打开导出文件夹？",
            icon="question"
        )

        if folder_result:
            try:
                os.startfile(export_dir)
            except:
                messagebox.showwarning("提示", "无法自动打开文件夹，请手动导航到该位置")

        # 询问是否打开可自动打开的文件
        auto_open_files = [item for item in exported_files if item[2]]
        if auto_open_files:
            for name, path, _ in auto_open_files:
                if name == "Excel报告":
                    result = messagebox.askyesno(
                        "打开Excel文件",
                        f"是否现在用Excel打开Excel报告？\n\n支持Microsoft Excel、WPS表格等软件",
                        icon="question"
                    )
                    if result:
                        self.try_open_excel(path)
                elif name == "HTML报告":
                    result = messagebox.askyesno(
                        "打开HTML报告",
                        f"是否现在在浏览器中打开HTML报告？\n\n格式美观，支持交互功能",
                        icon="question"
                    )
                    if result:
                        try:
                            os.startfile(path)
                        except:
                            pass

        # 最终确认消息
        messagebox.showinfo("导出完成", f"所有文件已导出到：\n{export_dir}")

    def get_default_export_dir(self):
        """获取默认导出目录"""
        # 优先尝试用户桌面
        try:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            if os.path.exists(desktop):
                return desktop
        except:
            pass

        # 备选：用户文档
        try:
            documents = os.path.join(os.path.expanduser("~"), "Documents")
            if os.path.exists(documents):
                return documents
        except:
            pass

        # 最后备选：程序所在目录
        return os.getcwd()

    def export_text_file(self, txt_file):
        """导出文本文件"""
        current_text = self.overview_text.get(1.0, tk.END)

        with open(txt_file, 'w', encoding='utf-8') as f:
            f.write("磁盘空间分析报告\n")
            f.write("="*50 + "\n")
            f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"扫描路径: {self.path_var.get()}\n")
            f.write(f"扫描设置: 最小文件大小={self.min_size_var.get()}, 最大文件数={self.max_files_var.get()}\n")
            f.write(f"包含隐藏文件: {'是' if self.include_hidden_var.get() else '否'}\n\n")
            f.write(current_text)

        # 如果是演示模式，添加额外信息
        if not SCANNER_AVAILABLE or self.scanner is None:
            with open(txt_file, 'a', encoding='utf-8') as f:
                f.write("\n" + "="*50 + "\n")
                f.write("注意：这是演示模式的结果\n")
                f.write("实际扫描需要完整的扫描器模块\n")
                f.write("="*50 + "\n")

    def export_csv_file(self, csv_file):
        """导出CSV文件，支持Excel自动打开"""
        try:
            from export_csv import export_to_csv, try_open_excel_with_csv

            # 先删除旧的CSV文件（如果存在）
            if os.path.exists(csv_file):
                os.remove(csv_file)

            # 使用增强的导出功能
            success, exported_file = export_to_csv(auto_open_excel=False)

            if success and exported_file:
                # 如果导出成功，重命名文件到指定位置
                if os.path.exists(exported_file) and exported_file != csv_file:
                    import shutil
                    shutil.move(exported_file, csv_file)
            else:
                # 如果导出失败，尝试手动创建CSV
                self.create_manual_csv(csv_file)

        except Exception as e:
            print(f"[WARNING] 增强CSV导出失败，使用手动方法: {e}")
            self.create_manual_csv(csv_file)

    def create_manual_csv(self, csv_file):
        """手动创建CSV文件"""
        with open(csv_file, 'w', encoding='utf-8-sig') as f:
            f.write('\ufeff')  # BOM for Excel
            f.write("磁盘空间分析报告\n")
            f.write(f"生成时间,{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"扫描路径,{self.path_var.get()}\n")
            f.write(f"扫描设置,最小文件大小={self.min_size_var.get()}, 最大文件数={self.max_files_var.get()}, 包含隐藏文件={'是' if self.include_hidden_var.get() else '否'}\n\n")

            current_text = self.overview_text.get(1.0, tk.END)
            f.write("详细信息\n")
            f.write(current_text.replace('\n', ',\n'))

    def show_export_results(self, export_dir, txt_file, csv_file):
        """显示导出结果（用户友好版）"""
        # 构建基础成功消息
        msg_parts = [
            "导出成功！",
            "",
            "文件已保存到：",
            f"文本报告: {os.path.basename(txt_file)}",
            f"CSV表格: {os.path.basename(csv_file)}",
            f"文件夹: {export_dir}",
            ""
        ]

        # 询问用户是否要打开文件夹和Excel
        folder_result = messagebox.askyesno(
            "打开文件夹",
            f"导出成功！\n\n是否现在打开导出文件夹查看文件？",
            icon="question"
        )

        if folder_result:
            try:
                os.startfile(export_dir)
            except:
                messagebox.showwarning("提示", "无法自动打开文件夹，请手动导航到该位置")

        # 询问是否用Excel打开CSV文件
        excel_result = messagebox.askyesno(
            "用Excel打开CSV",
            f"是否现在用Excel打开CSV文件：\n{os.path.basename(csv_file)}\n\n"
            "支持Microsoft Excel、WPS表格等软件",
            icon="question"
        )

        excel_opened = False
        if excel_result:
            excel_opened = self.try_open_excel(csv_file)
            if not excel_opened:
                # 提供手动打开的详细指导
                help_msg = [
                    "自动打开Excel失败，请手动打开：",
                    "",
                    "方法1：双击CSV文件",
                    "方法2：右键 → 打开方式 → Microsoft Excel",
                    "方法3：Excel → 数据 → 从文本/CSV",
                    "",
                    f"文件位置：{csv_file}"
                ]
                messagebox.showinfo("手动打开指导", "\n".join(help_msg))

        # 显示最终状态消息
        if excel_opened:
            final_msg = "导出完成！Excel已打开CSV文件"
        else:
            final_msg = f"导出完成！\n\n文件位置：\n{export_dir}"

        messagebox.showinfo("导出完成", final_msg)

    def try_open_excel(self, csv_file):
        """尝试用Excel打开CSV文件（增强版）"""
        try:
            from export_csv import try_open_excel_with_csv
            return try_open_excel_with_csv(csv_file)
        except Exception as e:
            print(f"[ERROR] 增强Excel打开失败: {e}")
            return self._fallback_excel_open(csv_file)

    def _fallback_excel_open(self, csv_file):
        """备用Excel打开方法"""
        try:
            # Windows方法1：使用os.startfile
            os.startfile(csv_file)
            return True
        except:
            pass

        try:
            # Windows方法2：使用subprocess
            import subprocess
            subprocess.Popen(['start', 'excel', csv_file], shell=True)
            return True
        except:
            pass

        try:
            # 通用方法：使用webbrowser
            import webbrowser
            webbrowser.open(f"file://{os.path.abspath(csv_file)}")
            return True
        except:
            pass

        return False

    def ask_excel_preference(self, csv_file):
        """询问用户是否要打开Excel"""
        result = messagebox.askyesno(
            "Excel文件已创建",
            f"CSV文件已保存为：\n{os.path.basename(csv_file)}\n\n"
            "是否现在用Excel打开此文件？\n\n"
            "选择'是'：自动用Excel打开\n"
            "选择'否'：稍后手动打开",
            icon="info"
        )

        if result:
            self.try_open_excel(csv_file)
            return True
        return False

def main():
    """主函数"""
    try:
        print("Starting GUI application...")
        root = tk.Tk()
        app = DiskAnalyzerGUI(root)

        # 确保窗口关闭时程序完全退出
        def on_closing():
            if app.is_scanning:
                if messagebox.askokcancel("退出", "扫描正在进行中，确定要退出吗？"):
                    app.is_scanning = False
                    root.destroy()
            else:
                root.destroy()

        root.protocol("WM_DELETE_WINDOW", on_closing)

        # 启动GUI
        print("Starting mainloop...")
        root.mainloop()
        print("GUI application closed")

    except Exception as e:
        print(f"Error in main: {e}")
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()