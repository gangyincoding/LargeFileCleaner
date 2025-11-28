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
import logging

# 尝试导入文件安全检查模块
try:
    from file_safety import FileSafetyChecker
    SAFETY_CHECKER_AVAILABLE = True
except ImportError:
    SAFETY_CHECKER_AVAILABLE = False
    print("[WARNING] File safety checker module not available")

# 尝试导入重复文件检测模块
try:
    from duplicate_finder import DuplicateFinder
    DUPLICATE_FINDER_AVAILABLE = True
    print("Duplicate finder imported successfully")
except ImportError as e:
    print(f"Duplicate finder import failed: {e}")
    DUPLICATE_FINDER_AVAILABLE = False

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
        self.root.title("磁盘空间分析工具 v1.4")
        self.root.geometry("800x600")
        self.root.resizable(True, True)

        # 初始化变量
        self.is_scanning = False
        self.scanner = None
        self.selected_files = set()  # 存储选中的文件路径
        self.selected_folders = set()  # 存储选中的文件夹路径
        self.folder_stats = {}  # 存储文件夹统计信息

        # 初始化文件安全检查器
        if SAFETY_CHECKER_AVAILABLE:
            try:
                self.safety_checker = FileSafetyChecker()
            except Exception as e:
                print(f"Safety checker initialization failed: {e}")
                self.safety_checker = None
        else:
            self.safety_checker = None

        # 初始化重复文件检测器
        if DUPLICATE_FINDER_AVAILABLE:
            try:
                self.duplicate_finder = DuplicateFinder()
            except Exception as e:
                print(f"Duplicate finder initialization failed: {e}")
                self.duplicate_finder = None
        else:
            self.duplicate_finder = None

        # 重复文件检测相关变量
        self.duplicate_groups = {}  # 存储重复文件组
        self.selected_duplicates = set()  # 存储选中要删除的重复文件

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

        # 创建顶部工具栏
        toolbar_frame = ttk.Frame(self.files_frame)
        toolbar_frame.pack(fill=tk.X, padx=5, pady=5)

        # 全选/取消全选按钮
        self.select_all_button = ttk.Button(toolbar_frame, text="全选",
                                           command=self.select_all_files, state=tk.DISABLED)
        self.select_all_button.pack(side=tk.LEFT, padx=(0, 5))

        self.deselect_all_button = ttk.Button(toolbar_frame, text="取消全选",
                                             command=self.deselect_all_files, state=tk.DISABLED)
        self.deselect_all_button.pack(side=tk.LEFT, padx=(0, 10))

        # 删除按钮（醒目的红色风格）
        self.delete_button = ttk.Button(toolbar_frame, text="删除选中文件 (慎重!)",
                                       command=self.delete_selected_files, state=tk.DISABLED)
        self.delete_button.pack(side=tk.LEFT, padx=(0, 10))

        # 状态标签
        self.file_selection_label = ttk.Label(toolbar_frame, text="未选择文件")
        self.file_selection_label.pack(side=tk.LEFT, padx=(10, 0))

        # 创建表格
        columns = ("选择", "排名", "文件名", "大小", "路径", "安全状态")
        self.files_tree = ttk.Treeview(self.files_frame, columns=columns, show="headings", height=15)

        # 设置列标题和宽度
        self.files_tree.heading("选择", text="☐")
        self.files_tree.column("选择", width=40, anchor=tk.CENTER)

        self.files_tree.heading("排名", text="排名")
        self.files_tree.column("排名", width=50, anchor=tk.CENTER)

        self.files_tree.heading("文件名", text="文件名")
        self.files_tree.column("文件名", width=200)

        self.files_tree.heading("大小", text="大小")
        self.files_tree.column("大小", width=100)

        self.files_tree.heading("路径", text="路径")
        self.files_tree.column("路径", width=250)

        self.files_tree.heading("安全状态", text="安全状态")
        self.files_tree.column("安全状态", width=100)

        # 绑定点击事件
        self.files_tree.bind("<Button-1>", self.on_tree_click)
        self.files_tree.bind("<Double-Button-1>", self.on_tree_double_click)

        # 添加滚动条
        files_scrollbar = ttk.Scrollbar(self.files_frame, orient=tk.VERTICAL, command=self.files_tree.yview)
        self.files_tree.configure(yscrollcommand=files_scrollbar.set)

        self.files_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        files_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # === 按文件夹查看页面 ===
        self.folders_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.folders_frame, text="按文件夹查看")

        # 创建顶部工具栏
        folder_toolbar_frame = ttk.Frame(self.folders_frame)
        folder_toolbar_frame.pack(fill=tk.X, padx=5, pady=5)

        # 全选/取消全选按钮
        self.folder_select_all_button = ttk.Button(folder_toolbar_frame, text="全选",
                                                   command=self.select_all_folders, state=tk.DISABLED)
        self.folder_select_all_button.pack(side=tk.LEFT, padx=(0, 5))

        self.folder_deselect_all_button = ttk.Button(folder_toolbar_frame, text="取消全选",
                                                     command=self.deselect_all_folders, state=tk.DISABLED)
        self.folder_deselect_all_button.pack(side=tk.LEFT, padx=(0, 10))

        # 删除文件夹按钮
        self.delete_folder_button = ttk.Button(folder_toolbar_frame, text="删除选中文件夹 (慎重!)",
                                              command=self.delete_selected_folders, state=tk.DISABLED)
        self.delete_folder_button.pack(side=tk.LEFT, padx=(0, 10))

        # 状态标签
        self.folder_selection_label = ttk.Label(folder_toolbar_frame, text="未选择文件夹")
        self.folder_selection_label.pack(side=tk.LEFT, padx=(10, 0))

        # 创建文件夹表格
        folder_columns = ("选择", "文件夹路径", "文件数量", "总大小", "安全状态")
        self.folders_tree = ttk.Treeview(self.folders_frame, columns=folder_columns, show="headings", height=15)

        # 设置列标题和宽度
        self.folders_tree.heading("选择", text="☐")
        self.folders_tree.column("选择", width=40, anchor=tk.CENTER)

        self.folders_tree.heading("文件夹路径", text="文件夹路径")
        self.folders_tree.column("文件夹路径", width=400)

        self.folders_tree.heading("文件数量", text="文件数量")
        self.folders_tree.column("文件数量", width=100, anchor=tk.CENTER)

        self.folders_tree.heading("总大小", text="总大小")
        self.folders_tree.column("总大小", width=120, anchor=tk.CENTER)

        self.folders_tree.heading("安全状态", text="安全状态")
        self.folders_tree.column("安全状态", width=120, anchor=tk.CENTER)

        # 绑定点击事件
        self.folders_tree.bind("<Button-1>", self.on_folder_tree_click)
        self.folders_tree.bind("<Double-Button-1>", self.on_folder_tree_double_click)

        # 添加滚动条
        folders_scrollbar = ttk.Scrollbar(self.folders_frame, orient=tk.VERTICAL, command=self.folders_tree.yview)
        self.folders_tree.configure(yscrollcommand=folders_scrollbar.set)

        self.folders_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        folders_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # === 空文件夹页面 ===
        self.empty_folders_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.empty_folders_frame, text="空文件夹清理")

        # 创建顶部工具栏
        empty_toolbar_frame = ttk.Frame(self.empty_folders_frame)
        empty_toolbar_frame.pack(fill=tk.X, padx=5, pady=5)

        # 扫描空文件夹按钮
        self.scan_empty_button = ttk.Button(empty_toolbar_frame, text="扫描空文件夹",
                                           command=self.scan_empty_folders)
        self.scan_empty_button.pack(side=tk.LEFT, padx=(0, 10))

        # 全选/取消全选按钮
        self.empty_select_all_button = ttk.Button(empty_toolbar_frame, text="全选",
                                                  command=self.select_all_empty_folders, state=tk.DISABLED)
        self.empty_select_all_button.pack(side=tk.LEFT, padx=(0, 5))

        self.empty_deselect_all_button = ttk.Button(empty_toolbar_frame, text="取消全选",
                                                    command=self.deselect_all_empty_folders, state=tk.DISABLED)
        self.empty_deselect_all_button.pack(side=tk.LEFT, padx=(0, 10))

        # 删除空文件夹按钮
        self.delete_empty_button = ttk.Button(empty_toolbar_frame, text="删除选中空文件夹",
                                             command=self.delete_empty_folders, state=tk.DISABLED)
        self.delete_empty_button.pack(side=tk.LEFT, padx=(0, 10))

        # 状态标签
        self.empty_selection_label = ttk.Label(empty_toolbar_frame, text="未扫描")
        self.empty_selection_label.pack(side=tk.LEFT, padx=(10, 0))

        # 创建空文件夹表格
        empty_columns = ("选择", "文件夹路径", "安全状态", "原因")
        self.empty_folders_tree = ttk.Treeview(self.empty_folders_frame, columns=empty_columns,
                                               show="headings", height=15)

        # 设置列标题和宽度
        self.empty_folders_tree.heading("选择", text="☐")
        self.empty_folders_tree.column("选择", width=40, anchor=tk.CENTER)

        self.empty_folders_tree.heading("文件夹路径", text="文件夹路径")
        self.empty_folders_tree.column("文件夹路径", width=500)

        self.empty_folders_tree.heading("安全状态", text="安全状态")
        self.empty_folders_tree.column("安全状态", width=120, anchor=tk.CENTER)

        self.empty_folders_tree.heading("原因", text="原因")
        self.empty_folders_tree.column("原因", width=200)

        # 绑定点击事件
        self.empty_folders_tree.bind("<Button-1>", self.on_empty_tree_click)
        self.empty_folders_tree.bind("<Double-Button-1>", self.on_empty_tree_double_click)

        # 添加滚动条
        empty_scrollbar = ttk.Scrollbar(self.empty_folders_frame, orient=tk.VERTICAL,
                                       command=self.empty_folders_tree.yview)
        self.empty_folders_tree.configure(yscrollcommand=empty_scrollbar.set)

        self.empty_folders_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        empty_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 初始化空文件夹相关变量
        self.selected_empty_folders = set()  # 存储选中的空文件夹路径
        self.empty_folder_list = []  # 存储所有空文件夹信息

        # === 重复文件检测页面 ===
        self.duplicates_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.duplicates_frame, text="重复文件检测")

        # 创建顶部工具栏
        dup_toolbar_frame = ttk.Frame(self.duplicates_frame)
        dup_toolbar_frame.pack(fill=tk.X, padx=5, pady=5)

        # 扫描重复文件按钮
        self.scan_dup_button = ttk.Button(dup_toolbar_frame, text="扫描重复文件",
                                          command=self.scan_duplicates)
        self.scan_dup_button.pack(side=tk.LEFT, padx=(0, 10))

        # 最小文件大小设置
        ttk.Label(dup_toolbar_frame, text="最小文件大小:").pack(side=tk.LEFT, padx=(0, 5))
        self.dup_min_size_var = tk.StringVar(value="1MB")
        dup_size_combo = ttk.Combobox(dup_toolbar_frame, textvariable=self.dup_min_size_var,
                                      values=["0KB", "100KB", "1MB", "10MB", "100MB"],
                                      state="readonly", width=10)
        dup_size_combo.pack(side=tk.LEFT, padx=(0, 15))

        # 全选/取消全选按钮
        self.dup_select_all_button = ttk.Button(dup_toolbar_frame, text="全选重复",
                                                command=self.select_all_duplicates, state=tk.DISABLED)
        self.dup_select_all_button.pack(side=tk.LEFT, padx=(0, 5))

        self.dup_deselect_all_button = ttk.Button(dup_toolbar_frame, text="取消全选",
                                                  command=self.deselect_all_duplicates, state=tk.DISABLED)
        self.dup_deselect_all_button.pack(side=tk.LEFT, padx=(0, 10))

        # 删除选中按钮
        self.delete_dup_button = ttk.Button(dup_toolbar_frame, text="删除选中文件",
                                           command=self.delete_duplicates, state=tk.DISABLED)
        self.delete_dup_button.pack(side=tk.LEFT, padx=(0, 10))

        # 状态标签
        self.dup_status_label = ttk.Label(dup_toolbar_frame, text="未扫描")
        self.dup_status_label.pack(side=tk.LEFT, padx=(10, 0))

        # 创建重复文件表格（树形视图）
        dup_columns = ("选择", "组", "文件名", "大小", "路径", "修改时间", "建议")
        self.duplicates_tree = ttk.Treeview(self.duplicates_frame, columns=dup_columns,
                                           show="tree headings", height=15)

        # 设置列标题和宽度
        self.duplicates_tree.heading("#0", text="重复组")
        self.duplicates_tree.column("#0", width=80)

        self.duplicates_tree.heading("选择", text="☐")
        self.duplicates_tree.column("选择", width=40, anchor=tk.CENTER)

        self.duplicates_tree.heading("组", text="组号")
        self.duplicates_tree.column("组", width=50, anchor=tk.CENTER)

        self.duplicates_tree.heading("文件名", text="文件名")
        self.duplicates_tree.column("文件名", width=200)

        self.duplicates_tree.heading("大小", text="大小")
        self.duplicates_tree.column("大小", width=100, anchor=tk.CENTER)

        self.duplicates_tree.heading("路径", text="路径")
        self.duplicates_tree.column("路径", width=300)

        self.duplicates_tree.heading("修改时间", text="修改时间")
        self.duplicates_tree.column("修改时间", width=150, anchor=tk.CENTER)

        self.duplicates_tree.heading("建议", text="建议")
        self.duplicates_tree.column("建议", width=80, anchor=tk.CENTER)

        # 绑定点击事件
        self.duplicates_tree.bind("<Button-1>", self.on_dup_tree_click)
        self.duplicates_tree.bind("<Double-Button-1>", self.on_dup_tree_double_click)

        # 添加滚动条
        dup_scrollbar = ttk.Scrollbar(self.duplicates_frame, orient=tk.VERTICAL,
                                      command=self.duplicates_tree.yview)
        self.duplicates_tree.configure(yscrollcommand=dup_scrollbar.set)

        self.duplicates_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        dup_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

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
            # 检查文件安全性
            safety_status = "未知"
            if self.safety_checker:
                warning_level = self.safety_checker.get_warning_level(file_path)
                is_system, reason = self.safety_checker.is_system_file(file_path)

                if warning_level == "danger":
                    safety_status = "🔴 系统文件"
                elif warning_level == "warning":
                    safety_status = "⚠️ 需警告"
                else:
                    safety_status = "✅ 可删除"

            # 插入到表格
            item_id = self.files_tree.insert("", tk.END, values=(
                "☐",  # 未选中状态
                i,
                file_path.name,
                self.scanner.format_size(file_size),
                str(file_path.parent),
                safety_status
            ))

            # 存储文件路径到item的tags中，方便后续获取
            self.files_tree.item(item_id, tags=(str(file_path),))

            # 根据安全状态设置行颜色
            if safety_status.startswith("🔴"):
                self.files_tree.tag_configure(item_id, background="#ffcccc")  # 红色背景
            elif safety_status.startswith("⚠️"):
                self.files_tree.tag_configure(item_id, background="#fff3cd")  # 黄色背景

        # 启用选择和删除按钮
        self.select_all_button.config(state=tk.NORMAL)
        self.deselect_all_button.config(state=tk.NORMAL)

        # === 文件夹分组统计 ===
        self._group_files_by_folder()
        self._display_folder_stats()

        # 启用文件夹选择和删除按钮
        self.folder_select_all_button.config(state=tk.NORMAL)
        self.folder_deselect_all_button.config(state=tk.NORMAL)

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

    def on_tree_click(self, event):
        """处理树视图点击事件"""
        try:
            # 获取点击位置
            region = self.files_tree.identify_region(event.x, event.y)
            if region != "cell":
                return

            # 获取点击的item和列
            item = self.files_tree.identify_row(event.y)
            column = self.files_tree.identify_column(event.x)

            if not item or column != "#1":  # #1 是第一列（选择列）
                return

            # 切换选择状态
            self.toggle_file_selection(item)

        except Exception as e:
            print(f"Tree click error: {e}")

    def on_tree_double_click(self, event):
        """处理双击事件 - 打开文件所在文件夹"""
        try:
            item = self.files_tree.identify_row(event.y)
            if not item:
                return

            # 获取文件路径
            file_path = self.get_file_path_from_item(item)
            if file_path and os.path.exists(file_path):
                # 打开文件所在文件夹并选中文件
                import subprocess
                subprocess.Popen(f'explorer /select,"{file_path}"')
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件位置：{str(e)}")

    def toggle_file_selection(self, item):
        """切换文件选择状态"""
        try:
            values = list(self.files_tree.item(item, "values"))
            file_path = self.get_file_path_from_item(item)

            if not file_path:
                return

            # 切换选择状态
            if values[0] == "☐":  # 未选中
                values[0] = "☑"
                self.selected_files.add(file_path)
            else:  # 已选中
                values[0] = "☐"
                self.selected_files.discard(file_path)

            # 更新显示
            self.files_tree.item(item, values=values)
            self.update_selection_label()

            # 如果有文件被选中，启用删除按钮
            if self.selected_files:
                self.delete_button.config(state=tk.NORMAL)
            else:
                self.delete_button.config(state=tk.DISABLED)

        except Exception as e:
            print(f"Toggle selection error: {e}")

    def get_file_path_from_item(self, item):
        """从item获取文件路径"""
        try:
            tags = self.files_tree.item(item, "tags")
            if tags:
                return tags[0]
            return None
        except:
            return None

    def select_all_files(self):
        """全选文件"""
        try:
            for item in self.files_tree.get_children():
                values = list(self.files_tree.item(item, "values"))
                file_path = self.get_file_path_from_item(item)

                if file_path:
                    values[0] = "☑"
                    self.selected_files.add(file_path)
                    self.files_tree.item(item, values=values)

            self.update_selection_label()
            if self.selected_files:
                self.delete_button.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("错误", f"全选失败：{str(e)}")

    def deselect_all_files(self):
        """取消全选"""
        try:
            for item in self.files_tree.get_children():
                values = list(self.files_tree.item(item, "values"))
                values[0] = "☐"
                self.files_tree.item(item, values=values)

            self.selected_files.clear()
            self.update_selection_label()
            self.delete_button.config(state=tk.DISABLED)

        except Exception as e:
            messagebox.showerror("错误", f"取消全选失败：{str(e)}")

    def update_selection_label(self):
        """更新选择状态标签"""
        count = len(self.selected_files)
        if count == 0:
            self.file_selection_label.config(text="未选择文件")
        else:
            self.file_selection_label.config(text=f"已选择 {count} 个文件")

    def delete_selected_files(self):
        """删除选中的文件"""
        if not self.selected_files:
            messagebox.showwarning("提示", "请先选择要删除的文件")
            return

        try:
            # 使用安全检查器批量检查文件
            if self.safety_checker:
                check_result = self.safety_checker.batch_check_files(self.selected_files)

                # 检查是否有危险文件（系统文件）
                if check_result["danger"]:
                    danger_files = "\n".join([os.path.basename(f) for f in check_result["danger"][:5]])
                    if len(check_result["danger"]) > 5:
                        danger_files += f"\n... 还有 {len(check_result['danger']) - 5} 个文件"

                    messagebox.showerror(
                        "禁止删除系统文件",
                        f"以下文件为系统关键文件，禁止删除：\n\n{danger_files}\n\n"
                        "为保护系统安全，已取消删除操作。\n请取消选择这些文件后再试。"
                    )
                    return

                # 显示警告信息
                warning_msg = self._build_delete_warning_message(check_result)
            else:
                warning_msg = self._build_simple_warning_message()

            # 显示确认对话框
            if not messagebox.askyesno("确认删除", warning_msg, icon="warning"):
                return

            # 执行删除
            self._perform_delete_operation()

        except Exception as e:
            messagebox.showerror("删除失败", f"删除过程中发生错误：{str(e)}")

    def _build_delete_warning_message(self, check_result):
        """构建删除警告消息"""
        safe_count = len(check_result["safe"])
        warning_count = len(check_result["warning"])
        total_count = safe_count + warning_count

        msg_parts = [
            "⚠️ 删除文件操作不可恢复！",
            "",
            f"即将删除 {total_count} 个文件：",
            f"  ✅ 普通文件: {safe_count} 个",
            f"  ⚠️ 重要用户数据: {warning_count} 个",
            ""
        ]

        if warning_count > 0:
            msg_parts.append("警告：部分文件位于重要用户目录中（如文档、桌面等）")
            msg_parts.append("")

        msg_parts.extend([
            "删除后文件将被移动到回收站",
            "",
            "是否确认删除？"
        ])

        return "\n".join(msg_parts)

    def _build_simple_warning_message(self):
        """构建简单的警告消息（当安全检查器不可用时）"""
        return (
            f"⚠️ 即将删除 {len(self.selected_files)} 个文件\n\n"
            "删除操作不可恢复！\n"
            "文件将被移动到回收站。\n\n"
            "是否确认删除？"
        )

    def _perform_delete_operation(self):
        """执行实际的删除操作"""
        success_count = 0
        failed_files = []
        deleted_files = []

        # 创建进度对话框
        progress_window = tk.Toplevel(self.root)
        progress_window.title("正在删除文件...")
        progress_window.geometry("400x150")
        progress_window.transient(self.root)
        progress_window.grab_set()

        # 居中显示
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - 200
        y = (progress_window.winfo_screenheight() // 2) - 75
        progress_window.geometry(f'400x150+{x}+{y}')

        progress_label = ttk.Label(progress_window, text="正在删除文件...", font=('Arial', 10))
        progress_label.pack(pady=20)

        progress_bar = ttk.Progressbar(progress_window, maximum=len(self.selected_files), length=350)
        progress_bar.pack(pady=10)

        status_label = ttk.Label(progress_window, text="", font=('Arial', 9))
        status_label.pack(pady=5)

        self.root.update()

        # 尝试导入send2trash库（用于安全删除到回收站）
        try:
            import send2trash
            use_recycle_bin = True
        except ImportError:
            use_recycle_bin = False

        # 删除文件
        for idx, file_path in enumerate(self.selected_files, 1):
            try:
                status_label.config(text=f"删除: {os.path.basename(file_path)}")
                progress_bar['value'] = idx
                self.root.update()

                if use_recycle_bin:
                    # 使用send2trash移动到回收站
                    send2trash.send2trash(file_path)
                else:
                    # 直接删除
                    os.remove(file_path)

                success_count += 1
                deleted_files.append(file_path)

            except Exception as e:
                failed_files.append((file_path, str(e)))

        progress_window.destroy()

        # 记录删除日志
        self._log_delete_operation(deleted_files, failed_files)

        # 从列表中移除已删除的文件
        self._remove_deleted_files_from_tree(deleted_files)

        # 显示结果
        self._show_delete_result(success_count, failed_files, use_recycle_bin)

        # 清空选择
        self.selected_files.clear()
        self.update_selection_label()
        self.delete_button.config(state=tk.DISABLED)

    def _log_delete_operation(self, deleted_files, failed_files):
        """记录删除操作到日志文件"""
        try:
            log_file = "delete_log.txt"
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(f"\n{'='*60}\n")
                f.write(f"删除操作时间: {timestamp}\n")
                f.write(f"成功删除: {len(deleted_files)} 个文件\n")
                f.write(f"删除失败: {len(failed_files)} 个文件\n")
                f.write(f"{'='*60}\n")

                if deleted_files:
                    f.write("\n成功删除的文件:\n")
                    for file_path in deleted_files:
                        f.write(f"  - {file_path}\n")

                if failed_files:
                    f.write("\n删除失败的文件:\n")
                    for file_path, error in failed_files:
                        f.write(f"  - {file_path}\n    错误: {error}\n")

                f.write(f"\n{'='*60}\n\n")

        except Exception as e:
            print(f"日志记录失败: {e}")

    def _remove_deleted_files_from_tree(self, deleted_files):
        """从树视图中移除已删除的文件"""
        try:
            deleted_set = set(deleted_files)
            for item in self.files_tree.get_children():
                file_path = self.get_file_path_from_item(item)
                if file_path in deleted_set:
                    self.files_tree.delete(item)
        except Exception as e:
            print(f"移除文件显示失败: {e}")

    def _show_delete_result(self, success_count, failed_files, use_recycle_bin):
        """显示删除结果"""
        if not failed_files:
            # 全部成功
            location = "回收站" if use_recycle_bin else "永久删除"
            messagebox.showinfo(
                "删除成功",
                f"成功删除 {success_count} 个文件\n\n"
                f"文件已移动到{location}\n\n"
                f"删除日志已保存到 delete_log.txt"
            )
        else:
            # 部分失败
            failed_list = "\n".join([f"{os.path.basename(f)}: {e}" for f, e in failed_files[:5]])
            if len(failed_files) > 5:
                failed_list += f"\n... 还有 {len(failed_files) - 5} 个文件"

            messagebox.showwarning(
                "删除部分成功",
                f"成功删除: {success_count} 个文件\n"
                f"删除失败: {len(failed_files)} 个文件\n\n"
                f"失败文件:\n{failed_list}\n\n"
                f"详细信息请查看 delete_log.txt"
            )

    def _group_files_by_folder(self):
        """将文件按父文件夹分组"""
        from collections import defaultdict
        from pathlib import Path

        self.folder_stats = {}

        if not self.scanner or not self.scanner.largest_files:
            return

        # 按文件夹分组
        folder_groups = defaultdict(lambda: {"files": [], "total_size": 0})

        for file_path, file_size in self.scanner.largest_files:
            parent_folder = str(file_path.parent)
            folder_groups[parent_folder]["files"].append((file_path, file_size))
            folder_groups[parent_folder]["total_size"] += file_size

        # 转换为列表并按总大小排序
        self.folder_stats = {
            folder: {
                "files": info["files"],
                "file_count": len(info["files"]),
                "total_size": info["total_size"]
            }
            for folder, info in folder_groups.items()
        }

    def _display_folder_stats(self):
        """显示文件夹统计信息"""
        # 清空表格
        for item in self.folders_tree.get_children():
            self.folders_tree.delete(item)

        # 按总大小排序
        sorted_folders = sorted(
            self.folder_stats.items(),
            key=lambda x: x[1]["total_size"],
            reverse=True
        )

        # 显示文件夹统计
        for folder_path, stats in sorted_folders:
            # 检查文件夹安全性
            safety_status = "未知"
            if self.safety_checker:
                is_safe, danger_files, warning_files = self.safety_checker.check_folder_safety(folder_path)

                if danger_files:
                    safety_status = "🔴 包含系统文件"
                elif warning_files:
                    safety_status = "⚠️ 需警告"
                else:
                    safety_status = "✅ 可删除"

            # 插入到表格
            item_id = self.folders_tree.insert("", tk.END, values=(
                "☐",  # 未选中状态
                folder_path,
                stats["file_count"],
                self.scanner.format_size(stats["total_size"]) if self.scanner else str(stats["total_size"]),
                safety_status
            ))

            # 存储文件夹路径到tags中
            self.folders_tree.item(item_id, tags=(folder_path,))

            # 根据安全状态设置行颜色
            if safety_status.startswith("🔴"):
                self.folders_tree.tag_configure(item_id, background="#ffcccc")  # 红色背景
            elif safety_status.startswith("⚠️"):
                self.folders_tree.tag_configure(item_id, background="#fff3cd")  # 黄色背景

    def on_folder_tree_click(self, event):
        """处理文件夹树视图点击事件"""
        try:
            region = self.folders_tree.identify_region(event.x, event.y)
            if region != "cell":
                return

            item = self.folders_tree.identify_row(event.y)
            column = self.folders_tree.identify_column(event.x)

            if not item or column != "#1":  # #1 是第一列（选择列）
                return

            # 切换选择状态
            self.toggle_folder_selection(item)

        except Exception as e:
            print(f"Folder tree click error: {e}")

    def on_folder_tree_double_click(self, event):
        """处理双击事件 - 打开文件夹"""
        try:
            item = self.folders_tree.identify_row(event.y)
            if not item:
                return

            folder_path = self.get_folder_path_from_item(item)
            if folder_path and os.path.exists(folder_path):
                import subprocess
                subprocess.Popen(f'explorer "{folder_path}"')
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹：{str(e)}")

    def toggle_folder_selection(self, item):
        """切换文件夹选择状态"""
        try:
            values = list(self.folders_tree.item(item, "values"))
            folder_path = self.get_folder_path_from_item(item)

            if not folder_path:
                return

            # 切换选择状态
            if values[0] == "☐":  # 未选中
                values[0] = "☑"
                self.selected_folders.add(folder_path)
            else:  # 已选中
                values[0] = "☐"
                self.selected_folders.discard(folder_path)

            # 更新显示
            self.folders_tree.item(item, values=values)
            self.update_folder_selection_label()

            # 如果有文件夹被选中，启用删除按钮
            if self.selected_folders:
                self.delete_folder_button.config(state=tk.NORMAL)
            else:
                self.delete_folder_button.config(state=tk.DISABLED)

        except Exception as e:
            print(f"Toggle folder selection error: {e}")

    def get_folder_path_from_item(self, item):
        """从item获取文件夹路径"""
        try:
            tags = self.folders_tree.item(item, "tags")
            if tags:
                return tags[0]
            return None
        except:
            return None

    def select_all_folders(self):
        """全选文件夹"""
        try:
            for item in self.folders_tree.get_children():
                values = list(self.folders_tree.item(item, "values"))
                folder_path = self.get_folder_path_from_item(item)

                if folder_path:
                    values[0] = "☑"
                    self.selected_folders.add(folder_path)
                    self.folders_tree.item(item, values=values)

            self.update_folder_selection_label()
            if self.selected_folders:
                self.delete_folder_button.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("错误", f"全选失败：{str(e)}")

    def deselect_all_folders(self):
        """取消全选文件夹"""
        try:
            for item in self.folders_tree.get_children():
                values = list(self.folders_tree.item(item, "values"))
                values[0] = "☐"
                self.folders_tree.item(item, values=values)

            self.selected_folders.clear()
            self.update_folder_selection_label()
            self.delete_folder_button.config(state=tk.DISABLED)

        except Exception as e:
            messagebox.showerror("错误", f"取消全选失败：{str(e)}")

    def update_folder_selection_label(self):
        """更新文件夹选择状态标签"""
        count = len(self.selected_folders)
        if count == 0:
            self.folder_selection_label.config(text="未选择文件夹")
        else:
            self.folder_selection_label.config(text=f"已选择 {count} 个文件夹")

    def delete_selected_folders(self):
        """删除选中的文件夹"""
        if not self.selected_folders:
            messagebox.showwarning("提示", "请先选择要删除的文件夹")
            return

        try:
            # 使用安全检查器检查所有文件夹
            danger_folders = []
            warning_folders = []
            safe_folders = []

            if self.safety_checker:
                for folder_path in self.selected_folders:
                    is_safe, danger_files, warning_files = self.safety_checker.check_folder_safety(folder_path)

                    if danger_files:
                        danger_folders.append(folder_path)
                    elif warning_files:
                        warning_folders.append(folder_path)
                    else:
                        safe_folders.append(folder_path)

                # 如果有危险文件夹，拒绝删除
                if danger_folders:
                    danger_list = "\n".join([os.path.basename(f) for f in danger_folders[:5]])
                    if len(danger_folders) > 5:
                        danger_list += f"\n... 还有 {len(danger_folders) - 5} 个文件夹"

                    messagebox.showerror(
                        "禁止删除系统文件夹",
                        f"以下文件夹包含系统关键文件，禁止删除：\n\n{danger_list}\n\n"
                        "为保护系统安全，已取消删除操作。\n请取消选择这些文件夹后再试。"
                    )
                    return

                # 构建警告消息
                warning_msg = self._build_folder_delete_warning(safe_folders, warning_folders)
            else:
                warning_msg = self._build_simple_folder_warning()

            # 显示确认对话框
            if not messagebox.askyesno("确认删除文件夹", warning_msg, icon="warning"):
                return

            # 执行删除
            self._perform_folder_delete_operation()

        except Exception as e:
            messagebox.showerror("删除失败", f"删除过程中发生错误：{str(e)}")

    def _build_folder_delete_warning(self, safe_folders, warning_folders):
        """构建文件夹删除警告消息"""
        total_count = len(safe_folders) + len(warning_folders)

        msg_parts = [
            "⚠️ 删除文件夹操作不可恢复！",
            "",
            f"即将删除 {total_count} 个文件夹及其所有内容：",
            f"  ✅ 普通文件夹: {len(safe_folders)} 个",
            f"  ⚠️ 包含重要用户数据的文件夹: {len(warning_folders)} 个",
            ""
        ]

        if warning_folders:
            msg_parts.append("警告：部分文件夹包含重要用户数据（如文档、桌面等）")
            msg_parts.append("")

        msg_parts.extend([
            "文件夹及其所有子文件、子文件夹将被移动到回收站",
            "",
            "是否确认删除？"
        ])

        return "\n".join(msg_parts)

    def _build_simple_folder_warning(self):
        """构建简单的文件夹删除警告"""
        return (
            f"⚠️ 即将删除 {len(self.selected_folders)} 个文件夹及其所有内容\n\n"
            "删除操作不可恢复！\n"
            "文件夹将被移动到回收站。\n\n"
            "是否确认删除？"
        )

    def _perform_folder_delete_operation(self):
        """执行文件夹删除操作"""
        success_count = 0
        failed_folders = []
        deleted_folders = []

        # 创建进度对话框
        progress_window = tk.Toplevel(self.root)
        progress_window.title("正在删除文件夹...")
        progress_window.geometry("400x150")
        progress_window.transient(self.root)
        progress_window.grab_set()

        # 居中显示
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - 200
        y = (progress_window.winfo_screenheight() // 2) - 75
        progress_window.geometry(f'400x150+{x}+{y}')

        progress_label = ttk.Label(progress_window, text="正在删除文件夹...", font=('Arial', 10))
        progress_label.pack(pady=20)

        progress_bar = ttk.Progressbar(progress_window, maximum=len(self.selected_folders), length=350)
        progress_bar.pack(pady=10)

        status_label = ttk.Label(progress_window, text="", font=('Arial', 9))
        status_label.pack(pady=5)

        self.root.update()

        # 尝试导入send2trash库
        try:
            import send2trash
            use_recycle_bin = True
        except ImportError:
            use_recycle_bin = False

        # 删除文件夹
        for idx, folder_path in enumerate(self.selected_folders, 1):
            try:
                status_label.config(text=f"删除: {os.path.basename(folder_path)}")
                progress_bar['value'] = idx
                self.root.update()

                if use_recycle_bin:
                    import send2trash
                    send2trash.send2trash(folder_path)
                else:
                    import shutil
                    shutil.rmtree(folder_path)

                success_count += 1
                deleted_folders.append(folder_path)

            except Exception as e:
                failed_folders.append((folder_path, str(e)))

        progress_window.destroy()

        # 记录删除日志
        self._log_folder_delete_operation(deleted_folders, failed_folders)

        # 从列表中移除已删除的文件夹
        self._remove_deleted_folders_from_tree(deleted_folders)

        # 显示结果
        self._show_folder_delete_result(success_count, failed_folders, use_recycle_bin)

        # 清空选择
        self.selected_folders.clear()
        self.update_folder_selection_label()
        self.delete_folder_button.config(state=tk.DISABLED)

    def _log_folder_delete_operation(self, deleted_folders, failed_folders):
        """记录文件夹删除操作到日志"""
        try:
            log_file = "delete_log.txt"
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(f"\n{'='*60}\n")
                f.write(f"文件夹删除操作时间: {timestamp}\n")
                f.write(f"成功删除: {len(deleted_folders)} 个文件夹\n")
                f.write(f"删除失败: {len(failed_folders)} 个文件夹\n")
                f.write(f"{'='*60}\n")

                if deleted_folders:
                    f.write("\n成功删除的文件夹:\n")
                    for folder_path in deleted_folders:
                        f.write(f"  - {folder_path}\n")

                if failed_folders:
                    f.write("\n删除失败的文件夹:\n")
                    for folder_path, error in failed_folders:
                        f.write(f"  - {folder_path}\n    错误: {error}\n")

                f.write(f"\n{'='*60}\n\n")

        except Exception as e:
            print(f"日志记录失败: {e}")

    def _remove_deleted_folders_from_tree(self, deleted_folders):
        """从树视图中移除已删除的文件夹"""
        try:
            deleted_set = set(deleted_folders)
            for item in self.folders_tree.get_children():
                folder_path = self.get_folder_path_from_item(item)
                if folder_path in deleted_set:
                    self.folders_tree.delete(item)
        except Exception as e:
            print(f"移除文件夹显示失败: {e}")

    def _show_folder_delete_result(self, success_count, failed_folders, use_recycle_bin):
        """显示文件夹删除结果"""
        if not failed_folders:
            location = "回收站" if use_recycle_bin else "永久删除"
            messagebox.showinfo(
                "删除成功",
                f"成功删除 {success_count} 个文件夹及其所有内容\n\n"
                f"文件夹已移动到{location}\n\n"
                f"删除日志已保存到 delete_log.txt"
            )
        else:
            failed_list = "\n".join([f"{os.path.basename(f)}: {e}" for f, e in failed_folders[:5]])
            if len(failed_folders) > 5:
                failed_list += f"\n... 还有 {len(failed_folders) - 5} 个文件夹"

            messagebox.showwarning(
                "删除部分成功",
                f"成功删除: {success_count} 个文件夹\n"
                f"删除失败: {len(failed_folders)} 个文件夹\n\n"
                f"失败文件夹:\n{failed_list}\n\n"
                f"详细信息请查看 delete_log.txt"
            )

    # ========== 空文件夹清理功能 ==========

    def scan_empty_folders(self):
        """扫描空文件夹"""
        if not self.safety_checker:
            messagebox.showerror("错误", "文件安全检查模块不可用")
            return

        scan_path = self.path_var.get()
        if not scan_path or not os.path.exists(scan_path):
            messagebox.showwarning("警告", "请选择有效的扫描路径")
            return

        # 确认开始扫描
        result = messagebox.askquestion(
            "扫描空文件夹",
            f"将扫描以下路径中的所有空文件夹：\n\n{scan_path}\n\n是否继续？",
            icon='question'
        )

        if result != 'yes':
            return

        # 禁用扫描按钮
        self.scan_empty_button.config(state=tk.DISABLED)
        self.empty_selection_label.config(text="扫描中...")

        # 在独立线程中执行扫描
        def scan_thread():
            try:
                # 进度回调
                def progress_callback(current_path):
                    # 更新状态标签（简化显示）
                    folder_name = os.path.basename(current_path)
                    self.root.after(0, lambda: self.empty_selection_label.config(
                        text=f"扫描中: {folder_name[:30]}..."
                    ))

                # 执行扫描
                empty_folders = self.safety_checker.find_empty_folders(
                    scan_path,
                    progress_callback=progress_callback
                )

                # 在主线程中更新UI
                self.root.after(0, lambda: self._display_empty_folders(empty_folders))

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("扫描失败", f"扫描空文件夹时出错:\n{str(e)}"))
                self.root.after(0, lambda: self.empty_selection_label.config(text="扫描失败"))
            finally:
                self.root.after(0, lambda: self.scan_empty_button.config(state=tk.NORMAL))

        thread = threading.Thread(target=scan_thread, daemon=True)
        thread.start()

    def _display_empty_folders(self, empty_folders):
        """显示空文件夹扫描结果"""
        # 清空现有内容
        for item in self.empty_folders_tree.get_children():
            self.empty_folders_tree.delete(item)

        self.empty_folder_list = empty_folders
        self.selected_empty_folders.clear()

        if not empty_folders:
            self.empty_selection_label.config(text="未发现空文件夹")
            messagebox.showinfo("扫描完成", "未发现空文件夹")
            return

        # 显示结果
        safe_count = 0
        unsafe_count = 0

        for folder_path, is_safe, reason in empty_folders:
            checkbox = "☐"
            if is_safe:
                status = "✅ 可删除"
                safe_count += 1
                tag = "safe"
            else:
                status = "🔴 不建议删除"
                unsafe_count += 1
                tag = "danger"

            item = self.empty_folders_tree.insert("", tk.END, values=(
                checkbox,
                folder_path,
                status,
                reason
            ))

            # 设置背景色
            if tag == "danger":
                self.empty_folders_tree.item(item, tags=("danger",))

        # 配置标签样式
        self.empty_folders_tree.tag_configure("danger", background="#ffcccc")

        # 更新状态标签
        self.empty_selection_label.config(
            text=f"找到 {len(empty_folders)} 个空文件夹 (可删除: {safe_count}, 不建议: {unsafe_count})"
        )

        # 启用按钮
        self.empty_select_all_button.config(state=tk.NORMAL)
        self.empty_deselect_all_button.config(state=tk.NORMAL)

        # 显示统计信息
        messagebox.showinfo(
            "扫描完成",
            f"扫描完成！\n\n"
            f"共找到 {len(empty_folders)} 个空文件夹\n"
            f"• 可安全删除: {safe_count}\n"
            f"• 不建议删除: {unsafe_count}\n\n"
            f"请在列表中选择要删除的空文件夹"
        )

    def select_all_empty_folders(self):
        """全选空文件夹（只选择可安全删除的）"""
        self.selected_empty_folders.clear()

        for item in self.empty_folders_tree.get_children():
            values = self.empty_folders_tree.item(item)['values']
            status = values[2]  # 安全状态

            # 只选择可删除的文件夹
            if "可删除" in status:
                folder_path = values[1]
                self.selected_empty_folders.add(folder_path)
                self.empty_folders_tree.set(item, "选择", "☑")

        self._update_empty_selection_label()

    def deselect_all_empty_folders(self):
        """取消全选空文件夹"""
        self.selected_empty_folders.clear()

        for item in self.empty_folders_tree.get_children():
            self.empty_folders_tree.set(item, "选择", "☐")

        self._update_empty_selection_label()

    def on_empty_tree_click(self, event):
        """处理空文件夹树的点击事件"""
        region = self.empty_folders_tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.empty_folders_tree.identify_column(event.x)
            item = self.empty_folders_tree.identify_row(event.y)

            if item and column == "#1":  # 选择列
                self._toggle_empty_folder_selection(item)

    def on_empty_tree_double_click(self, event):
        """双击打开文件夹位置"""
        item = self.empty_folders_tree.selection()
        if item:
            values = self.empty_folders_tree.item(item[0])['values']
            folder_path = values[1]

            try:
                # 打开父文件夹并选中该文件夹
                parent_folder = os.path.dirname(folder_path)
                if os.path.exists(parent_folder):
                    os.startfile(parent_folder)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件夹:\n{str(e)}")

    def _toggle_empty_folder_selection(self, item):
        """切换空文件夹选择状态"""
        values = self.empty_folders_tree.item(item)['values']
        folder_path = values[1]
        status = values[2]

        # 不建议删除的文件夹不能选择
        if "不建议" in status:
            messagebox.showwarning(
                "警告",
                f"该文件夹不建议删除:\n\n{folder_path}\n\n原因: {values[3]}"
            )
            return

        # 切换选择状态
        if folder_path in self.selected_empty_folders:
            self.selected_empty_folders.remove(folder_path)
            self.empty_folders_tree.set(item, "选择", "☐")
        else:
            self.selected_empty_folders.add(folder_path)
            self.empty_folders_tree.set(item, "选择", "☑")

        self._update_empty_selection_label()

    def _update_empty_selection_label(self):
        """更新空文件夹选择状态标签"""
        count = len(self.selected_empty_folders)
        if count == 0:
            self.empty_selection_label.config(
                text=f"找到 {len(self.empty_folder_list)} 个空文件夹，未选择"
            )
            self.delete_empty_button.config(state=tk.DISABLED)
        else:
            self.empty_selection_label.config(text=f"已选择 {count} 个空文件夹")
            self.delete_empty_button.config(state=tk.NORMAL)

    def delete_empty_folders(self):
        """删除选中的空文件夹"""
        if not self.selected_empty_folders:
            messagebox.showwarning("警告", "请先选择要删除的空文件夹")
            return

        # 确认删除
        result = messagebox.askquestion(
            "确认删除",
            f"确定要删除 {len(self.selected_empty_folders)} 个空文件夹吗？\n\n"
            f"文件夹将被移动到回收站（如果可用）\n\n"
            f"此操作无法通过本工具撤销！",
            icon='warning'
        )

        if result != 'yes':
            return

        # 尝试导入send2trash
        try:
            from send2trash import send2trash
            use_recycle_bin = True
        except ImportError:
            # 询问是否永久删除
            result = messagebox.askquestion(
                "回收站不可用",
                "无法使用回收站功能，文件夹将被永久删除！\n\n是否继续？",
                icon='warning'
            )
            if result != 'yes':
                return
            use_recycle_bin = False

        # 创建进度窗口
        progress_window = tk.Toplevel(self.root)
        progress_window.title("删除空文件夹")
        progress_window.geometry("400x150")
        progress_window.transient(self.root)
        progress_window.grab_set()

        ttk.Label(progress_window, text="正在删除空文件夹...",
                 font=('Arial', 12)).pack(pady=20)

        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(progress_window, variable=progress_var,
                                      maximum=100, length=350)
        progress_bar.pack(pady=10)

        status_label = ttk.Label(progress_window, text="准备中...")
        status_label.pack(pady=10)

        # 在独立线程中执行删除
        deleted_folders = []
        failed_folders = []

        def delete_thread():
            total = len(self.selected_empty_folders)
            for i, folder_path in enumerate(self.selected_empty_folders, 1):
                try:
                    # 更新进度
                    progress = (i / total) * 100
                    folder_name = os.path.basename(folder_path)
                    self.root.after(0, lambda p=progress, n=folder_name: (
                        progress_var.set(p),
                        status_label.config(text=f"删除: {n}")
                    ))

                    # 执行删除
                    if use_recycle_bin:
                        send2trash(folder_path)
                    else:
                        import shutil
                        shutil.rmtree(folder_path)

                    deleted_folders.append(folder_path)

                except Exception as e:
                    failed_folders.append((folder_path, str(e)))

                time.sleep(0.05)  # 短暂延迟，让用户看到进度

            # 完成后关闭进度窗口并显示结果
            self.root.after(0, lambda: progress_window.destroy())
            self.root.after(0, lambda: self._show_empty_delete_result(
                deleted_folders, failed_folders, use_recycle_bin
            ))

        thread = threading.Thread(target=delete_thread, daemon=True)
        thread.start()

    def _show_empty_delete_result(self, deleted_folders, failed_folders, use_recycle_bin):
        """显示空文件夹删除结果"""
        # 记录日志
        self._log_empty_delete_operation(deleted_folders, failed_folders)

        # 从树视图中移除已删除的文件夹
        self._remove_deleted_empty_from_tree(deleted_folders)

        # 清空选择
        self.selected_empty_folders.clear()
        self._update_empty_selection_label()

        # 显示结果
        if not failed_folders:
            location = "回收站" if use_recycle_bin else "永久删除"
            messagebox.showinfo(
                "删除成功",
                f"成功删除 {len(deleted_folders)} 个空文件夹\n\n"
                f"文件夹已移动到{location}\n\n"
                f"删除日志已保存到 delete_log.txt"
            )
        else:
            failed_list = "\n".join([f"{os.path.basename(f)}: {e}" for f, e in failed_folders[:5]])
            if len(failed_folders) > 5:
                failed_list += f"\n... 还有 {len(failed_folders) - 5} 个文件夹"

            messagebox.showwarning(
                "删除部分成功",
                f"成功删除: {len(deleted_folders)} 个空文件夹\n"
                f"删除失败: {len(failed_folders)} 个空文件夹\n\n"
                f"失败文件夹:\n{failed_list}\n\n"
                f"详细信息请查看 delete_log.txt"
            )

    def _log_empty_delete_operation(self, deleted_folders, failed_folders):
        """记录空文件夹删除操作到日志"""
        try:
            log_file = "delete_log.txt"
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(f"\n{'='*60}\n")
                f.write(f"空文件夹删除操作时间: {timestamp}\n")
                f.write(f"成功删除: {len(deleted_folders)} 个空文件夹\n")
                f.write(f"删除失败: {len(failed_folders)} 个空文件夹\n")
                f.write(f"{'='*60}\n")

                if deleted_folders:
                    f.write("\n成功删除的空文件夹:\n")
                    for folder_path in deleted_folders:
                        f.write(f"  - {folder_path}\n")

                if failed_folders:
                    f.write("\n删除失败的空文件夹:\n")
                    for folder_path, error in failed_folders:
                        f.write(f"  - {folder_path}\n    错误: {error}\n")

                f.write(f"\n{'='*60}\n\n")

        except Exception as e:
            print(f"日志记录失败: {e}")

    def _remove_deleted_empty_from_tree(self, deleted_folders):
        """从树视图中移除已删除的空文件夹"""
        try:
            deleted_set = set(deleted_folders)
            for item in self.empty_folders_tree.get_children():
                values = self.empty_folders_tree.item(item)['values']
                folder_path = values[1]
                if folder_path in deleted_set:
                    self.empty_folders_tree.delete(item)

            # 更新empty_folder_list
            self.empty_folder_list = [
                (path, safe, reason) for path, safe, reason in self.empty_folder_list
                if path not in deleted_set
            ]

        except Exception as e:
            print(f"移除文件夹显示失败: {e}")

    # ========== 重复文件检测功能 ==========

    def scan_duplicates(self):
        """扫描重复文件"""
        if not self.duplicate_finder:
            messagebox.showerror("错误", "重复文件检测模块不可用")
            return

        scan_path = self.path_var.get()
        if not scan_path or not os.path.exists(scan_path):
            messagebox.showwarning("警告", "请选择有效的扫描路径")
            return

        # 解析最小文件大小
        min_size_str = self.dup_min_size_var.get()
        min_size = self._parse_size(min_size_str)

        # 确认开始扫描
        result = messagebox.askquestion(
            "扫描重复文件",
            f"将扫描以下路径中的重复文件：\n\n{scan_path}\n\n"
            f"最小文件大小: {min_size_str}\n\n"
            f"注意：大文件扫描可能需要较长时间\n\n是否继续？",
            icon='question'
        )

        if result != 'yes':
            return

        # 禁用扫描按钮
        self.scan_dup_button.config(state=tk.DISABLED)
        self.dup_status_label.config(text="扫描中...")

        # 在独立线程中执行扫描
        def scan_thread():
            try:
                # 进度回调
                def progress_callback(message, phase, progress):
                    # 更新状态标签
                    self.root.after(0, lambda: self.dup_status_label.config(
                        text=f"{phase}: {message[:30]}..."
                    ))

                # 执行扫描
                duplicate_groups = self.duplicate_finder.find_duplicates(
                    scan_path,
                    min_size=min_size,
                    progress_callback=progress_callback
                )

                # 在主线程中更新UI
                self.root.after(0, lambda: self._display_duplicates(duplicate_groups))

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("扫描失败", f"扫描重复文件时出错:\n{str(e)}"))
                self.root.after(0, lambda: self.dup_status_label.config(text="扫描失败"))
            finally:
                self.root.after(0, lambda: self.scan_dup_button.config(state=tk.NORMAL))

        thread = threading.Thread(target=scan_thread, daemon=True)
        thread.start()

    def _parse_size(self, size_str):
        """解析文件大小字符串为字节数"""
        size_str = size_str.upper().replace(" ", "")

        if "KB" in size_str:
            return int(float(size_str.replace("KB", "")) * 1024)
        elif "MB" in size_str:
            return int(float(size_str.replace("MB", "")) * 1024 * 1024)
        elif "GB" in size_str:
            return int(float(size_str.replace("GB", "")) * 1024 * 1024 * 1024)
        else:
            return 0

    def _display_duplicates(self, duplicate_groups):
        """显示重复文件扫描结果"""
        # 清空现有内容
        for item in self.duplicates_tree.get_children():
            self.duplicates_tree.delete(item)

        self.duplicate_groups = duplicate_groups
        self.selected_duplicates.clear()

        if not duplicate_groups:
            self.dup_status_label.config(text="未发现重复文件")
            messagebox.showinfo("扫描完成", "未发现重复文件")
            return

        # 获取统计信息
        stats = self.duplicate_finder.get_statistics()

        # 显示结果 - 使用树形结构
        group_num = 0
        total_duplicates = 0

        for file_hash, files in duplicate_groups.items():
            group_num += 1

            # 获取智能保留建议
            suggested = self.duplicate_finder.get_smart_keep_suggestion(files)

            # 计算这一组的浪费空间
            file_size = files[0]['size']
            wasted = file_size * (len(files) - 1)

            # 创建组节点
            group_text = f"组 {group_num} ({len(files)} 个文件, 浪费 {self.duplicate_finder.format_size(wasted)})"
            group_item = self.duplicates_tree.insert("", tk.END, text=group_text, open=True)

            # 添加该组的所有文件
            for file_info in files:
                is_suggested = (file_info == suggested)
                checkbox = "☐"
                suggestion = "✓ 保留" if is_suggested else ""

                # 格式化修改时间
                mtime_str = datetime.fromtimestamp(file_info['mtime']).strftime('%Y-%m-%d %H:%M:%S')

                child_item = self.duplicates_tree.insert(
                    group_item, tk.END,
                    values=(
                        checkbox,
                        group_num,
                        file_info['name'],
                        self.duplicate_finder.format_size(file_info['size']),
                        file_info['path'],
                        mtime_str,
                        suggestion
                    ),
                    tags=("suggested" if is_suggested else "duplicate",)
                )

                # 存储文件路径以便后续操作
                self.duplicates_tree.set(child_item, "#0", "")  # 清除树列的值

                if not is_suggested:
                    total_duplicates += 1

        # 配置标签样式
        self.duplicates_tree.tag_configure("suggested", background="#e8f5e9")  # 浅绿色
        self.duplicates_tree.tag_configure("duplicate", background="#fff9c4")  # 浅黄色

        # 更新状态标签
        self.dup_status_label.config(
            text=f"找到 {len(duplicate_groups)} 组重复文件，共 {stats['total_duplicates']} 个文件，"
                 f"可节省 {stats['wasted_space_formatted']}"
        )

        # 启用按钮
        self.dup_select_all_button.config(state=tk.NORMAL)
        self.dup_deselect_all_button.config(state=tk.NORMAL)

        # 显示统计信息
        messagebox.showinfo(
            "扫描完成",
            f"扫描完成！\n\n"
            f"扫描文件总数: {stats['total_scanned']}\n"
            f"重复文件组数: {stats['duplicate_groups']}\n"
            f"重复文件数量: {stats['total_duplicates']}\n"
            f"浪费空间: {stats['wasted_space_formatted']}\n\n"
            f"建议：保留标记为 ✓保留 的文件，删除其他重复文件"
        )

    def select_all_duplicates(self):
        """全选重复文件（不选择建议保留的）"""
        self.selected_duplicates.clear()

        for group_item in self.duplicates_tree.get_children():
            for child_item in self.duplicates_tree.get_children(group_item):
                values = self.duplicates_tree.item(child_item)['values']
                suggestion = values[6]  # 建议列

                # 只选择非建议保留的文件
                if suggestion != "✓ 保留":
                    file_path = values[4]  # 路径列
                    self.selected_duplicates.add(file_path)
                    self.duplicates_tree.set(child_item, "选择", "☑")

        self._update_dup_selection_label()

    def deselect_all_duplicates(self):
        """取消全选重复文件"""
        self.selected_duplicates.clear()

        for group_item in self.duplicates_tree.get_children():
            for child_item in self.duplicates_tree.get_children(group_item):
                self.duplicates_tree.set(child_item, "选择", "☐")

        self._update_dup_selection_label()

    def on_dup_tree_click(self, event):
        """处理重复文件树的点击事件"""
        region = self.duplicates_tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.duplicates_tree.identify_column(event.x)
            item = self.duplicates_tree.identify_row(event.y)

            # 只处理子项（文件），不处理组节点
            if item and column == "#1":  # 选择列
                parent = self.duplicates_tree.parent(item)
                if parent:  # 确保是子项
                    self._toggle_dup_selection(item)

    def on_dup_tree_double_click(self, event):
        """双击打开文件所在位置"""
        item = self.duplicates_tree.selection()
        if item:
            # 检查是否是文件项（有父节点）
            parent = self.duplicates_tree.parent(item[0])
            if parent:
                values = self.duplicates_tree.item(item[0])['values']
                file_path = values[4]  # 路径列

                try:
                    # 打开文件所在文件夹并选中该文件
                    import subprocess
                    subprocess.run(['explorer', '/select,', file_path])
                except Exception as e:
                    messagebox.showerror("错误", f"无法打开文件位置:\n{str(e)}")

    def _toggle_dup_selection(self, item):
        """切换重复文件选择状态"""
        values = self.duplicates_tree.item(item)['values']
        file_path = values[4]  # 路径列
        suggestion = values[6]  # 建议列

        # 建议保留的文件不能选择删除
        if suggestion == "✓ 保留":
            messagebox.showwarning(
                "警告",
                f"该文件被建议保留，不应删除:\n\n{file_path}\n\n"
                f"建议删除同组的其他文件"
            )
            return

        # 切换选择状态
        if file_path in self.selected_duplicates:
            self.selected_duplicates.remove(file_path)
            self.duplicates_tree.set(item, "选择", "☐")
        else:
            self.selected_duplicates.add(file_path)
            self.duplicates_tree.set(item, "选择", "☑")

        self._update_dup_selection_label()

    def _update_dup_selection_label(self):
        """更新重复文件选择状态标签"""
        count = len(self.selected_duplicates)

        if count == 0:
            stats = self.duplicate_finder.get_statistics()
            self.dup_status_label.config(
                text=f"找到 {len(self.duplicate_groups)} 组重复文件，未选择"
            )
            self.delete_dup_button.config(state=tk.DISABLED)
        else:
            # 计算选中文件的总大小
            total_size = 0
            for group_item in self.duplicates_tree.get_children():
                for child_item in self.duplicates_tree.get_children(group_item):
                    values = self.duplicates_tree.item(child_item)['values']
                    file_path = values[4]
                    if file_path in self.selected_duplicates:
                        # 从文件路径获取大小
                        try:
                            total_size += os.path.getsize(file_path)
                        except:
                            pass

            self.dup_status_label.config(
                text=f"已选择 {count} 个重复文件，可释放 {self.duplicate_finder.format_size(total_size)}"
            )
            self.delete_dup_button.config(state=tk.NORMAL)

    def delete_duplicates(self):
        """删除选中的重复文件"""
        if not self.selected_duplicates:
            messagebox.showwarning("警告", "请先选择要删除的重复文件")
            return

        # 确认删除
        result = messagebox.askquestion(
            "确认删除",
            f"确定要删除 {len(self.selected_duplicates)} 个重复文件吗？\n\n"
            f"文件将被移动到回收站（如果可用）\n\n"
            f"此操作无法通过本工具撤销！",
            icon='warning'
        )

        if result != 'yes':
            return

        # 尝试导入send2trash
        try:
            from send2trash import send2trash
            use_recycle_bin = True
        except ImportError:
            # 询问是否永久删除
            result = messagebox.askquestion(
                "回收站不可用",
                "无法使用回收站功能，文件将被永久删除！\n\n是否继续？",
                icon='warning'
            )
            if result != 'yes':
                return
            use_recycle_bin = False

        # 创建进度窗口
        progress_window = tk.Toplevel(self.root)
        progress_window.title("删除重复文件")
        progress_window.geometry("400x150")
        progress_window.transient(self.root)
        progress_window.grab_set()

        ttk.Label(progress_window, text="正在删除重复文件...",
                 font=('Arial', 12)).pack(pady=20)

        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(progress_window, variable=progress_var,
                                      maximum=100, length=350)
        progress_bar.pack(pady=10)

        status_label = ttk.Label(progress_window, text="准备中...")
        status_label.pack(pady=10)

        # 在独立线程中执行删除
        deleted_files = []
        failed_files = []

        def delete_thread():
            total = len(self.selected_duplicates)
            for i, file_path in enumerate(self.selected_duplicates, 1):
                try:
                    # 更新进度
                    progress = (i / total) * 100
                    file_name = os.path.basename(file_path)
                    self.root.after(0, lambda p=progress, n=file_name: (
                        progress_var.set(p),
                        status_label.config(text=f"删除: {n}")
                    ))

                    # 执行删除
                    if use_recycle_bin:
                        send2trash(file_path)
                    else:
                        os.remove(file_path)

                    deleted_files.append(file_path)

                except Exception as e:
                    failed_files.append((file_path, str(e)))

                time.sleep(0.05)  # 短暂延迟，让用户看到进度

            # 完成后关闭进度窗口并显示结果
            self.root.after(0, lambda: progress_window.destroy())
            self.root.after(0, lambda: self._show_dup_delete_result(
                deleted_files, failed_files, use_recycle_bin
            ))

        thread = threading.Thread(target=delete_thread, daemon=True)
        thread.start()

    def _show_dup_delete_result(self, deleted_files, failed_files, use_recycle_bin):
        """显示重复文件删除结果"""
        # 记录日志
        self._log_dup_delete_operation(deleted_files, failed_files)

        # 从树视图中移除已删除的文件
        self._remove_deleted_dup_from_tree(deleted_files)

        # 清空选择
        self.selected_duplicates.clear()
        self._update_dup_selection_label()

        # 显示结果
        if not failed_files:
            location = "回收站" if use_recycle_bin else "永久删除"
            messagebox.showinfo(
                "删除成功",
                f"成功删除 {len(deleted_files)} 个重复文件\n\n"
                f"文件已移动到{location}\n\n"
                f"删除日志已保存到 delete_log.txt"
            )
        else:
            failed_list = "\n".join([f"{os.path.basename(f)}: {e}" for f, e in failed_files[:5]])
            if len(failed_files) > 5:
                failed_list += f"\n... 还有 {len(failed_files) - 5} 个文件"

            messagebox.showwarning(
                "删除部分成功",
                f"成功删除: {len(deleted_files)} 个文件\n"
                f"删除失败: {len(failed_files)} 个文件\n\n"
                f"失败文件:\n{failed_list}\n\n"
                f"详细信息请查看 delete_log.txt"
            )

    def _log_dup_delete_operation(self, deleted_files, failed_files):
        """记录重复文件删除操作到日志"""
        try:
            log_file = "delete_log.txt"
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(f"\n{'='*60}\n")
                f.write(f"重复文件删除操作时间: {timestamp}\n")
                f.write(f"成功删除: {len(deleted_files)} 个重复文件\n")
                f.write(f"删除失败: {len(failed_files)} 个文件\n")
                f.write(f"{'='*60}\n")

                if deleted_files:
                    f.write("\n成功删除的重复文件:\n")
                    for file_path in deleted_files:
                        f.write(f"  - {file_path}\n")

                if failed_files:
                    f.write("\n删除失败的文件:\n")
                    for file_path, error in failed_files:
                        f.write(f"  - {file_path}\n    错误: {error}\n")

                f.write(f"\n{'='*60}\n\n")

        except Exception as e:
            print(f"日志记录失败: {e}")

    def _remove_deleted_dup_from_tree(self, deleted_files):
        """从树视图中移除已删除的重复文件"""
        try:
            deleted_set = set(deleted_files)

            # 遍历所有组
            for group_item in list(self.duplicates_tree.get_children()):
                # 遍历组内的文件
                children_to_delete = []
                for child_item in self.duplicates_tree.get_children(group_item):
                    values = self.duplicates_tree.item(child_item)['values']
                    file_path = values[4]
                    if file_path in deleted_set:
                        children_to_delete.append(child_item)

                # 删除文件项
                for child_item in children_to_delete:
                    self.duplicates_tree.delete(child_item)

                # 如果组内文件少于2个，删除整个组
                if len(self.duplicates_tree.get_children(group_item)) < 2:
                    self.duplicates_tree.delete(group_item)

        except Exception as e:
            print(f"移除文件显示失败: {e}")

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