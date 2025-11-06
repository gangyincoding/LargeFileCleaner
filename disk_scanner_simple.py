#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
磁盘扫描器 - BAT界面版本（简化版）
直接在命令行中显示扫描进度和结果
"""

import os
import sys
import time
from pathlib import Path
from collections import defaultdict
from datetime import datetime

class DiskScanner:
    def __init__(self):
        self.total_files = 0
        self.total_size = 0
        self.largest_files = []
        self.file_types = defaultdict(lambda: {'count': 0, 'size': 0})
        self.scanned_files = 0
        self.start_time = 0

        # 文件类型过滤器
        self.file_type_filter = None
        self.file_type_mapping = self._create_file_type_mapping()

        # 进度回调函数
        self.progress_callback = None

    def format_size(self, size_bytes):
        """格式化文件大小"""
        if size_bytes == 0:
            return "0 B"

        size_names = ["B", "KB", "MB", "GB", "TB"]
        i = 0
        size = float(size_bytes)

        while size >= 1024.0 and i < len(size_names) - 1:
            size /= 1024.0
            i += 1

        return f"{size:.1f} {size_names[i]}"

    def _create_file_type_mapping(self):
        """创建文件类型映射"""
        return {
            "文档文件": [".txt", ".doc", ".docx", ".pdf", ".rtf", ".odt"],
            "图片文件": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp", ".svg"],
            "视频文件": [".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".webm", ".m4v"],
            "音频文件": [".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a"],
            "压缩文件": [".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz"],
            "程序文件": [".exe", ".dll", ".msi", ".app", ".deb", ".rpm"],
            "代码文件": [".py", ".js", ".html", ".css", ".java", ".cpp", ".c", ".php", ".rb", ".go"],
            "光盘镜像": [".iso", ".img", ".dmg", ".vhd"],
            "种子文件": [".torrent", ".magnet"],
            "设计文件": [".psd", ".ai", ".sketch", ".fig"],
            "电子书": [".epub", ".mobi", ".azw", ".azw3"]
        }

    def set_file_type_filter(self, selected_types):
        """设置文件类型过滤器
        Args:
            selected_types: list of selected type names, or None for all files
        """
        self.file_type_filter = selected_types

    def set_progress_callback(self, callback):
        """设置进度回调函数
        Args:
            callback: 回调函数，接受参数 (progress, scanned_files, total_files)
        """
        self.progress_callback = callback

    def should_scan_file(self, file_path):
        """判断文件是否应该被扫描
        Args:
            file_path: Path object of the file
        Returns:
            bool: True if file should be scanned
        """
        # 如果没有设置过滤器，扫描所有文件
        if self.file_type_filter is None or "全部文件" in self.file_type_filter:
            return True

        file_ext = file_path.suffix.lower()

        # 检查是否属于选中的文件类型
        for type_name, extensions in self.file_type_mapping.items():
            if type_name in self.file_type_filter:
                if file_ext in extensions:
                    return True

        # 如果选择了"其他文件"类型，处理未知扩展名
        if "其他文件" in self.file_type_filter:
            known_extensions = set()
            for extensions in self.file_type_mapping.values():
                known_extensions.update(extensions)
            if file_ext not in known_extensions:
                return True

        return False

    def get_available_file_types(self):
        """获取所有可用的文件类型"""
        return ["全部文件"] + list(self.file_type_mapping.keys()) + ["其他文件"]

    def get_file_type(self, file_path):
        """获取文件类型"""
        suffix = file_path.suffix.lower()
        if not suffix:
            return "无扩展名"

        type_mapping = {
            '.txt': '文本文件', '.doc': 'Word文档', '.docx': 'Word文档',
            '.pdf': 'PDF文档', '.jpg': '图片', '.jpeg': '图片', '.png': '图片',
            '.gif': '图片', '.bmp': '图片', '.mp4': '视频', '.avi': '视频',
            '.mkv': '视频', '.mov': '视频', '.mp3': '音频', '.wav': '音频',
            '.flac': '音频', '.zip': '压缩文件', '.rar': '压缩文件', '.7z': '压缩文件',
            '.exe': '程序文件', '.dll': '程序文件', '.msi': '安装包',
            '.py': '代码文件', '.js': '代码文件', '.html': '网页文件',
            '.css': '样式文件', '.iso': '光盘镜像', '.torrent': '种子文件'
        }

        return type_mapping.get(suffix, f'其他文件({suffix})')

    def scan_directory(self, directory_path, min_file_size_kb=1, max_files=100, include_hidden=False):
        """扫描目录"""
        self.start_time = time.time()
        min_file_size = min_file_size_kb * 1024

        print(f"[*] 开始扫描目录: {directory_path}")
        print(f"[配置] 最小文件大小: {min_file_size_kb} KB")
        print(f"[配置] 最大显示文件数: {max_files}")
        print(f"[配置] 包含隐藏文件: {'是' if include_hidden else '否'}")
        print("-" * 60)

        try:
            path = Path(directory_path)
            if not path.exists() or not path.is_dir():
                print(f"[错误] 目录不存在或无效 - {directory_path}")
                return False

            # 首先计算总文件数（用于进度显示）
            total_files_estimate = 0
            for root, dirs, files in os.walk(directory_path):
                total_files_estimate += len(files)

            print(f"[信息] 预估文件总数: {total_files_estimate:,}")
            print("-" * 60)

            # 开始扫描
            file_sizes = []

            for root, dirs, files in os.walk(directory_path):
                for file in files:
                    try:
                        file_path = Path(root) / file

                        # 检查隐藏文件
                        if not include_hidden and file.startswith('.'):
                            continue

                        # 检查文件类型过滤器
                        if not self.should_scan_file(file_path):
                            continue

                        if file_path.is_file():
                            file_size = file_path.stat().st_size

                            # 只统计大于指定大小的文件
                            if file_size >= min_file_size:
                                file_sizes.append((file_path, file_size))

                                # 按类型统计
                                file_type = self.get_file_type(file_path)
                                self.file_types[file_type]['count'] += 1
                                self.file_types[file_type]['size'] += file_size

                                self.total_files += 1
                                self.total_size += file_size

                            # 显示进度
                            self.scanned_files += 1

                            # 更频繁地更新进度（每10个文件或每1%进度）
                            current_progress = (self.scanned_files / total_files_estimate * 100) if total_files_estimate > 0 else 0

                            # 更新条件：每10个文件 或 进度变化达到1%
                            if (self.scanned_files % 10 == 0) or (current_progress >= int(self.scanned_files / total_files_estimate * 100 + 1) if total_files_estimate > 0 else 0):
                                progress = min(current_progress, 100)  # 确保不超过100%
                                print(f"[进度] {progress:.1f}% ({self.scanned_files:,}/{total_files_estimate:,})")

                                # 调用GUI进度回调
                                if self.progress_callback:
                                    try:
                                        self.progress_callback(progress, self.scanned_files, total_files_estimate)
                                    except:
                                        pass  # 忽略回调错误，不影响扫描

                    except (OSError, PermissionError) as e:
                        continue
                    except Exception as e:
                        continue

            # 排序并获取最大的文件
            file_sizes.sort(key=lambda x: x[1], reverse=True)
            self.largest_files = file_sizes[:max_files]

            # 确保最终进度是100%
            if self.progress_callback:
                try:
                    self.progress_callback(100, self.scanned_files, total_files_estimate)
                except:
                    pass

            return True

        except Exception as e:
            print(f"[错误] 扫描过程中发生严重错误: {e}")
            return False

    def display_results(self, scan_path, min_file_size_kb, max_files, include_hidden):
        """显示扫描结果"""
        scan_time = time.time() - self.start_time

        print("\n" + "=" * 70)
        print("扫描完成！")
        print("=" * 70)

        # 基本信息
        print(f"\n[统计] 扫描结果:")
        print(f"   扫描路径: {scan_path}")
        print(f"   扫描文件数: {self.scanned_files:,}")
        print(f"   符合条件文件: {self.total_files:,}")
        print(f"   总大小: {self.format_size(self.total_size)}")
        print(f"   扫描耗时: {scan_time:.2f} 秒")

        # 文件类型统计
        if self.file_types:
            print(f"\n[类型] 文件类型分布 (按大小排序):")
            print("-" * 70)
            print(f"{'类型':<15} {'数量':<8} {'大小':<12} {'占比':<8}")
            print("-" * 70)

            # 按大小排序
            sorted_types = sorted(self.file_types.items(), key=lambda x: x[1]['size'], reverse=True)

            for file_type, stats in sorted_types[:10]:  # 显示前10种类型
                percentage = (stats['size'] / self.total_size * 100) if self.total_size > 0 else 0
                print(f"{file_type:<15} {stats['count']:<8} {self.format_size(stats['size']):<12} {percentage:>5.1f}%")

        # 最大的文件
        if self.largest_files:
            print(f"\n[排行] 最大的文件 (前{min(len(self.largest_files), max_files)}个):")
            print("-" * 100)
            print(f"{'排名':<4} {'文件名':<40} {'大小':<12} {'路径':<43}")
            print("-" * 100)

            for i, (file_path, file_size) in enumerate(self.largest_files, 1):
                name = file_path.name
                if len(name) > 38:
                    name = name[:35] + "..."

                path = str(file_path.parent)
                if len(path) > 40:
                    path = "..." + path[-40:]

                print(f"{i:<4} {name:<40} {self.format_size(file_size):<12} {path}")

        # 清理建议
        print(f"\n[建议] 清理建议:")
        if self.total_size == 0:
            print("   没有找到符合条件的文件，建议降低最小文件大小阈值。")
        else:
            print("   1. 检查视频文件和压缩包，通常占用空间最大")
            print("   2. 查看重复文件和临时文件")
            print("   3. 清理不需要的程序安装包")
            print("   4. 重要文件删除前请先备份")

        print("\n" + "=" * 70)

    def save_results(self, scan_path, min_file_size_kb, max_files, include_hidden):
        """保存结果到文件"""
        try:
            with open('scan_results.txt', 'w', encoding='utf-8') as f:
                scan_time = time.time() - self.start_time

                f.write("磁盘空间分析报告\n")
                f.write("=" * 50 + "\n")
                f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"扫描路径: {scan_path}\n")
                f.write(f"扫描设置: 最小{min_file_size_kb}KB, 最多{max_files}个文件, 包含隐藏文件{'是' if include_hidden else '否'}\n")
                f.write(f"扫描耗时: {scan_time:.2f} 秒\n")
                f.write(f"符合条件的文件数: {self.total_files:,}\n")
                f.write(f"总大小: {self.format_size(self.total_size)}\n\n")

                if self.file_types:
                    f.write("文件类型统计:\n")
                    f.write("-" * 40 + "\n")
                    sorted_types = sorted(self.file_types.items(), key=lambda x: x[1]['size'], reverse=True)

                    for file_type, stats in sorted_types:
                        percentage = (stats['size'] / self.total_size * 100) if self.total_size > 0 else 0
                        f.write(f"{file_type}: {stats['count']}个文件, {self.format_size(stats['size'])} ({percentage:.1f}%)\n")

                if self.largest_files:
                    f.write(f"\n最大的文件 (前{len(self.largest_files)}个):\n")
                    f.write("-" * 60 + "\n")

                    for i, (file_path, file_size) in enumerate(self.largest_files, 1):
                        f.write(f"{i}. {file_path.name} - {self.format_size(file_size)}\n")
                        f.write(f"   路径: {file_path}\n\n")

            print(f"[成功] 结果已保存到: scan_results.txt")
            return True

        except Exception as e:
            print(f"[错误] 保存结果时出错: {e}")
            return False

def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("[错误] 请提供扫描路径")
        return

    # 解析命令行参数
    scan_path = sys.argv[1]
    min_file_size_kb = int(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[2].isdigit() else 1
    max_files = int(sys.argv[3]) if len(sys.argv) > 3 and sys.argv[3].isdigit() else 100
    include_hidden = sys.argv[4].lower() == 'y' if len(sys.argv) > 4 else False

    # 创建扫描器并开始扫描
    scanner = DiskScanner()

    if scanner.scan_directory(scan_path, min_file_size_kb, max_files, include_hidden):
        scanner.display_results(scan_path, min_file_size_kb, max_files, include_hidden)
        scanner.save_results(scan_path, min_file_size_kb, max_files, include_hidden)
    else:
        print("[错误] 扫描失败！")

if __name__ == "__main__":
    main()