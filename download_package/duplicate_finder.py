#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
重复文件检测模块
用于扫描和识别磁盘中的重复文件
"""

import os
import hashlib
from pathlib import Path
from collections import defaultdict
from datetime import datetime


class DuplicateFinder:
    """重复文件检测器"""

    def __init__(self):
        self.duplicate_groups = {}  # {hash: [file_info, ...]}
        self.total_scanned = 0
        self.total_duplicates = 0
        self.total_wasted_space = 0

    def calculate_file_hash(self, file_path, hash_algorithm='md5'):
        """
        计算文件的哈希值

        Args:
            file_path: 文件路径
            hash_algorithm: 哈希算法 ('md5' 或 'sha256')

        Returns:
            str: 文件哈希值，失败返回None
        """
        try:
            if hash_algorithm == 'md5':
                hasher = hashlib.md5()
            elif hash_algorithm == 'sha256':
                hasher = hashlib.sha256()
            else:
                hasher = hashlib.md5()

            # 分块读取文件，避免大文件占用过多内存
            with open(file_path, 'rb') as f:
                while True:
                    chunk = f.read(8192)  # 8KB 块
                    if not chunk:
                        break
                    hasher.update(chunk)

            return hasher.hexdigest()

        except (PermissionError, OSError, IOError) as e:
            # 无法读取的文件跳过
            return None

    def find_duplicates(self, root_path, min_size=0, progress_callback=None):
        """
        在指定路径下查找重复文件

        Args:
            root_path: 搜索的根路径
            min_size: 最小文件大小（字节），小于此大小的文件不检查
            progress_callback: 进度回调函数 callback(current_file, phase, progress)

        Returns:
            dict: 重复文件组 {hash: [file_info, ...]}
        """
        root_path = Path(root_path)

        # 第一阶段：按文件大小分组（预筛选）
        size_groups = defaultdict(list)
        file_list = []

        if progress_callback:
            progress_callback("正在扫描文件...", "scan", 0)

        # 递归遍历所有文件
        try:
            for dirpath, dirnames, filenames in os.walk(root_path):
                for filename in filenames:
                    file_path = os.path.join(dirpath, filename)
                    try:
                        file_stat = os.stat(file_path)
                        file_size = file_stat.st_size

                        # 跳过小文件
                        if file_size < min_size:
                            continue

                        # 跳过空文件
                        if file_size == 0:
                            continue

                        file_info = {
                            'path': file_path,
                            'size': file_size,
                            'mtime': file_stat.st_mtime,
                            'name': filename
                        }

                        file_list.append(file_info)
                        size_groups[file_size].append(file_info)

                    except (PermissionError, OSError):
                        continue

        except Exception as e:
            print(f"扫描文件时出错: {e}")

        # 过滤：只保留有多个文件的大小组（可能重复）
        potential_duplicates = {
            size: files for size, files in size_groups.items()
            if len(files) > 1
        }

        if progress_callback:
            progress_callback(
                f"发现 {len(file_list)} 个文件，{sum(len(f) for f in potential_duplicates.values())} 个可能重复",
                "analyze",
                50
            )

        # 第二阶段：计算哈希值
        hash_groups = defaultdict(list)
        processed = 0
        total_to_hash = sum(len(files) for files in potential_duplicates.values())

        for size, files in potential_duplicates.items():
            for file_info in files:
                # 计算哈希值
                file_hash = self.calculate_file_hash(file_info['path'])

                if file_hash:
                    file_info['hash'] = file_hash
                    hash_groups[file_hash].append(file_info)

                processed += 1

                if progress_callback and total_to_hash > 0:
                    progress = 50 + int((processed / total_to_hash) * 50)
                    progress_callback(
                        file_info['name'],
                        "hash",
                        progress
                    )

        # 第三阶段：过滤出真正的重复文件组
        self.duplicate_groups = {
            file_hash: files for file_hash, files in hash_groups.items()
            if len(files) > 1
        }

        # 计算统计信息
        self.total_scanned = len(file_list)
        self.total_duplicates = sum(len(files) for files in self.duplicate_groups.values())

        # 计算浪费的空间（总大小 - 每组保留一个文件）
        self.total_wasted_space = 0
        for files in self.duplicate_groups.values():
            if files:
                file_size = files[0]['size']
                # 浪费空间 = 文件大小 × (重复数量 - 1)
                self.total_wasted_space += file_size * (len(files) - 1)

        if progress_callback:
            progress_callback(
                f"检测完成！找到 {len(self.duplicate_groups)} 组重复文件",
                "complete",
                100
            )

        return self.duplicate_groups

    def get_smart_keep_suggestion(self, duplicate_group):
        """
        智能推荐保留哪个文件

        策略：
        1. 优先保留修改时间最新的
        2. 如果时间相同，保留路径最短的
        3. 如果路径长度相同，保留文件名最短的

        Args:
            duplicate_group: 重复文件列表

        Returns:
            dict: 推荐保留的文件信息
        """
        if not duplicate_group:
            return None

        # 按策略排序
        sorted_files = sorted(
            duplicate_group,
            key=lambda f: (
                -f['mtime'],  # 最新的优先（负号表示降序）
                len(f['path']),  # 路径短的优先
                len(f['name'])  # 文件名短的优先
            )
        )

        return sorted_files[0]

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
        return f"{size:.2f} {size_names[i]}"

    def get_statistics(self):
        """
        获取检测统计信息

        Returns:
            dict: 统计信息
        """
        return {
            'total_scanned': self.total_scanned,
            'total_duplicates': self.total_duplicates,
            'duplicate_groups': len(self.duplicate_groups),
            'wasted_space': self.total_wasted_space,
            'wasted_space_formatted': self.format_size(self.total_wasted_space)
        }


# 测试代码
if __name__ == "__main__":
    finder = DuplicateFinder()

    # 测试路径
    test_path = "E:\\WorkSpaceCode\\LargeFileCleaner\\测试空文件夹"

    print("开始检测重复文件...")
    print("="*60)

    def progress_callback(message, phase, progress):
        print(f"[{phase}] {progress}% - {message}")

    duplicates = finder.find_duplicates(test_path, min_size=0, progress_callback=progress_callback)

    print("\n检测完成！")
    print("="*60)

    stats = finder.get_statistics()
    print(f"扫描文件总数: {stats['total_scanned']}")
    print(f"重复文件数量: {stats['total_duplicates']}")
    print(f"重复文件组数: {stats['duplicate_groups']}")
    print(f"浪费空间: {stats['wasted_space_formatted']}")

    print("\n重复文件详情：")
    print("="*60)

    for i, (file_hash, files) in enumerate(duplicates.items(), 1):
        print(f"\n组 {i} (哈希: {file_hash[:8]}...)")
        print(f"文件大小: {finder.format_size(files[0]['size'])}")
        print(f"重复数量: {len(files)}")

        # 智能推荐
        suggested = finder.get_smart_keep_suggestion(files)
        print(f"建议保留: {suggested['path']}")

        print("所有文件:")
        for file_info in files:
            mtime_str = datetime.fromtimestamp(file_info['mtime']).strftime('%Y-%m-%d %H:%M:%S')
            marker = " ← 推荐保留" if file_info == suggested else ""
            print(f"  - {file_info['path']}")
            print(f"    修改时间: {mtime_str}{marker}")
