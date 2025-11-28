#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件安全检测模块
用于检测文件是否为系统关键文件，防止误删
"""

import os
from pathlib import Path

class FileSafetyChecker:
    """文件安全检查器"""

    def __init__(self):
        # C盘系统关键目录列表（不可删除）
        self.system_dirs = [
            "C:\\Windows",
            "C:\\Program Files",
            "C:\\Program Files (x86)",
            "C:\\ProgramData",
            "C:\\System Volume Information",
            "C:\\$Recycle.Bin",
            "C:\\Recovery",
            "C:\\Boot",
            "C:\\bootmgr",
            "C:\\hiberfil.sys",
            "C:\\pagefile.sys",
            "C:\\swapfile.sys"
        ]

        # 系统文件扩展名（高风险）
        self.system_extensions = [
            ".sys", ".dll", ".exe", ".drv", ".ocx",
            ".cpl", ".scr", ".msi", ".vxd", ".inf"
        ]

        # 用户数据目录（可以删除，但需要特别警告）
        self.user_data_dirs = [
            "Documents", "Desktop", "Pictures",
            "Videos", "Music", "Downloads"
        ]

    def is_system_file(self, file_path):
        """
        检查文件是否为系统文件

        Args:
            file_path: 文件路径（字符串或Path对象）

        Returns:
            tuple: (是否为系统文件, 原因描述)
        """
        try:
            path = Path(file_path)
            path_str = str(path.absolute()).upper()

            # 1. 检查是否在C盘根目录的系统文件
            if path.drive.upper() == "C:":
                # 检查是否在系统关键目录中
                for sys_dir in self.system_dirs:
                    sys_dir_upper = sys_dir.upper()
                    if path_str.startswith(sys_dir_upper):
                        return True, f"位于系统关键目录: {sys_dir}"

                # 检查是否是C盘根目录下的系统文件
                if len(path.parts) == 2:  # C:\ + 文件名
                    return True, "C盘根目录下的文件（可能是系统文件）"

                # 检查C:\Users\<username>下的系统目录
                if "C:\\USERS" in path_str:
                    parts = path.parts
                    # C:\Users\username\AppData 下的文件
                    if "APPDATA" in [p.upper() for p in parts]:
                        return True, "用户AppData目录（包含应用程序数据）"

            # 2. 检查文件扩展名
            ext = path.suffix.lower()
            if ext in self.system_extensions:
                # 如果不在用户目录中的系统扩展名文件，认为是系统文件
                if not self._is_in_user_directory(path):
                    return True, f"系统文件扩展名: {ext}"

            # 3. 检查文件属性（系统属性和隐藏属性）
            if os.path.exists(file_path):
                import stat
                file_stat = os.stat(file_path)
                # Windows系统文件属性检查
                if hasattr(stat, 'FILE_ATTRIBUTE_SYSTEM'):
                    if file_stat.st_file_attributes & stat.FILE_ATTRIBUTE_SYSTEM:
                        return True, "文件具有系统属性标记"

            return False, "可安全删除"

        except Exception as e:
            # 如果检查出错，为安全起见，标记为系统文件
            return True, f"无法验证安全性: {str(e)}"

    def _is_in_user_directory(self, path):
        """检查文件是否在用户目录中"""
        path_str = str(path).upper()
        for user_dir in self.user_data_dirs:
            if user_dir.upper() in path_str:
                return True
        return False

    def is_safe_to_delete(self, file_path):
        """
        判断文件是否可以安全删除

        Args:
            file_path: 文件路径

        Returns:
            bool: True表示可以安全删除，False表示不应删除
        """
        is_system, _ = self.is_system_file(file_path)
        return not is_system

    def get_warning_level(self, file_path):
        """
        获取删除文件的警告级别

        Args:
            file_path: 文件路径

        Returns:
            str: "safe", "warning", "danger"
        """
        is_system, reason = self.is_system_file(file_path)

        if is_system:
            return "danger"

        # 检查是否是重要的用户数据目录
        path_str = str(file_path).upper()
        for user_dir in self.user_data_dirs:
            if user_dir.upper() in path_str:
                return "warning"

        return "safe"

    def batch_check_files(self, file_list):
        """
        批量检查文件列表

        Args:
            file_list: 文件路径列表

        Returns:
            dict: {
                "safe": [可安全删除的文件],
                "warning": [需要警告的文件],
                "danger": [禁止删除的文件]
            }
        """
        result = {
            "safe": [],
            "warning": [],
            "danger": []
        }

        for file_path in file_list:
            level = self.get_warning_level(file_path)
            result[level].append(file_path)

        return result

    def check_folder_safety(self, folder_path):
        """
        检查文件夹是否可以安全删除（递归检查所有子文件和子文件夹）

        Args:
            folder_path: 文件夹路径

        Returns:
            tuple: (是否安全, 危险文件列表, 警告文件列表)
        """
        import os
        from pathlib import Path

        folder_path = Path(folder_path)
        danger_files = []
        warning_files = []

        try:
            # 首先检查文件夹本身是否在系统关键位置
            is_system, reason = self.is_system_file(folder_path)
            if is_system:
                return False, [str(folder_path)], []

            # 递归检查文件夹中的所有文件
            if folder_path.exists() and folder_path.is_dir():
                for root, dirs, files in os.walk(folder_path):
                    # 检查所有文件
                    for file in files:
                        file_path = os.path.join(root, file)
                        level = self.get_warning_level(file_path)

                        if level == "danger":
                            danger_files.append(file_path)
                        elif level == "warning":
                            warning_files.append(file_path)

                    # 检查所有子文件夹
                    for dir_name in dirs:
                        dir_path = os.path.join(root, dir_name)
                        is_system, _ = self.is_system_file(dir_path)
                        if is_system:
                            danger_files.append(dir_path)

            # 如果有危险文件，则不安全
            if danger_files:
                return False, danger_files, warning_files

            # 没有危险文件，但有警告文件
            if warning_files:
                return True, [], warning_files

            # 完全安全
            return True, [], []

        except Exception as e:
            # 出错时保守处理，标记为不安全
            return False, [str(folder_path)], []

    def is_folder_in_system_area(self, folder_path):
        """
        检查文件夹是否位于系统关键区域

        Args:
            folder_path: 文件夹路径

        Returns:
            bool: True表示在系统区域
        """
        is_system, _ = self.is_system_file(folder_path)
        return is_system

    def is_empty_folder(self, folder_path):
        """
        检查文件夹是否为空（包括递归检查子文件夹）

        一个文件夹被认为是空的，如果：
        1. 完全没有文件和子文件夹
        2. 或者只包含空子文件夹（递归判断）

        Args:
            folder_path: 文件夹路径

        Returns:
            bool: True表示文件夹为空
        """
        try:
            folder_path = Path(folder_path)

            if not folder_path.exists() or not folder_path.is_dir():
                return False

            # 获取文件夹中的所有内容
            items = list(folder_path.iterdir())

            # 如果完全为空，返回True
            if not items:
                return True

            # 如果有内容，递归检查每个子项
            for item in items:
                if item.is_file():
                    # 如果有任何文件，就不是空文件夹
                    return False
                elif item.is_dir():
                    # 如果子文件夹不为空，就不是空文件夹
                    if not self.is_empty_folder(item):
                        return False

            # 所有子文件夹都为空，则当前文件夹也为空
            return True

        except (PermissionError, OSError):
            # 无法访问的文件夹，为安全起见认为不是空文件夹
            return False

    def is_safe_to_delete_empty_folder(self, folder_path):
        """
        检查空文件夹是否可以安全删除

        排除以下情况：
        1. 系统关键目录
        2. 应用程序可能需要的特殊目录

        Args:
            folder_path: 文件夹路径

        Returns:
            tuple: (是否可以安全删除, 原因描述)
        """
        try:
            folder_path = Path(folder_path)
            path_str = str(folder_path.absolute()).upper()

            # 检查是否为系统文件夹
            is_system, reason = self.is_system_file(folder_path)
            if is_system:
                return False, reason

            # 检查是否在C盘根目录（特殊保护）
            if folder_path.drive.upper() == "C:" and len(folder_path.parts) <= 2:
                return False, "C盘根目录下的文件夹，为安全起见不建议删除"

            # 一些特殊的可能需要的空目录（应用程序可能会检查这些目录的存在）
            special_folder_names = [
                ".GIT", ".SVN", ".HG",  # 版本控制系统
                "NODE_MODULES", "__PYCACHE__",  # 包管理器
                "APPDATA", "PROGRAMDATA"  # 系统数据目录
            ]

            folder_name_upper = folder_path.name.upper()
            if folder_name_upper in special_folder_names:
                return False, f"特殊目录 ({folder_path.name})，可能被应用程序使用"

            # 检查父目录是否为系统目录
            for parent in folder_path.parents:
                parent_str = str(parent).upper()
                for sys_dir in self.system_dirs:
                    if parent_str == sys_dir.upper():
                        return False, f"位于系统目录 {sys_dir} 下，不建议删除"

            return True, "可以安全删除"

        except Exception as e:
            return False, f"无法验证安全性: {str(e)}"

    def find_empty_folders(self, root_path, progress_callback=None):
        """
        在指定路径下查找所有空文件夹

        Args:
            root_path: 搜索的根路径
            progress_callback: 进度回调函数 callback(current_path)

        Returns:
            list: [(文件夹路径, 是否可安全删除, 原因), ...]
        """
        empty_folders = []

        try:
            root_path = Path(root_path)

            if not root_path.exists() or not root_path.is_dir():
                return empty_folders

            # 使用 os.walk 遍历所有子目录，从底部开始（bottomup=True）
            # 这样可以先处理最深层的文件夹
            for dirpath, dirnames, filenames in os.walk(root_path, topdown=False):
                folder = Path(dirpath)

                # 进度回调
                if progress_callback:
                    progress_callback(str(folder))

                # 检查是否为空文件夹
                if self.is_empty_folder(folder):
                    # 检查是否可以安全删除
                    is_safe, reason = self.is_safe_to_delete_empty_folder(folder)
                    empty_folders.append((str(folder), is_safe, reason))

            return empty_folders

        except Exception as e:
            print(f"查找空文件夹时出错: {e}")
            return empty_folders


# 测试代码
if __name__ == "__main__":
    checker = FileSafetyChecker()

    # 测试一些路径
    test_paths = [
        "C:\\Windows\\System32\\kernel32.dll",
        "C:\\Users\\test\\Downloads\\movie.mp4",
        "C:\\Users\\test\\Desktop\\document.pdf",
        "D:\\Data\\large_file.zip",
        "C:\\Program Files\\app\\program.exe"
    ]

    print("文件安全性检测测试：")
    print("="*60)

    for path in test_paths:
        is_system, reason = checker.is_system_file(path)
        level = checker.get_warning_level(path)

        print(f"\n路径: {path}")
        print(f"系统文件: {is_system}")
        print(f"原因: {reason}")
        print(f"警告级别: {level}")
        print("-"*60)
