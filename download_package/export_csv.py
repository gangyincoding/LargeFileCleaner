#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将扫描结果导出为CSV格式，支持Excel自动打开
"""

import csv
import os
import subprocess
from datetime import datetime

def try_open_excel_with_csv(csv_file):
    """尝试用Excel打开CSV文件，支持多种方式"""
    try:
        # 方法1：尝试使用Excel直接打开
        try:
            os.startfile(csv_file)
            return True
        except:
            pass

        # 方法2：尝试通过Excel可执行文件路径打开
        excel_paths = [
            r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE",
            r"C:\Program Files\Microsoft Office\Office15\EXCEL.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE",
            r"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE",
        ]

        for excel_path in excel_paths:
            if os.path.exists(excel_path):
                try:
                    subprocess.Popen([excel_path, csv_file])
                    return True
                except:
                    pass

        # 方法3：尝试使用默认程序打开
        try:
            os.system(f'start "" "{csv_file}"')
            return True
        except:
            pass

        # 方法4：尝试使用WPS或其他办公软件
        wps_paths = [
            r"C:\Program Files (x86)\Kingsoft\WPS Office\ksolaunch.exe",
            r"C:\Program Files\Kingsoft\WPS Office\ksolaunch.exe",
        ]

        for wps_path in wps_paths:
            if os.path.exists(wps_path):
                try:
                    subprocess.Popen([wps_path, csv_file])
                    return True
                except:
                    pass

        return False
    except Exception as e:
        print(f"[ERROR] 尝试打开Excel时出错：{e}")
        return False


def export_to_csv(auto_open_excel=True):
    """
    导出扫描结果为CSV
    :param auto_open_excel: 是否自动用Excel打开CSV文件
    :return: tuple (是否成功, CSV文件路径)
    """
    try:
        # 读取扫描结果
        with open('scan_results.txt', 'r', encoding='utf-8') as f:
            content = f.read()

        # 解析扫描结果
        lines = content.split('\n')
        scan_path = ""
        scan_time = ""
        total_files = 0
        total_size = ""

        for line in lines:
            if "扫描路径:" in line:
                scan_path = line.split(":", 1)[1].strip()
            elif "扫描耗时:" in line:
                scan_time = line.split(":", 1)[1].strip()
            elif "符合条件的文件数:" in line:
                total_files = line.split(":", 1)[1].strip().replace(",", "")
            elif "总大小:" in line:
                total_size = line.split(":", 1)[1].strip()

        # 创建CSV文件
        csv_filename = f"磁盘分析报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

        with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)

            # 写入标题和基本信息
            writer.writerow(['磁盘空间分析报告'])
            writer.writerow([f'生成时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'])
            writer.writerow([f'扫描路径: {scan_path}'])
            writer.writerow([f'扫描耗时: {scan_time}'])
            writer.writerow([f'总文件数: {total_files}'])
            writer.writerow([f'总大小: {total_size}'])
            writer.writerow([])  # 空行

            # 写入文件类型统计
            writer.writerow(['文件类型统计'])
            writer.writerow(['文件类型', '文件数量', '占用大小', '占比'])

            # 简单解析文件类型信息
            in_type_section = False
            for line in lines:
                if "文件类型统计:" in line:
                    in_type_section = True
                    continue
                elif "最大的文件" in line:
                    in_type_section = False
                    continue

                if in_type_section and line.strip() and ":" in line:
                    try:
                        parts = line.split(":")
                        if len(parts) >= 2:
                            file_type = parts[0].strip()
                            details = parts[1].strip()
                            if "个文件" in details:
                                count_part = details.split("个文件")[0].strip()
                                size_part = details.split("个文件")[1].strip()
                                size = size_part.split("(")[0].strip()
                                percentage = size_part.split("(")[1].replace(")", "").strip() if "(" in size_part else ""

                                writer.writerow([file_type, count_part, size, percentage])
                    except:
                        continue

            writer.writerow([])  # 空行

            # 写入最大文件列表
            writer.writerow(['最大文件列表'])
            writer.writerow(['排名', '文件名', '大小', '路径'])

            in_file_section = False
            file_rank = 1
            for line in lines:
                if "最大的文件" in line:
                    in_file_section = True
                    continue
                elif "=" in line and in_file_section:
                    in_file_section = False
                    continue

                if in_file_section and line.strip() and (line.strip().endswith(".txt") or
                    line.strip().endswith(".exe") or line.strip().endswith(".mp4") or
                    line.strip().endswith(".zip") or line.strip().endswith(".pdf") or
                    "." in line.strip()):
                    try:
                        # 尝试解析文件信息行
                        if ". " in line and " - " in line:
                            parts = line.split(". ", 1)
                            if len(parts) == 2:
                                file_info = parts[1].split(" - ")
                                if len(file_info) == 2:
                                    filename = file_info[0].strip()
                                    filesize = file_info[1].strip()
                                    writer.writerow([file_rank, filename, filesize, ""])
                                    file_rank += 1
                    except:
                        continue

        print(f"[SUCCESS] CSV文件已导出: {csv_filename}")

        # 尝试用Excel打开CSV文件
        if auto_open_excel:
            print("[INFO] 正在尝试用Excel打开CSV文件...")
            if try_open_excel_with_csv(csv_filename):
                print("[SUCCESS] Excel已自动打开CSV文件")
            else:
                print("[WARNING] 无法自动打开Excel，请手动打开CSV文件")
                print(f"[INFO] 文件位置: {os.path.abspath(csv_filename)}")

        return True, csv_filename

    except Exception as e:
        print(f"[ERROR] 导出CSV失败: {e}")
        return False, ""

if __name__ == "__main__":
    export_to_csv()