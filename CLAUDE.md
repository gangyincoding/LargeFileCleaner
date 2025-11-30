# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

这是一个专业的磁盘空间分析工具（Disk Space Analyzer），基于 Python 和 Tkinter 开发，提供图形化界面用于快速扫描和识别占用磁盘空间的大文件，支持文件删除、文件夹批量删除、空文件夹清理和重复文件检测。

**主要语言**: Python 3.6+
**目标平台**: Windows 7/8/10/11
**部署方式**: 绿色软件，无需安装
**当前版本**: v1.4

## 核心架构

### 模块划分

项目采用模块化架构，核心组件分为：

1. **GUI 层** (`disk_analyzer_gui_stable.py`)
   - 基于 Tkinter 的图形界面
   - 负责用户交互、进度显示和结果展示
   - 使用多线程防止界面阻塞
   - 包含演示模式（demo mode）作为后备方案

2. **扫描引擎** (`disk_scanner_simple.py`)
   - `DiskScanner` 类：核心扫描逻辑
   - 支持文件类型过滤（13种类型分类）
   - 提供进度回调机制
   - 包含文件大小格式化工具函数

3. **导出模块**
   - `export_excel.py`: Excel 格式导出（依赖 openpyxl，支持降级模式）
   - `export_csv.py`: CSV 格式导出
   - 所有导出模块都支持在依赖缺失时降级处理

4. **安全检查模块** (`file_safety.py`) ⭐v1.1+
   - `FileSafetyChecker` 类：文件和文件夹安全性检查
   - 系统文件识别和保护
   - 文件夹递归安全检查
   - 空文件夹检测（v1.3+）
   - 支持三级安全警告（safe/warning/danger）

5. **重复文件检测模块** (`duplicate_finder.py`) ⭐v1.4
   - `DuplicateFinder` 类：重复文件检测引擎
   - MD5哈希算法实现
   - 两阶段检测算法（大小分组 + 哈希计算）
   - 智能保留建议
   - 统计信息计算

### 核心功能模块

#### 1. 大文件扫描（v1.0）
- 递归扫描指定目录
- 按大小排序文件
- 文件类型过滤
- 多线程扫描

#### 2. 文件删除（v1.1）
- 单文件选择和删除
- 系统文件保护
- 回收站支持
- 删除日志记录

#### 3. 文件夹批量删除（v1.2）
- 按文件夹分组显示
- 文件夹统计信息
- 递归安全检查
- 批量删除整个文件夹

#### 4. 空文件夹清理（v1.3）
- 独立扫描空文件夹
- 递归空文件夹检测
- 智能安全过滤
- 批量清理功能
- **空文件夹定义**：
  - 完全空：没有任何文件和子文件夹
  - 递归空：只包含空子文件夹
- **自动保护目录**：
  - 系统关键目录（Windows、Program Files等）
  - 版本控制目录（.git、.svn等）
  - 包管理器目录（node_modules、__pycache__等）
  - C盘根目录下的文件夹

#### 5. 重复文件检测（v1.4） ⭐最新
- MD5哈希精确识别
- 两阶段检测算法优化性能
- 智能保留建议
- 树形分组展示
- **检测算法**：
  - 阶段1：按文件大小分组（快速预筛选）
  - 阶段2：计算MD5哈希值（精确识别）
- **智能保留策略**（按优先级）：
  1. 修改时间最新的文件
  2. 路径最短的文件
  3. 文件名最短的文件
- **性能优化**：
  - 分块读取（8KB块）避免内存问题
  - 只对大小相同的文件计算哈希
  - 支持设置最小文件大小过滤

### 依赖处理策略

项目采用**优雅降级**设计模式：
- 所有非标准库依赖（如 openpyxl）都通过 try-except 包裹
- 依赖缺失时自动切换到演示模式或兼容格式
- GUI 在扫描器不可用时使用内置的简化版本

### 文件类型过滤系统

扫描引擎支持 13 种文件类型分类：
- 文档文件、图片文件、视频文件、音频文件
- 压缩文件、程序文件、代码文件、系统文件
- 临时文件、备份文件、数据库文件、配置文件、其他文件

过滤器通过扩展名映射实现，支持多选和全选/清除操作。

## 开发命令

### 启动应用
```bash
# 推荐方式（通过 BAT 启动脚本）
DiskCleaner_GUI_Stable.bat

# 直接启动 Python GUI
python disk_analyzer_gui_stable.py
```

### 项目结构
```
LargeFileCleaner/
├── disk_analyzer_gui_stable.py   # GUI 主程序（v1.4 - 包含重复文件检测）
├── disk_scanner_simple.py        # 扫描引擎
├── file_safety.py                # 安全检查模块（v1.1+ - 文件/文件夹/空文件夹检查）
├── duplicate_finder.py           # 重复文件检测模块（v1.4）
├── export_excel.py               # Excel 导出
├── export_csv.py                 # CSV 导出
├── DiskCleaner_GUI_Stable.bat    # Windows 启动脚本
├── README.md                     # 项目说明文档
├── CLAUDE.md                     # 开发指南（本文件）
├── AGENTS.md                     # AI Agent开发指南
├── CHANGELOG.md                  # 版本更新历史
├── LICENSE                       # MIT开源协议
├── requirements.txt              # Python依赖列表
├── 使用说明.txt                  # 快速开始指南
├── 软件信息.txt                  # 软件详细信息
├── index.html                    # GitHub Pages项目主页
└── docs/                         # 开发文档目录
    ├── README.md
    └── development-reports/      # 开发报告文档
```

## 重要技术细节

### 进度条设计
- **独占一行布局**：进度条占据整行，百分比文字居中显示
- 使用 `ttk.Progressbar` 组件，长度 500px
- 通过专用 `Frame` 和 `columnconfigure` 实现自适应宽度
- 详见 `进度条布局优化说明.md`

### 多线程处理
- 扫描操作在独立线程中执行
- GUI 更新通过 `root.after()` 确保线程安全
- 使用标志位 `is_scanning` 控制扫描状态

### 编码规范
- 所有 Python 文件使用 UTF-8 编码（`# -*- coding: utf-8 -*-`）
- 完善支持中文文件名和路径
- 导出文件统一使用 UTF-8 编码

### 错误处理
- 所有 I/O 操作都包含异常处理
- BAT 脚本包含 Python 环境检测
- GUI 包含完整的错误消息提示

## 测试和验证

由于项目特性，测试主要通过手动方式进行：

1. **启动测试**: 通过 BAT 脚本验证 Python 环境检测
2. **扫描测试**: 测试不同大小目录的扫描性能
3. **导出测试**: 验证三种导出格式（Excel/CSV/HTML）
4. **依赖降级测试**: 在无 openpyxl 环境下测试降级行为

## 发布流程

该项目使用 GitHub Release 和 GitHub Pages 进行发布：
- `index.html` 作为项目介绍页面（GitHub Pages）
- 通过 GitHub Release 发布ZIP包（用户版和完整版）
- 支持从 Release 一键下载

打包发布时确保：
1. 使用 PyInstaller 打包 EXE 文件
2. 创建用户版ZIP包（包含EXE和文档）
3. 创建完整版ZIP包（包含用户版+源代码）
4. 上传ZIP包到 GitHub Release
5. 更新 `index.html` 中的下载链接
6. 更新版本号和日期信息

## Git 工作流

主要分支：`main`
提交规范：
- `feat:` 新功能
- `fix:` 错误修复
- `docs:` 文档更新
- `chore:` 杂项工作（如清理缓存）

## 注意事项

1. **不要添加敏感路径**：扫描结果可能包含用户隐私路径，注意 `.gitignore` 配置
2. **保持向后兼容**：工具面向普通用户，需要支持较低版本 Python (3.6+)
3. **依赖最小化**：尽量使用标准库，可选依赖需要降级处理
4. **Windows 路径处理**：使用 `pathlib.Path` 处理路径兼容性
