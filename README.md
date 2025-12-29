# PPT 转 PDF 合并工具 / PPT to PDF Merger Tool

一个功能强大的 PowerPoint 文件合并工具，支持将多个 PPT/PPTX 文件合并为一个文件，并可转换为 PDF 格式。

A powerful PowerPoint file merger tool that supports merging multiple PPT/PPTX files into one file, with the option to convert to PDF format.

---

## 功能特性 / Features

### 核心功能 / Core Features

- **PPT 文件合并** / **PPT File Merging**
  - 支持将多个 PPT/PPTX 文件合并为一个文件
  - 自动生成目录页，显示每个文件的页数和起始页码
  - 支持拖拽排序，自定义合并顺序
  - Supports merging multiple PPT/PPTX files into one file
  - Automatically generates a table of contents page showing page count and starting page number for each file
  - Supports drag-and-drop sorting to customize merge order

- **PDF 转换与合并** / **PDF Conversion and Merging**
  - 将 PPT 文件转换为 PDF 并合并（Windows 平台）
  - 支持两种预设模式："博士组会" 和 "大模型和开放世界组组会"
  - 自动生成带目录的合并 PDF 文件
  - Convert PPT files to PDF and merge them (Windows platform)
  - Supports two preset modes: "博士组会" and "大模型和开放世界组组会"
  - Automatically generates merged PDF files with table of contents

- **用户界面** / **User Interface**
  - 直观的图形界面，操作简单
  - 自动保存上次选择的目录
  - 支持批量选择和添加文件
  - Intuitive graphical interface with simple operations
  - Automatically saves the last selected directory
  - Supports batch selection and file addition

---

## 系统要求 / System Requirements

### 操作系统 / Operating System
- **Windows**: Windows 7 或更高版本（需要安装 Microsoft PowerPoint）
- **macOS**: macOS 10.12 或更高版本
- **Windows**: Windows 7 or later (requires Microsoft PowerPoint installation)
- **macOS**: macOS 10.12 or later

### Python 环境 / Python Environment
- Python 3.7 或更高版本
- Python 3.7 or later

---

## 安装说明 / Installation

### 1. 克隆或下载项目 / Clone or Download the Project

```bash
# 如果使用 Git
git clone <repository-url>
cd -PPT-

# 或直接下载 ZIP 文件并解压
# Or download ZIP file and extract
```

### 2. 安装依赖库 / Install Dependencies

#### Windows 平台 / Windows Platform

```bash
# 基础依赖
pip install PyPDF2 reportlab

# PPT 合并功能（需要 PowerPoint）
pip install pywin32

# 可选：美化界面
pip install ttkbootstrap
```

#### macOS 平台 / macOS Platform

```bash
# 基础依赖
pip install PyPDF2 reportlab

# PPT 合并功能
pip install python-pptx

# 可选：美化界面
pip install ttkbootstrap
```

#### 使用 Conda / Using Conda

```bash
# 激活 conda 环境
conda activate base

# 安装依赖
pip install PyPDF2 reportlab python-pptx ttkbootstrap
```

### 3. Windows 平台额外配置 / Additional Windows Configuration

如果使用 PDF 转换功能，需要确保存在 `单个ppt转为pdf.vbs` 脚本文件（该文件需要单独提供）。

If using PDF conversion functionality, ensure the `单个ppt转为pdf.vbs` script file exists (this file needs to be provided separately).

---

## 使用方法 / Usage

### Windows 平台 / Windows Platform

#### 方法 1：直接运行 Python 脚本 / Method 1: Run Python Script Directly

```bash
python ppt_pdf_merger.py
```

#### 方法 2：使用批处理文件（如果存在）/ Method 2: Use Batch File (if available)

双击 `启动PPT合并工具.bat`（如果存在）

Double-click `启动PPT合并工具.bat` (if available)

### macOS 平台 / macOS Platform

#### 方法 1：使用启动脚本（推荐）/ Method 1: Use Launch Script (Recommended)

双击 `mac 下启动PPT合并工具.command` 文件

Double-click the `mac 下启动PPT合并工具.command` file

#### 方法 2：命令行运行 / Method 2: Command Line

```bash
# 确保在项目目录中
cd /path/to/-PPT-

# 使用 conda（推荐）
conda activate base
python ppt_pdf_merger.py

# 或直接使用系统 Python
python3 ppt_pdf_merger.py
```

---

## 操作指南 / User Guide

### 基本操作流程 / Basic Workflow

1. **选择目录** / **Select Directory**
   - 点击 "选择目录" 按钮，选择包含 PPT 文件的文件夹
   - Click the "选择目录" button to select the folder containing PPT files

2. **选择文件** / **Select Files**
   - 在左侧 "可选 PPT 文件" 列表中选择要合并的文件
   - 点击 "添加 →" 按钮将文件添加到右侧列表
   - 或点击 "全选 →" 添加所有文件
   - Select files from the left "可选 PPT 文件" list
   - Click "添加 →" to add files to the right list
   - Or click "全选 →" to add all files

3. **调整顺序** / **Adjust Order**
   - 在右侧 "已选 PPT 文件" 列表中，可以拖拽文件调整合并顺序
   - Drag and drop files in the right "已选 PPT 文件" list to adjust merge order

4. **执行合并** / **Execute Merge**
   - **合并为 PPT**: 将选中的 PPT 文件合并为一个 PPTX 文件
   - **博士组会** / **大模型和开放世界组组会**: 将 PPT 转换为 PDF 并合并（Windows 平台）
   - **Merge as PPT**: Merge selected PPT files into one PPTX file
   - **博士组会** / **大模型和开放世界组组会**: Convert PPT to PDF and merge (Windows platform)

### 功能说明 / Feature Details

#### 合并为 PPT / Merge as PPT
- 将所有选中的 PPT 文件合并为一个 PPTX 文件
- 自动在文件开头插入目录页
- 目录页显示每个文件的名称、页数和起始页码
- 输出文件名格式：`YYYYMMDD合并PPT.pptx`
- Merges all selected PPT files into one PPTX file
- Automatically inserts a table of contents page at the beginning
- The table of contents shows each file's name, page count, and starting page number
- Output file name format: `YYYYMMDD合并PPT.pptx`

#### PDF 转换与合并 / PDF Conversion and Merging
- 先将每个 PPT 文件转换为 PDF（使用 VBS 脚本，仅 Windows）
- 然后合并所有 PDF 文件
- 在合并后的 PDF 开头添加目录页
- 输出文件名格式：`YYYYMMDD[模式名称].pdf`
- First converts each PPT file to PDF (using VBS script, Windows only)
- Then merges all PDF files
- Adds a table of contents page at the beginning of the merged PDF
- Output file name format: `YYYYMMDD[Mode Name].pdf`

---

## 项目结构 / Project Structure

```
-PPT-/
├── ppt_pdf_merger.py          # 主程序文件 / Main program file
├── mac 下启动PPT合并工具.command  # macOS 启动脚本 / macOS launch script
├── ppt_merger_settings.json   # 配置文件（自动生成）/ Config file (auto-generated)
├── README.md                  # 说明文档 / Documentation
└── 单个ppt转为pdf.vbs         # PPT 转 PDF 脚本（Windows，需单独提供）/ PPT to PDF script (Windows, needs to be provided separately)
```

---

## 依赖库说明 / Dependencies

### 必需依赖 / Required Dependencies

- **tkinter**: Python 标准库，用于 GUI 界面
- **PyPDF2**: PDF 文件处理
- **reportlab**: PDF 生成（用于创建目录页）
- **tkinter**: Python standard library for GUI interface
- **PyPDF2**: PDF file processing
- **reportlab**: PDF generation (for creating table of contents)

### 平台特定依赖 / Platform-Specific Dependencies

- **Windows**: `pywin32` - 用于 PowerPoint COM 接口
- **macOS**: `python-pptx` - 用于 PPT 文件处理
- **Windows**: `pywin32` - For PowerPoint COM interface
- **macOS**: `python-pptx` - For PPT file processing

### 可选依赖 / Optional Dependencies

- **ttkbootstrap**: 美化界面样式（如果未安装，将使用默认样式）
- **ttkbootstrap**: Beautify interface styles (if not installed, default styles will be used)

---

## 注意事项 / Notes

### Windows 平台 / Windows Platform

1. **PowerPoint 要求** / **PowerPoint Requirements**
   - 合并 PPT 功能需要安装 Microsoft PowerPoint
   - PPT merging functionality requires Microsoft PowerPoint installation

2. **PDF 转换** / **PDF Conversion**
   - PDF 转换功能需要 `单个ppt转为pdf.vbs` 脚本文件
   - 该脚本使用 PowerPoint COM 接口进行转换
   - PDF conversion requires the `单个ppt转为pdf.vbs` script file
   - The script uses PowerPoint COM interface for conversion

3. **文件路径** / **File Paths**
   - 避免使用包含特殊字符的路径
   - Avoid using paths with special characters

### macOS 平台 / macOS Platform

1. **PPT 合并** / **PPT Merging**
   - 使用 `python-pptx` 库进行 PPT 合并
   - 某些复杂的 PPT 格式可能无法完美保留
   - Uses `python-pptx` library for PPT merging
   - Some complex PPT formats may not be perfectly preserved

2. **PDF 转换** / **PDF Conversion**
   - macOS 平台暂不支持 PDF 转换功能
   - PDF conversion is not supported on macOS platform

3. **启动脚本** / **Launch Script**
   - 如果双击 `.command` 文件无法运行，请检查文件权限：
   - If double-clicking the `.command` file doesn't work, check file permissions:
   ```bash
   chmod +x "mac 下启动PPT合并工具.command"
   ```

### 通用注意事项 / General Notes

1. **文件备份** / **File Backup**
   - 建议在合并前备份原始文件
   - It's recommended to backup original files before merging

2. **文件大小** / **File Size**
   - 合并大量或大型 PPT 文件可能需要较长时间
   - Merging many or large PPT files may take some time

3. **目录页** / **Table of Contents**
   - 目录页会自动插入到合并文件的第一页
   - The table of contents page is automatically inserted as the first page of the merged file

---

## 常见问题 / FAQ

### Q: 为什么 macOS 上无法使用 PDF 转换功能？
**A:** PDF 转换功能依赖 Windows 的 PowerPoint COM 接口和 VBS 脚本，macOS 平台暂不支持。可以使用其他工具先将 PPT 转换为 PDF，然后手动合并。

**Q: Why can't I use PDF conversion on macOS?**
**A:** PDF conversion relies on Windows PowerPoint COM interface and VBS scripts, which are not supported on macOS. You can use other tools to convert PPT to PDF first, then merge manually.

### Q: 合并后的文件在哪里？
**A:** 合并后的文件会保存在您选择的目录中，文件名包含日期和模式名称。

**Q: Where is the merged file saved?**
**A:** The merged file is saved in the directory you selected, with a filename containing the date and mode name.

### Q: 如何修改输出文件名？
**A:** 目前输出文件名是自动生成的，格式为 `YYYYMMDD合并PPT.pptx` 或 `YYYYMMDD[模式名称].pdf`。如需修改，可以手动重命名生成的文件。

**Q: How to modify the output filename?**
**A:** Currently, the output filename is auto-generated in the format `YYYYMMDD合并PPT.pptx` or `YYYYMMDD[Mode Name].pdf`. You can manually rename the generated file if needed.

### Q: 支持哪些 PPT 格式？
**A:** 支持 `.ppt` 和 `.pptx` 格式。

**Q: What PPT formats are supported?**
**A:** Supports `.ppt` and `.pptx` formats.

---

## 许可证 / License

本项目为个人使用工具，请根据实际需求使用。

This project is a personal use tool. Please use according to your actual needs.

---

## 更新日志 / Changelog

### 当前版本 / Current Version
- 支持 Windows 和 macOS 平台
- 支持 PPT 文件合并
- 支持 PDF 转换与合并（Windows）
- 支持拖拽排序
- 自动生成目录页
- Supports Windows and macOS platforms
- Supports PPT file merging
- Supports PDF conversion and merging (Windows)
- Supports drag-and-drop sorting
- Automatically generates table of contents

---

## 贡献 / Contributing

欢迎提交问题报告和改进建议。

Issues and improvement suggestions are welcome.

---

## 联系方式 / Contact

如有问题或建议，请通过项目仓库提交 Issue。

For questions or suggestions, please submit an Issue through the project repository.

