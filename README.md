# ToPDF

一个简单的 Windows 桌面工具，用于批量将 Word 文档转换为 PDF。

选择文件夹，点击开始转换，程序会在后台调用 Microsoft Word，将生成的 PDF 保存到桌面上的新文件夹。

## 下载与使用

前往 [Releases 页面](https://github.com/promisetsuri/ToPDF/releases)，下载发布附件中的 `WordToPDF.exe`。

1. 双击 `WordToPDF.exe` 打开程序。
2. 点击 **选择文件夹**，或直接输入 Word 文档所在的文件夹地址。
3. 如需转换子文件夹中的文档，勾选 **包含子文件夹**。
4. 点击 **开始转换**，在窗口中查看进度和转换结果。
5. 完成后点击 **打开输出文件夹**，查看 PDF 和转换日志。

> 使用 exe 无需安装 Python，但电脑必须安装桌面版 Microsoft Word，并已完成激活和首次启动。本程序不使用 WPS 转换。

## 功能

- 批量转换 `.doc` 和 `.docx` 文件。
- 支持包含子文件夹，并保留原有目录结构。
- 每次任务在桌面创建独立输出文件夹，避免覆盖之前的结果。
- 同名 `.doc` 和 `.docx` 生成不同名称的 PDF。
- 自动跳过以 `~$` 开头的 Word 临时文件。
- 后台转换，窗口显示进度和错误；单个文件失败后继续处理其他文件。
- 支持在当前文件处理结束后停止任务。
- 尝试关闭每个文档和本次任务创建的 Word 实例。
- 禁用文档宏，不修改源文档。

## 输出位置

PDF 默认保存到 Windows 桌面的新文件夹中，例如：

```text
桌面/
└── Word转PDF_20261013_132000/
    ├── 文档A.pdf
    ├── 文档B.pdf
    ├── 子文件夹/
    │   └── 文档C.pdf
    └── 转换日志.txt
```

程序通过 Windows 获取实际桌面路径，支持重定向到 OneDrive 的桌面。`转换日志.txt` 包含任务统计和错误信息。

## 从源码运行

需要 Windows、Python 3.10 或更高版本，以及桌面版 Microsoft Word。Python 安装时需保留 Tcl/Tk 支持；Tkinter 不需要单独通过 pip 安装。

下载或克隆仓库后，在项目目录打开 PowerShell：

```powershell
git clone https://github.com/promisetsuri/ToPDF.git
cd ToPDF
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe topdf.py
```

如果已经下载并进入项目目录，跳过前两条命令；已有虚拟环境时跳过创建步骤。请逐条执行命令。

## 打包 exe

在 Windows 项目目录中执行：

```powershell
.\.venv\Scripts\python.exe -m pip install --upgrade pyinstaller
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed --name WordToPDF --hidden-import=win32timezone topdf.py
```

生成的程序位于 `dist\WordToPDF.exe`。打包后请在目标电脑测试；Microsoft Word 不会随 exe 一起打包。

## 常见问题

### 提示 `No module named win32com`

依赖可能安装到了另一个 Python 环境。使用同一个虚拟环境安装和运行：

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe topdf.py
```

### Word 无法启动或部分文档转换失败

先确认 Microsoft Word 能手动正常打开。密码文档不支持输入密码；损坏文档、受保护视图或需要交互确认的文件可能无法转换。查看窗口中的错误和输出目录中的 `转换日志.txt`。

### 点击停止后没有立即结束

停止操作会等待当前文件的 Word 调用完成。如果 Word 无响应或被对话框阻塞，程序可能一直等待；本版本没有强制超时终止功能。检查 Word 是否有待处理提示。

### PDF 排版与预期不同

转换结果受本机 Word 版本、字体及文档内容影响。可先用 Word 手动打开原文件，检查排版和字体。

## 反馈

欢迎通过 [Issues](https://github.com/promisetsuri/ToPDF/issues) 反馈问题。请说明 Windows 和 Word 版本、操作步骤，以及窗口或日志中的错误信息。

提交截图或示例文档前，请移除个人信息和敏感内容。
