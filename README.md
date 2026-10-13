# Word 转 PDF 桌面工具

用 Tkinter 创建的 Windows 桌面应用，通过本机 Microsoft Word 批量转换 `.doc`、`.docx`。

## 环境

- Windows 10/11，安装桌面版 Microsoft Word，并先手动打开一次完成激活和首次启动。
- 源码运行需要 Python 3.10 或更高版本（安装时保留 Tcl/Tk）。Tkinter 无需另行 pip 安装。
- exe 使用者不需要安装 Python，但仍然需要 Microsoft Word。WPS 不适用本程序。

## 运行（PowerShell）

在项目文件夹打开终端。已有 `.venv` 时跳过创建步骤。

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe topdf.py
```

选择或粘贴文件夹地址，点击“开始转换”。默认只处理该文件夹中的 Word 文件；勾选“包含子文件夹”可递归处理并保留目录结构。跳过 Word 临时锁文件 `~$...`。

每次在 Windows 实际桌面（支持 OneDrive 重定向）创建独立输出文件夹。同名 `.doc` 和 `.docx` 生成不同 PDF，避免覆盖。窗口显示进度和错误，输出目录内保存 `转换日志.txt`。单个文件失败后继续处理其他文件。

“停止”会等待当前文件结束；转换中关闭窗口会请求停止，任务结束后再关闭窗口。程序使用独立的 Word 实例，后台线程初始化 COM，文档在 finally 中关闭，任务结束时退出 Word，且禁用文档宏。

## 打包成可双击的 exe

在 Windows 项目目录、同一个虚拟环境中运行：

```powershell
.\.venv\Scripts\python.exe -m pip install --upgrade pyinstaller
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed --name WordToPDF --hidden-import=win32timezone topdf.py
```

生成文件为 `dist\WordToPDF.exe`。双击出现窗口，可复制给安装了 Microsoft Word 的 Windows 用户。打包后应在目标电脑测试普通文档、中文路径和错误文档。若 exe 启动异常，重新打包时去掉 `--windowed`，从终端运行查看错误。

参数说明：[PyInstaller 官方文档](https://pyinstaller.org/en/stable/usage.html)。Word 导出接口：[Microsoft 官方文档](https://learn.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat)。

## 上传 GitHub：网页方式

1. 登录 GitHub，点击右上角 `+` → `New repository`，填写 `word-to-pdf`。根据需要选择 Public 或 Private，再点击 `Create repository`。
2. 在仓库页选择 `uploading an existing file`（已有仓库用 `Add file` → `Upload files`）。上传 `topdf.py`、`requirements.txt`、`README.md`、`.gitignore`。隐藏的 `.gitignore` 可用网页 `Add file` → `Create new file` 创建并粘贴内容。
3. 填写提交说明，例如 `Add Word to PDF desktop app`，点击 `Commit changes`。
4. exe 可在仓库的 `Releases` → `Create a new release` 中作为附件发布，例如标签 `v1.0.0`。不上传 `.venv`、`build`、`dist` 文件夹或自己的文档。

## 上传 GitHub：Git 命令方式

先安装 Git，在 GitHub 创建一个空仓库（不要勾选自动创建 README），再在本项目目录运行。将下面的 `YOUR_USERNAME` 和仓库名替换为自己的值。

```powershell
git init
git add topdf.py requirements.txt README.md .gitignore
git commit -m "Add Word to PDF desktop app"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/word-to-pdf.git
git push -u origin main
```

首次提交若要求身份设置，按 Git 的提示填写自己的用户名和邮箱；首次推送按提示在浏览器登录 GitHub。不要把密码或令牌写进项目。以上命令用于新项目，已有仓库请沿用原来的远程地址。

## 故障排查和边界

- `No module named win32com`：用上面的 `.venv\Scripts\python.exe` 安装依赖并运行，不要混用解释器。
- Word 无法启动：确认安装的是 Microsoft Word 桌面版，完成激活后重新尝试。
- 密码文档会按失败记录，本程序不提供密码输入。损坏、受保护视图或需要交互确认的文件可能无法转换。
- 程序在正常成功、异常及停止路径中尝试关闭文档和专用 Word 实例。若 Word 本身无响应或弹出模态对话框，COM 调用可能一直等待；本版本没有强制超时终止功能。停止按钮无法中断正在执行的 COM 调用。可检查 Word 提示或任务管理器中的后台 Word 进程，先确认没有其他待保存文档再处理。
- PDF 外观受本机字体、Word 版本和文档内容影响。
- 发布为开源项目时，可自行选择并添加 LICENSE 文件。本项目未替你选择许可证。
