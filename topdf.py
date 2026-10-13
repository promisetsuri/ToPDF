"""Windows Word → PDF desktop application. Run with Python or PyInstaller."""
import os
import queue
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText


def desktop_path():
    # Windows shell resolves redirected desktops, including OneDrive.
    from win32com.shell import shell, shellcon
    return Path(shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOPDIRECTORY, 0, 0))


def find_documents(source, recursive):
    items = source.rglob('*') if recursive else source.iterdir()
    return sorted((p for p in items if p.is_file()
                   and p.suffix.lower() in {'.doc', '.docx'}
                   and not p.name.startswith('~$')), key=lambda p: str(p).casefold())


def reserve_output(desktop):
    base = 'Word转PDF_' + datetime.now().strftime('%Y%m%d_%H%M%S')
    for index in range(10000):
        path = desktop / (base if index == 0 else f'{base}_{index}')
        try:
            path.mkdir()
            return path
        except FileExistsError:
            continue
    raise RuntimeError('无法创建新的输出文件夹')


def convert_batch(source, files, output, events, stop):
    """All COM objects stay on this worker thread; Tk is updated via a queue."""
    word = None
    initialized = False
    success = failed = processed = 0
    errors = []
    fatal = None
    used = set()
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        initialized = True
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3  # Disable document macros.
        for item in files:
            if stop.is_set():
                break
            document = None
            relative = item.relative_to(source)
            target = output / relative.with_suffix('.pdf')
            # foo.doc and foo.docx must not overwrite one another.
            original = target
            number = 2
            while str(target).casefold() in used or target.exists():
                target = original.with_name(f'{original.stem}_{number}.pdf')
                number += 1
            used.add(str(target).casefold())
            events.put(('log', f'正在转换：{relative}'))
            try:
                target.parent.mkdir(parents=True, exist_ok=True)
                document = word.Documents.Open(
                    FileName=str(item), ConfirmConversions=False,
                    ReadOnly=True, AddToRecentFiles=False, Visible=False,
                    NoEncodingDialog=True,
                    PasswordDocument='__no_password_provided__',
                    WritePasswordDocument='__no_password_provided__')
                document.ExportAsFixedFormat(OutputFileName=str(target), ExportFormat=17)
                if not target.is_file() or target.stat().st_size == 0:
                    raise RuntimeError('Word 未生成有效的 PDF 文件')
                success += 1
                events.put(('log', f'成功：{target.relative_to(output)}'))
            except Exception as exc:
                failed += 1
                detail = f'失败：{relative}\n原因：{exc}'
                errors.append(detail)
                events.put(('log', detail))
                try:
                    target.unlink(missing_ok=True)
                except OSError:
                    pass
            finally:
                if document is not None:
                    try:
                        document.Close(SaveChanges=0)
                    except Exception as exc:
                        detail = f'文档关闭异常：{relative}：{exc}'
                        errors.append(detail)
                        events.put(('log', detail))
                    document = None
            processed += 1
            events.put(('progress', processed))
    except Exception as exc:
        fatal = f'任务无法继续：{exc}\n请确认已安装并完成 Microsoft Word 首次启动。'
        errors.append(fatal)
        events.put(('log', fatal))
    finally:
        if word is not None:
            try:
                word.Quit(SaveChanges=0)
            except Exception as exc:
                detail = f'Word 退出异常：{exc}'
                errors.append(detail)
                events.put(('log', detail))
            word = None
        if initialized:
            pythoncom.CoUninitialize()
        summary = (f'成功 {success}，失败 {failed}，未处理 {len(files) - processed}\n'
                   f'输出：{output}')
        try:
            (output / '转换日志.txt').write_text(
                summary + '\n\n' + '\n\n'.join(errors), encoding='utf-8')
        except OSError as exc:
            events.put(('log', f'日志保存失败：{exc}'))
        events.put(('done', (summary, fatal)))


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Word 转 PDF')
        self.geometry('760x540')
        self.minsize(600, 430)
        self.events = queue.Queue()
        self.stop = threading.Event()
        self.running = False
        self.output = None
        self.source = tk.StringVar()
        self.recursive = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value='选择 Word 文件所在的文件夹，然后开始转换。')
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text='Word 文件夹').pack(anchor='w')
        row = ttk.Frame(frame)
        row.pack(fill='x', pady=(6, 10))
        self.entry = ttk.Entry(row, textvariable=self.source)
        self.entry.pack(side='left', fill='x', expand=True)
        self.browse = ttk.Button(row, text='选择文件夹', command=self.choose)
        self.browse.pack(side='left', padx=(8, 0))
        self.check = ttk.Checkbutton(frame, text='包含子文件夹（保留目录结构）', variable=self.recursive)
        self.check.pack(anchor='w')
        ttk.Label(frame, text='PDF 保存到桌面新建的“Word转PDF_时间”文件夹。').pack(anchor='w', pady=8)
        buttons = ttk.Frame(frame)
        buttons.pack(fill='x', pady=6)
        self.start_button = ttk.Button(buttons, text='开始转换', command=self.start)
        self.start_button.pack(side='left')
        self.cancel_button = ttk.Button(buttons, text='停止', command=self.cancel, state='disabled')
        self.cancel_button.pack(side='left', padx=8)
        self.open_button = ttk.Button(buttons, text='打开输出文件夹', command=self.open_output, state='disabled')
        self.open_button.pack(side='left')
        self.progress = ttk.Progressbar(frame, mode='determinate')
        self.progress.pack(fill='x', pady=8)
        ttk.Label(frame, textvariable=self.status, wraplength=700).pack(anchor='w')
        self.log = ScrolledText(frame, height=15, state='disabled', wrap='word')
        self.log.pack(fill='both', expand=True, pady=(10, 0))
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.after(100, self.poll)

    def choose(self):
        folder = filedialog.askdirectory(title='选择 Word 文件夹')
        if folder:
            self.source.set(folder)

    def append(self, text):
        self.log.configure(state='normal')
        self.log.insert('end', text + '\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def start(self):
        if self.running:
            return
        try:
            raw = self.source.get().strip().strip('"')
            if not raw:
                raise ValueError('请先选择或填写源文件夹。')
            source = Path(raw).expanduser().resolve()
            if not source.is_dir():
                raise ValueError('源文件夹不存在。')
            files = find_documents(source, self.recursive.get())
            if not files:
                raise ValueError('文件夹内没有可转换的 .doc 或 .docx 文件。')
            import pythoncom
            import win32com.client
            output = reserve_output(desktop_path())
        except Exception as exc:
            messagebox.showerror('无法开始', str(exc))
            return
        self.output = output
        self.total = len(files)
        self.running = True
        self.stop.clear()
        self.progress.configure(maximum=self.total, value=0)
        self.status.set(f'正在启动 Word，共 {self.total} 个文件……')
        for widget in (self.entry, self.browse, self.check, self.start_button):
            widget.configure(state='disabled')
        self.cancel_button.configure(state='normal')
        self.open_button.configure(state='normal')
        self.append(f'新任务：{source}\n输出：{output}')
        threading.Thread(target=convert_batch,
                         args=(source, files, output, self.events, self.stop),
                         daemon=False).start()

    def cancel(self):
        self.stop.set()
        self.cancel_button.configure(state='disabled')
        self.status.set('将在当前文件转换结束后停止，并关闭 Word。')

    def open_output(self):
        if self.output:
            try:
                os.startfile(str(self.output))
            except OSError as exc:
                messagebox.showerror('无法打开文件夹', str(exc))

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == 'log':
                    self.append(value)
                elif kind == 'progress':
                    self.progress.configure(value=value)
                    if not self.stop.is_set():
                        self.status.set(f'已处理 {value} / {self.total}')
                elif kind == 'done':
                    self.running = False
                    summary, fatal = value
                    title = '任务出错' if fatal else ('已停止' if self.stop.is_set() else '转换完成')
                    self.status.set(title + '：' + summary)
                    self.append(title + '\n' + summary)
                    for widget in (self.entry, self.browse, self.check, self.start_button):
                        widget.configure(state='normal')
                    self.cancel_button.configure(state='disabled')
        except queue.Empty:
            pass
        self.after(100, self.poll)

    def close(self):
        if self.running:
            if messagebox.askyesno('正在转换', '停止后才能关闭窗口。是否在当前文件结束后停止？'):
                self.cancel()
            return
        self.destroy()


if __name__ == '__main__':
    App().mainloop()
