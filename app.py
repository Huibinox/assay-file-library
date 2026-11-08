import json
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import uuid
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from storage import Library, PATTERN, BusyError, ConflictError, atomic_json

BASE = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
CONFIG = Path(os.environ.get('LOCALAPPDATA', str(BASE))) / 'AssayFileLibrary' / 'settings.json'


def expiry_style(value, today=None):
    if not value:
        return '', '未设置'
    days = (date.fromisoformat(value) - (today or date.today())).days
    return ('red' if days <= 5 else 'yellow' if days <= 30 else '',
            f'已过期 {-days} 天' if days < 0 else f'剩余 {days} 天')


def convert_word(source, directory):
    pdf = Path(directory) / 'preview.pdf'
    script = Path(directory) / 'convert.ps1'
    script.write_text("""$ErrorActionPreference='Stop'
$w=$null; $d=$null
try {
 $w=New-Object -ComObject Word.Application
 $w.Visible=$false; $w.AutomationSecurity=3; $w.DisplayAlerts=0
 $d=$w.Documents.Open($args[0],$false,$true)
 $d.Repaginate()
 $n=$d.ComputeStatistics(2)
 $d.ExportAsFixedFormat($args[1],17)
 Write-Output $n
} finally {
 if($d){$d.Close(0)}
 if($w){$w.Quit()}
}
""", encoding='utf-8-sig')
    result = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
                             str(script), str(source), str(pdf)], capture_output=True,
                            text=True, encoding='utf-8', errors='replace', timeout=180,
                            creationflags=0x08000000)
    if result.returncode or not pdf.exists():
        raise RuntimeError('Word 转换失败')
    return pdf, int(result.stdout.strip().splitlines()[-1])


class App:
    def __init__(self, root, library=None):
        self.root = root
        self.library = library
        self.projects = {}
        self.records = {}
        self.busy = False
        self.events = queue.Queue()
        self.note = tk.StringVar(value='请选择资料库。每台电脑安装程序，选择同一个服务器文件夹即可。')
        root.title('Assay File 库 · v0.2.0 共享盘版')
        root.geometry('1200x720')
        root.minsize(1000, 560)
        root.protocol('WM_DELETE_WINDOW', self.close)
        style = ttk.Style()
        style.configure('Treeview', rowheight=32, font=('Microsoft YaHei UI', 10))
        top = ttk.Frame(root, padding=16)
        top.pack(fill='x')
        ttk.Label(top, text='Assay File 库', font=('Microsoft YaHei UI', 20, 'bold')).pack(side='left')
        ttk.Button(top, text='选择资料库', command=self.choose_library).pack(side='right', padx=4)
        ttk.Button(top, text='打开资料文件夹', command=self.open_library).pack(side='right', padx=4)
        self.location = tk.StringVar(value='尚未选择资料库')
        ttk.Label(root, textvariable=self.location, padding=(16, 0, 16, 12), wraplength=1100).pack(fill='x')
        bar = ttk.Frame(root, padding=(16, 0, 16, 12))
        bar.pack(fill='x')
        ttk.Label(bar, text='搜索项目号 / 记录名称  ').pack(side='left')
        self.query = tk.StringVar()
        ttk.Entry(bar, textvariable=self.query, width=30).pack(side='left')
        self.query.trace_add('write', lambda *_: self.render())
        for label, fn in [('新增项目', self.add_project), ('导入记录', self.import_record),
                          ('填写最近有效期', self.expiry), ('刷新', self.refresh)]:
            ttk.Button(bar, text=label, command=fn).pack(side='left', padx=4)
        frame = ttk.Frame(root)
        frame.pack(fill='both', expand=True, padx=16)
        self.tree = ttk.Treeview(frame, columns=('pages', 'updated', 'expiry', 'status', 'print'), selectmode='browse')
        self.tree.heading('#0', text='项目号 / 试验记录')
        self.tree.column('#0', width=390)
        for key, title, width in [('pages', '页数', 60), ('updated', '导入 / 更新日期', 150),
                                  ('expiry', '最近有效期', 120), ('status', '状态', 140), ('print', '操作', 65)]:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor='center')
        scroll = ttk.Scrollbar(frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        self.tree.pack(fill='both', expand=True)
        self.tree.tag_configure('yellow', background='#fff0b3')
        self.tree.tag_configure('red', background='#ffd2d2')
        self.tree.bind('<ButtonRelease-1>', self.row_click)
        actions = ttk.Frame(root, padding=16)
        actions.pack(fill='x')
        for label, fn in [('预览 PDF', self.preview), ('打开 Word 副本', self.open_file),
                          ('打印所选记录', self.print_file), ('更新所选记录', self.update), ('查看历史', self.history)]:
            ttk.Button(actions, text=label, command=fn).pack(side='left', padx=4)
        ttk.Label(root, textvariable=self.note, padding=(16, 0, 16, 12), wraplength=1100).pack(fill='x')
        if library:
            self.location.set('当前资料库：' + str(library.root))
            self.refresh()
        else:
            root.after(100, self.restore_library)
        root.after(100, self.poll)
        root.after(15000, self.tick)

    def restore_library(self):
        try:
            value = json.loads(CONFIG.read_text(encoding='utf-8')).get('library')
        except (OSError, ValueError):
            value = None
        if value:
            # Do not recreate a remembered network location if the share is unavailable.
            if not Path(value).is_dir():
                self.note.set('上次选择的资料库不可访问，请连接共享盘后重新选择。')
                return
            self.attach(value)
        else:
            self.choose_library()

    def attach(self, value):
        try:
            lib = Library(value)
            projects = lib.list_projects()
            if (lib.root / 'library.db').exists() and not projects:
                messagebox.showwarning('旧版资料库', '检测到 v0.1 的 library.db。新版使用独立的 projects 文件夹。\n旧文档不会自动迁移或删除；请将旧版当前 Word 重新导入，新旧数据可以并存。')
            self.library = lib
            self.projects = {p['id']: p for p in projects}
            CONFIG.parent.mkdir(parents=True, exist_ok=True)
            atomic_json(CONFIG, {'library': str(lib.root)})
            self.location.set('当前资料库：' + str(lib.root))
            self.render()
            self.note.set('资料库已连接。每 15 秒刷新；导入前先在本机准备文档，再安全写入共享库。')
        except (OSError, ValueError) as exc:
            messagebox.showerror('无法连接资料库', str(exc))

    def choose_library(self):
        if self.busy:
            return messagebox.showinfo('正在处理', '请等待当前导入或更新完成。')
        value = filedialog.askdirectory(title='选择统一资料库文件夹（可选公司共享盘）', mustexist=True)
        if value:
            self.attach(value)

    def open_library(self):
        if self.library:
            self.launch(self.library.root)

    def ready(self):
        if not self.library:
            messagebox.showinfo('选择资料库', '请先点击右上角“选择资料库”。')
            return False
        if self.busy:
            messagebox.showinfo('正在处理', '请等待当前导入或更新完成。')
            return False
        return True

    def refresh(self):
        if not self.library:
            return
        try:
            projects = self.library.list_projects()
            self.projects = {p['id']: p for p in projects}
            self.render()
        except (OSError, ValueError) as exc:
            self.note.set('资料库读取失败，请检查共享盘连接后刷新。当前列表可能已过时：' + str(exc))

    def tick(self):
        self.refresh()
        self.root.after(15000, self.tick)

    def selection(self):
        items = self.tree.selection()
        return items[0] if items else ''

    def project(self):
        selected = self.selection()
        if selected.startswith('p:'):
            return selected[2:]
        if selected.startswith('r:') and selected[2:] in self.records:
            return self.records[selected[2:]][0]
        return None

    def record(self):
        selected = self.selection()
        if not selected.startswith('r:') or selected[2:] not in self.records:
            messagebox.showinfo('选择记录', '请先展开项目，选择一份试验记录。')
            return None
        pid, cached = self.records[selected[2:]]
        try:
            data = self.library.load(pid)
            record = next(r for r in data['records'] if r['id'] == cached['id'])
            return pid, record
        except (OSError, ValueError, StopIteration) as exc:
            messagebox.showerror('无法读取记录', '请检查共享盘连接并刷新。\n' + str(exc))
            return None

    def render(self):
        selected = self.selection()
        opened = {i for i in self.tree.get_children() if self.tree.item(i, 'open')}
        self.tree.delete(*self.tree.get_children())
        self.records = {}
        q = self.query.get().strip().lower()
        for pid, project in sorted(self.projects.items()):
            matching = [r for r in project['records'] if q in pid.lower() or q in r['name'].lower()]
            if q and q not in pid.lower() and not matching:
                continue
            tag, status = expiry_style(project.get('expiry'))
            parent = 'p:' + pid
            self.tree.insert('', 'end', iid=parent, text=pid,
                             values=('', '', project.get('expiry') or '未设置', status),
                             tags=(tag,), open=bool(q) or parent in opened)
            for r in matching:
                self.records[r['id']] = (pid, r)
                self.tree.insert(parent, 'end', iid='r:' + r['id'], text=r['name'],
                                 values=(r.get('pages') or '—', r['updated'], '',
                                         'PDF 就绪' if r.get('pdf') else '无 PDF，可用 Word', '打印'))
        if selected and self.tree.exists(selected):
            self.tree.selection_set(selected)

    def row_click(self, event):
        row = self.tree.identify_row(event.y)
        if row.startswith('r:') and self.tree.identify_column(event.x) == '#5':
            self.tree.selection_set(row)
            self.print_file()

    def report_error(self, exc):
        title = '稍后再试' if isinstance(exc, BusyError) else '内容已更新' if isinstance(exc, ConflictError) else '操作失败'
        messagebox.showerror(title, str(exc))

    def add_project(self):
        if not self.ready():
            return
        pid = simpledialog.askstring('新增项目', '项目号：123456-654321-FSA 或带 -HGR')
        if not pid:
            return
        try:
            pid = pid.strip().upper()
            self.library.create(pid)
            self.refresh()
            self.tree.selection_set('p:' + pid)
        except (OSError, ValueError, BusyError) as exc:
            self.report_error(exc)

    def expiry(self):
        if not self.ready():
            return
        pid = self.project()
        if not pid:
            return messagebox.showinfo('选择项目', '请先选择项目。')
        try:
            project = self.library.load(pid)
            value = simpledialog.askstring('最近有效期', '输入 YYYY-MM-DD；留空清除', initialvalue=project.get('expiry') or '')
            if value is None:
                return
            value = date.fromisoformat(value.strip()).isoformat() if value.strip() else None
            self.library.set_expiry(pid, value, project['revision'])
            self.refresh()
        except (OSError, ValueError, BusyError, ConflictError) as exc:
            self.report_error(exc)

    def import_record(self):
        if not self.ready():
            return
        selected_project = self.project()
        paths = filedialog.askopenfilenames(title='导入 Word（可多选）', filetypes=[('Word 文档', '*.docx *.doc')])
        jobs = []
        for path in paths:
            match = re.search(PATTERN, Path(path).stem.upper())
            pid = match.group() if match else selected_project
            if not pid:
                pid = simpledialog.askstring('所属项目', f'{Path(path).name}\n请输入项目号')
                pid = pid.strip().upper() if pid else ''
            if not re.fullmatch(PATTERN, pid):
                messagebox.showerror('未导入', '项目号格式不正确。')
                continue
            name = simpledialog.askstring('记录名称', f'项目：{pid}\n例如：验证－4样本、正式－11样本', initialvalue=Path(path).stem)
            if name and name.strip():
                jobs.append((path, pid, uuid.uuid4().hex, name.strip(), None))
        self.start_jobs(jobs)

    def update(self):
        if not self.ready():
            return
        item = self.record()
        if not item:
            return
        pid, r = item
        path = filedialog.askopenfilename(filetypes=[('Word 文档', '*.docx *.doc')])
        if path:
            self.start_jobs([(path, pid, r['id'], r['name'], r['version'])])

    def start_jobs(self, jobs):
        if not jobs:
            return
        self.busy = True
        library = self.library
        self.note.set('正在本机准备文档和 PDF，完成后写入资料库；请勿关闭程序。')
        def worker():
            try:
                for source, pid, rid, name, expected in jobs:
                    try:
                        with tempfile.TemporaryDirectory(prefix='assay-') as td:
                            local = Path(td) / ('input' + Path(source).suffix.lower())
                            shutil.copyfile(source, local)
                            try:
                                pdf, pages = convert_word(local, td)
                            except Exception:
                                reply = threading.Event()
                                answer = []
                                self.events.put(('convert_failed', name, reply, answer))
                                reply.wait()
                                if not answer[0]:
                                    continue
                                pdf, pages = None, None
                            library.commit_record(pid, rid, name, local, pdf, pages,
                                                  datetime.now().strftime('%Y-%m-%d %H:%M'), expected)
                            self.events.put(('saved', name))
                    except Exception as exc:
                        self.events.put(('error', exc))
            finally:
                self.events.put(('done',))
        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == 'convert_failed':
                    _, name, reply, answer = event
                    answer.append(messagebox.askyesno('PDF 转换未完成', f'{name}\n\n请检查是否安装 Microsoft Word。\n是否仍保存 Word？\n若正在更新，确认后将使用新版 Word，不再使用旧版 PDF。'))
                    reply.set()
                elif event[0] == 'saved':
                    self.note.set('已安全保存：' + event[1])
                    self.refresh()
                elif event[0] == 'error':
                    self.report_error(event[1])
                    self.note.set('部分操作未完成，请核对提示；原当前版本不会被不完整文件替换。')
                elif event[0] == 'done':
                    self.busy = False
                    self.refresh()
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def launch(self, path, verb='open'):
        try:
            os.startfile(path, verb)
            return True
        except OSError as exc:
            messagebox.showerror('无法打开', str(exc))
            return False

    def local_copy(self, relative):
        source = self.library.resolve(relative)
        cache = Path(tempfile.gettempdir()) / 'AssayFileLibrary' / uuid.uuid4().hex
        cache.mkdir(parents=True)
        dest = cache / source.name
        shutil.copyfile(source, dest)
        return dest

    def open_file(self):
        item = self.record()
        if item:
            try:
                self.launch(self.local_copy(item[1]['source']))
                self.note.set('已打开本机副本。编辑后请使用“更新所选记录”导入，库中文档不会被直接修改。')
            except (OSError, ValueError) as exc:
                self.report_error(exc)

    def preview(self):
        item = self.record()
        if item:
            try:
                if item[1]['pdf']:
                    self.launch(self.local_copy(item[1]['pdf']))
                else:
                    messagebox.showinfo('无 PDF', '该记录尚无 PDF，可打开 Word 副本查看。')
            except (OSError, ValueError) as exc:
                self.report_error(exc)

    def print_file(self):
        item = self.record()
        if not item:
            return
        pid, r = item
        if messagebox.askokcancel('确认打印', f'项目：{pid}\n记录：{r["name"]}\n更新时间：{r["updated"]}\n\n将通过系统关联程序发送到默认打印机。'):
            try:
                # Recheck after the confirmation dialog to avoid silently printing an old revision.
                current = next(x for x in self.library.load(pid)['records'] if x['id'] == r['id'])
                if current['version'] != r['version']:
                    raise ConflictError('确认期间该记录已更新，请刷新后重新打印。')
                if self.launch(self.local_copy(r['pdf'] or r['source']), 'print'):
                    self.note.set('已请求系统打印，请在打印队列确认；当前打印的是确认时选中的版本。')
            except (OSError, ValueError, ConflictError) as exc:
                self.report_error(exc)

    def history(self):
        item = self.record()
        if not item:
            return
        pid, r = item
        win = tk.Toplevel(self.root)
        win.title('历史版本：' + r['name'])
        win.geometry('650x350')
        rows = r.get('history', [])
        listing = tk.Listbox(win, font=('Microsoft YaHei UI', 11))
        listing.pack(fill='both', expand=True, padx=12, pady=12)
        for n, entry in enumerate(rows, 1):
            listing.insert('end', f'{n}. {entry["updated"]}  {entry["name"]}')
        if not rows:
            listing.insert('end', '尚无历史版本')
        def view():
            selected = listing.curselection()
            if rows and selected:
                entry = rows[selected[0]]
                try:
                    self.launch(self.local_copy(entry['pdf'] or entry['source']))
                except (OSError, ValueError) as exc:
                    self.report_error(exc)
        ttk.Button(win, text='打开所选历史版本副本', command=view).pack(pady=10)

    def close(self):
        if self.busy:
            messagebox.showinfo('正在保存', '请等待当前任务结束后关闭，以免中断上传。')
            return
        self.root.destroy()


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
