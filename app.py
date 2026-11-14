import hashlib
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
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from storage import Library, PATTERN, BusyError, ConflictError, DuplicateNameError, name_key, atomic_json

BASE = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
CONFIG = Path(os.environ.get('LOCALAPPDATA', str(BASE))) / 'AssayFileLibrary' / 'settings.json'


def import_name(stem, pid):
    match = re.search(PATTERN, stem.upper())
    label = stem[match.end():].strip(' _-') if match and match.group() == pid else stem.strip()
    standard = re.fullmatch(r'(SA|Qualification)(_.+)?', label, re.IGNORECASE)
    if standard:
        purpose = 'SA' if standard[1].lower() == 'sa' else 'Qualification'
        return purpose + (standard[2] or '')
    legacy = re.fullmatch(r'Daily\s+Assay\s+File(?:[ _-]+(.*))?', label, re.IGNORECASE)
    if not legacy:
        return None
    detail = (legacy[1] or '').strip(' _-')
    qualification = re.match(r'(?:qualify|qualification)(?=$|[ _-])', detail, re.IGNORECASE)
    if qualification:
        detail = detail[qualification.end():].strip(' _-')
    purpose = 'Qualification' if qualification else 'SA'
    return purpose + ('_' + detail if detail else '')


def expiry_style(value, today=None):
    if not value:
        return '', '未设置'
    days = (date.fromisoformat(value) - (today or date.today())).days
    return ('red' if days <= 5 else 'yellow' if days <= 30 else '',
            f'已过期 {-days} 天' if days < 0 else f'剩余 {days} 天')


def saved_word_pages(source):
    if Path(source).suffix.lower() != '.docx':
        return None
    try:
        with zipfile.ZipFile(source) as doc:
            properties = ET.fromstring(doc.read('docProps/app.xml'))
        pages = int(properties.findtext('{*}Pages', '0'))
        return pages if pages > 0 else None
    except (OSError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile):
        return None


def reagent_expiry(source):
    if source.suffix.lower() != '.docx':
        return None
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    try:
        with zipfile.ZipFile(source) as doc:
            body = ET.fromstring(doc.read('word/document.xml'))
    except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile):
        return None
    dates = []
    for table in body.findall('.//w:tbl', ns):
        rows = [[''.join(t.text or '' for t in cell.findall('.//w:t', ns))
                 for cell in row.findall('./w:tc', ns)]
                for row in table.findall('./w:tr', ns)]
        if not rows or not rows[0]:
            continue
        headers = [cell.casefold() for cell in rows[0]]
        if not any('reagent' in cell or '试剂名称' in cell for cell in headers):
            continue
        for column, header in enumerate(headers):
            if not any(label in header for label in ('exp.', 'exp date', 'expiry', '有效期', '过期日期', '过期日')):
                continue
            for row in rows[1:]:
                if column >= len(row):
                    continue
                for match in re.finditer(r'(?<!\d)(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?!\d)', row[column]):
                    try:
                        dates.append(date(*map(int, match.groups())))
                    except ValueError:
                        continue
    return min(dates).isoformat() if dates else None


def count_word_pages(source, directory):
    script = Path(directory) / 'word-pages.ps1'
    script.write_text("""$ErrorActionPreference='Stop'
$w=$null; $d=$null
try {
 $w=New-Object -ComObject Word.Application
 $w.Visible=$false; $w.AutomationSecurity=3; $w.DisplayAlerts=0
 $d=$w.Documents.Open($args[0],$false,$true)
 $d.Repaginate()
 $n=$d.ComputeStatistics(2)
 Write-Output $n
} finally {
 if($d){$d.Close(0)}
 if($w){$w.Quit()}
}
""", encoding='utf-8-sig')
    result = subprocess.run(['powershell', '-NoProfile', '-File',
                             str(script), str(source)], capture_output=True,
                            text=True, encoding='utf-8', errors='replace', timeout=180,
                            creationflags=0x08000000)
    pages = next((int(line) for line in result.stdout.splitlines() if re.fullmatch(r'[1-9][0-9]*', line.strip())), None)
    return pages


class App:
    def __init__(self, root, library=None):
        self.root = root
        self.library = library
        self.projects = {}
        self.records = {}
        self.busy = False
        self.events = queue.Queue()
        self.inbox_seen = {}
        self.inbox_pages = {}
        self.expiry_cache = {}
        self.live_seen = {}
        self.live_done = {}
        self.note = tk.StringVar(value='请选择资料库。每台电脑安装程序，选择同一个服务器文件夹即可。')
        root.title('Assay File 库 · v0.3.1 共享盘版')
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
        for key, title, width in [('pages', 'Word 页数', 80), ('updated', '导入 / 更新日期', 150),
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
        self.tree.bind('<Double-1>', self.open_on_double_click)
        actions = ttk.Frame(root, padding=16)
        actions.pack(fill='x')
        for label, fn in [('打开 Word', self.open_file),
                          ('打印所选记录', self.print_file), ('更新所选记录', self.update), ('查看历史', self.history), ('删除所选项', self.delete_selected), ('回收站', self.recycle_bin)]:
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
            self.inbox_seen.clear()
            self.inbox_pages.clear()
            self.expiry_cache.clear()
            self.live_seen.clear()
            self.live_done.clear()
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
        if self.library and not self.busy:
            self.scan_live()
            if not self.busy:
                self.scan_inbox()
        self.root.after(15000, self.tick)

    def scan_live(self):
        observed = {}
        jobs = []
        try:
            for project in self.library.list_projects():
                for record in project['records']:
                    if not record.get('snapshot'):
                        continue
                    path = self.library.resolve(record['source'])
                    try:
                        stat = path.stat()
                    except FileNotFoundError:
                        continue
                    signature = (stat.st_size, stat.st_mtime_ns)
                    observed[path] = signature
                    if stat.st_size and self.live_seen.get(path) == signature and self.live_done.get(path) != signature:
                        jobs.append((path, project['id'], record['id'], record['name'], record['version']))
            self.live_seen = observed
            self.start_jobs(jobs, automatic=True)
        except (OSError, ValueError) as exc:
            self.note.set(f'读取共享库失败：{exc}')

    def scan_inbox(self):
        library = self.library
        try:
            projects = {p['id']: p for p in library.list_projects(include_deleted=True)}
            folders = [(library.root, None)]
            folders.extend((folder, folder.name) for folder in library.projects_dir.iterdir()
                           if re.fullmatch(PATTERN, folder.name) and folder.is_dir())
            managed = {r['source'].casefold() for project in projects.values() for r in project['records']}
            folders.extend((folder, folder.name) for folder in library.root.iterdir()
                           if re.fullmatch(PATTERN, folder.name) and folder.is_dir())
            observed = {}
            jobs = []
            notices = []
            pending_names = set()
            for folder, project_id in folders:
                for path in sorted(folder.iterdir()):
                    if path.name.startswith('~$') or path.suffix.lower() not in ('.doc', '.docx') or not path.is_file():
                        continue
                    match = re.search(PATTERN, path.stem.upper())
                    if project_id and match and match.group() != project_id:
                        notices.append(f'{path.name} 的项目号与所在文件夹不符')
                        continue
                    pid = project_id or (match.group() if match else None)
                    if not pid:
                        notices.append(f'{path.name} 缺少项目号')
                        continue
                    relative = path.relative_to(library.root).as_posix()
                    data = projects.get(pid, {'id': pid, 'records': []})
                    if relative.casefold() in managed or relative.casefold() in data.get('inbox_imports', []) or data.get('deleted'):
                        continue
                    stat = path.stat()
                    signature = (stat.st_size, stat.st_mtime_ns)
                    observed[relative] = signature
                    if not stat.st_size or self.inbox_seen.get(relative) != signature:
                        continue
                    name = import_name(path.stem, pid)
                    if name is None:
                        notices.append(f'{path.name} 无法识别用途，请手动导入并选择记录名称')
                        continue
                    try:
                        library.check_name(data, name)
                    except DuplicateNameError:
                        notices.append(f'{path.name} 与项目 {pid} 的记录重名')
                        continue
                    key = (pid, name_key(name))
                    if key not in pending_names:
                        jobs.append((path, pid, uuid.uuid4().hex, name, None))
                        pending_names.add(key)
                    else:
                        notices.append(f'{path.name} 与待导入文件重名')
            self.inbox_seen = observed
            if notices:
                self.note.set('未自动导入：' + '；'.join(notices[:3]) + '。请核对文件名或使用“更新所选记录”。')
            self.start_jobs(jobs, automatic=True)
        except (OSError, ValueError) as exc:
            self.note.set(f'自动扫描失败，下次将重试：{exc}')

    def word_pages(self, record):
        if record.get('pages'):
            return record['pages']
        source = record['source']
        if source not in self.inbox_pages:
            self.inbox_pages[source] = saved_word_pages(self.library.resolve(source))
        return self.inbox_pages[source] or '—'

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
            self.library.require_active(data)
            record = next(r for r in data['records'] if r['id'] == cached['id'])
            return pid, record
        except (OSError, ValueError, StopIteration, ConflictError) as exc:
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
                source = r['source']
                missing = bool(r.get('snapshot')) and not self.library.resolve(source).is_file()
                key = (source, r['version'])
                if key not in self.expiry_cache:
                    self.expiry_cache[key] = reagent_expiry(self.library.resolve(r.get('snapshot', source)))
                expiry = self.expiry_cache[key]
                record_tag, record_status = expiry_style(expiry)
                self.tree.insert(parent, 'end', iid='r:' + r['id'], text=r['name'],
                                 values=(self.word_pages(r), r['updated'], expiry or '—',
                                         'Word 文件缺失' if missing else record_status if expiry else '未识别有效期', '打印'),
                                 tags=(record_tag,))
        if selected and self.tree.exists(selected):
            self.tree.selection_set(selected)

    def open_on_double_click(self, event):
        row = self.tree.identify_row(event.y)
        if row.startswith('r:') and self.tree.identify_column(event.x) != '#5':
            self.tree.selection_set(row)
            self.open_file()

    def row_click(self, event):
        row = self.tree.identify_row(event.y)
        if row.startswith('r:') and self.tree.identify_column(event.x) == '#5':
            self.tree.selection_set(row)
            self.print_file()

    def report_error(self, exc):
        title = '同名记录提醒' if isinstance(exc, DuplicateNameError) else '稍后再试' if isinstance(exc, BusyError) else '内容已更新' if isinstance(exc, ConflictError) else '操作失败'
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
            name = 'SA'
            while True:
                name = simpledialog.askstring('记录名称', f'项目：{pid}\n填写 SA 或 Qualification，可加 _4样本 或 _Panel2', initialvalue=name)
                if name is None:
                    break
                try:
                    self.library.check_name(self.library.load(pid), name)
                    if any(job[1] == pid and name_key(job[3]) == name_key(name) for job in jobs):
                        raise DuplicateNameError('本次批量导入已有同名记录，请更改名称或取消此项。')
                    jobs.append((path, pid, uuid.uuid4().hex, name.strip(), None))
                    break
                except (DuplicateNameError, ValueError) as exc:
                    self.report_error(exc)
                except OSError as exc:
                    self.report_error(exc)
                    break
        self.start_jobs(jobs)

    def update(self):
        if not self.ready():
            return
        item = self.record()
        if not item:
            return
        pid, r = item
        if r.get('snapshot'):
            self.open_file()
            return
        path = filedialog.askopenfilename(filetypes=[('Word 文档', '*.docx *.doc')])
        if path:
            self.start_jobs([(path, pid, r['id'], r['name'], r['version'])])

    def start_jobs(self, jobs, automatic=False):
        if not jobs:
            return
        self.busy = True
        library = self.library
        self.note.set('正在本机准备 Word 文档，完成后写入资料库；请勿关闭程序。')
        def worker():
            try:
                for source, pid, rid, name, expected in jobs:
                    try:
                        with tempfile.TemporaryDirectory(prefix='assay-') as td:
                            local = Path(td) / ('input' + Path(source).suffix.lower())
                            before = Path(source).stat() if automatic else None
                            shutil.copyfile(source, local)
                            if before is not None:
                                after = Path(source).stat()
                                if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
                                    continue
                                if local.suffix == '.docx':
                                    with zipfile.ZipFile(local) as doc:
                                        doc.getinfo('word/document.xml')
                            if automatic and expected is not None:
                                with local.open('rb') as stream:
                                    digest = hashlib.file_digest(stream, 'sha256').hexdigest()
                                current = next((r for r in library.load(pid)['records'] if r['id'] == rid), None)
                                if current and digest == current.get('digest'):
                                    self.events.put(('captured', source, (before.st_size, before.st_mtime_ns)))
                                    continue
                            pages = None
                            try:
                                pages = count_word_pages(local, td)
                            except (OSError, subprocess.SubprocessError):
                                pass
                            pages = pages or saved_word_pages(local)
                            if automatic and expected is not None:
                                library.capture_live(pid, rid, expected, local, pages,
                                                     datetime.now().strftime('%Y-%m-%d %H:%M'))
                                self.events.put(('captured', source, (before.st_size, before.st_mtime_ns)))
                                self.events.put(('saved', name))
                                continue
                            inbox = source.relative_to(library.root).as_posix().casefold() if automatic else None
                            result = library.commit_record(pid, rid, name, local, pages,
                                                           datetime.now().strftime('%Y-%m-%d %H:%M'), expected, inbox=inbox, live=expected is None, incoming=source)
                            if result:
                                self.events.put(('saved', name))
                    except Exception as exc:
                        self.events.put(('auto_error' if automatic else 'error', exc))
            finally:
                self.events.put(('done',))
        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == 'captured':
                    self.live_done[event[1]] = event[2]
                elif event[0] == 'saved':
                    self.note.set('已安全保存：' + event[1])
                    self.refresh()
                elif event[0] == 'error':
                    self.report_error(event[1])
                    self.note.set('部分操作未完成，请核对提示；原当前版本不会被不完整文件替换。')
                elif event[0] == 'auto_error':
                    self.note.set(f'自动导入暂未完成：{event[1]}')
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
                record = item[1]
                path = self.library.resolve(record['source']) if record.get('snapshot') else self.local_copy(record['source'])
                self.launch(path)
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
                data = self.library.load(pid)
                self.library.require_active(data)
                current = next((x for x in data['records'] if x['id'] == r['id']), None)
                if current is None or current['version'] != r['version']:
                    raise ConflictError('确认期间该记录已更新，请刷新后重新打印。')
                if self.launch(self.local_copy(r['source']), 'print'):
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
                    self.launch(self.local_copy(entry['source']))
                except (OSError, ValueError) as exc:
                    self.report_error(exc)
        ttk.Button(win, text='打开所选历史版本副本', command=view).pack(pady=10)

    def delete_selected(self):
        if not self.ready():
            return
        pid = self.project()
        if not pid:
            return messagebox.showinfo('选择项目或记录', '请选择要删除的项目或记录。')
        try:
            if self.selection().startswith('p:'):
                data = self.library.load(pid)
                if not messagebox.askyesno('删除整个项目',
                        f'{pid}\n包含 {len(data["records"])} 份当前记录及其历史版本。\n\n确定移入回收站？之后可以恢复。'):
                    return
                self.library.delete_project(pid, data['revision'])
            else:
                item = self.record()
                if not item:
                    return
                pid, record = item
                if not messagebox.askyesno('删除记录',
                        f'项目：{pid}\n记录：{record["name"]}\n\n将此记录及其历史版本移入回收站？'):
                    return
                self.library.delete_record(pid, record['id'], record['version'])
            self.refresh()
            self.note.set('已移入回收站，可恢复。文件仍保留，此操作不释放磁盘空间。')
        except (OSError, ValueError, BusyError, ConflictError) as exc:
            self.report_error(exc)

    def recycle_bin(self):
        if not self.ready():
            return
        library = self.library
        try:
            items = library.recycle_items()
        except (OSError, ValueError) as exc:
            return self.report_error(exc)
        win = tk.Toplevel(self.root)
        win.title('回收站')
        win.geometry('850x400')
        ttk.Label(win, text='删除项可恢复。文件仍在库中，不释放磁盘空间；本版不提供永久清除。', padding=12).pack(fill='x')
        listing = tk.Listbox(win, font=('Microsoft YaHei UI', 10))
        listing.pack(fill='both', expand=True, padx=12)
        for item in items:
            listing.insert('end', f'{item["deleted"]} | {item["pid"]} | {item["name"]}')
        if not items:
            listing.insert('end', '回收站为空')
        def restore():
            selected = listing.curselection()
            if not items or not selected:
                return
            item = items[selected[0]]
            if not messagebox.askyesno('恢复', f'恢复 {item["pid"]} / {item["name"]}？', parent=win):
                return
            try:
                library.restore(item['pid'], item['token'], item.get('revision'))
                self.refresh()
                self.note.set('已恢复删除项。')
                win.destroy()
            except (OSError, ValueError, BusyError, ConflictError) as exc:
                self.report_error(exc)
        ttk.Button(win, text='恢复所选项', command=restore).pack(pady=12)

    def close(self):
        if self.busy:
            messagebox.showinfo('正在保存', '请等待当前任务结束后关闭，以免中断上传。')
            return
        self.root.destroy()


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
