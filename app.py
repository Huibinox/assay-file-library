import os
import re
import sqlite3
import shutil
import subprocess
import threading
import uuid
import sys
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

BASE = Path(sys.executable).resolve().parent if getattr(sys,'frozen',False) else Path(__file__).resolve().parent
LIB = BASE / 'Assay File库'
LIB.mkdir(exist_ok=True)
DB = sqlite3.connect(LIB / 'library.db')
DB.execute('CREATE TABLE IF NOT EXISTS projects(id TEXT PRIMARY KEY, expiry TEXT)')
DB.execute('CREATE TABLE IF NOT EXISTS records(id TEXT PRIMARY KEY, project TEXT, name TEXT, source TEXT, pdf TEXT, pages INTEGER, updated TEXT)')
DB.execute('CREATE TABLE IF NOT EXISTS history(record TEXT, source TEXT, pdf TEXT, updated TEXT)')
DB.commit()
PATTERN = r'\d{6}-\d{6}-FSA(?:-HGR)?'


class App:
    def __init__(self, root):
        self.root = root
        root.title('Assay File 库 · 初版')
        root.geometry('1120x700')
        root.minsize(900, 500)
        ttk.Style().configure('Treeview', rowheight=30, font=('Microsoft YaHei UI', 10))
        top = ttk.Frame(root, padding=16)
        top.pack(fill='x')
        ttk.Label(top, text='Assay File 库', font=('Microsoft YaHei UI', 20, 'bold')).pack(side='left')
        ttk.Button(top, text='打开资料文件夹', command=lambda: os.startfile(LIB)).pack(side='right')
        bar = ttk.Frame(root, padding=(16, 0, 16, 12))
        bar.pack(fill='x')
        self.query = tk.StringVar()
        ttk.Label(bar, text='搜索项目号 / 记录名称  ').pack(side='left')
        ttk.Entry(bar, textvariable=self.query, width=35).pack(side='left')
        self.query.trace_add('write', lambda *_: self.refresh())
        for label, fn in [('新增项目', self.add_project), ('导入记录', self.import_record), ('填写最近有效期', self.expiry)]:
            ttk.Button(bar, text=label, command=fn).pack(side='left', padx=6)
        self.tree = ttk.Treeview(root, columns=('pages', 'updated', 'expiry', 'status', 'print'), selectmode='browse')
        self.tree.heading('#0', text='项目号 / 试验记录（展开项目选择记录）')
        self.tree.column('#0', width=440)
        for key, title, width in [('pages','页数',65),('updated','导入 / 更新日期',150),('expiry','最近有效期',130),('status','状态',160)]:
            self.tree.heading(key,text=title)
            self.tree.column(key,width=width,anchor='center')
        self.tree.heading('print',text='操作')
        self.tree.column('print',width=70,anchor='center')
        self.tree.bind('<ButtonRelease-1>',self.row_click)
        self.tree.pack(fill='both', expand=True, padx=16)
        self.tree.tag_configure('yellow', background='#fff0b3')
        self.tree.tag_configure('red', background='#ffd2d2')
        self.tree.bind('<Double-1>', lambda e: self.open_file())
        actions = ttk.Frame(root, padding=16)
        actions.pack(fill='x')
        for label, fn in [('预览 PDF',self.preview),('打开 Word',self.open_file),('打印所选记录',self.print_file),('更新所选记录',self.update),('查看历史',self.history)]:
            ttk.Button(actions,text=label,command=fn).pack(side='left',padx=4)
        self.note = tk.StringVar(value='导入时复制原文件；PDF 转换需要本机安装 Microsoft Word。打印前请预览并确认打印机。')
        ttk.Label(root,textvariable=self.note,padding=(16,0,16,12),wraplength=1050).pack(fill='x')
        self.refresh()
        self.tick()

    def tick(self):
        self.refresh()
        self.root.after(60000,self.tick)

    def selection(self):
        s=self.tree.selection()
        return s[0] if s else None

    def project(self):
        s=self.selection()
        if not s: return None
        return s[2:] if s.startswith('p:') else DB.execute('SELECT project FROM records WHERE id=?',(s[2:],)).fetchone()[0]

    def record(self):
        s=self.selection()
        if not s or not s.startswith('r:'):
            messagebox.showinfo('选择记录','请先展开项目，选择一份试验记录。')
            return None
        return DB.execute('SELECT * FROM records WHERE id=?',(s[2:],)).fetchone()

    def refresh(self):
        selected=self.selection()
        opened={i for i in self.tree.get_children() if self.tree.item(i,'open')}
        self.tree.delete(*self.tree.get_children())
        q=self.query.get().strip().lower()
        for pid,expiry in DB.execute('SELECT * FROM projects ORDER BY id').fetchall():
            rows=DB.execute('SELECT * FROM records WHERE project=? ORDER BY name',(pid,)).fetchall()
            matching=[r for r in rows if q in pid.lower() or q in r[2].lower()]
            if q and q not in pid.lower() and not matching: continue
            tag=''; status='未设置'
            if expiry:
                days=(date.fromisoformat(expiry)-date.today()).days
                status=f'已过期 {-days} 天' if days<0 else f'剩余 {days} 天'
                tag='red' if days<=5 else 'yellow' if days<=30 else ''
            parent='p:'+pid
            self.tree.insert('', 'end',iid=parent,text=pid,values=('', '', expiry or '未设置',status),tags=(tag,),open=bool(q) or parent in opened)
            for r in matching:
                self.tree.insert(parent,'end',iid='r:'+r[0],text=r[2],values=(r[5] or '—',r[6], '', 'PDF 就绪' if r[4] else '待转换 PDF','打印'))
        if selected and self.tree.exists(selected): self.tree.selection_set(selected)

    def row_click(self,event):
        row=self.tree.identify_row(event.y)
        if row.startswith('r:') and self.tree.identify_column(event.x)=='#5':
            self.tree.selection_set(row)
            self.print_file()

    def add_project(self):
        pid=simpledialog.askstring('新增项目','项目号：123456-654321-FSA 或带 -HGR')
        if not pid:return
        pid=pid.strip().upper()
        if not re.fullmatch(PATTERN,pid):return messagebox.showerror('格式错误','请输入六位公司号-六位项目号-FSA，可选后缀 -HGR。')
        DB.execute('INSERT OR IGNORE INTO projects VALUES (?,NULL)',(pid,)); DB.commit(); self.refresh()
        self.tree.selection_set('p:'+pid)

    def expiry(self):
        pid=self.project()
        if not pid:return messagebox.showinfo('选择项目','请先选择项目。')
        old=DB.execute('SELECT expiry FROM projects WHERE id=?',(pid,)).fetchone()[0]
        value=simpledialog.askstring('最近有效期','输入 YYYY-MM-DD；留空清除',initialvalue=old or '')
        if value is None:return
        try: value=date.fromisoformat(value.strip()).isoformat() if value.strip() else None
        except ValueError:return messagebox.showerror('日期错误','请使用 YYYY-MM-DD，例如 2026-12-31。')
        DB.execute('UPDATE projects SET expiry=? WHERE id=?',(value,pid));DB.commit();self.refresh()

    def import_record(self):
        files=filedialog.askopenfilenames(title='导入 Word（可多选）',filetypes=[('Word 文档','*.docx *.doc')])
        for path in files:
            match=re.search(PATTERN,Path(path).stem.upper())
            pid=match.group() if match else self.project()
            if not pid:
                pid=simpledialog.askstring('所属项目',f'{Path(path).name}\n请输入项目号')
                pid=pid.strip().upper() if pid else ''
            if not re.fullmatch(PATTERN,pid):
                messagebox.showerror('未导入','项目号格式不正确。');continue
            name=simpledialog.askstring('记录名称','例如：验证－4样本、正式－11样本',initialvalue=Path(path).stem)
            if not name or not name.strip():continue
            DB.execute('INSERT OR IGNORE INTO projects VALUES (?,NULL)',(pid,));DB.commit()
            self.save(path,pid,name.strip())

    def save(self,path,pid,name,old=None):
        rid=old[0] if old else uuid.uuid4().hex
        folder=LIB/pid/rid/datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        folder.mkdir(parents=True)
        dest=folder/Path(path).name
        shutil.copy2(path,dest)
        now=datetime.now().strftime('%Y-%m-%d %H:%M')
        if old:DB.execute('INSERT INTO history VALUES (?,?,?,?)',(rid,old[3],old[4],old[6]))
        DB.execute('INSERT OR REPLACE INTO records VALUES (?,?,?,?,?,?,?)',(rid,pid,name,str(dest),'',None,now));DB.commit();self.refresh()
        self.note.set('文档已保存，正在使用 Word 转换 PDF…')
        def convert():
            pdf=dest.with_suffix('.pdf')
            script="$ErrorActionPreference='Stop'; $w=$null; $d=$null; try { $w=New-Object -ComObject Word.Application; $w.Visible=$false; $w.AutomationSecurity=3; $d=$w.Documents.Open($args[0],$false,$true); $d.Repaginate(); $n=$d.ComputeStatistics(2); $d.ExportAsFixedFormat($args[1],17); Write-Output $n } finally { if($d){$d.Close(0)}; if($w){$w.Quit()} }"
            scriptpath=folder/'convert.ps1'
            scriptpath.write_text(script,encoding='utf-8-sig')
            try:
                result=subprocess.run(['powershell','-NoProfile','-ExecutionPolicy','Bypass','-File',str(scriptpath),str(dest),str(pdf)],capture_output=True,text=True,timeout=180,creationflags=0x08000000)
                if result.returncode or not pdf.exists():raise RuntimeError('Word 转换未完成，请确认安装 Microsoft Word，且文档可正常打开。')
                pages=int(result.stdout.strip().splitlines()[-1])
                self.root.after(0,lambda:self.converted(rid,dest,pdf,pages))
            except Exception:
                self.root.after(0,lambda:self.note.set('Word 已保存在库中；PDF 转换失败。可打开 Word 手动打印，或安装 Word 后重新更新记录。'))
        threading.Thread(target=convert,daemon=True).start()

    def converted(self,rid,dest,pdf,pages):
        DB.execute('UPDATE records SET pdf=?,pages=? WHERE id=? AND source=?',(str(pdf),pages,rid,str(dest)));DB.commit();self.refresh();self.note.set('PDF 已生成，页数已更新。')

    def update(self):
        r=self.record()
        if not r:return
        path=filedialog.askopenfilename(filetypes=[('Word 文档','*.docx *.doc')])
        if path:self.save(path,r[1],r[2],r)

    def launch(self,path,verb='open'):
        try:os.startfile(path,verb)
        except OSError as e:messagebox.showerror('无法打开',str(e))

    def open_file(self):
        r=self.record()
        if r:self.launch(r[3])

    def preview(self):
        r=self.record()
        if r:
            if r[4]:self.launch(r[4])
            else:messagebox.showinfo('PDF 尚未就绪','可先打开 Word 查看。')

    def print_file(self):
        r=self.record()
        if r and messagebox.askokcancel('确认打印',f'项目：{r[1]}\n记录：{r[2]}\n\n将通过系统关联程序发送到默认打印机，请确认所选记录。'):
            self.launch(r[4] or r[3],'print')
            self.note.set('已请求系统关联程序打印；请在打印队列确认是否成功。')

    def history(self):
        r=self.record()
        if r:os.startfile(LIB/r[1]/r[0])


if __name__=='__main__':
    root=tk.Tk()
    App(root)
    root.mainloop()
