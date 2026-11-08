"""Per-project JSON storage with SMB-compatible Windows byte-range locks."""
import contextlib
import json
import os
from pathlib import Path
import re
import shutil
import tempfile
import uuid

PATTERN = r'\d{6}-\d{6}-FSA(?:-HGR)?'


class BusyError(Exception):
    pass


class ConflictError(Exception):
    pass


def atomic_json(path, data):
    path = Path(path)
    fd, tmp = tempfile.mkstemp(prefix='.pending-', dir=path.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)


class Library:
    def __init__(self, root):
        self.root = Path(root).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self.projects_dir = self.root / 'projects'
        self.projects_dir.mkdir(exist_ok=True)

    def folder(self, pid):
        if not re.fullmatch(PATTERN, pid):
            raise ValueError('项目号格式不正确。')
        return self.projects_dir / pid

    def load(self, pid):
        path = self.folder(pid) / 'project.json'
        if not path.exists():
            return {'schema': 2, 'id': pid, 'expiry': None, 'revision': 0, 'records': []}
        data = json.loads(path.read_text(encoding='utf-8'))
        if data.get('schema') != 2 or data.get('id') != pid:
            raise ValueError(f'项目数据格式异常：{pid}')
        return data

    def list_projects(self):
        # Never return an incomplete list if a read fails.
        return [self.load(p.name) for p in sorted(self.projects_dir.iterdir())
                if p.is_dir() and re.fullmatch(PATTERN, p.name) and (p / 'project.json').exists()]

    @contextlib.contextmanager
    def lock(self, pid):
        folder = self.folder(pid)
        folder.mkdir(parents=True, exist_ok=True)
        f = open(folder / '.write.lock', 'a+b')
        locked = False
        try:
            if os.name == 'nt':
                import msvcrt
                f.seek(0)
                try:
                    msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
                except OSError as exc:
                    raise BusyError('该项目正在编辑，请稍后再试。') from exc
            else:
                import fcntl
                try:
                    fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                except OSError as exc:
                    raise BusyError('该项目正在编辑，请稍后再试。') from exc
            locked = True
            yield
        finally:
            if locked:
                if os.name == 'nt':
                    f.seek(0)
                    msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    fcntl.flock(f.fileno(), fcntl.LOCK_UN)
            f.close()

    def write(self, data):
        data['revision'] += 1
        atomic_json(self.folder(data['id']) / 'project.json', data)

    def create(self, pid):
        with self.lock(pid):
            data = self.load(pid)
            if not (self.folder(pid) / 'project.json').exists():
                self.write(data)

    def set_expiry(self, pid, expiry, expected_revision):
        with self.lock(pid):
            data = self.load(pid)
            if data['revision'] != expected_revision:
                raise ConflictError('项目已被其他电脑更新，请刷新后重新填写有效期。')
            data['expiry'] = expiry
            self.write(data)

    def resolve(self, relative):
        path = (self.root / relative).resolve()
        if not path.is_relative_to(self.root):
            raise ValueError('文档路径不在资料库内。')
        return path

    def commit_record(self, pid, rid, name, source, pdf, pages, updated, expected=None):
        """Copy immutable files first; atomically publish only after all copies succeed."""
        with self.lock(pid):
            data = self.load(pid)
            old = next((r for r in data['records'] if r['id'] == rid), None)
            if expected is not None and (old is None or old['version'] != expected):
                raise ConflictError('该记录已被其他电脑更新，本次不会覆盖。请刷新并核对最新版本。')
            if expected is None and old is not None:
                raise ConflictError('记录编号已存在，请重新导入。')
            version = uuid.uuid4().hex
            folder = self.folder(pid) / 'records' / rid / version
            folder.mkdir(parents=True)
            dest = folder / ('document' + Path(source).suffix.lower())
            shutil.copyfile(source, dest)
            pdf_dest = folder / 'document.pdf'
            if pdf:
                shutil.copyfile(pdf, pdf_dest)
            entry = {'id': rid, 'name': name, 'version': version,
                     'source': dest.relative_to(self.root).as_posix(),
                     'pdf': pdf_dest.relative_to(self.root).as_posix() if pdf else '',
                     'pages': pages if pdf else None, 'updated': updated,
                     'history': [] if old is None else old.get('history', []) +
                     [{k: v for k, v in old.items() if k != 'history'}]}
            if old:
                data['records'][data['records'].index(old)] = entry
            else:
                data['records'].append(entry)
            self.write(data)
            return entry
