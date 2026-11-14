import hashlib
"""Per-project JSON storage with SMB-compatible Windows byte-range locks."""
import contextlib
import json
import os
from pathlib import Path
import re
import shutil
import tempfile
import uuid
from datetime import datetime

PATTERN = r'\d{6}-\d{6}-FSA(?:-HGR)?'


def name_key(name):
    return name.strip().casefold()


class BusyError(Exception):
    pass


class ConflictError(Exception):
    pass


class DuplicateNameError(ConflictError):
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
        if data.get('schema') not in (2, 3) or data.get('id') != pid:
            raise ValueError(f'项目数据格式异常：{pid}')
        return data

    def list_projects(self, include_deleted=False):
        # Never return an incomplete list if a read fails.
        projects = [self.load(p.name) for p in sorted(self.projects_dir.iterdir())
                    if p.is_dir() and re.fullmatch(PATTERN, p.name) and (p / 'project.json').exists()]
        return [p for p in projects if include_deleted or not p.get('deleted')]

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
        data['schema'] = 3
        data['revision'] += 1
        atomic_json(self.folder(data['id']) / 'project.json', data)

    def create(self, pid):
        with self.lock(pid):
            data = self.load(pid)
            self.require_active(data)
            if not (self.folder(pid) / 'project.json').exists():
                self.write(data)

    def set_expiry(self, pid, expiry, expected_revision):
        with self.lock(pid):
            data = self.load(pid)
            self.require_active(data)
            if data['revision'] != expected_revision:
                raise ConflictError('项目已被其他电脑更新，请刷新后重新填写有效期。')
            data['expiry'] = expiry
            self.write(data)

    def resolve(self, relative):
        path = (self.root / relative).resolve()
        if not path.is_relative_to(self.root):
            raise ValueError('文档路径不在资料库内。')
        return path

    def commit_record(self, pid, rid, name, source, pages, updated, expected=None, *, inbox=None, live=False, incoming=None):
        """Copy immutable files first; atomically publish only after all copies succeed."""
        with self.lock(pid):
            data = self.load(pid)
            self.require_active(data)
            if inbox is not None and inbox in data.get('inbox_imports', []):
                return None
            old = next((r for r in data['records'] if r['id'] == rid), None)
            if expected is not None and (old is None or old['version'] != expected):
                raise ConflictError('该记录已被其他电脑更新，本次不会覆盖。请刷新并核对最新版本。')
            if expected is None and old is not None:
                raise ConflictError('记录编号已存在，请重新导入。')
            # An existing legacy duplicate may still update its own document.
            if old is None or name_key(old['name']) != name_key(name):
                self.check_name(data, name, rid)
            version = uuid.uuid4().hex
            folder = self.folder(pid) / 'records' / rid / version
            folder.mkdir(parents=True)
            dest = folder / ('document' + Path(source).suffix.lower())
            shutil.copyfile(source, dest)
            entry = {'id': rid, 'name': name, 'version': version,
                     'source': dest.relative_to(self.root).as_posix(),
                     'pages': pages, 'updated': updated,
                     'history': [] if old is None else old.get('history', []) +
                     [{k: v for k, v in old.items() if k != 'history'}]}
            if live:
                if not re.fullmatch(r'(?:SA|Qualification)(?:_[^<>:"/\\|?*\x00-\x1f]+)?', name) or name.endswith((' ', '.')):
                    raise ValueError('用途须为 SA 或 Qualification，补充说明须符合 Windows 文件命名规则。')
                display = self.root / pid / (pid + '_' + name + dest.suffix)
                display.parent.mkdir(exist_ok=True)
                adopted = incoming is not None and Path(incoming).resolve() == display.resolve()
                if not adopted:
                    if display.exists():
                        raise DuplicateNameError(f'最新版文件已存在：{display.name}')
                    fd, temporary = tempfile.mkstemp(dir=display.parent)
                    os.close(fd)
                    try:
                        shutil.copyfile(dest, temporary)
                        os.rename(temporary, display)
                    finally:
                        Path(temporary).unlink(missing_ok=True)
                with dest.open('rb') as stream:
                    digest = hashlib.file_digest(stream, 'sha256').hexdigest()
                entry.update(snapshot=entry['source'], source=display.relative_to(self.root).as_posix(), digest=digest)
            if old:
                data['records'][data['records'].index(old)] = entry
            else:
                data['records'].append(entry)
            if inbox is not None:
                data.setdefault('inbox_imports', []).append(inbox)
            try:
                self.write(data)
            except Exception:
                if live and not adopted:
                    display.unlink(missing_ok=True)
                raise
            return entry

    def capture_live(self, pid, rid, expected, source, pages, updated):
        with self.lock(pid):
            data = self.load(pid)
            self.require_active(data)
            entry = next((r for r in data['records'] if r['id'] == rid), None)
            if not entry or entry['version'] != expected:
                raise ConflictError('记录已变化，下次扫描将重试。')
            with source.open('rb') as stream:
                digest = hashlib.file_digest(stream, 'sha256').hexdigest()
            if digest == entry['digest']:
                return
            current = self.resolve(entry['source'])
            previous = self.resolve(entry['snapshot'])
            version = uuid.uuid4().hex
            history = current.parent / 'history' / current.stem
            history.mkdir(parents=True, exist_ok=True)
            archived = history / (entry['updated'].replace(':', '').replace(' ', '_') + '_' + expected[:8] + current.suffix)
            shutil.copyfile(previous, archived)
            snapshot = previous.with_name(version + current.suffix)
            shutil.copyfile(source, snapshot)
            entry['history'].append({'source': archived.relative_to(self.root).as_posix(),
                                     'updated': entry['updated'], 'name': entry['name']})
            entry.update(version=version, snapshot=snapshot.relative_to(self.root).as_posix(),
                         digest=digest, pages=pages, updated=updated)
            self.write(data)

    @staticmethod
    def require_active(data):
        if data.get('deleted'):
            raise ConflictError('项目已在回收站中，请先恢复项目。')

    @staticmethod
    def check_name(data, name, exclude_id=None):
        if not name.strip():
            raise ValueError('记录名称不能为空。')
        if any(r['id'] != exclude_id and name_key(r['name']) == name_key(name)
               for r in data['records']):
            raise DuplicateNameError(f'项目 {data["id"]} 已有同名记录“{name.strip()}”。\n'
                                     '如需替换文档，请使用“更新所选记录”；如为不同记录，请更改名称。')

    def move_live_files(self, data, records, restoring=False):
        moved = []
        try:
            for record in records:
                if not record.get('snapshot'):
                    continue
                if restoring:
                    destination = self.resolve(record['source'])
                    origin = self.resolve(record.pop('recycled_source'))
                else:
                    origin = self.resolve(record['source'])
                    destination = self.folder(data['id']) / 'trash' / uuid.uuid4().hex / origin.name
                    record['recycled_source'] = destination.relative_to(self.root).as_posix()
                if destination.exists():
                    raise DuplicateNameError(f'目标文件已存在：{destination.name}')
                destination.parent.mkdir(parents=True, exist_ok=True)
                os.rename(origin, destination)
                moved.append((origin, destination))
            self.write(data)
        except Exception:
            for origin, destination in reversed(moved):
                os.rename(destination, origin)
            raise

    def delete_project(self, pid, expected_revision):
        with self.lock(pid):
            data = self.load(pid)
            self.require_active(data)
            if data['revision'] != expected_revision:
                raise ConflictError('项目已更新，请刷新后再删除。')
            data['deleted'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.move_live_files(data, data['records'])

    def delete_record(self, pid, rid, expected_version):
        with self.lock(pid):
            data = self.load(pid)
            self.require_active(data)
            record = next((r for r in data['records'] if r['id'] == rid), None)
            if not record or record['version'] != expected_version:
                raise ConflictError('记录已更新或删除，请刷新后再操作。')
            data['records'].remove(record)
            data.setdefault('trash', []).append({'token': uuid.uuid4().hex, 'record': record,
                'deleted': datetime.now().strftime('%Y-%m-%d %H:%M:%S')})
            self.move_live_files(data, [record])

    def recycle_items(self):
        items = []
        for data in self.list_projects(include_deleted=True):
            if data.get('deleted'):
                items.append({'pid': data['id'], 'token': None, 'name': '整个项目',
                              'deleted': data['deleted'], 'revision': data['revision']})
            else:
                for item in data.get('trash', []):
                    items.append({'pid': data['id'], 'token': item['token'],
                                  'name': item['record']['name'], 'deleted': item['deleted']})
        return items

    def restore(self, pid, token=None, expected_revision=None):
        with self.lock(pid):
            data = self.load(pid)
            if token is None:
                if not data.get('deleted') or data['revision'] != expected_revision:
                    raise ConflictError('回收站内容已变化，请重新打开回收站。')
                data.pop('deleted')
                for record in data['records']:
                    record['version'] = uuid.uuid4().hex
            else:
                self.require_active(data)
                item = next((t for t in data.get('trash', []) if t['token'] == token), None)
                if item is None:
                    raise ConflictError('该记录已被恢复，请刷新。')
                record = dict(item['record'])
                self.check_name(data, record['name'])
                if any(r['id'] == record['id'] for r in data['records']):
                    raise ConflictError('记录编号冲突，无法恢复。')
                record['version'] = uuid.uuid4().hex
                data['records'].append(record)
                data['trash'].remove(item)
            self.move_live_files(data, data['records'] if token is None else [record], restoring=True)
