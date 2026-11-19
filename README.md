# Assay File Library

A Windows document library for assay records, with project search, version history, expiry reminders, and printing. The v0.3.3 shared-drive trial described here is intended for small teams using a company SMB file server.

The software interface remains in Chinese. Chinese control labels and filename examples below refer to the existing application.

## Download and use

Download the latest ZIP from [Releases](https://github.com/Huibinox/assay-file-library/releases). Extract a separate copy on each computer, run `python app.py` using Python 3 with Tkinter, and click Select Library (`选择资料库`) in the top-right corner. Everyone selects the same shared folder.

Run the program locally and keep the library on the shared drive. Each computer stores its selected path in `%LOCALAPPDATA%/AssayFileLibrary/settings.json`. If the share becomes unavailable, the application asks you to reconnect rather than switching libraries automatically.

Word files containing only a project number are imported as SA. Notes are retained as SA_note; explicit validation markers are classified as Qualification. Conflicting purpose markers require manual confirmation.

## v0.3.0: Edit the shared library directly

Run `python app.py` and select the shared library. Try a test folder first. During import, enter a purpose such as SA, Qualification, `SA_4样本`, or Qualification_Panel2. Word files matching the naming rules can also be detected automatically.

The current document is stored at `<library>/<project number>/<project number>_<purpose>[_<additional note>].docx`. Double-click a record or open its current file in File Explorer, then save directly in Word.

While running, the application scans every 15 seconds. It updates page counts and reagent expiry dates after two consecutive checks show a stable file. Older versions are stored in the project's `history` directory. Intermediate saves made while the application is closed or between scans may not be archived.

Delete Selected (`删除所选项`) moves records or projects to the recycle bin. Restore returns Word files to the project folder. Externally deleted or renamed files appear as missing; history remains available, and renamed files are not merged automatically.

Records using the older layout still open as copies. New imports use the new layout. All team members should use the same updated version.

## Earlier versions

### Automatic Word detection and page counts

Place .docx or .doc files in the library root with a project number in the filename, for example `123456-654321-FSA 正式4样本.docx`. Files can also be placed in `projects/<project number>/`. While running, the application scans every 15 seconds and imports files after two consecutive stable checks.

Original dropped files are retained. Each source path is automatically imported once; use Update Selected Record (`更新所选记录`) to update an existing record. Duplicate names and missing project numbers are reported in the footer.

DOCX page counts use the value saved by Word. Missing page-count metadata and legacy DOC files display a dash.

### v0.2.2: Duplicate record names

Current record names must be unique within a project, ignoring leading/trailing spaces and English letter case. Rename or cancel an import when prompted. Duplicate names within a batch are also checked. The name is checked again under the project lock during saving to prevent concurrent duplicate imports.

Use Update Selected Record to replace an existing file. Different projects may use the same record name. Restore also checks for duplicate current names; conflicting items remain in the recycle bin. Pre-existing duplicates are not merged or deleted automatically and can still be updated. Upgrade every computer so older clients cannot bypass these checks.

### v0.2.1: Deletion and recycle bin

Select a project or record and click Delete Selected, then confirm. Deleting a whole project hides its current records and history. Open Recycle Bin (`回收站`) to restore items. Deletion and restoration use project locks and conflict checks.

Deletion is recoverable and does not free disk space; this version has no permanent purge. Restoring a project does not restore records that had already been deleted individually.

When upgrading from v0.2.0, select the existing shared library without reimporting. The newer writer upgrades projects to schema 3, which older versions refuse to read so they cannot ignore deletion state. Close old clients, back up the shared library, and upgrade all computers together. Do not write to the upgraded library using an older version.

## Earlier-version features

- Project identifiers follow a six-digit company number, a six-digit project number, and FSA, with an optional -HGR suffix.
- Projects can contain multiple records, such as validation records for four samples or routine records for eleven samples.
- Search by project number or record name; import multiple Word files and print individual records.
- Enter the nearest project expiry date manually: 6–30 days remaining is yellow; 0–5 days and expired dates are red.
- Display Word page counts and preserve older versions on update.
- Open Word as a local temporary copy to protect the shared current file.
- Refresh every 15 seconds or manually. If a record changes during print confirmation, review it again.

The copy-based opening behaviour above describes the earlier layout; v0.3.0 introduces direct editing for the new layout.

## Shared-drive storage

In the earlier layout, each project uses `projects/<project number>/project.json`, with Word documents and versions in its `records/` directory. Document paths are relative to the library, allowing different drive mappings on different computers.

Writes use Windows file-range locks at project level. Different projects can be updated independently; a busy project asks users to retry. The operating system holds the lock and releases it when the process exits. It does not depend on the operator's name. Do not delete `.write.lock` manually.

Word files are prepared locally and copied in full to the shared library before the project index is atomically replaced. Document versions and project changes are checked for conflicts; stale updates do not silently overwrite newer ones.

All five computers must use the updated version and its write mechanism. Do not edit project JSON or overwrite library files manually. The server must support SMB file-range locks and same-directory renaming, and users need permission to create, write, and rename files. This version has no account permission system. Company share permissions apply, and users with direct file access can still modify files in File Explorer.

## Word, printing, and backups

Printing depends on the Windows Word file association and printer. A submitted request does not confirm physical output. If the system print command is unsupported, open the Word copy and print manually.

For the earlier copy-based workflow, reimport edited documents with Update Selected Record. The update date records the import/update operation, not the effective date of a controlled document.

Back up the entire library, preferably while users have stopped editing, or use a consistent server snapshot. Preview and printing create copies in the local system temporary directory; clean these according to company endpoint policies.

## Upgrading from v0.1

The old `library.db` is neither modified nor deleted and is no longer used by the newer version. Preserve a backup. Have one person reimport current Word documents into a new library and enter expiry dates. Old version history remains in the old library. Begin shared maintenance only after everyone has switched. This is not an automatic migration tool for the old SQLite library.

## Run from source and build

Requires Python 3 and tkinter.

```powershell
python app.py
```

Optional build:

```powershell
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name AssayFileLibrary app.py
```

Tests:

```powershell
python -m unittest discover -v
```

## Verification scope

Local Windows checks covered cross-process write locks, lock release after abnormal exits, stale-version conflicts, independent project writes, preservation of the current version after copy/index replacement failures, relative paths, and library relocation. The interface and EXE startup were also checked.

Network interruption, server restart, and five simultaneous clients have not been verified on the company's actual SMB server. Actual assay-record pagination and physical printing have also not been verified. Before shared use, test de-identified documents on two computers: view while another client updates, update the same record concurrently, and disconnect/reconnect the share. Server-side lock recovery after a network interruption depends on the SMB implementation; local checks do not validate the server.
