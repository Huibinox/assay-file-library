# Assay File Library

Windows 单机版试验记录文档库，初版 v0.1。

## 功能

- 按 `123456-654321-FSA` 或 `123456-654321-FSA-HGR` 管理项目。
- 每个项目可添加多份记录，例如“验证－4样本”“正式－11样本”。
- 按项目号或记录名称搜索，支持多选导入 Word。
- 导入时复制文件到本地库，更新记录时保留历史版本。
- 手动填写项目最近有效期：6～30 天黄色，0～5 天及已过期红色。
- 使用本机 Microsoft Word 生成 PDF、读取页数；预览或请求系统打印。

## 使用

在 Releases 下载 ZIP，解压到有写入权限的固定文件夹，双击 `AssayFileLibrary.exe`。运行 EXE 无需安装 Python。

源代码运行需要 Python 3 和 tkinter：

```powershell
python app.py
```

自动转换 PDF 需要 Windows 桌面版 Microsoft Word。打印依赖系统关联的文档阅读器和打印机。PDF 阅读器不支持系统打印命令时，请先预览，再手动打印。

## 本地数据

程序在自身目录下创建 `Assay File库/`，保存 SQLite 数据库、Word 副本、PDF 和各历史版本。原文档不会移动。备份时先关闭程序，再复制整个资料库。

该目录被 Git 忽略，仓库及发布包不包含用户实验文档或数据库。每次导入/更新会记录操作时间，该日期不代表文档生效日期。

## 构建 Windows EXE

```powershell
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name AssayFileLibrary app.py
```

EXE 在 `dist/` 目录。

## 初版验证与限制

已检查源码语法、Tk 界面启动、SQLite 基础存储、搜索、临期红色标记及打印操作栏，并完成 EXE 启动检查。

尚未用实际实验记录和打印机验证 Word 转换、分页、排版及纸张输出。初次使用请先用脱敏文档测试。打印按钮只提交请求，不能确认打印完成。

初版适合单机，不支持多人同时写入共享数据库。修改文档后需通过“更新所选记录”重新导入，避免 Word 与 PDF 不一致。
