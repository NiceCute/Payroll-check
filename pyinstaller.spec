# -*- mode: python ; coding: utf-8 -*-
"""
pyinstaller.spec - Cấu hình đóng gói ChiHoLuongTool V1 thành ứng dụng Windows (.exe).
Chạy bằng lệnh:
    pyinstaller --noconfirm --clean pyinstaller.spec
hoặc chạy trực tiếp file `build.bat`.
"""

from pathlib import Path

block_cipher = None
project_root = Path(SPECPATH).resolve()

datas = [
    (str(project_root / "templates" / "*.json"), "templates"),
]

hiddenimports = [
    "core",
    "core.models",
    "core.normalize",
    "core.excel_reader",
    "core.source_reader",
    "core.target_reader",
    "core.template_store",
    "core.reconciler",
    "core.safety",
    "core.excel_writer",
    "core.verifier",
    "core.reporter",
    "core.workflow",
    "gui",
    "gui.main_window",
    "gui.preview_table",
    "gui.mapping_dialog",
    "openpyxl",
    "xlrd",
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
]

a = Analysis(
    [str(project_root / "app.py")],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "mypy", "pyflakes"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ChiHoLuongTool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="ChiHoLuongTool",
)
