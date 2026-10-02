# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

# customtkinter (temi/font) e tkinterdnd2 (libreria nativa tkdnd) hanno file di dati
# che PyInstaller non rileva da solo.
datas = collect_data_files('customtkinter') + collect_data_files('tkinterdnd2')
binaries = collect_dynamic_libs('tkinterdnd2')

a = Analysis(
    ['cribis_export.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=['tkinterdnd2', 'openpyxl', 'xlrd', 'xlwt'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='cribis_export',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
