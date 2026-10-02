# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules

# customtkinter (temi/font), tkinterdnd2 (libreria nativa tkdnd) e selenium (script JS e
# selenium-manager) hanno file di dati che PyInstaller non rileva da solo.
# selenium importa i moduli del browser in modo dinamico (lazy): vanno inclusi tutti.
datas = collect_data_files('customtkinter') + collect_data_files('tkinterdnd2') + collect_data_files('selenium')
binaries = collect_dynamic_libs('tkinterdnd2')

a = Analysis(
    ['cribis_export.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=['tkinterdnd2', 'openpyxl', 'xlrd', 'xlwt'] + collect_submodules('selenium'),
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
