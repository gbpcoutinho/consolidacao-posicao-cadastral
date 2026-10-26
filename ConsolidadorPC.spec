# -*- mode: python ; coding: utf-8 -*-
# PyInstaller, modo pasta (--onedir). Gere com scripts\gerar_exe.ps1.
import os

raiz = SPECPATH

a = Analysis(
    [os.path.join(raiz, 'consolidador_app.py')],
    pathex=[raiz],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Nada de testes nem ferramentas de desenvolvimento no executavel.
    excludes=['pytest', '_pytest', 'tests', 'IPython', 'matplotlib', 'jinja2'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ConsolidadorPC',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='ConsolidadorPC',
)
