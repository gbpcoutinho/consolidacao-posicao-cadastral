# -*- mode: python ; coding: utf-8 -*-
# PyInstaller, modo pasta (--onedir). Gere com scripts\gerar_exe.ps1.
import os

from PyInstaller.utils.hooks import collect_submodules

raiz = SPECPATH

# Logo e icone (gerados por scripts/gerar_recursos.py, fora do git). Sem eles o
# executavel sai igual, so sem icone e sem logo na janela.
recursos = os.path.join(raiz, 'app', 'recursos')
tem_recursos = os.path.isdir(recursos)
icone = os.path.join(recursos, 'icone.ico')
tem_icone = os.path.isfile(icone)
if not tem_icone:
    print('AVISO: %s nao existe; o executavel sai sem icone proprio. '
          'Rode scripts/gerar_recursos.py antes.' % icone)

a = Analysis(
    [os.path.join(raiz, 'consolidador_app.py')],
    pathex=[raiz],
    binaries=[],
    datas=[(recursos, os.path.join('app', 'recursos'))] if tem_recursos else [],
    # msoffcrypto importa os formatos (ooxml, ole...) por dentro do pacote.
    hiddenimports=collect_submodules('msoffcrypto') + ['tkinterdnd2'],
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
    icon=icone if tem_icone else None,
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
