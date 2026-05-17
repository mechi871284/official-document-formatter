# -*- mode: python ; coding: utf-8 -*-

import os
import sys

spec_dir = SPECPATH
src_dir = os.path.join(spec_dir, '..', 'src')
sys.path.insert(0, src_dir)

from PyInstaller.utils.hooks import collect_submodules

a = Analysis(
    [os.path.join(src_dir, 'main.py')],
    pathex=[src_dir],
    binaries=[],
    datas=[],
    hiddenimports=[
        'formatters',
        'formatters.base_handler',
        'formatters.docx_handler',
        'formatters.doc_handler',
        'formatters.wps_handler',
        'formatters.router',
        'formatters.robustness',
        'formatters.cross_platform',
    ],
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
    name='公文格式化工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
    distpath=os.path.join(spec_dir, '..'),
)
