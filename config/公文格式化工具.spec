# -*- mode: python ; coding: utf-8 -*-

import os
import sys

# 获取项目根目录的绝对路径
spec_dir = SPECPATH
src_dir = os.path.join(spec_dir, '..', 'src')

# 将src目录添加到Python路径
sys.path.insert(0, src_dir)

# 使用PyInstaller工具自动收集win32com所有子模块
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
    ] + collect_submodules('win32com'),
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
    icon=os.path.join(spec_dir, '..', 'assets', 'icons', '公文格式化工具图标设计.ico'),
    distpath=os.path.join(spec_dir, '..'),
)
