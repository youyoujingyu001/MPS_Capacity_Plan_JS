# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec —— macOS Apple Silicon (arm64)
在 Mac 上执行:  pyinstaller MPS-Capacity.app.spec --noconfirm

目标产物: dist/MPS-Capacity.app  (arm64, 内置 pandas/openpyxl/numpy)
"""
import os, sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None
base = os.path.abspath('.')

datas = []
# 核心脚本打包进去 (GUI 通过 sys.argv[0] 同目录引用)
datas.append((os.path.join(base, 'MPS-Capacity_V2.0.py'), '.'))
# pandas 的 data 目录 (有 i18n / 时区数据); 2 元组由 Analysis 自动转为 3 元组
datas += collect_data_files('pandas')

hiddenimports = [
    'openpyxl', 'et_xmlfile',
    'pandas',
]
hiddenimports += collect_submodules('pandas')
hiddenimports += collect_submodules('openpyxl')

binaries = []

a = Analysis(
    ['MPS-GUI.py'],
    pathex=[base],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter.test', 'pandas.tests', 'numpy.tests'],
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
    name='MPS-Capacity',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,   # 纯窗口 App: 不弹终端窗口 (核心脚本在进程内 import, 日志显示在 GUI 里)
    disable_windowed_traceback=False,
    argv_emulation=False,
    # 目标架构由 CI 环境变量 TARGET_ARCH 决定 (arm64 / x86_64); 未设置时用当前构建机架构
    target_arch=os.environ.get('TARGET_ARCH') or None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='MPS-Capacity',
)

app = BUNDLE(
    coll,
    name='MPS-Capacity.app',
    icon=None,
    bundle_identifier='com.user.mps-capacity',
    info_plist={
        'NSHighResolutionCapable': True,
        'LSUIElement': False,
        'NSRequiresAquaSystemAppearance': False,
        'LSMinimumSystemVersion': '12.0',
    },
)
