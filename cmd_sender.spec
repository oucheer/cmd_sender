# -*- mode: python ; coding: utf-8 -*-
import os

block_cipher = None

# 获取程序根目录
root_dir = os.path.dirname(os.path.abspath(SPEC))

a = Analysis(
    ['complete_command_sender.py'],
    pathex=[root_dir],
    binaries=[],
    datas=[
        ('modules', 'modules'),
        ('cmd_sender.ico', '.'),
    ],
    hiddenimports=[
        'pyautogui',
        'pyperclip',
        'serial',
        'serial.tools.list_ports',
        'win32gui',
        'win32con',
        'win32process',
        'psutil',
        'keyboard',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='cmd_sender',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='cmd_sender.ico',
)
