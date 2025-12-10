# -*- mode: python ; coding: utf-8 -*-
"""
MyNote PyInstaller 打包配置文件
使用方法: pyinstaller build_config.py
"""

block_cipher = None

a = Analysis(
    ['note.py'],  # 主程序入口
    pathex=[],
    binaries=[],
    datas=[
        ('assets/note.ico', 'assets'),  # 包含图标文件到打包
    ],
    hiddenimports=[
        'win32api',
        'win32con',
        'win32gui',
        'ctypes.wintypes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'PIL',
        'PyQt5',
        'PyQt6',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='MyNote',  # 生成的exe文件名
    debug=False,  # 不输出调试信息
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # 启用UPX压缩
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    target_arch=None,
    cofile_name=None,
    icon='assets/note.ico',  # 程序图标
    version_file='version_info.txt'  # 版本信息文件
)
