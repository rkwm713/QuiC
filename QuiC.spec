# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['__main__.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('main.py', '.'),
        ('compare.py', '.'),
        ('spida_writer.py', '.'),
        ('editable_tree.py', '.'),
        ('logo.png', '.'),
    ],
    hiddenimports=[
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'ttkbootstrap',
        'ttkbootstrap.constants',
        'tkintermapview',
        'pandas',
        'openpyxl',
        'et_xmlfile',
        'json',
        'pathlib',
        'traceback',
        'sys',
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageTk'
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
    name='QuiC',
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
    icon=['assets\\quic.ico'],
)
