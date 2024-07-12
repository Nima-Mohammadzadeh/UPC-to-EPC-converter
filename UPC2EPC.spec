# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['UPC2EPC.py'],
    pathex=['C:\\Users\\Jason\\OneDrive\\Documents\\UPC2EPC Convertor'],
    binaries=[],
    datas=[
	('C:\\Users\\Jason\\OneDrive\\Documents\\UPC2EPC Convertor\\Roll Tracker v.3.xlsx', '.'),
        ('C:\\Users\\Jason\\OneDrive\\Documents\\UPC2EPC Convertor\\Templates', 'Templates'),
        ('C:\\Users\\Jason\\OneDrive\\Documents\\UPC2EPC Convertor\\download.png', '.')
	],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='UPC2EPC',
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
)
