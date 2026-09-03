# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Video_Resizer.py'],
    pathex=[],
    binaries=[],
    datas=[('tesseract', '.'), ('tessdata', 'tessdata')],
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name='Video Resizer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app_icon.png'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Video Resizer',
)
app = BUNDLE(
    coll,
    name='Video Resizer.app',
    icon='app_icon.png',
    bundle_identifier=None,
)
