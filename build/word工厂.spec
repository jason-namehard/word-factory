# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['E:/Zspace/projects/word-factory/wordfactory_app.py'],
    pathex=['E:/Zspace/projects/word-factory'],
    binaries=[],
    datas=[('E:/Zspace/projects/word-factory/wordfactory/gui/web/index.html', 'wordfactory/gui/web')],
    hiddenimports=['wordfactory', 'wordfactory.gui.server', 'wordfactory.paths'],
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
    name='word工厂',
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
)
