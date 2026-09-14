# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['office_app/main.py'],
    pathex=['.'],
    binaries=[],
    datas=[('templates/article.xsl', 'templates'), ('config/journal_defaults.json', 'config'), ('C:/Users/Sec. Investigadores/SCIELO_XML/.venv/Lib/site-packages/customtkinter', 'customtkinter')],
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
    name='Convertidor SCIELO FACPyA',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='Convertidor SCIELO FACPyA',
)
