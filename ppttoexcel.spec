
# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['ppttoexcel2.py'],
    pathex=[],
    binaries=[],
    datas=[('pptGuide2.ui', '.'), ('logger.py', '.')],
    hiddenimports=[
        'PyQt5.sip',
        'openpyxl.styles',
        'openpyxl.workbook',
        'python-pptx',
        'pptx',
        'pandas'
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
    name='ppttoexcelV2',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # UPX 비활성화 (호환성 문제 방지)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='logo.ico',  # 리스트가 아닌 문자열로 수정
)