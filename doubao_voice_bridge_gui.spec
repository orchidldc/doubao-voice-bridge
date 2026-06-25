# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['doubao_voice_bridge_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('tools', 'tools'), ('assets/doubao_d.ico', 'assets')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'PyQt5',
        'PySide6',
        'matplotlib',
        'IPython',
        'cv2',
        'numpy',
        'PIL',
        'cryptography',
        'OpenSSL',
        'bcrypt',
    ],
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
    name='DouBaoVoiceBridge',
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
    icon='assets/doubao_d.ico',
    uac_admin=True,
)
