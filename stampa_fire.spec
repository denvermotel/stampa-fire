# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec - STAMPA FIRE
# Genera un singolo .exe autonomo (onefile), senza console, con richiesta UAC.
# Compatibile con PyInstaller 6+.

a = Analysis(
    ['stampa_fire.py'],
    pathex=[],
    binaries=[],
    datas=[('assets', 'assets')],
    hiddenimports=[
        'openpyxl',
        'openpyxl.cell._writer',
        'openpyxl.styles.stylesheet',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='StampaFire',
    icon='assets/icona.ico',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,       # nessuna finestra console
    uac_admin=True,      # richiede privilegi amministratore all'avvio (UAC)
)
