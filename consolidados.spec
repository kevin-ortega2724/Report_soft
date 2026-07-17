# -*- mode: python ; coding: utf-8 -*-
# Genera el ejecutable de Windows para ReportSoft - Consolidados.
# Compilar SIEMPRE en Windows (PyInstaller no hace cross-compile):
#   .venv\Scripts\activate
#   pyinstaller consolidados.spec --clean
#
# El .exe queda en dist\ReportSoft-Consolidados\ (modo onedir).
# data\ y config\ NO se empaquetan: deben quedar junto al .exe, igual que
# hoy quedan junto a run.py (ver utils.obtener_directorio_base()).

block_cipher = None

hiddenimports = [
    "PyQt5.QtPrintSupport",
    "matplotlib.backends.backend_qt5agg",
    "openpyxl",
    "bs4",
]

a = Analysis(
    ["run.py"],
    pathex=["src"],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ReportSoft-Consolidados",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    name="ReportSoft-Consolidados",
)
