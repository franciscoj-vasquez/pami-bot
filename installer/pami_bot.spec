# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec — KINETICA
Genera dos ejecutables en un único directorio:
  - "KINETICA.exe"   GUI principal (sin consola)
  - "bot_runner.exe" proceso del bot lanzado por la GUI (con consola)
"""
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files

# El spec vive en installer/ — ROOT apunta a la raíz del proyecto
ROOT   = Path(SPECPATH).parent
SRC    = ROOT / "src"

import customtkinter as _ctk
CTK_DIR = Path(_ctk.__file__).parent

# Playwright incluye un driver Node.js (~80 MB) que necesita viajar con el bot
pw_datas = collect_data_files("playwright", includes=["driver/**/*"])

# ── bot_runner ────────────────────────────────────────────────────────────────
a_bot = Analysis(
    [str(SRC / "bot.py")],
    pathex=[str(SRC), str(ROOT)],
    binaries=[],
    datas=pw_datas,
    hiddenimports=["playwright", "playwright.sync_api"],
    hookspath=[],
    runtime_hooks=[],
    excludes=["customtkinter", "tkinter", "_tkinter"],
)
pyz_bot = PYZ(a_bot.pure)
exe_bot = EXE(
    pyz_bot, a_bot.scripts, [],
    exclude_binaries=True,
    name="bot_runner",
    debug=False, strip=False, upx=True,
    console=False,
)

# ── GUI ───────────────────────────────────────────────────────────────────────
a_gui = Analysis(
    [str(SRC / "gui.py")],
    pathex=[str(SRC), str(ROOT)],
    binaries=[],
    datas=[(str(CTK_DIR), "customtkinter")],
    hiddenimports=[
        "customtkinter",
        "PIL._tkinter_finder",
        "PIL._imagingtk",
        "keyring.backends.Windows",
        "keyring.backends.fail",
        "requests",
        "licencia",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=["playwright"],
)
pyz_gui = PYZ(a_gui.pure)
exe_gui = EXE(
    pyz_gui, a_gui.scripts, [],
    exclude_binaries=True,
    name="KINETICA",
    debug=False, strip=False, upx=True,
    console=False,
    icon=str(ROOT / "installer" / "icon.ico") if (ROOT / "installer" / "icon.ico").exists() else None,
)

# ── Directorio final (ambos ejecutables comparten librerías) ──────────────────
coll = COLLECT(
    exe_gui,
    a_gui.binaries, a_gui.zipfiles, a_gui.datas,
    exe_bot,
    a_bot.binaries, a_bot.zipfiles, a_bot.datas,
    strip=False, upx=True, upx_exclude=[],
    name="KINETICA",
)
