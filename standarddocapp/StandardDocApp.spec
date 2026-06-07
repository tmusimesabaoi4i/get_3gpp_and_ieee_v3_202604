# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for StandardDocApp.

Build with::

    python -m PyInstaller --noconfirm --clean StandardDocApp.spec

(run from the ``standarddocapp/`` folder; ``build_exe.bat`` does this for
you).

This spec is the *recommended* way to build the distributable exe:

* runs the absolute-import-safe entry script
  (``src/standarddocapp/__main__.py``) so the bundled exe does **not**
  hit ``ImportError: attempted relative import with no known parent
  package`` at startup;
* bundles the static assets folder (``src/standarddocapp/assets``) so an
  optional ``app.ico`` (or future logo files) ships with the exe;
* explicitly collects every submodule of ``stdharvest`` / ``stdsearch``
  / ``standarddocapp`` so dynamic imports reach the bundle;
* uses ``collect_all`` for the libraries that ship data files or have
  many lazy submodules (``lxml``, ``openpyxl``, ``mammoth``, ``bs4``,
  ``win32com``, ``PIL``);
* if ``src/standarddocapp/assets/app.ico`` is present, it is used as the
  ``.exe`` icon automatically (drop your own ``app.ico`` there to
  customize without editing this spec).
"""
from pathlib import Path

from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None
PROJECT_ROOT = Path(SPECPATH).resolve()
REPO_ROOT = PROJECT_ROOT.parent

ENTRY = str(PROJECT_ROOT / "src" / "standarddocapp" / "__main__.py")
ASSETS_SRC = PROJECT_ROOT / "src" / "standarddocapp" / "assets"
ICON_PATH = ASSETS_SRC / "app.ico"

datas = []
binaries = []
hiddenimports = []


def _bundle(name: str) -> None:
    """Pull every data file / binary / submodule of *name* into the build."""
    d, b, h = collect_all(name)
    datas.extend(d)
    binaries.extend(b)
    hiddenimports.extend(h)


for pkg in (
    "lxml",
    "openpyxl",
    "mammoth",
    "bs4",
    "win32com",
    "PIL",  # image plugins incl. WMF/EMF rasteriser used for HTML images
):
    _bundle(pkg)

# Sibling source packages: prefer collect_submodules (no data files, just
# Python modules). They live in REPO_ROOT/{stdharvest,stdsearch}/src and are
# made importable via the ``pathex`` argument below.
for pkg in ("stdharvest", "stdsearch", "standarddocapp"):
    hiddenimports.extend(collect_submodules(pkg))

hiddenimports.extend(
    [
        "requests",
        "lxml.etree",
        "lxml._elementpath",
        "pythoncom",
        "pywintypes",
        "win32com.client",
        "win32com.gen_py",
        "tkinter",
        "tkinter.ttk",
        "tkinter.scrolledtext",
        "tkinter.filedialog",
        "tkinter.messagebox",
    ]
)

if ASSETS_SRC.exists():
    datas.append((str(ASSETS_SRC), "standarddocapp/assets"))


a = Analysis(
    [ENTRY],
    pathex=[
        str(PROJECT_ROOT / "src"),
        str(REPO_ROOT / "stdharvest" / "src"),
        str(REPO_ROOT / "stdsearch" / "src"),
    ],
    binaries=binaries,
    datas=datas,
    hiddenimports=sorted(set(hiddenimports)),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "playwright",
        "selenium",
        "pytest",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="StandardDocApp",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ICON_PATH) if ICON_PATH.exists() else None,
)
