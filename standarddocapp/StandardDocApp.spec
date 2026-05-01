# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for StandardDocApp.

Build with::

    pyinstaller --noconfirm StandardDocApp.spec

from the ``standarddocapp/`` folder.

The build expects ``stdharvest`` and ``stdsearch`` to be importable. When
running this spec from a virtual environment in which both packages have been
installed editable (``pip install -e ../stdharvest -e ../stdsearch``), no extra
configuration is needed.
"""

from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules

block_cipher = None
HERE = Path(SPECPATH).resolve()
REPO_ROOT = HERE.parent

# Collect every submodule of the sibling packages so dynamic imports work.
hiddenimports = []
for pkg in ("stdharvest", "stdsearch", "standarddocapp"):
    hiddenimports.extend(collect_submodules(pkg))

hiddenimports.extend(
    [
        # third-party libs that PyInstaller sometimes misses
        "openpyxl",
        "requests",
        "mammoth",
        "bs4",
        "lxml",
        "lxml.etree",
        "lxml._elementpath",
        # pywin32 / Office COM
        "pythoncom",
        "pywintypes",
        "win32com",
        "win32com.client",
        "win32com.gen_py",
    ]
)

# Tk needs to ship with the binary; PyInstaller handles this automatically,
# but we add tkinter just to be explicit.
hiddenimports.append("tkinter")
hiddenimports.append("tkinter.ttk")
hiddenimports.append("tkinter.scrolledtext")
hiddenimports.append("tkinter.filedialog")
hiddenimports.append("tkinter.messagebox")

datas = []

a = Analysis(
    [str(HERE / "src" / "standarddocapp" / "__main__.py")],
    pathex=[
        str(HERE / "src"),
        str(REPO_ROOT / "stdharvest" / "src"),
        str(REPO_ROOT / "stdsearch" / "src"),
    ],
    binaries=[],
    datas=datas,
    hiddenimports=sorted(set(hiddenimports)),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
)
