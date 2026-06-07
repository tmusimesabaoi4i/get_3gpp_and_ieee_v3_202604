"""End-to-end build helper for StandardDocApp.exe.

Run from a Python environment that has access to a working pip; this script
will:

1. install ``stdharvest``, ``stdsearch`` and ``standarddocapp`` in editable mode
   (so the spec file can resolve them);
2. install/upgrade ``pyinstaller``;
3. invoke PyInstaller against ``StandardDocApp.spec``.

The resulting executable is written to ``standarddocapp/dist/StandardDocApp.exe``.

Usage (from the repository root)::

    python standarddocapp\\build_tools\\build_app.py

or directly::

    cd standarddocapp
    python build_tools\\build_app.py

Notes
-----
* The PyInstaller entry script is
  ``standarddocapp/build_tools/standarddocapp_launcher.py`` (chosen via the
  ``StandardDocApp.spec`` file). Using the package's ``__main__.py`` directly
  fails at runtime with ``ImportError: attempted relative import with no
  known parent package`` because PyInstaller treats the entry script as a
  top-level module. The launcher uses absolute imports and forwards to
  ``standarddocapp.app.launch``.
* The spec file picks up ``src/standarddocapp/assets/app.ico`` automatically
  if you drop one in. No code changes needed to customize the .exe icon.
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path


HERE = Path(__file__).resolve().parent
APP_DIR = HERE.parent
REPO_ROOT = APP_DIR.parent
SPEC = APP_DIR / "StandardDocApp.spec"


def run(cmd: list[str], cwd: Path | None = None) -> None:
    print("$", " ".join(cmd))
    subprocess.check_call(cmd, cwd=str(cwd) if cwd else None)


def pip_install_editable(target: Path) -> None:
    if not (target / "pyproject.toml").exists():
        raise FileNotFoundError(f"pyproject.toml not found in {target}")
    run([sys.executable, "-m", "pip", "install", "-e", str(target)])


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build StandardDocApp.exe")
    parser.add_argument(
        "--skip-deps", action="store_true",
        help="Skip pip install steps (assume packages already installed).",
    )
    parser.add_argument(
        "--clean", action="store_true",
        help="Remove dist/ and build/ before running.",
    )
    parser.add_argument(
        "extra", nargs=argparse.REMAINDER,
        help="Extra args passed straight through to PyInstaller "
             "(use after --, e.g. `-- --log-level=DEBUG`).",
    )
    args = parser.parse_args(argv)

    if not SPEC.exists():
        print(f"Spec not found: {SPEC}", file=sys.stderr)
        return 2

    if args.clean:
        for d in (APP_DIR / "dist", APP_DIR / "build"):
            if d.exists():
                print(f"Cleaning {d}")
                shutil.rmtree(d, ignore_errors=True)

    if not args.skip_deps:
        run([sys.executable, "-m", "pip", "install", "-U", "pip", "wheel"])
        pip_install_editable(REPO_ROOT / "stdharvest")
        pip_install_editable(REPO_ROOT / "stdsearch")
        pip_install_editable(APP_DIR)
        run([sys.executable, "-m", "pip", "install", "-U", "pyinstaller>=6.0"])

    extra = [a for a in (args.extra or []) if a != "--"]
    run(
        [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", *extra, str(SPEC)],
        cwd=APP_DIR,
    )

    exe = APP_DIR / "dist" / "StandardDocApp.exe"
    if exe.exists():
        print(f"\nBuild complete: {exe}")
    else:
        print(
            "\nBuild finished but StandardDocApp.exe was not found in dist/. "
            "Check PyInstaller output above.",
            file=sys.stderr,
        )
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
