"""PyInstaller entry point for StandardDocApp.

This file is intentionally placed *outside* of the ``standarddocapp`` package.

PyInstaller invokes the entry script as a top-level module
(``__name__ == "__main__"``) with no parent package, so any relative imports
inside that script (e.g. ``from .app import launch``) fail at runtime with::

    ImportError: attempted relative import with no known parent package

To avoid that, we use this small launcher with **absolute** imports as the
PyInstaller entry script. The ``standarddocapp`` package itself stays clean
and ``python -m standarddocapp`` still works for development.
"""
from __future__ import annotations

import sys


def main() -> int:
    from standarddocapp.app import launch
    return launch(sys.argv[1:])


if __name__ == "__main__":
    sys.exit(main())
