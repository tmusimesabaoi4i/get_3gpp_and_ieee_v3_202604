"""Entry point: ``python -m standarddocapp`` and the PyInstaller exe both
execute this module.

IMPORTANT
---------
This file uses **absolute imports** on purpose. Relative imports (``from
.app import launch``) work for ``python -m standarddocapp`` but fail at
runtime when PyInstaller treats the entry script as a top-level module
(no parent package). Keep these imports absolute so the same file can
serve both flows.
"""
from __future__ import annotations

import sys

from standarddocapp.app import launch


def main(argv: list[str] | None = None) -> int:
    return launch(argv if argv is not None else sys.argv[1:])


if __name__ == "__main__":
    sys.exit(main())
