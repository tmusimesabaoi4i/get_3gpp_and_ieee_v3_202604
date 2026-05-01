"""Entry point: `python -m standarddocapp` launches the GUI."""
from __future__ import annotations

import sys

from .app import launch


def main(argv: list[str] | None = None) -> int:
    return launch(argv)


if __name__ == "__main__":
    sys.exit(main())
