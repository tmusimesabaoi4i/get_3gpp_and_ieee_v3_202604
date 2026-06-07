"""Validate ``standarddocapp/src/standarddocapp/assets/app.ico``.

Run from the repository root (or with ``--icon`` pointing at any .ico file)::

    python tools/check_icon.py
    python tools/check_icon.py --icon path/to/app.ico

Exit status is non-zero if the file is missing, not a valid ICO, or is
missing any of the recommended sizes (16/24/32/48/64/128/256). Used by
``standarddocapp/build_exe.bat`` as a pre-flight check.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_ICON = REPO_ROOT / "standarddocapp" / "src" / "standarddocapp" / "assets" / "app.ico"

REQUIRED_SIZES = {
    (16, 16),
    (24, 24),
    (32, 32),
    (48, 48),
    (64, 64),
    (128, 128),
    (256, 256),
}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate the StandardDocApp .ico file.")
    parser.add_argument(
        "--icon",
        type=Path,
        default=DEFAULT_ICON,
        help=f"Path to app.ico (default: {DEFAULT_ICON})",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Fail if any of the recommended sizes is missing (default).",
    )
    args = parser.parse_args(argv)

    icon_path: Path = args.icon
    if not icon_path.exists():
        print(f"ERROR: app.ico not found: {icon_path}", file=sys.stderr)
        return 1

    try:
        from PIL import Image
    except ImportError:
        print(
            "ERROR: Pillow is not installed. Run `python -m pip install pillow` first.",
            file=sys.stderr,
        )
        return 2

    try:
        img = Image.open(icon_path)
        img.load()
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: failed to open ICO: {exc}", file=sys.stderr)
        return 3

    if img.format != "ICO":
        print(f"ERROR: not an ICO file (Pillow says format={img.format!r})", file=sys.stderr)
        return 4

    sizes = sorted(img.ico.sizes())
    print(f"ICO path:        {icon_path}")
    print(f"format:          {img.format}")
    print(f"primary size:    {img.size}")
    print(f"available sizes: {sizes}")

    actual = set(img.ico.sizes())
    missing = REQUIRED_SIZES - actual
    if missing:
        print(
            f"ERROR: missing icon sizes: {sorted(missing)}",
            file=sys.stderr,
        )
        print(
            "Hint: regenerate with ImageMagick:\n"
            "  magick convert app.png -define icon:auto-resize=256,128,64,48,32,16 app.ico",
            file=sys.stderr,
        )
        return 5

    print("OK: app.ico contains all recommended sizes (16/24/32/48/64/128/256).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
