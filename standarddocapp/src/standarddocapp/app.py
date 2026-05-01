"""Main Tk app: notebook with Harvest / Search / About tabs."""
from __future__ import annotations

import logging
import sys
import tkinter as tk
from importlib.resources import as_file, files
from tkinter import messagebox, ttk

from . import APP_NAME, __version__
from .about_tab import AboutTab
from .harvest_tab import HarvestTab
from .osutil import open_in_shell
from .paths import app_log_dir
from .search_tab import SearchTab
from .sysinfo import collect_env_info
from .widgets import StatusBar


# Suppress: never crash app startup just because the icon failed to load.
_log = logging.getLogger(__name__)


def _apply_window_icon(root: tk.Tk) -> None:
    """Set the *window* icon (title bar + taskbar) from the bundled .ico.

    The PyInstaller spec ``icon=`` only sets the file icon shown in Explorer
    for ``StandardDocApp.exe`` itself. The tkinter window icon is independent
    and has to be applied here, otherwise the title bar / taskbar still show
    the default Tk feather logo.

    Looks up ``standarddocapp/assets/app.ico`` via ``importlib.resources`` so
    the same code path works for ``python -m standarddocapp`` (icon read from
    the source tree) and for the PyInstaller-bundled exe (icon read from
    ``sys._MEIPASS/standarddocapp/assets/app.ico``).
    """
    try:
        ref = files("standarddocapp.assets") / "app.ico"
        if not ref.is_file():
            return
        with as_file(ref) as ico_path:
            root.iconbitmap(default=str(ico_path))
    except Exception as exc:  # noqa: BLE001
        _log.debug("could not apply window icon: %s", exc)


def _apply_app_user_model_id() -> None:
    """Tell Windows this exe is its own application for taskbar grouping.

    Without this, Windows treats us as a generic "Python" host and may share
    the taskbar slot / icon with other Python apps.
    """
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            f"StandardDocApp.{__version__}"
        )
    except Exception as exc:  # noqa: BLE001
        _log.debug("could not set AppUserModelID: %s", exc)


def _configure_root_logging() -> None:
    """Make sure stdharvest/stdsearch loggers reach our queue handlers.

    We don't add a stdout StreamHandler because the app may be a windowed exe
    with no console; per-job FileHandlers are added on demand by JobRunner.
    """
    for name in ("stdharvest", "stdsearch", "standarddocapp.html_check"):
        lg = logging.getLogger(name)
        lg.setLevel(logging.INFO)
        lg.propagate = False


def _apply_theme(root: tk.Tk) -> None:
    style = ttk.Style(root)
    preferred = ("vista", "xpnative", "winnative", "clam", "alt", "default")
    available = set(style.theme_names())
    for name in preferred:
        if name in available:
            style.theme_use(name)
            break
    style.configure("TNotebook.Tab", padding=(14, 6))
    style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
    style.configure("Status.TLabel", foreground="#555")
    style.configure("StatusRight.TLabel", foreground="#888")


def _build_menu(root: tk.Tk, env) -> None:
    bar = tk.Menu(root)

    file_menu = tk.Menu(bar, tearoff=0)
    file_menu.add_command(
        label="Open log folder",
        command=lambda: _safe_open(env.log_dir),
    )
    if env.readme_path:
        file_menu.add_command(
            label=f"Open README ({env.readme_path.name})",
            command=lambda: _safe_open(env.readme_path),
        )
    file_menu.add_separator()
    file_menu.add_command(label="Quit", command=root.destroy)
    bar.add_cascade(label="File", menu=file_menu)

    help_menu = tk.Menu(bar, tearoff=0)
    help_menu.add_command(
        label="About...",
        command=lambda: messagebox.showinfo(
            "About",
            f"{APP_NAME} v{__version__}\n\n"
            "Integrated GUI for stdharvest (collect) and stdsearch (search).\n"
            f"OS: {env.os_summary}\n"
            f"Python: {env.python_version.split()[0]}\n"
            f"Office: {'Available' if env.office_available else 'Not detected'}\n"
            f"LibreOffice: {'Available' if env.soffice_path else 'Not detected'}",
        ),
    )
    bar.add_cascade(label="Help", menu=help_menu)
    root.config(menu=bar)


def _safe_open(path) -> None:
    try:
        open_in_shell(path)
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Error", f"Failed to open {path}: {exc}")


def build_root() -> tk.Tk:
    root = tk.Tk()
    root.title(f"{APP_NAME} {__version__}")
    root.geometry("1100x760")
    root.minsize(960, 620)

    _apply_theme(root)
    _apply_window_icon(root)

    env = collect_env_info(APP_NAME, __version__, app_log_dir())
    _build_menu(root, env)

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=4, pady=(4, 0))

    harvest = HarvestTab(notebook)
    search = SearchTab(notebook)
    about = AboutTab(notebook, env)

    notebook.add(harvest, text="  収集 / Harvest  ")
    notebook.add(search, text="  検索 / Search  ")
    notebook.add(about, text="  設定・ログ / About  ")

    status = StatusBar(root)
    status.pack(fill=tk.X, side=tk.BOTTOM)

    badges = []
    badges.append("Office: " + ("OK" if env.office_available else "-"))
    badges.append("LibreOffice: " + ("OK" if env.soffice_path else "-"))
    badges.append(f"Logs: {env.log_dir}")
    status.set_left(" | ".join(badges))
    status.set_right(f"{APP_NAME} v{__version__}")

    def _on_close() -> None:
        running = []
        for tab, name in ((harvest, "Harvest"), (search, "Search"),
                          (about, "Validator")):
            runner = getattr(tab, "_runner", None) or getattr(tab, "_validator_runner", None)
            if runner is not None and getattr(runner, "is_running", False):
                running.append(name)
        if running:
            ok = messagebox.askyesno(
                "Quit",
                "The following job(s) are still running and will be aborted:\n"
                f"  - {', '.join(running)}\n\nQuit anyway?",
            )
            if not ok:
                return
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", _on_close)
    return root


def launch(argv: list[str] | None = None) -> int:
    _configure_root_logging()
    _apply_app_user_model_id()
    root = build_root()
    try:
        root.mainloop()
    except KeyboardInterrupt:
        return 130
    return 0


if __name__ == "__main__":
    sys.exit(launch())
