"""Search tab - GUI wrapper around stdsearch.

UX guarantees:

* All "Open ..." buttons re-read the current Excel before resolving the
  output directory, so editing the workbook (B1 ProjectName / B2 JobFolder /
  B3 OutputFolder) instantly redirects the actions.
* Auto-detects when the chosen Excel changes on disk and reloads its
  preview (a "Reload" button forces a refresh too).
* Surfaces a per-tab elapsed-time / state line during runs.
"""
from __future__ import annotations

import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional, Tuple

from .log_panel import LogPanel
from .osutil import open_in_shell
from .paths import app_log_dir
from .runner import JobMessage, JobRunner


def _today_compact() -> str:
    try:
        from stdharvest.utils import now_compact_date  # type: ignore
        return now_compact_date()
    except Exception:
        return datetime.now().strftime("%Y%m%d")


class _SearchSnapshot:
    """Lightweight view of the search workbook used by Open buttons."""

    __slots__ = ("project_name", "job_folder", "output_folder")

    def __init__(self, project_name: str, job_folder: Optional[Path],
                 output_folder: Optional[Path]) -> None:
        self.project_name = project_name
        self.job_folder = job_folder
        self.output_folder = output_folder

    @property
    def output_root(self) -> Optional[Path]:
        if self.output_folder is not None:
            return self.output_folder
        if self.job_folder is not None:
            return self.job_folder / "search"
        return None

    def latest_output_dir(self) -> Optional[Path]:
        root = self.output_root
        if root is None or not root.exists() or not self.project_name:
            return None
        today = root / f"{_today_compact()}_{self.project_name}"
        if today.exists():
            return today
        suffix = f"_{self.project_name}"
        candidates: List[Path] = sorted(
            (p for p in root.iterdir() if p.is_dir() and p.name.endswith(suffix)),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        return candidates[0] if candidates else None


def _read_search_excel(path: Path) -> Tuple[_SearchSnapshot, list]:
    """Re-import every time so the lazy import works after install."""
    from stdsearch.excel_io import read_search_excel  # type: ignore
    project = read_search_excel(path)
    snap = _SearchSnapshot(
        project_name=project.project_name,
        job_folder=Path(project.job_folder) if project.job_folder else None,
        output_folder=Path(project.output_folder) if project.output_folder else None,
    )
    return snap, list(project.rules)


class SearchTab(ttk.Frame):
    """`stdsearch` GUI: pick Excel -> run -> open results."""

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=8)
        self._excel_path: Optional[Path] = None
        self._excel_mtime: Optional[float] = None

        self._runner = JobRunner(logger_names=("stdsearch",))
        self._job_started: Optional[float] = None

        self._build_ui()
        self._poll_job()
        self._poll_excel_mtime()

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        self.columnconfigure(1, weight=1)

        ttk.Label(self, text="Excel (sample_search.xlsx):").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=4
        )
        self._excel_var = tk.StringVar()
        ttk.Entry(self, textvariable=self._excel_var, state="readonly").grid(
            row=0, column=1, sticky="ew", pady=4
        )
        btns = ttk.Frame(self)
        btns.grid(row=0, column=2, sticky="e", padx=(8, 0), pady=4)
        ttk.Button(btns, text="Browse...", command=self._browse_excel, width=10).pack(
            side=tk.LEFT
        )
        ttk.Button(btns, text="Reload", command=self._reload_excel, width=8).pack(
            side=tk.LEFT, padx=(4, 0)
        )

        info = ttk.LabelFrame(self, text="Project Preview", padding=6)
        info.grid(row=1, column=0, columnspan=3, sticky="ew", pady=6)
        info.columnconfigure(1, weight=1)

        labels = [
            ("ProjectName:", "_project_var"),
            ("JobFolder:", "_jobfolder_var"),
            ("OutputFolder:", "_outfolder_var"),
            ("manifest.json:", "_manifest_var"),
            ("Latest output dir (resolved):", "_resolved_var"),
        ]
        for r, (label, attr) in enumerate(labels):
            ttk.Label(info, text=label).grid(row=r, column=0, sticky="w", padx=(0, 6))
            var = tk.StringVar(value="-")
            setattr(self, attr, var)
            ttk.Label(info, textvariable=var, foreground="#222").grid(
                row=r, column=1, sticky="w"
            )

        self._preview_status = tk.StringVar(value="")
        ttk.Label(info, textvariable=self._preview_status, foreground="#a60").grid(
            row=len(labels), column=0, columnspan=2, sticky="w", pady=(4, 0)
        )

        rules_frame = ttk.LabelFrame(self, text="Search Rules", padding=4)
        rules_frame.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=6)
        rules_frame.columnconfigure(0, weight=1)
        rules_frame.rowconfigure(0, weight=1)
        cols = ("use", "name", "match", "scope", "terms")
        self._tree = ttk.Treeview(
            rules_frame, columns=cols, show="headings", height=6,
        )
        self._tree.heading("use", text="Use")
        self._tree.heading("name", text="Rule")
        self._tree.heading("match", text="Match")
        self._tree.heading("scope", text="Scope")
        self._tree.heading("terms", text="Terms")
        self._tree.column("use", width=50, anchor="center", stretch=False)
        self._tree.column("name", width=160)
        self._tree.column("match", width=70, anchor="center", stretch=False)
        self._tree.column("scope", width=90, anchor="center", stretch=False)
        self._tree.column("terms", width=420, stretch=True)
        self._tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(rules_frame, orient="vertical", command=self._tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self._tree.configure(yscrollcommand=sb.set)

        actions = ttk.Frame(self)
        actions.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 4))
        actions.columnconfigure(4, weight=1)

        self._run_btn = ttk.Button(
            actions, text="Run search", command=self._on_run,
            state=tk.DISABLED, width=14,
        )
        self._run_btn.grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)
        self._open_html_btn = ttk.Button(
            actions, text="Open search_results.html",
            command=self._open_html, state=tk.DISABLED, width=24,
        )
        self._open_html_btn.grid(row=0, column=1, sticky="w", padx=6, pady=2)
        self._open_csv_btn = ttk.Button(
            actions, text="Open search_results.csv",
            command=self._open_csv, state=tk.DISABLED, width=24,
        )
        self._open_csv_btn.grid(row=0, column=2, sticky="w", padx=6, pady=2)
        self._open_dir_btn = ttk.Button(
            actions, text="Open output folder",
            command=self._open_output_dir, state=tk.DISABLED, width=18,
        )
        self._open_dir_btn.grid(row=0, column=3, sticky="w", padx=6, pady=2)

        self._state_var = tk.StringVar(value="Idle")
        ttk.Label(actions, textvariable=self._state_var, foreground="#357").grid(
            row=1, column=0, columnspan=3, sticky="w", padx=(0, 6), pady=(4, 0),
        )
        self._progress = ttk.Progressbar(actions, mode="indeterminate", length=200)
        self._progress.grid(row=1, column=4, sticky="e", pady=(4, 0))

        self._log = LogPanel(self, title="stdsearch log")
        self._log.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        self.rowconfigure(4, weight=1)

    # ------------------------------------------------------------------ Excel
    def _browse_excel(self) -> None:
        initial = str(self._excel_path.parent) if self._excel_path else ""
        path = filedialog.askopenfilename(
            title="Select sample_search.xlsx",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
            initialdir=initial,
        )
        if not path:
            return
        self._set_excel(Path(path))

    def _reload_excel(self) -> None:
        if self._excel_path is None:
            messagebox.showinfo("No Excel", "Pick an Excel workbook first.")
            return
        self._refresh_preview(force=True)

    def _set_excel(self, path: Path) -> None:
        self._excel_path = path
        try:
            self._excel_mtime = path.stat().st_mtime
        except OSError:
            self._excel_mtime = None
        self._excel_var.set(str(path))
        self._refresh_preview(force=True)

    def _poll_excel_mtime(self) -> None:
        try:
            if self._excel_path is not None and self._excel_path.exists():
                mt = self._excel_path.stat().st_mtime
                if self._excel_mtime is not None and mt != self._excel_mtime:
                    self._excel_mtime = mt
                    self._log.append("Excel changed on disk - reloading preview.")
                    self._refresh_preview(force=False)
        finally:
            self.after(1500, self._poll_excel_mtime)

    def _refresh_preview(self, force: bool) -> None:
        if self._excel_path is None:
            return
        for r in self._tree.get_children():
            self._tree.delete(r)
        if not self._excel_path.exists():
            self._preview_status.set(f"Excel not found: {self._excel_path}")
            self._run_btn.configure(state=tk.DISABLED)
            return
        try:
            snap, rules = _read_search_excel(self._excel_path)
        except Exception as exc:  # noqa: BLE001
            self._project_var.set("(failed to read Excel)")
            self._jobfolder_var.set(str(exc))
            self._outfolder_var.set("-")
            self._manifest_var.set("-")
            self._resolved_var.set("-")
            self._preview_status.set(f"Failed to read Excel: {exc}")
            self._run_btn.configure(state=tk.DISABLED)
            return

        self._project_var.set(snap.project_name)
        self._jobfolder_var.set(str(snap.job_folder) if snap.job_folder else "-")
        if snap.output_folder is not None:
            self._outfolder_var.set(str(snap.output_folder))
        elif snap.job_folder is not None:
            self._outfolder_var.set(f"(default) {snap.job_folder / 'search'}")
        else:
            self._outfolder_var.set("-")

        manifest = (snap.job_folder / "logs" / "manifest.json") if snap.job_folder else None
        if manifest and manifest.exists():
            self._manifest_var.set(f"OK: {manifest}")
        elif manifest:
            self._manifest_var.set(f"NOT FOUND: {manifest}")
        else:
            self._manifest_var.set("-")

        resolved = snap.latest_output_dir()
        if resolved is not None:
            self._resolved_var.set(str(resolved))
            self._open_dir_btn.configure(state=tk.NORMAL)
            self._open_html_btn.configure(
                state=tk.NORMAL if (resolved / "search_results.html").exists() else tk.DISABLED
            )
            self._open_csv_btn.configure(
                state=tk.NORMAL if (resolved / "search_results.csv").exists() else tk.DISABLED
            )
        else:
            self._resolved_var.set("(none yet - run a search to create)")
            self._open_dir_btn.configure(state=tk.DISABLED)
            self._open_html_btn.configure(state=tk.DISABLED)
            self._open_csv_btn.configure(state=tk.DISABLED)

        for rule in rules:
            self._tree.insert(
                "", tk.END,
                values=(
                    "yes" if rule.use else "no",
                    rule.rule_name,
                    rule.match_type,
                    rule.scope,
                    " | ".join(rule.terms),
                ),
            )

        problems = []
        if not snap.project_name:
            problems.append("ProjectName empty")
        if snap.job_folder is None or not snap.job_folder.exists():
            problems.append(f"JobFolder missing ({snap.job_folder})")
        elif manifest and not manifest.exists():
            problems.append("manifest.json missing - run harvest first")

        if problems:
            self._preview_status.set(" / ".join(problems))
            self._run_btn.configure(state=tk.DISABLED)
        else:
            self._preview_status.set("")
            if not self._runner.is_running:
                self._run_btn.configure(state=tk.NORMAL)
        if force:
            self._log.append(f"Loaded Excel: {self._excel_path}")

    # ------------------------------------------------------------------ run
    def _on_run(self) -> None:
        if self._runner.is_running:
            messagebox.showinfo("Busy", "A search job is already running.")
            return
        if self._excel_path is None or not self._excel_path.exists():
            messagebox.showwarning("Excel required", "Select sample_search.xlsx first.")
            return
        try:
            from stdsearch.cli import run as search_run  # type: ignore
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(
                "stdsearch not installed",
                f"Failed to import stdsearch: {exc}\n\n"
                "Install it with: pip install -e ./stdsearch",
            )
            return

        excel = self._excel_path.resolve()
        self._log.append(f"=== Starting search: {excel} ===")
        self._run_btn.configure(state=tk.DISABLED)
        self._open_html_btn.configure(state=tk.DISABLED)
        self._open_csv_btn.configure(state=tk.DISABLED)
        self._open_dir_btn.configure(state=tk.DISABLED)
        self._progress.configure(mode="indeterminate")
        self._progress.start(80)
        self._job_started = time.time()
        self._tick_state(running=True)
        self._runner.start(
            target=lambda: search_run(excel),
            log_file_dir=app_log_dir() / "search",
            log_file_prefix="search",
        )

    def _tick_state(self, running: bool) -> None:
        if running and self._runner.is_running and self._job_started is not None:
            elapsed = int(time.time() - self._job_started)
            self._state_var.set(
                f"Running... {elapsed // 60:02d}:{elapsed % 60:02d}"
            )
            self.after(500, lambda: self._tick_state(True))

    def _poll_job(self) -> None:
        try:
            for msg in self._runner.drain():
                self._handle_message(msg)
        finally:
            self.after(120, self._poll_job)

    def _handle_message(self, msg: JobMessage) -> None:
        if msg.kind == "log":
            self._log.append(msg.text)
            return
        if msg.kind in ("done", "error"):
            self._progress.stop()
            self._progress.configure(mode="determinate")
            self._progress["value"] = 0
            elapsed = (
                int(time.time() - self._job_started) if self._job_started else 0
            )
            self._state_var.set(
                f"{'Done' if msg.kind == 'done' else 'Failed'} "
                f"({elapsed // 60:02d}:{elapsed % 60:02d})"
            )
            self._job_started = None
            self._log.append(f"=== {msg.text} ===")
            self._refresh_preview(force=False)

    # ------------------------------------------------------------------ open
    def _resolve_now(self) -> Optional[_SearchSnapshot]:
        if self._excel_path is None or not self._excel_path.exists():
            messagebox.showwarning("No Excel", "Pick an Excel workbook first.")
            return None
        try:
            snap, _ = _read_search_excel(self._excel_path)
            return snap
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(
                "Excel read failed",
                f"Failed to read {self._excel_path}: {exc}",
            )
            return None

    def _open_html(self) -> None:
        snap = self._resolve_now()
        if snap is None:
            return
        latest = snap.latest_output_dir()
        if latest is None:
            messagebox.showwarning(
                "Nothing to open",
                "No search output folder for this project yet. Run a search first.",
            )
            return
        target = latest / "search_results.html"
        if not target.exists():
            messagebox.showwarning("Missing", f"{target} does not exist.")
            return
        try:
            open_in_shell(target)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Error", f"Failed to open HTML: {exc}")

    def _open_csv(self) -> None:
        snap = self._resolve_now()
        if snap is None:
            return
        latest = snap.latest_output_dir()
        if latest is None:
            messagebox.showwarning(
                "Nothing to open",
                "No search output folder for this project yet. Run a search first.",
            )
            return
        target = latest / "search_results.csv"
        if not target.exists():
            messagebox.showwarning("Missing", f"{target} does not exist.")
            return
        try:
            open_in_shell(target)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Error", f"Failed to open CSV: {exc}")

    def _open_output_dir(self) -> None:
        snap = self._resolve_now()
        if snap is None:
            return
        latest = snap.latest_output_dir()
        if latest is None:
            messagebox.showwarning(
                "Nothing to open",
                "No search output folder for this project yet. Run a search first.",
            )
            return
        try:
            open_in_shell(latest)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Error", f"Failed to open folder: {exc}")
