"""About / Settings / Logs tab.

Layout:

1. **Environment** - app version, Python, OS, Office/LibreOffice detection.
2. **Sample Excel generators** - includes a "Build sample download.xlsx..."
   dialog that lets the user pick SourceType / save folder / JobName /
   ProxyURL / OutputRootFolder / reference Excel.
3. **HTML validator** - point at a stdharvest job folder and verify every
   generated HTML parses, has a body, and has no broken local links.
4. **Quick actions** - all buttons placed via grid with explicit padding
   so they never overlap at any window size.
"""
from __future__ import annotations

import logging
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional

from .html_check import format_report_lines, validate_job_html
from .log_panel import LogPanel
from .osutil import open_in_shell, reveal_in_shell
from .paths import app_log_dir
from .runner import JobMessage, JobRunner
from .sample_dialog import SampleExcelDialog
from .samples import (
    build_search_sample,
    default_output_root,
)
from .sysinfo import EnvInfo
from .widgets import ScrollableFrame


_VALIDATOR_LOGGER_NAME = "standarddocapp.html_check"
_validator_logger = logging.getLogger(_VALIDATOR_LOGGER_NAME)
_validator_logger.setLevel(logging.INFO)
_validator_logger.propagate = False


class AboutTab(ttk.Frame):
    def __init__(self, master: tk.Misc, env: EnvInfo) -> None:
        super().__init__(master)
        self.env = env
        self._validator_runner = JobRunner(logger_names=(_VALIDATOR_LOGGER_NAME,))
        self._validator_job_folder: Optional[Path] = None
        self._build_ui()
        self._poll_validator()

    # ---------------------------------------------------------------- layout
    def _build_ui(self) -> None:
        scroll = ScrollableFrame(self)
        scroll.pack(fill=tk.BOTH, expand=True)
        root = scroll.inner
        root.columnconfigure(0, weight=1)

        self._build_header(root, row=0)
        self._build_env_section(root, row=1)
        self._build_samples_section(root, row=2)
        self._build_validator_section(root, row=3)
        self._build_quick_actions(root, row=4)

    # ---------------------------------------------------------------- header
    def _build_header(self, parent: ttk.Frame, row: int) -> None:
        header = ttk.Frame(parent)
        header.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(
            header,
            text=f"{self.env.app_name}  v{self.env.app_version}",
            font=("Segoe UI", 14, "bold"),
        ).pack(side=tk.LEFT)
        ttk.Label(
            header,
            text="Integrated GUI for stdharvest + stdsearch",
            foreground="#555",
        ).pack(side=tk.LEFT, padx=12)

    # ---------------------------------------------------------------- env
    def _build_env_section(self, parent: ttk.Frame, row: int) -> None:
        info = ttk.LabelFrame(parent, text="Environment", padding=8)
        info.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        info.columnconfigure(1, weight=1)

        rows = [
            ("Python", self.env.python_version),
            ("Python executable", self.env.python_executable),
            ("OS", self.env.os_summary),
            ("Microsoft Office",
             ("Available - " if self.env.office_available else "Not available - ") + self.env.office_detail),
            ("LibreOffice (soffice)",
             (str(self.env.soffice_path) if self.env.soffice_path else "Not found")
             + " | " + self.env.soffice_detail),
            ("App log directory", str(self.env.log_dir)),
        ]
        for i, (k, v) in enumerate(rows):
            ttk.Label(info, text=f"{k}:", font=("Segoe UI", 9, "bold")).grid(
                row=i, column=0, sticky="nw", padx=(0, 8), pady=2
            )
            ttk.Label(
                info, text=str(v), wraplength=720, justify=tk.LEFT,
            ).grid(row=i, column=1, sticky="ew", pady=2)

        if self.env.office_available:
            badge_text = "PDF: Office is preferred. LibreOffice is used as fallback."
            badge_fg = "#0a7"
            badge_font = ("Segoe UI", 9, "italic")
        elif self.env.soffice_path:
            badge_text = "PDF: Office not detected; LibreOffice fallback will be used."
            badge_fg = "#a60"
            badge_font = ("Segoe UI", 9, "italic")
        else:
            badge_text = "WARNING: Neither Office nor LibreOffice was detected. PDF conversion will fail."
            badge_fg = "#c00"
            badge_font = ("Segoe UI", 9, "bold")
        ttk.Label(info, text=badge_text, foreground=badge_fg, font=badge_font).grid(
            row=len(rows), column=0, columnspan=2, sticky="w", pady=(6, 0)
        )

    # ---------------------------------------------------------------- samples
    def _build_samples_section(self, parent: ttk.Frame, row: int) -> None:
        frame = ttk.LabelFrame(parent, text="Sample Excel Generators", padding=8)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        frame.columnconfigure(0, weight=1)

        intro = (
            "「Build sample download.xlsx...」 を押すと、保存先・SourceType・JobName・ProxyURL・"
            "OutputRootFolder・参照元Excel(任意) をその場で指定して sample_download.xlsx を作れます。"
            " OutputRootFolder の初期値は ユーザーの Downloads フォルダです。"
            " 参照元 Excel を指定すると、その A 列のハイパーリンクが C 列に書き込まれます。"
        )
        ttk.Label(frame, text=intro, wraplength=820, justify=tk.LEFT, foreground="#444").grid(
            row=0, column=0, sticky="ew", pady=(0, 8)
        )

        btn_grid = ttk.Frame(frame)
        btn_grid.grid(row=1, column=0, sticky="ew", pady=2)
        for c in range(3):
            btn_grid.columnconfigure(c, weight=1, uniform="samples")

        ttk.Button(
            btn_grid, text="Build sample download.xlsx... (3GPP)",
            command=lambda: self._open_dialog("3gpp"),
        ).grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        ttk.Button(
            btn_grid, text="Build sample download.xlsx... (IEEE)",
            command=lambda: self._open_dialog("ieee"),
        ).grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(
            btn_grid, text="Build sample search.xlsx...",
            command=self._save_search_sample,
        ).grid(row=0, column=2, sticky="ew", padx=4, pady=4)

        self._samples_status = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self._samples_status, foreground="#0a7",
                  wraplength=820, justify=tk.LEFT).grid(
            row=2, column=0, sticky="ew", pady=(6, 0)
        )

    def _open_dialog(self, source_type: str) -> None:
        result = SampleExcelDialog.run(
            self.winfo_toplevel(),
            source_type=source_type,
            default_dir=default_output_root(),
        )
        if result is None:
            return
        self._samples_status.set(f"Wrote: {result}")
        if messagebox.askyesno("Done", f"Wrote {result}\n\nReveal in Explorer?"):
            try:
                reveal_in_shell(result)
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Error", f"Failed to reveal: {exc}")

    def _save_search_sample(self) -> None:
        initial = str(default_output_root())
        path_str = filedialog.asksaveasfilename(
            title="Save sample_search.xlsx",
            defaultextension=".xlsx",
            initialfile="sample_search.xlsx",
            initialdir=initial,
            filetypes=[("Excel workbook", "*.xlsx")],
            parent=self.winfo_toplevel(),
        )
        if not path_str:
            return
        try:
            built = build_search_sample(Path(path_str))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Sample build failed", f"{exc}")
            return
        self._samples_status.set(f"Wrote: {built}")
        if messagebox.askyesno("Done", f"Wrote {built}\n\nReveal in Explorer?"):
            try:
                reveal_in_shell(built)
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Error", f"Failed to reveal: {exc}")

    # ---------------------------------------------------------------- validator
    def _build_validator_section(self, parent: ttk.Frame, row: int) -> None:
        frame = ttk.LabelFrame(parent, text="HTML Integrity Validator", padding=8)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        frame.columnconfigure(1, weight=1)

        intro = (
            "stdharvest が生成した job フォルダを選択すると、index.html / 個別HTML / "
            "combined / combine_full をすべて構文チェックし、内部リンク切れと "
            "アンカー不整合を検出します。読み取り専用で、HTML を書き換えません。"
        )
        ttk.Label(frame, text=intro, wraplength=820, justify=tk.LEFT, foreground="#444").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 8)
        )

        ttk.Label(frame, text="Job folder:", font=("Segoe UI", 9, "bold")).grid(
            row=1, column=0, sticky="w", padx=(0, 8)
        )
        self._job_folder_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self._job_folder_var, state="readonly").grid(
            row=1, column=1, sticky="ew", pady=2
        )
        ttk.Button(frame, text="Browse...", command=self._browse_job_folder, width=14).grid(
            row=1, column=2, sticky="e", padx=(8, 0)
        )

        action_row1 = ttk.Frame(frame)
        action_row1.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        action_row1.columnconfigure(2, weight=1)
        self._validator_btn = ttk.Button(
            action_row1, text="Validate HTML", command=self._on_validate,
            state=tk.DISABLED, width=18,
        )
        self._validator_btn.grid(row=0, column=0, sticky="w", padx=(0, 6))
        ttk.Button(
            action_row1, text="Open job folder",
            command=self._open_validator_job_folder, width=18,
        ).grid(row=0, column=1, sticky="w", padx=6)
        self._validator_count_var = tk.StringVar(value="")
        ttk.Label(action_row1, textvariable=self._validator_count_var, foreground="#555").grid(
            row=0, column=2, sticky="e", padx=6
        )
        self._validator_progress = ttk.Progressbar(action_row1, mode="determinate", length=240)
        self._validator_progress.grid(row=0, column=3, sticky="e")

        self._validator_log = LogPanel(frame, title="Validator log")
        self._validator_log.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))

    def _browse_job_folder(self) -> None:
        initial = ""
        if self._validator_job_folder is not None:
            initial = str(self._validator_job_folder)
        elif self.env.repo_root:
            initial = str(self.env.repo_root)
        path = filedialog.askdirectory(
            title="Select stdharvest job folder",
            initialdir=initial,
            parent=self.winfo_toplevel(),
        )
        if not path:
            return
        self._validator_job_folder = Path(path)
        self._job_folder_var.set(str(self._validator_job_folder))
        self._validator_btn.configure(state=tk.NORMAL)
        if not (self._validator_job_folder / "html").exists():
            self._validator_log.append(
                f"Note: {self._validator_job_folder} has no 'html/' subfolder yet. "
                "Validation will report 0 files."
            )

    def _on_validate(self) -> None:
        if self._validator_runner.is_running:
            messagebox.showinfo("Busy", "A validation job is already running.")
            return
        if self._validator_job_folder is None or not self._validator_job_folder.exists():
            messagebox.showwarning("Folder required", "Select a job folder first.")
            return

        job_folder = self._validator_job_folder
        self._validator_log.append(f"=== Validating: {job_folder} ===")
        self._validator_btn.configure(state=tk.DISABLED)
        self._validator_progress.configure(mode="indeterminate")
        self._validator_progress.start(80)
        self._validator_count_var.set("scanning...")

        progress_lock = threading.Lock()
        progress_state = {"done": 0, "total": 0}

        def _progress(done: int, total: int, _fr) -> None:
            with progress_lock:
                progress_state["done"] = done
                progress_state["total"] = total

        def _tick():
            if not self._validator_runner.is_running:
                return
            with progress_lock:
                d, t = progress_state["done"], progress_state["total"]
            if t > 0:
                if str(self._validator_progress["mode"]) == "indeterminate":
                    self._validator_progress.stop()
                    self._validator_progress.configure(mode="determinate", maximum=t)
                self._validator_progress["value"] = d
                self._validator_count_var.set(f"{d}/{t} files")
            self.after(150, _tick)
        self.after(150, _tick)

        def _target() -> int:
            log = logging.getLogger(_VALIDATOR_LOGGER_NAME)
            t0 = datetime.now()
            log.info("Validating %s", job_folder)
            report = validate_job_html(job_folder, progress=_progress)
            for line in format_report_lines(report):
                log.info(line)
            elapsed = (datetime.now() - t0).total_seconds()
            log.info("Elapsed: %.2fs", elapsed)
            return 0 if report.summary()["broken"] == 0 else 1

        self._validator_runner.start(
            target=_target,
            log_file_dir=app_log_dir() / "validator",
            log_file_prefix="validator",
        )

    def _poll_validator(self) -> None:
        try:
            for msg in self._validator_runner.drain():
                self._handle_validator_message(msg)
        finally:
            self.after(120, self._poll_validator)

    def _handle_validator_message(self, msg: JobMessage) -> None:
        if msg.kind == "log":
            self._validator_log.append(msg.text)
            return
        if msg.kind in ("done", "error"):
            self._validator_progress.stop()
            self._validator_progress.configure(mode="determinate")
            self._validator_btn.configure(state=tk.NORMAL)
            self._validator_log.append(f"=== {msg.text} ===")

    def _open_validator_job_folder(self) -> None:
        path = self._validator_job_folder
        if not path or not path.exists():
            messagebox.showwarning("Folder missing", "Select an existing job folder first.")
            return
        try:
            open_in_shell(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Error", f"Failed to open folder: {exc}")

    # ---------------------------------------------------------------- quick
    def _build_quick_actions(self, parent: ttk.Frame, row: int) -> None:
        frame = ttk.LabelFrame(parent, text="Quick Actions", padding=8)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 12))

        # Use a grid with 3 stretching columns so buttons distribute evenly
        # and never overlap regardless of window width.
        cols = 3
        for c in range(cols):
            frame.columnconfigure(c, weight=1, uniform="quick")

        items: List[tuple[str, callable]] = []
        items.append(("Open app log folder", lambda: self._open(self.env.log_dir)))
        items.append(("Open Downloads folder", lambda: self._open(default_output_root())))
        if self.env.readme_path:
            items.append(
                (f"Open README ({self.env.readme_path.name})",
                 lambda: self._open(self.env.readme_path))
            )
        for p in self.env.sample_paths:
            items.append((f"Reveal sample: {p.name}", lambda p=p: self._reveal(p)))

        for idx, (label, cb) in enumerate(items):
            r, c = divmod(idx, cols)
            ttk.Button(frame, text=label, command=cb).grid(
                row=r, column=c, sticky="ew", padx=4, pady=4,
            )

        if not items:
            ttk.Label(
                frame,
                text="(クイックアクションは利用できません。)",
                foreground="#888",
            ).grid(row=0, column=0, columnspan=cols, sticky="w")

    def _open(self, path: Path) -> None:
        try:
            open_in_shell(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Error", f"Failed to open {path}: {exc}")

    def _reveal(self, path: Path) -> None:
        try:
            reveal_in_shell(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Error", f"Failed to reveal {path}: {exc}")
