"""Modal dialog that drives :func:`samples.build_harvest_sample`.

Lets the user pick:

* SourceType  (3gpp / ieee)
* Output folder for the workbook
* Output file name (default ``sample_download.xlsx``)
* JobName
* ProxyURL
* OutputRootFolder (defaults to the user's Downloads folder)
* Reference Excel (optional) - column A's hyperlinks become the C column
  of the produced workbook.

Any field left blank reverts to the canonical sample defaults from
``samples.py``.
"""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional, Tuple

from .samples import (
    DEFAULT_3GPP_URLS,
    DEFAULT_IEEE_URLS,
    DEFAULT_JOB_NAMES,
    build_harvest_sample,
    default_output_root,
    read_links_from_excel,
)


class SampleExcelDialog(tk.Toplevel):
    """Modal sample-builder dialog.

    Use :meth:`run` for the synchronous flow:

        path = SampleExcelDialog.run(parent, source_type="ieee")
    """

    def __init__(self, master: tk.Misc, *, source_type: str = "3gpp",
                 default_dir: Optional[Path] = None) -> None:
        super().__init__(master)
        self.title("Build sample download.xlsx")
        self.transient(master)
        self.resizable(False, False)
        self.grab_set()

        self._result: Optional[Path] = None

        self._source_var = tk.StringVar(value=source_type if source_type in ("3gpp", "ieee") else "3gpp")
        self._save_dir_var = tk.StringVar(
            value=str(default_dir) if default_dir else str(default_output_root())
        )
        self._filename_var = tk.StringVar(value="sample_download.xlsx")
        self._jobname_var = tk.StringVar(value=DEFAULT_JOB_NAMES[self._source_var.get()])
        self._proxy_var = tk.StringVar(value="")
        self._proxyuser_var = tk.StringVar(value="")
        self._proxypass_var = tk.StringVar(value="")
        self._outroot_var = tk.StringVar(value=str(default_output_root()))
        self._refexcel_var = tk.StringVar(value="")

        self._loaded_urls: Optional[List[Tuple[str, str]]] = None

        self._build_ui()
        self._update_url_preview()
        self._center_on(master)

    # ---------------------------------------------------------------- helpers
    def _center_on(self, master: tk.Misc) -> None:
        try:
            self.update_idletasks()
            mx = master.winfo_rootx()
            my = master.winfo_rooty()
            mw = master.winfo_width()
            mh = master.winfo_height()
            sw = self.winfo_width()
            sh = self.winfo_height()
            x = mx + max(0, (mw - sw) // 2)
            y = my + max(0, (mh - sh) // 3)
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _row(self, parent: ttk.Frame, row: int, label: str, var: tk.Variable,
             *, width: int = 56, browse: Optional[str] = None) -> None:
        ttk.Label(parent, text=label, font=("Segoe UI", 9, "bold")).grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=3
        )
        ttk.Entry(parent, textvariable=var, width=width).grid(
            row=row, column=1, sticky="ew", pady=3
        )
        if browse == "dir":
            ttk.Button(
                parent, text="Browse folder...", width=18,
                command=lambda v=var: self._pick_dir(v),
            ).grid(row=row, column=2, sticky="e", padx=(8, 0))
        elif browse == "file":
            ttk.Button(
                parent, text="Browse Excel...", width=18,
                command=self._pick_reference_excel,
            ).grid(row=row, column=2, sticky="e", padx=(8, 0))
        else:
            ttk.Frame(parent, width=1).grid(row=row, column=2)

    # ---------------------------------------------------------------- UI
    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)

        intro = ttk.Label(
            frame,
            text=(
                "未指定項目は既存サンプルのデフォルト値が使われます。\n"
                "参照元 Excel を指定すると、その A 列のハイパーリンク (またはセル文字列) が "
                "sample_download.xlsx の C 列に書き込まれます。"
            ),
            wraplength=640, justify=tk.LEFT, foreground="#444",
        )
        intro.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        # SourceType
        ttk.Label(frame, text="SourceType:", font=("Segoe UI", 9, "bold")).grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=3
        )
        src_combo = ttk.Combobox(
            frame, textvariable=self._source_var,
            values=("3gpp", "ieee"), state="readonly", width=10,
        )
        src_combo.grid(row=1, column=1, sticky="w", pady=3)
        src_combo.bind("<<ComboboxSelected>>", self._on_source_changed)

        self._row(frame, 2, "保存先フォルダ:", self._save_dir_var, browse="dir")
        self._row(frame, 3, "ファイル名:", self._filename_var)
        self._row(frame, 4, "JobName:", self._jobname_var)
        self._row(frame, 5, "ProxyURL:", self._proxy_var)
        self._row(frame, 6, "ProxyUser (任意):", self._proxyuser_var)
        self._row(frame, 7, "ProxyPassword (任意):", self._proxypass_var)
        self._row(frame, 8, "OutputRootFolder:", self._outroot_var, browse="dir")
        self._row(frame, 9, "参照元 Excel (任意):", self._refexcel_var, browse="file")

        # URL preview
        prev_box = ttk.LabelFrame(frame, text="C 列に書き込まれる URL (preview)", padding=6)
        prev_box.grid(row=10, column=0, columnspan=3, sticky="ew", pady=(8, 4))
        prev_box.columnconfigure(0, weight=1)
        self._preview_text = tk.Text(
            prev_box, height=8, width=88, wrap=tk.NONE, state=tk.DISABLED,
            font=("Consolas", 9),
        )
        self._preview_text.grid(row=0, column=0, sticky="ew")
        sb = ttk.Scrollbar(prev_box, orient="vertical", command=self._preview_text.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self._preview_text.configure(yscrollcommand=sb.set)
        self._preview_status = tk.StringVar(value="")
        ttk.Label(prev_box, textvariable=self._preview_status, foreground="#357").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )

        # OK / Cancel
        actions = ttk.Frame(frame)
        actions.grid(row=11, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        actions.columnconfigure(0, weight=1)
        ttk.Button(actions, text="Cancel", command=self._on_cancel, width=12).grid(
            row=0, column=1, sticky="e", padx=(0, 6)
        )
        ttk.Button(actions, text="Build", command=self._on_build, width=14).grid(
            row=0, column=2, sticky="e"
        )

        # Reactivity
        self._refexcel_var.trace_add("write", lambda *_: self._update_url_preview())

    # ---------------------------------------------------------------- events
    def _on_source_changed(self, _e=None) -> None:
        src = self._source_var.get()
        # Bump JobName to the canonical default when the user hasn't customised it.
        if self._jobname_var.get() in DEFAULT_JOB_NAMES.values() or not self._jobname_var.get().strip():
            self._jobname_var.set(DEFAULT_JOB_NAMES.get(src, ""))
        self._update_url_preview()

    def _pick_dir(self, var: tk.StringVar) -> None:
        initial = var.get() or str(default_output_root())
        path = filedialog.askdirectory(title="Select folder", initialdir=initial, parent=self)
        if path:
            var.set(path)

    def _pick_reference_excel(self) -> None:
        initial = self._refexcel_var.get() or self._save_dir_var.get() or ""
        path = filedialog.askopenfilename(
            title="Select reference Excel (column A)",
            filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")],
            initialdir=initial,
            parent=self,
        )
        if path:
            self._refexcel_var.set(path)

    def _update_url_preview(self) -> None:
        ref_path_str = self._refexcel_var.get().strip()
        if ref_path_str:
            try:
                links = read_links_from_excel(Path(ref_path_str))
                self._loaded_urls = links
                if links:
                    self._preview_status.set(
                        f"参照元 Excel から {len(links)} 件読み込みました"
                    )
                else:
                    self._preview_status.set(
                        "参照元 Excel の A 列にURL/ハイパーリンクが見つかりません (デフォルトを使用)"
                    )
            except FileNotFoundError:
                self._loaded_urls = None
                self._preview_status.set(f"見つかりません: {ref_path_str}")
            except Exception as exc:  # noqa: BLE001
                self._loaded_urls = None
                self._preview_status.set(f"読み込み失敗: {exc}")
        else:
            self._loaded_urls = None
            self._preview_status.set("参照元 Excel 未指定 → デフォルト URL を使用します")

        rows = self._loaded_urls
        if not rows:
            src = self._source_var.get()
            if src == "3gpp":
                base = DEFAULT_3GPP_URLS
            elif src == "ieee":
                base = DEFAULT_IEEE_URLS
            else:
                base = ()
            rows = [("(default)", u) for u in base]

        self._preview_text.configure(state=tk.NORMAL)
        self._preview_text.delete("1.0", tk.END)
        for i, (title, url) in enumerate(rows, start=1):
            display_title = title if title else "(no title)"
            self._preview_text.insert(tk.END, f"{i:2d}. {display_title}\n      {url}\n")
        self._preview_text.configure(state=tk.DISABLED)

    # ---------------------------------------------------------------- finalize
    def _on_cancel(self) -> None:
        self._result = None
        self.destroy()

    def _on_build(self) -> None:
        save_dir = Path(self._save_dir_var.get().strip() or str(default_output_root()))
        filename = self._filename_var.get().strip() or "sample_download.xlsx"
        if not filename.lower().endswith(".xlsx"):
            filename += ".xlsx"
        out_path = save_dir / filename

        try:
            save_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(
                "Folder error", f"Cannot create folder: {save_dir}\n{exc}", parent=self,
            )
            return

        urls = self._loaded_urls if self._loaded_urls else None

        try:
            built = build_harvest_sample(
                out_path,
                source_type=self._source_var.get(),
                output_root=self._outroot_var.get().strip() or None,
                job_name=self._jobname_var.get().strip() or None,
                proxy_url=self._proxy_var.get().strip(),
                proxy_user=self._proxyuser_var.get().strip(),
                proxy_password=self._proxypass_var.get(),
                urls=urls,
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Build failed", f"{exc}", parent=self)
            return

        self._result = built
        self.destroy()

    # ---------------------------------------------------------------- entry
    @classmethod
    def run(cls, master: tk.Misc, *, source_type: str = "3gpp",
            default_dir: Optional[Path] = None) -> Optional[Path]:
        dlg = cls(master, source_type=source_type, default_dir=default_dir)
        master.wait_window(dlg)
        return dlg._result
