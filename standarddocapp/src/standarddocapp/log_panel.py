"""Reusable log display widget (ScrolledText with auto-scroll & clear)."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText


class LogPanel(ttk.Frame):
    """A bordered, auto-scrolling, monospace log display."""

    def __init__(self, master: tk.Misc, title: str = "Log") -> None:
        super().__init__(master)
        header = ttk.Frame(self)
        header.pack(fill=tk.X)
        ttk.Label(header, text=title, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)

        self._auto_scroll = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            header, text="Auto-scroll", variable=self._auto_scroll
        ).pack(side=tk.RIGHT)
        ttk.Button(header, text="Clear", command=self.clear, width=8).pack(
            side=tk.RIGHT, padx=4
        )

        self._text = ScrolledText(
            self,
            height=14,
            wrap=tk.NONE,
            state=tk.DISABLED,
            font=("Consolas", 9),
            background="#1e1e1e",
            foreground="#dcdcdc",
            insertbackground="#dcdcdc",
        )
        self._text.pack(fill=tk.BOTH, expand=True, pady=(4, 0))
        self._text.tag_configure("ERROR", foreground="#f48771")
        self._text.tag_configure("WARNING", foreground="#dcdc7a")
        self._text.tag_configure("INFO", foreground="#dcdcdc")
        self._text.tag_configure("DEBUG", foreground="#9090a0")

    def append(self, line: str) -> None:
        if not line.endswith("\n"):
            line = line + "\n"
        tag = "INFO"
        for level in ("ERROR", "WARNING", "DEBUG"):
            if f"[{level}]" in line:
                tag = level
                break
        self._text.configure(state=tk.NORMAL)
        self._text.insert(tk.END, line, tag)
        if self._auto_scroll.get():
            self._text.see(tk.END)
        self._text.configure(state=tk.DISABLED)

    def clear(self) -> None:
        self._text.configure(state=tk.NORMAL)
        self._text.delete("1.0", tk.END)
        self._text.configure(state=tk.DISABLED)
