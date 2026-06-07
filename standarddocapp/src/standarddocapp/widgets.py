"""Reusable Tk widgets shared by tabs."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Callable, Optional


class ScrollableFrame(ttk.Frame):
    """A vertically scrollable container that hosts a regular ttk.Frame.

    Children should be packed/gridded into ``self.inner``.
    """

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master)
        self._canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.inner = ttk.Frame(self._canvas, padding=8)
        self._win = self._canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>", self._on_inner_configure)
        self._canvas.bind("<Configure>", self._on_canvas_configure)

        self._canvas.bind("<Enter>", self._activate_wheel)
        self._canvas.bind("<Leave>", self._deactivate_wheel)

    def _on_inner_configure(self, _event: tk.Event) -> None:
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self._canvas.itemconfigure(self._win, width=event.width)

    def _activate_wheel(self, _e: tk.Event) -> None:
        self._canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self._canvas.bind_all("<Button-4>", self._on_mousewheel_x11)
        self._canvas.bind_all("<Button-5>", self._on_mousewheel_x11)

    def _deactivate_wheel(self, _e: tk.Event) -> None:
        self._canvas.unbind_all("<MouseWheel>")
        self._canvas.unbind_all("<Button-4>")
        self._canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event: tk.Event) -> None:
        delta = -1 * int(event.delta / 120) if event.delta else 0
        if delta:
            self._canvas.yview_scroll(delta, "units")

    def _on_mousewheel_x11(self, event: tk.Event) -> None:
        if event.num == 4:
            self._canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self._canvas.yview_scroll(1, "units")


class CollapsibleFrame(ttk.Frame):
    """A LabelFrame-style header you can fold/unfold."""

    def __init__(
        self, master: tk.Misc, title: str, expanded: bool = True,
        on_toggle: Optional[Callable[[bool], None]] = None,
    ) -> None:
        super().__init__(master)
        self._expanded = tk.BooleanVar(value=expanded)
        self._on_toggle = on_toggle
        self._toggle_btn = ttk.Button(
            self, text="", style="Toolbutton", command=self._toggle,
        )
        self._toggle_btn.grid(row=0, column=0, sticky="w")
        self._title_lbl = ttk.Label(self, text=title, font=("Segoe UI", 10, "bold"))
        self._title_lbl.grid(row=0, column=1, sticky="w", padx=4)
        self.columnconfigure(1, weight=1)
        self.body = ttk.Frame(self)
        self.body.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(4, 0))
        self._refresh()

    def _refresh(self) -> None:
        if self._expanded.get():
            self._toggle_btn.configure(text="\u25bc")  # down triangle
            self.body.grid()
        else:
            self._toggle_btn.configure(text="\u25b6")  # right triangle
            self.body.grid_remove()
        if self._on_toggle is not None:
            self._on_toggle(self._expanded.get())

    def _toggle(self) -> None:
        self._expanded.set(not self._expanded.get())
        self._refresh()


class StatusBar(ttk.Frame):
    """Footer status bar with left text + right text."""

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master)
        ttk.Separator(self, orient="horizontal").pack(fill=tk.X)
        bar = ttk.Frame(self, padding=(8, 2))
        bar.pack(fill=tk.X)
        self._left_var = tk.StringVar(value="")
        self._right_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self._left_var, foreground="#555").pack(side=tk.LEFT)
        ttk.Label(bar, textvariable=self._right_var, foreground="#888").pack(side=tk.RIGHT)

    def set_left(self, text: str) -> None:
        self._left_var.set(text)

    def set_right(self, text: str) -> None:
        self._right_var.set(text)
