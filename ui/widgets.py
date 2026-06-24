"""tkinter 위젯 헬퍼."""
from __future__ import annotations

import tkinter as tk

from . import theme as T


def make_button(parent, text, command, *, bg=T.ACCENT, fg=T.CRUST,
                hover=T.BLUE, font=T.FONT_B, padx=14, pady=6):
    b = tk.Button(
        parent, text=text, command=command,
        bg=bg, fg=fg, activebackground=hover, activeforeground=fg,
        font=font, relief="flat", bd=0, padx=padx, pady=pady, cursor="hand2",
    )
    b.bind("<Enter>", lambda e: b.config(bg=hover))
    b.bind("<Leave>", lambda e: b.config(bg=bg))
    return b
