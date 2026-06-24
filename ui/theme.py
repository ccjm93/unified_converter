"""Catppuccin Mocha 기반 단일 테마 팔레트/폰트/ttk 스타일."""
from __future__ import annotations

# Catppuccin Mocha
BG = "#1e1e2e"        # base
SURFACE = "#313244"   # surface0
SURFACE2 = "#45475a"  # surface1
TEXT = "#cdd6f4"      # text
SUBTEXT = "#a6adc8"   # subtext0
ACCENT = "#cba6f7"    # mauve
BLUE = "#89b4fa"
GREEN = "#a6e3a1"
YELLOW = "#f9e2af"
RED = "#f38ba8"
CRUST = "#11111b"

FONT_H = ("Segoe UI", 17, "bold")
FONT_B = ("Segoe UI", 10, "bold")
FONT_N = ("Segoe UI", 10)
MONO = ("Consolas", 9)


def apply_ttk(root) -> None:
    """진행바 등 ttk 위젯 스타일을 등록한다."""
    from tkinter import ttk

    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(
        "Accent.Horizontal.TProgressbar",
        troughcolor=SURFACE,
        background=ACCENT,
        bordercolor=BG,
        lightcolor=ACCENT,
        darkcolor=ACCENT,
        thickness=10,
    )
