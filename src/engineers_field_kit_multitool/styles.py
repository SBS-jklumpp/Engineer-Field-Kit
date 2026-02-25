import tkinter as tk
from tkinter import ttk

# Cohesive modern palette: dark slate surfaces + muted amber accent.
DARK_BG = "#121821"
DARK_PANEL = "#1a2330"
DARK_PANEL_2 = "#212c3b"
DARK_BORDER = "#2f3a4a"
DARK_TEXT = "#e7ecf3"
DARK_MUTED = "#9ea9b8"
DARK_ACCENT = "#caa06a"
DARK_ACCENT_HOVER = "#d6ad79"
DARK_WARN = "#d8a873"
DARK_OK = "#31b28a"

LIGHT_BG = "#f5f5f5"
LIGHT_PANEL = "#ffffff"
LIGHT_PANEL_2 = "#e9ecef"
LIGHT_BORDER = "#c7ced8"
LIGHT_TEXT = "#1f2937"
LIGHT_MUTED = "#5b6472"
LIGHT_ACCENT = "#0f62fe"
LIGHT_ACCENT_HOVER = "#3d7bff"
LIGHT_WARN = "#c46f00"
LIGHT_OK = "#0f7b5f"

DEFAULT_FONT = ("Segoe UI", 9)


def apply_theme(root: tk.Tk, field_mode: bool = False, dark_mode: bool = True):
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    if dark_mode:
        bg = DARK_BG
        panel = DARK_PANEL
        panel_2 = DARK_PANEL_2
        border = DARK_BORDER
        text = DARK_TEXT
        muted = DARK_MUTED
        accent = DARK_ACCENT
        accent_hover = DARK_ACCENT_HOVER
    else:
        bg = LIGHT_BG
        panel = LIGHT_PANEL
        panel_2 = LIGHT_PANEL_2
        border = LIGHT_BORDER
        text = LIGHT_TEXT
        muted = LIGHT_MUTED
        accent = LIGHT_ACCENT
        accent_hover = LIGHT_ACCENT_HOVER

    font_size = 11 if field_mode else 9
    strong_text = "#f4f7fc" if (field_mode and dark_mode) else text
    strong_muted = "#bdc8d8" if (field_mode and dark_mode) else muted
    strong_border = "#59677d" if (field_mode and dark_mode) else border
    button_padding = (14, 9) if field_mode else (10, 6)
    primary_padding = (16, 10) if field_mode else (12, 7)
    tab_padding = [16, 9] if field_mode else [14, 7]
    row_height = 30 if field_mode else 24

    root.configure(bg=bg)

    # Shared defaults keep spacing and typography consistent across ttk widgets.
    style.configure(
        ".",
        background=bg,
        foreground=strong_text,
        fieldbackground=panel,
        bordercolor=strong_border,
        borderwidth=1,
        relief="flat",
        font=("Segoe UI", font_size),
    )
    style.configure("TFrame", background=bg)
    style.configure(
        "TLabelframe",
        background=bg,
        foreground=strong_text,
        bordercolor=strong_border,
        borderwidth=1,
        relief="flat",
    )
    style.configure("TLabelframe.Label", background=bg, foreground=accent, font=("Segoe UI Semibold", font_size))
    style.configure("TLabel", background=bg, foreground=strong_text)

    style.configure("TButton", background=panel_2, foreground=strong_text, bordercolor=strong_border, padding=button_padding, relief="flat")
    style.map(
        "TButton",
        background=[("active", "#2a3648" if dark_mode else "#dbe3ee"), ("pressed", "#172130" if dark_mode else "#d1dae6"), ("disabled", panel)],
        foreground=[("disabled", strong_muted)],
    )
    style.configure(
        "Primary.TButton",
        background=accent,
        foreground="#141b24" if dark_mode else "#ffffff",
        bordercolor=accent,
        padding=primary_padding,
        relief="flat",
        font=("Segoe UI Semibold", font_size),
    )
    style.map(
        "Primary.TButton",
        background=[("active", accent_hover), ("pressed", "#bf935a" if dark_mode else "#0b4fd1"), ("disabled", panel)],
        foreground=[("disabled", strong_muted)],
    )

    style.configure("TEntry", fieldbackground=panel, foreground=strong_text, bordercolor=strong_border, insertcolor=strong_text, padding=4)
    style.configure("TCombobox", fieldbackground=panel, background=panel, foreground=strong_text, bordercolor=strong_border, padding=3)
    style.map(
        "TCombobox",
        fieldbackground=[("readonly", panel), ("active", panel_2)],
        foreground=[("readonly", strong_text)],
    )
    style.configure("TSpinbox", fieldbackground=panel, foreground=strong_text, bordercolor=strong_border, padding=2)
    style.configure("TCheckbutton", background=bg, foreground=strong_text, padding=(2, 1))
    style.map("TCheckbutton", background=[("active", bg)], foreground=[("active", strong_text)])
    style.configure("TRadiobutton", background=bg, foreground=strong_text, padding=(2, 1))
    style.map("TRadiobutton", background=[("active", bg)], foreground=[("active", strong_text)])

    style.configure("TNotebook", background=bg, bordercolor=strong_border, tabmargins=[4, 7, 4, 0])
    style.configure("TNotebook.Tab", background=panel_2, foreground=strong_muted, padding=tab_padding, bordercolor=strong_border, relief="flat")
    style.map(
        "TNotebook.Tab",
        background=[("selected", panel), ("active", "#2a3648" if dark_mode else "#dbe3ee")],
        foreground=[("selected", strong_text), ("active", strong_text)],
    )

    style.configure("Treeview", background=panel, fieldbackground=panel, foreground=strong_text, bordercolor=strong_border, rowheight=row_height)
    style.map("Treeview", background=[("selected", "#3a4a60" if dark_mode else "#cfe2ff")], foreground=[("selected", "#f7fafc" if dark_mode else "#111827")])
    style.configure("Treeview.Heading", background=panel_2, foreground=strong_text, bordercolor=strong_border, padding=(8, 6))
    style.map("Treeview.Heading", background=[("active", "#2c3a4f" if dark_mode else "#dbe3ee")])

    style.configure("Vertical.TScrollbar", background=panel_2, troughcolor=bg, bordercolor=strong_border, arrowcolor=strong_muted)
    style.configure("Horizontal.TScrollbar", background=panel_2, troughcolor=bg, bordercolor=strong_border, arrowcolor=strong_muted)
