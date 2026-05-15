import tkinter as tk
from tkinter import ttk

from ui.widgets import ToolTip

PAD = {"padx": 8, "pady": 4}


def make_field(parent, row, label, var, width=40, tooltip=None):
    ttk.Label(parent, text=label).grid(row=row, column=0, sticky="e", **PAD)
    entry = ttk.Entry(parent, textvariable=var, width=width)
    entry.grid(row=row, column=1, sticky="w", **PAD)
    if tooltip:
        ToolTip(entry, tooltip)
    return entry


def make_label_section(parent, row, text):
    ttk.Separator(parent, orient="horizontal").grid(
        row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=(12, 2))
    ttk.Label(parent, text=text, font=("Segoe UI", 10, "bold"),
              foreground="#1a5276").grid(
        row=row + 1, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 4))
    return row + 2
