"""Theme colours and the custom round widgets.

Grey ground based on the Yaru Files icon, a shade darker, with dark-green
actions. Fields are grey too, set apart only by a darker border.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

BG = "#888888"           # window background (darker Files-icon grey)
CARD = BG                # no panels: everything sits on the one grey
FIELD = BG               # fields are the same grey, outlined only
FIELD_BORDER = "#5E5E5E"
GREEN = "#0B5D3B"        # primary dark green
GREEN_HOVER = "#0F7549"
GREEN_PRESSED = "#084A2F"
GREEN_LIGHT = "#9FB5A6"  # hover tint behind outline buttons
TEXT = "#1C1C1C"
MUTED = "#2A2A2A"
BORDER = "#6E6E6E"
DISABLED = "#737373"
DISABLED_TEXT = "#B5B5B5"
TRACK_OFF = "#666666"
TRACK_DISABLED = "#7A7A7A"
TROUGH = "#727272"


def pick_font_family(root) -> str:
    from tkinter import font as tkfont
    available = set(tkfont.families(root))
    for family in ("Segoe UI", "Ubuntu", "Cantarell", "Noto Sans",
                   "DejaVu Sans", "Helvetica"):
        if family in available:
            return family
    return "TkDefaultFont"


def round_rect(canvas, x1, y1, x2, y2, r, **kw):
    """Draw a rounded rectangle (a pill when r = height / 2)."""
    points = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
              x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
              x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return canvas.create_polygon(points, smooth=True, **kw)


class RoundButton(tk.Canvas):
    """Pill-shaped button with hover/press states; supports config(state=)."""

    def __init__(self, parent, text, command, width=150, height=42,
                 filled=True, font=None):
        super().__init__(parent, width=width, height=height,
                         bg=parent["bg"], highlightthickness=0,
                         cursor="hand2")
        self._text, self._command = text, command
        self._filled, self._font = filled, font
        self._bw, self._bh = width, height
        self._state, self._hover, self._pressed = "normal", False, False
        self.bind("<Enter>", lambda e: self._set(hover=True))
        self.bind("<Leave>", lambda e: self._set(hover=False, pressed=False))
        self.bind("<ButtonPress-1>", lambda e: self._set(pressed=True))
        self.bind("<ButtonRelease-1>", self._release)
        self._draw()

    def _set(self, **kw):
        for key, value in kw.items():
            setattr(self, "_" + key, value)
        self._draw()

    def _release(self, event):
        was_pressed = self._pressed
        self._set(pressed=False)
        inside = 0 <= event.x <= self._bw and 0 <= event.y <= self._bh
        if was_pressed and inside and self._state == "normal" and self._command:
            self._command()

    def _draw(self):
        self.delete("all")
        if self._state == "disabled":
            fill, outline, fg = DISABLED, DISABLED, DISABLED_TEXT
        elif self._filled:
            fill = (GREEN_PRESSED if self._pressed
                    else GREEN_HOVER if self._hover else GREEN)
            outline, fg = fill, "#FFFFFF"
        else:
            fill = GREEN_LIGHT if (self._hover or self._pressed) else self["bg"]
            outline, fg = GREEN, GREEN
        round_rect(self, 2, 2, self._bw - 2, self._bh - 2, (self._bh - 4) / 2,
                   fill=fill, outline=outline, width=2)
        self.create_text(self._bw / 2, self._bh / 2, text=self._text,
                         fill=fg, font=self._font)

    def configure(self, cnf=None, **kw):
        if "state" in kw:
            self._state = kw.pop("state")
            super().configure(
                cursor="hand2" if self._state == "normal" else "arrow")
            self._draw()
        if cnf or kw:
            return super().configure(cnf, **kw)

    config = configure


class Toggle(tk.Frame):
    """iOS-style switch bound to a BooleanVar, with label and hint line."""

    def __init__(self, parent, text, variable, hint="", enabled=True,
                 font=None, hint_font=None):
        super().__init__(parent, bg=parent["bg"])
        self.var, self.enabled = variable, enabled
        cursor = "hand2" if enabled else "arrow"
        self.cv = tk.Canvas(self, width=46, height=26, bg=parent["bg"],
                            highlightthickness=0, cursor=cursor)
        self.cv.grid(row=0, column=0, rowspan=2, padx=(0, 12), sticky="n")
        label = tk.Label(self, text=text, bg=parent["bg"],
                         fg=TEXT if enabled else MUTED, font=font,
                         cursor=cursor)
        label.grid(row=0, column=1, sticky="w")
        if hint:
            tk.Label(self, text=hint, bg=parent["bg"], fg=MUTED,
                     font=hint_font).grid(row=1, column=1, sticky="w")
        for widget in (self.cv, label):
            widget.bind("<Button-1>", self._toggle)
        variable.trace_add("write", lambda *a: self._draw())
        self._draw()

    def _draw(self):
        c = self.cv
        c.delete("all")
        on = bool(self.var.get())
        track = (GREEN if on else TRACK_OFF) if self.enabled else TRACK_DISABLED
        round_rect(c, 1, 1, 45, 25, 12, fill=track, outline="")
        x = 32 if on else 14
        c.create_oval(x - 10, 3, x + 10, 23, fill="#FFFFFF", outline="")

    def _toggle(self, _event=None):
        if self.enabled:
            self.var.set(not self.var.get())


def apply_theme(root):
    root.configure(bg=BG)
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Green.Horizontal.TProgressbar", troughcolor=TROUGH,
                    background=GREEN, bordercolor=TROUGH,
                    lightcolor=GREEN, darkcolor=GREEN, thickness=10)
    style.configure("Log.Vertical.TScrollbar", troughcolor=CARD,
                    background=BORDER, bordercolor=CARD,
                    arrowcolor=GREEN)
