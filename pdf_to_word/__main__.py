"""Entry point: ``python -m pdf_to_word``."""
from __future__ import annotations

import logging
import sys

from .config import PROJECT_DIR, Settings


def setup_logging() -> None:
    logging.basicConfig(
        filename=str(PROJECT_DIR / "ocr_gui.log"),
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )


def main() -> int:
    setup_logging()
    import tkinter as tk
    from tkinter import messagebox

    try:
        settings = Settings.from_env()
        root = tk.Tk(className="pdf-to-word")  # matches StartupWMClass
        from .gui.app import OCRApp
        OCRApp(root, settings)
        root.after(200, root.focus_force)
        root.mainloop()
        return 0
    except Exception as e:
        logging.exception("Fatal error")
        messagebox.showerror("Startup Error", f"Fatal error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
