"""Launcher kept so existing desktop shortcuts keep working.

The application now lives in the ``pdf_to_word`` package; run it with
``python -m pdf_to_word`` or through this file.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pdf_to_word.__main__ import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
