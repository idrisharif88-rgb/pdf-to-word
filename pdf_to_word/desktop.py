"""Small OS integrations: completion sound and revealing the output file."""
from __future__ import annotations

import logging
import os
import platform
import shutil
import subprocess

log = logging.getLogger(__name__)

SYSTEM = platform.system()


def notify_sound() -> None:
    """Play a short completion sound; silently skip if unavailable."""
    try:
        if SYSTEM == "Windows":
            import winsound  # type: ignore
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
        elif shutil.which("canberra-gtk-play"):
            subprocess.Popen(["canberra-gtk-play", "-i", "complete"])
        elif shutil.which("paplay"):
            subprocess.Popen(
                ["paplay", "/usr/share/sounds/freedesktop/stereo/complete.oga"])
    except Exception as e:
        log.warning("Could not play notification sound: %s", e)


def reveal_in_file_manager(path: str) -> None:
    """Open the folder containing ``path`` in the desktop file manager."""
    path = os.path.abspath(path)
    if SYSTEM == "Windows":
        subprocess.Popen(["explorer", "/select,", path])
    elif SYSTEM == "Darwin":
        subprocess.Popen(["open", "-R", path])
    else:
        subprocess.Popen(["xdg-open", os.path.dirname(path)])
