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


class NoNativeDialog(Exception):
    """No native file chooser on this system; caller falls back to Tk."""


def native_file_dialog(title: str, initial: str = "", save: bool = False,
                       filters=()) -> "str | None":
    """Show the desktop's own file chooser (GNOME/GTK via zenity).

    Returns the chosen path, or None if the user cancelled. ``initial`` is a
    folder or a file path to start from. ``filters`` is a sequence of
    (label, [patterns]) pairs. Blocks, so call it off the Tk thread.
    Raises NoNativeDialog when zenity is unavailable or fails.
    """
    zenity = shutil.which("zenity") if SYSTEM == "Linux" else None
    if not zenity:
        raise NoNativeDialog()
    cmd = [zenity, "--file-selection", f"--title={title}"]
    if save:
        cmd.append("--save")  # zenity asks before overwriting
    if initial:
        if os.path.isdir(initial):
            initial = initial.rstrip("/") + "/"  # trailing slash = open folder
        cmd.append(f"--filename={initial}")
    for label, patterns in filters:
        cmd.append(f"--file-filter={label} | {' '.join(patterns)}")
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True)
    except OSError as e:
        raise NoNativeDialog() from e
    if proc.returncode == 0:
        return proc.stdout.strip() or None
    if proc.returncode == 1:  # cancelled or closed
        return None
    log.warning("zenity failed (%s): %s", proc.returncode, proc.stderr.strip())
    raise NoNativeDialog()
