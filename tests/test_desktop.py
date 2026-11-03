import subprocess

import pytest

from pdf_to_word import desktop


class Done:
    def __init__(self, code, out=""):
        self.returncode, self.stdout, self.stderr = code, out, ""


@pytest.fixture
def linux(monkeypatch):
    monkeypatch.setattr(desktop, "SYSTEM", "Linux")
    monkeypatch.setattr(desktop.shutil, "which", lambda name: "/usr/bin/zenity")


def test_open_dialog_builds_command(linux, monkeypatch, tmp_path):
    seen = {}

    def fake_run(cmd, **kw):
        seen["cmd"] = cmd
        return Done(0, str(tmp_path / "a.pdf") + "\n")
    monkeypatch.setattr(desktop.subprocess, "run", fake_run)
    path = desktop.native_file_dialog("Pick", str(tmp_path),
                                      filters=[("PDF", ["*.pdf", "*.PDF"])])
    assert path == str(tmp_path / "a.pdf")
    assert f"--filename={tmp_path}/" in seen["cmd"]
    assert "--file-filter=PDF | *.pdf *.PDF" in seen["cmd"]
    assert "--save" not in seen["cmd"]


def test_save_dialog_and_cancel(linux, monkeypatch):
    calls = []
    monkeypatch.setattr(desktop.subprocess, "run",
                        lambda cmd, **kw: calls.append(cmd) or Done(1))
    assert desktop.native_file_dialog("Save", "/x/book.docx", save=True) is None
    assert "--save" in calls[0] and "--filename=/x/book.docx" in calls[0]


def test_falls_back_when_missing_or_broken(monkeypatch):
    monkeypatch.setattr(desktop, "SYSTEM", "Linux")
    monkeypatch.setattr(desktop.shutil, "which", lambda name: None)
    with pytest.raises(desktop.NoNativeDialog):
        desktop.native_file_dialog("Pick")
    monkeypatch.setattr(desktop.shutil, "which", lambda name: "/usr/bin/zenity")
    monkeypatch.setattr(desktop.subprocess, "run", lambda cmd, **kw: Done(255))
    with pytest.raises(desktop.NoNativeDialog):
        desktop.native_file_dialog("Pick")
