"""Tkinter front end. All conversion work lives in ``pipeline``; this module
only collects options, starts a job on a worker thread and shows progress.
"""
from __future__ import annotations

import dataclasses
import logging
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from typing import Optional

from .. import pdf_tools
from ..config import Settings
from ..desktop import (NoNativeDialog, native_file_dialog, notify_sound,
                       reveal_in_file_manager)
from ..pipeline import ConversionJob, JobCancelled
from ..text import parse_page_range
from .widgets import (ACCENT, BG, BORDER, CARD, FIELD, FIELD_BORDER, GREEN, MUTED, TEXT,
                      RoundButton, Toggle, apply_theme, pick_font_family)

log = logging.getLogger(__name__)
LOG_LINES_KEPT = 200
ASSETS = Path(__file__).resolve().parent.parent / "assets"


class OCRApp:
    def __init__(self, root: tk.Tk, settings: Settings):
        self.root = root
        self.settings = settings
        self.job: Optional[ConversionJob] = None
        self._picking = False  # one file chooser at a time
        apply_theme(root)
        root.title("PDF to Word — Arabic OCR")
        self._set_window_icon()
        root.geometry("920x800")
        root.minsize(820, 720)

        family = pick_font_family(root)
        self.f_title = (family, 18, "bold")
        self.f_sub = (family, 10)
        self.f_card = (family, 12, "bold")
        self.f_label = (family, 10, "bold")
        self.f_body = (family, 10)
        self.f_hint = (family, 9)
        self.f_btn = (family, 11, "bold")
        self._build()

    def _set_window_icon(self) -> None:
        """Window and dock icon (Scan to text); kept on self so Tk's
        PhotoImage is not garbage-collected."""
        try:
            self._icons = [tk.PhotoImage(file=str(ASSETS / f"icon-{n}.png"))
                           for n in (256, 64, 32)]
            self.root.iconphoto(True, *self._icons)
        except tk.TclError:
            log.warning("Window icon not set", exc_info=True)

    # ================================================================ layout
    def _build(self) -> None:
        root, s = self.root, self.settings
        header = tk.Frame(root, bg=BG)
        header.pack(fill="x")
        tk.Label(header, text="PDF  →  Word", bg=BG, fg=ACCENT,
                 font=self.f_title).pack(anchor="w", padx=24, pady=(16, 0))
        tk.Label(header, text="Arabic OCR · Google Document AI", bg=BG,
                 fg=MUTED, font=self.f_sub).pack(anchor="w", padx=24,
                                                     pady=(0, 16))

        body = tk.Frame(root, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=16)
        body.columnconfigure(0, weight=1, uniform="col")
        body.columnconfigure(1, weight=1, uniform="col")
        body.rowconfigure(4, weight=1)

        files = self._card(body, "Files")
        files.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        self.file_entry = self._field_row(files, "Input file (PDF or image)",
                                          "Browse", self.browse_file)
        self.output_entry = self._field_row(files, "Output Word file",
                                            "Change", self.choose_output)

        pages = self._card(body, "Pages")
        pages.grid(row=1, column=0, sticky="nsew", padx=(0, 6), pady=(0, 12))
        self._label(pages, "Page range", "all, 1-5, 2,5,7")
        self.page_entry = self._entry(pages)
        self.page_entry.insert(0, "all")
        self.page_entry.pack(fill="x", padx=16, pady=(0, 10), ipady=6)
        self._label(pages, "Keep as pictures", "e.g. 1-6,140-146")
        self.image_pages_entry = self._entry(pages)
        self.image_pages_entry.insert(0, s.image_pages)
        self.image_pages_entry.pack(fill="x", padx=16, pady=(0, 16), ipady=6)

        options = self._card(body, "Options")
        options.grid(row=1, column=1, sticky="nsew", padx=(6, 0), pady=(0, 12))
        self.quality_var = tk.BooleanVar(value=s.enable_quality_scores)
        self.cleanup_var = tk.BooleanVar(value=s.cleanup)
        self.exact_var = tk.BooleanVar(value=s.exact_layout)
        has_form = s.form_parser_available
        self.tables_var = tk.BooleanVar(value=s.extract_tables and has_form)
        for text, var, hint, enabled in (
            ("Quality scoring", self.quality_var,
             "Image quality scores from OCR", True),
            ("Smart text cleanup", self.cleanup_var,
             "Tidy spacing and OCR noise", True),
            ("Exact layout", self.exact_var,
             "Looks like the scan; text does not reflow", True),
            ("Extract tables (Form Parser)", self.tables_var,
             "Extra paid request per batch" if has_form
             else "Set GCP_FORM_PROCESSOR_ID in .env to enable", has_form),
        ):
            Toggle(options, text, var, hint=hint, enabled=enabled,
                   font=self.f_label, hint_font=self.f_hint)\
                .pack(anchor="w", padx=16, pady=(0, 10))

        actions = tk.Frame(body, bg=BG)
        actions.grid(row=2, column=0, columnspan=2, pady=(4, 10))
        self.run_btn = RoundButton(actions, "Run OCR", self.run_ocr,
                                   width=170, height=44, font=self.f_btn)
        self.run_btn.grid(row=0, column=0, padx=8)
        self.cancel_btn = RoundButton(actions, "Cancel", self.cancel_ocr,
                                      width=130, height=44, filled=False,
                                      font=self.f_btn)
        self.cancel_btn.grid(row=0, column=1, padx=8)
        self.cancel_btn.config(state="disabled")

        status = tk.Frame(body, bg=BG)
        status.grid(row=3, column=0, columnspan=2, sticky="ew")
        self.status_label = tk.Label(status, text="Ready.", anchor="w",
                                     fg=ACCENT, bg=BG, font=self.f_label)
        self.status_label.pack(fill="x", pady=(0, 6))
        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(status, variable=self.progress_var, maximum=100,
                        style="Green.Horizontal.TProgressbar")\
            .pack(fill="x", pady=(0, 12))

        log_card = self._card(body, "Activity")
        log_card.grid(row=4, column=0, columnspan=2, sticky="nsew")
        inner = tk.Frame(log_card, bg=CARD)
        inner.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        self.log_box = tk.Text(inner, height=8, bg=FIELD, fg=TEXT,
                               font=("DejaVu Sans Mono", 9), relief="flat",
                               highlightthickness=1, highlightbackground=FIELD_BORDER,
                               padx=8, pady=6, state="disabled")
        scroll = ttk.Scrollbar(inner, command=self.log_box.yview,
                               style="Log.Vertical.TScrollbar")
        self.log_box.config(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_box.pack(side="left", fill="both", expand=True)

    def _card(self, parent, title):
        card = tk.Frame(parent, bg=CARD)  # a group, not a box
        tk.Label(card, text=title, bg=CARD, fg=ACCENT, font=self.f_card)\
            .pack(anchor="w", padx=16, pady=(12, 8))
        return card

    def _label(self, parent, text, hint=""):
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill="x", padx=16, pady=(0, 4))
        tk.Label(row, text=text, bg=CARD, fg=TEXT, font=self.f_label)\
            .pack(side="left")
        if hint:
            tk.Label(row, text=f"  {hint}", bg=CARD, fg=MUTED,
                     font=self.f_hint).pack(side="left")

    def _entry(self, parent):
        return tk.Entry(parent, bg=FIELD, fg=TEXT, insertbackground=ACCENT,
                        relief="flat", highlightthickness=1,
                        highlightbackground=FIELD_BORDER, highlightcolor=ACCENT,
                        font=self.f_body)

    def _field_row(self, parent, label, button_text, command):
        self._label(parent, label)
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill="x", padx=16, pady=(0, 12))
        entry = self._entry(row)
        entry.pack(side="left", fill="x", expand=True, ipady=7, padx=(0, 10))
        RoundButton(row, button_text, command, width=110, height=36,
                    filled=False, font=self.f_label).pack(side="right")
        return entry

    # ============================================== thread-safe UI updates
    def _ui(self, fn, *args) -> None:
        """Run ``fn`` on the Tk thread (Tk is not thread-safe)."""
        self.root.after(0, fn, *args)

    def log(self, msg: str) -> None:
        self._ui(self._append_log, msg)

    def _append_log(self, msg: str) -> None:
        self.log_box.config(state="normal")
        self.log_box.insert("end", msg + "\n")
        lines = int(self.log_box.index("end-1c").split(".")[0])
        if lines > LOG_LINES_KEPT:
            self.log_box.delete("1.0", f"{lines - LOG_LINES_KEPT + 20}.0")
        self.log_box.see("end")
        self.log_box.config(state="disabled")

    def set_status(self, msg: str) -> None:
        self._ui(lambda: self.status_label.config(text=msg))

    def set_progress(self, value: float) -> None:
        self._ui(self.progress_var.set, value)

    def _set_running(self, running: bool) -> None:
        self.run_btn.config(state="disabled" if running else "normal")
        self.cancel_btn.config(state="normal" if running else "disabled")

    # ================================================================ actions
    # --- file pickers: the desktop's own chooser, Tk's as a fallback ---
    def _pick_file(self, title: str, initial: str, save: bool, filters,
                   on_chosen) -> None:
        """Open a chooser without freezing the window; ``on_chosen(path)``
        runs on the Tk thread only when a file was picked."""
        if self._picking:
            return
        self._picking = True

        def work():
            try:
                path = native_file_dialog(title, initial, save, filters)
            except NoNativeDialog:
                self._ui(tk_fallback)
                return
            self._ui(finish, path)

        def tk_fallback():
            types = [(label, " ".join(p)) for label, p in filters]
            folder = initial if os.path.isdir(initial) else os.path.dirname(initial)
            kwargs = dict(title=title, filetypes=types,
                          initialdir=folder or os.path.expanduser("~"))
            if save:
                path = filedialog.asksaveasfilename(
                    defaultextension=".docx",
                    initialfile=os.path.basename(initial)
                    if not os.path.isdir(initial) else "", **kwargs)
            else:
                path = filedialog.askopenfilename(**kwargs)
            finish(path)

        def finish(path):
            self._picking = False
            if path:
                on_chosen(path)

        threading.Thread(target=work, daemon=True).start()

    def _start_folder(self, entry) -> str:
        current = os.path.expanduser(entry.get().strip())
        folder = os.path.dirname(current) if current else ""
        if folder and os.path.isdir(folder):
            return folder
        downloads = os.path.expanduser("~/Downloads")
        return downloads if os.path.isdir(downloads) else os.path.expanduser("~")

    def browse_file(self) -> None:
        patterns = [f"*{s}" for s in pdf_tools.SUPPORTED_SUFFIXES]
        patterns += [p.upper() for p in patterns]
        self._pick_file("Choose a PDF or image", self._start_folder(self.file_entry),
                        save=False,
                        filters=[("PDF or image", patterns), ("All files", ["*"])],
                        on_chosen=self._set_input)

    def _set_input(self, filename: str) -> None:
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, filename)
        stem = os.path.splitext(os.path.basename(filename))[0]
        output = os.path.join(os.path.dirname(filename), f"{stem}.docx")
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, output)
        self.log(f"Output folder: {os.path.dirname(filename)}")

    def choose_output(self) -> None:
        current = os.path.expanduser(self.output_entry.get().strip())
        initial = current if current and os.path.isdir(os.path.dirname(current)) \
            else self._start_folder(self.output_entry)
        self._pick_file("Save Word file as", initial, save=True,
                        filters=[("Word document", ["*.docx"])],
                        on_chosen=self._set_output)

    def _set_output(self, filename: str) -> None:
        if not filename.lower().endswith(".docx"):
            filename += ".docx"
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, filename)

    def _job_settings(self) -> Settings:
        """This run's settings: a copy, so a running job never sees edits."""
        return dataclasses.replace(
            self.settings,
            enable_quality_scores=self.quality_var.get(),
            cleanup=self.cleanup_var.get(),
            exact_layout=self.exact_var.get(),
            extract_tables=self.tables_var.get(),
            image_pages=self.image_pages_entry.get().strip(),
        )

    def run_ocr(self) -> None:
        if self.job is not None:
            return
        input_path = self.file_entry.get().strip()
        output_path = self.output_entry.get().strip()
        if not input_path or not output_path:
            messagebox.showerror("Error", "Please select input and output files.")
            return
        if not os.path.isfile(input_path):
            messagebox.showerror("Error", f"File not found:\n{input_path}")
            return
        try:
            pdf_tools.mime_type_for(input_path)
            pages = parse_page_range(self.page_entry.get(),
                                     pdf_tools.page_count(input_path))
        except ValueError as e:
            messagebox.showerror("Page Range", str(e))
            return
        except Exception as e:
            log.exception("Could not read input")
            messagebox.showerror("Error", f"Could not read the file:\n{e}")
            return
        if not pages:
            messagebox.showwarning("Page Range", "No valid pages selected.")
            return

        self.job = ConversionJob(
            self._job_settings(), input_path, output_path, pages,
            on_log=self.log, on_status=self.set_status,
            on_progress=self.set_progress)
        self.log(f"Starting OCR for {len(pages)} page(s): "
                 f"{pages[0]}–{pages[-1]}")
        self.set_status("Starting...")
        self.set_progress(0)
        self._set_running(True)
        threading.Thread(target=self._run_job, args=(self.job,),
                         daemon=True).start()

    def _run_job(self, job: ConversionJob) -> None:
        try:
            result = job.run()
        except JobCancelled:
            self.set_status("Canceled. Finished parts are kept for resume.")
            self.log("Processing canceled.")
        except Exception as e:
            log.exception("Conversion failed")
            self.set_status("Failed.")
            self._ui(messagebox.showerror, "Error", str(e))
        else:
            self.set_status("Done.")
            notify_sound()
            how = "Converted without OCR" if result.native else "OCR completed!"
            self._ui(messagebox.showinfo, "Success",
                     f"{how}\nFinal file:\n{result.output_path}")
            try:
                reveal_in_file_manager(str(result.output_path))
            except Exception as e:
                self.log(f"Auto-open failed: {e}")
        finally:
            self._ui(self._job_finished)

    def _job_finished(self) -> None:
        self.job = None
        self._set_running(False)

    def cancel_ocr(self) -> None:
        if self.job is not None:
            self.job.cancel()
            self.set_status("Canceling... waiting for running batches.")
            self.log("User requested cancel.")
