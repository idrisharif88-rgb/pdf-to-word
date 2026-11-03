"""The conversion job, independent of any GUI.

    job = ConversionJob(settings, "book.pdf", "book.docx", pages)
    result = job.run()                 # blocking; call from a worker thread
    job.cancel()                       # from any thread

Callbacks (``on_log``, ``on_status``, ``on_progress``) are invoked from
worker threads; a GUI must marshal them onto its own thread.
"""
from __future__ import annotations

import gc
import logging
import multiprocessing
import threading
from concurrent.futures import FIRST_EXCEPTION, ThreadPoolExecutor, wait
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, List, Optional, Sequence

from . import pdf_tools
from .builder import build_batch_document
from .config import Settings
from .merge import existing_parts, merge_parts, part_path, save_atomically
from .text import batches

log = logging.getLogger(__name__)


class JobCancelled(Exception):
    pass


@dataclass
class JobResult:
    output_path: Path
    native: bool  # converted directly, without OCR


def _noop(*_args) -> None:
    pass


class ConversionJob:
    def __init__(self, settings: Settings, input_path: str, output_path: str,
                 pages: Sequence[int],
                 on_log: Callable[[str], None] = _noop,
                 on_status: Callable[[str], None] = _noop,
                 on_progress: Callable[[float], None] = _noop,
                 client_factory=None):
        self.s = settings
        self.input_path = input_path
        self.output_path = Path(output_path).resolve()
        self.pages = list(pages)
        self.on_log, self.on_status, self.on_progress = on_log, on_status, on_progress
        self._client_factory = client_factory
        self._cancel = threading.Event()

    def cancel(self) -> None:
        self._cancel.set()

    @property
    def cancelled(self) -> bool:
        return self._cancel.is_set()

    def _log(self, message: str) -> None:
        log.info(message)
        self.on_log(message)

    # ----------------------------------------------------------------
    def run(self) -> JobResult:
        if not self.pages:
            raise ValueError("No pages selected.")
        out_dir = self.output_path.parent
        out_dir.mkdir(parents=True, exist_ok=True)

        if self.s.native_convert and pdf_tools.has_text_layer(self.input_path, self.s):
            self._log("PDF has a text layer - converting directly, no OCR needed.")
            self.on_status("Converting native PDF...")
            pdf_tools.convert_native_pdf(self.input_path, str(self.output_path),
                                         self.pages)
            self.on_progress(100)
            return JobResult(self.output_path, native=True)

        return JobResult(self._run_ocr(out_dir), native=False)

    def _make_client(self):
        if self._client_factory:
            return self._client_factory(self.s)
        from .ocr_client import DocumentAIClient
        return DocumentAIClient(self.s)

    def _run_ocr(self, out_dir: Path) -> Path:
        base = self.output_path.stem
        mime = pdf_tools.mime_type_for(self.input_path)
        groups = batches(self.pages, self.s.batch_size)
        total = len(groups)

        done = existing_parts(out_dir, base)
        todo = [(n, g) for n, g in enumerate(groups, start=1) if n not in done]
        if done:
            self._log(f"Resuming: {len(done)} part(s) already on disk.")
        else:
            self._log("Starting new OCR job...")
        self._log("Form Parser (tables): " +
                  ("ON" if self.s.use_form_parser else "OFF"))

        workers = self.s.max_workers or max(2, min(6, multiprocessing.cpu_count() - 1))
        self._log(f"Using {workers} threads")
        completed = total - len(todo)
        self.on_progress(completed / total * 100)

        if todo:
            client = self._make_client()
            executor = ThreadPoolExecutor(max_workers=workers)
            try:
                futures = {executor.submit(self._process_part, client, mime,
                                           n, g, out_dir, base): n
                           for n, g in todo}
                pending = set(futures)
                while pending:
                    finished, pending = wait(pending, timeout=0.5,
                                             return_when=FIRST_EXCEPTION)
                    for future in finished:
                        future.result()  # re-raises a worker's error
                        completed += 1
                        self.on_progress(completed / total * 100)
                        self.on_status(f"Processed {completed}/{total} parts...")
                        self._log(f"Completed part {futures[future]} ({completed}/{total})")
                    if self.cancelled:
                        raise JobCancelled()
            finally:
                # Drop queued batches at once instead of paying for them.
                executor.shutdown(wait=True, cancel_futures=True)

        on_disk = existing_parts(out_dir, base)
        missing = [n for n in range(1, total + 1) if n not in on_disk]
        if missing:
            raise RuntimeError(f"Part file(s) missing after OCR: {missing}")
        parts = [on_disk[n] for n in range(1, total + 1)]
        merged = out_dir / f"{base}_merged.docx"
        self.on_status("Merging parts...")
        merge_parts(parts, merged, self.s, on_log=self._log)
        self._log(f"Merged file saved: {merged}")
        for part in parts:
            try:
                part.unlink()
            except OSError as e:
                self._log(f"Could not delete {part.name}: {e}")
        self.on_progress(100)
        return merged

    def _process_part(self, client, mime: str, part_num: int,
                      pages: List[int], out_dir: Path, base: str) -> None:
        if self.cancelled:
            return
        self.on_status(f"Processing part {part_num} (pages {pages[0]}–{pages[-1]})...")
        self._log(f"Starting part {part_num} for pages {pages}")
        ocr_doc = client.ocr(self.input_path, mime, pages)
        form_doc = client.tables(self.input_path, mime, pages)
        if self.cancelled:
            return
        doc = build_batch_document(self.s, self.input_path, pages,
                                   ocr_doc, form_doc)
        target = part_path(out_dir, base, part_num)
        save_atomically(doc, target)
        self._log(f"Saved {target.name}")
        del doc
        gc.collect()
