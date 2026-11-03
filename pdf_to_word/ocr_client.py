"""The only module that talks to Google Document AI.

One client is created per job and shared by the worker threads (gRPC clients
are thread-safe). Every call has a timeout and retries transient errors such
as 504 Deadline Exceeded or 503 Unavailable with exponential back-off.
"""
from __future__ import annotations

import logging
import os
from typing import Optional, Sequence

from google.api_core import exceptions as gexc
from google.api_core.client_options import ClientOptions
from google.api_core.retry import Retry, if_exception_type

from .config import Settings
from .pdf_tools import PDF_MIME, extract_pages

log = logging.getLogger(__name__)

TRANSIENT = (gexc.DeadlineExceeded, gexc.ServiceUnavailable,
             gexc.InternalServerError, gexc.TooManyRequests,
             gexc.Aborted)


class DocumentAIClient:
    def __init__(self, settings: Settings, client=None):
        # Imported here so tests and the GUI start without the SDK loaded.
        from google.cloud import documentai  # type: ignore

        missing = [name for name, value in (
            ("GCP_PROJECT_ID", settings.gcp_project_id),
            ("GCP_PROCESSOR_ID", settings.gcp_processor_id)) if not value]
        if missing:
            raise ValueError(
                f"Missing {', '.join(missing)}. Copy .env.example to .env "
                "and fill in your own Google Cloud IDs.")
        self._dai = documentai
        self.settings = settings
        self._client = client or self._make_client()
        self._retry = Retry(
            predicate=if_exception_type(*TRANSIENT),
            initial=2.0, maximum=60.0, multiplier=2.0,
            timeout=settings.retry_deadline,
            on_error=lambda e: log.warning("Document AI retrying after: %s", e),
        )

    def _make_client(self):
        s = self.settings
        opts = ClientOptions(
            api_endpoint=f"{s.gcp_location}-documentai.googleapis.com")
        key = s.google_application_credentials
        if key:
            if not os.path.isfile(key):
                raise FileNotFoundError(
                    f"Service-account key not found: {key}\n"
                    "Fix GOOGLE_APPLICATION_CREDENTIALS in .env, or unset it "
                    "to use 'gcloud auth application-default login'.")
            from google.oauth2 import service_account
            creds = service_account.Credentials.from_service_account_file(key)
            return self._dai.DocumentProcessorServiceClient(
                credentials=creds, client_options=opts)
        return self._dai.DocumentProcessorServiceClient(client_options=opts)

    # ----------------------------------------------------------------
    def _payload(self, path: str, mime_type: str,
                 pages: Sequence[int]) -> bytes:
        """Only the batch's pages are uploaded, never the whole PDF."""
        if mime_type == PDF_MIME:
            return extract_pages(path, pages)
        with open(path, "rb") as fh:
            return fh.read()

    def _process(self, processor_id: str, content: bytes, mime_type: str,
                 process_options=None):
        dai = self._dai
        s = self.settings
        request = dai.ProcessRequest(
            name=self._client.processor_path(
                s.gcp_project_id, s.gcp_location, processor_id),
            raw_document=dai.RawDocument(content=content, mime_type=mime_type),
            process_options=process_options,
        )
        result = self._client.process_document(
            request=request, retry=self._retry, timeout=s.request_timeout)
        return result.document

    def ocr(self, path: str, mime_type: str, pages: Sequence[int]):
        """Arabic/English OCR of ``pages``; returns a documentai.Document
        whose pages are in the same order as ``pages``."""
        dai = self._dai
        content = self._payload(path, mime_type, pages)
        options = dai.ProcessOptions(ocr_config=dai.OcrConfig(
            enable_native_pdf_parsing=False,
            enable_image_quality_scores=self.settings.enable_quality_scores,
            enable_symbol=True,
            hints=dai.OcrConfig.Hints(language_hints=["ar", "en"]),
        ))
        return self._process(self.settings.gcp_processor_id, content,
                             mime_type, options)

    def tables(self, path: str, mime_type: str,
               pages: Sequence[int]) -> Optional[object]:
        """Form Parser run over the same pages, only to recover tables.

        A failure here is not fatal: the batch keeps its text, minus tables.
        """
        if not self.settings.use_form_parser:
            return None
        try:
            content = self._payload(path, mime_type, pages)
            return self._process(self.settings.gcp_form_processor_id,
                                 content, mime_type)
        except Exception as e:
            log.warning("Form Parser call failed, tables skipped: %s", e)
            return None
