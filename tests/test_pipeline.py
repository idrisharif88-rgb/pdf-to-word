import dataclasses

import pytest
from docx import Document
from docx.oxml.ns import qn
from pypdf import PdfReader

from conftest import BODY, make_ocr_document, standard_page
from pdf_to_word.builder import build_batch_document
from pdf_to_word.config import Settings
from pdf_to_word.ocr_client import DocumentAIClient
from pdf_to_word.pdf_tools import extract_pages
from pdf_to_word.pipeline import ConversionJob, JobCancelled

BASE = Settings(native_convert=False, extract_images=False, add_toc=False,
                batch_size=5, max_workers=2)


class FakeClient:
    """Stands in for DocumentAIClient; records the batches requested."""

    def __init__(self, settings, fail_on=None, on_call=None):
        self.settings = settings
        self.calls = []
        self.table_calls = 0
        self.fail_on = fail_on
        self.on_call = on_call

    def ocr(self, path, mime, pages):
        self.calls.append(list(pages))
        if self.on_call:
            self.on_call(pages)
        if self.fail_on and self.fail_on in pages:
            raise RuntimeError("boom")
        return make_ocr_document([standard_page(n) for n in pages])

    def tables(self, path, mime, pages):
        if self.settings.use_form_parser:
            self.table_calls += 1
        return None


def body_texts(path):
    return [p.text for p in Document(str(path)).paragraphs if p.text.strip()]


def run_job(settings, pdf, out, pages, client):
    job = ConversionJob(settings, str(pdf), str(out), pages,
                        client_factory=lambda s: client)
    return job, job.run()


def test_builder_drops_running_heads_and_finds_headings():
    doc = build_batch_document(BASE, "x.pdf", [1, 2, 3],
                               make_ocr_document([standard_page(n) for n in (1, 2, 3)]))
    texts = [(p.style.name, p.text) for p in doc.paragraphs if p.text.strip()]
    assert not any("اسم الكتاب" in t for _, t in texts)
    assert ("Heading 1", "الفصل 1") in texts
    body = [p for p in doc.paragraphs if p.text.startswith(BODY)]
    assert len(body) == 6
    bidi = body[0]._element.pPr.find(qn("w:bidi"))
    assert bidi.get(qn("w:val")) == "1"


def test_full_job_merges_all_parts_in_order(sample_pdf, tmp_path):
    client = FakeClient(BASE)
    pages = list(range(1, 13))
    _, result = run_job(BASE, sample_pdf, tmp_path / "book.docx", pages, client)

    assert sorted(map(tuple, client.calls)) == [(1, 2, 3, 4, 5), (6, 7, 8, 9, 10), (11, 12)]
    assert result.output_path.name == "book_merged.docx"
    headings = [t for t in body_texts(result.output_path) if t.startswith("الفصل")]
    assert headings == [f"الفصل {n}" for n in pages]  # nothing lost or reordered
    assert not list(tmp_path.glob("book_part_*"))      # temp parts removed


def test_resume_skips_finished_parts(sample_pdf, tmp_path):
    pages = list(range(1, 13))
    failing = FakeClient(BASE, fail_on=11)
    with pytest.raises(RuntimeError):
        run_job(BASE, sample_pdf, tmp_path / "book.docx", pages, failing)
    assert sorted(p.name for p in tmp_path.glob("book_part_*.docx")) == \
        ["book_part_1.docx", "book_part_2.docx"]

    retry = FakeClient(BASE)
    _, result = run_job(BASE, sample_pdf, tmp_path / "book.docx", pages, retry)
    assert retry.calls == [[11, 12]]  # only the missing batch is paid for
    headings = [t for t in body_texts(result.output_path) if t.startswith("الفصل")]
    assert headings == [f"الفصل {n}" for n in pages]


def test_cancel_stops_queued_batches(sample_pdf, tmp_path):
    settings = dataclasses.replace(BASE, max_workers=1, batch_size=2)
    holder = {}
    client = FakeClient(settings, on_call=lambda pages: holder["job"].cancel())
    job = ConversionJob(settings, str(sample_pdf), str(tmp_path / "b.docx"),
                        list(range(1, 13)), client_factory=lambda s: client)
    holder["job"] = job
    with pytest.raises(JobCancelled):
        job.run()
    assert len(client.calls) == 1  # 5 queued batches never sent


@pytest.mark.parametrize("enabled,form_id,expected", [
    (True, "form123", 3), (False, "form123", 0), (True, "", 0)])
def test_form_parser_toggle(sample_pdf, tmp_path, enabled, form_id, expected):
    settings = dataclasses.replace(BASE, extract_tables=enabled,
                                   gcp_form_processor_id=form_id)
    client = FakeClient(settings)
    run_job(settings, sample_pdf, tmp_path / "book.docx",
            list(range(1, 13)), client)
    assert client.table_calls == expected


def test_extract_pages_uploads_only_the_batch(sample_pdf):
    from io import BytesIO
    data = extract_pages(str(sample_pdf), [3, 4, 5])
    assert len(PdfReader(BytesIO(data)).pages) == 3
    assert len(data) < sample_pdf.stat().st_size


class RecordingGrpc:
    def __init__(self):
        self.requests = []

    def processor_path(self, project, location, processor):
        return f"projects/{project}/locations/{location}/processors/{processor}"

    def process_document(self, request, retry=None, timeout=None):
        self.requests.append((request, retry, timeout))

        class R:
            document = make_ocr_document([standard_page(1)])
        return R()


def test_client_requests_have_timeout_retry_and_arabic_hints(sample_pdf):
    settings = dataclasses.replace(BASE, gcp_form_processor_id="form1",
                                   extract_tables=True, request_timeout=123)
    grpc = RecordingGrpc()
    client = DocumentAIClient(settings, client=grpc)
    client.ocr(str(sample_pdf), "application/pdf", [1, 2])
    client.tables(str(sample_pdf), "application/pdf", [1, 2])

    (ocr_req, retry, timeout), (form_req, _, _) = grpc.requests
    assert ocr_req.name.endswith(settings.gcp_processor_id)
    assert list(ocr_req.process_options.ocr_config.hints.language_hints) == ["ar", "en"]
    assert form_req.name.endswith("/form1")
    assert timeout == 123 and retry is not None


def test_client_skips_form_parser_when_disabled(sample_pdf):
    settings = dataclasses.replace(BASE, gcp_form_processor_id="form1",
                                   extract_tables=False)
    grpc = RecordingGrpc()
    client = DocumentAIClient(settings, client=grpc)
    assert client.tables(str(sample_pdf), "application/pdf", [1]) is None
    assert grpc.requests == []
