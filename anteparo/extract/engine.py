"""Pick a strategy per document, run it, attach printed totals."""
from __future__ import annotations

import re

from .classes import is_summary_row
from .models import DocResult
from .pdfio import open_pdf, page_text, has_text_layer
from .prose import parse_prose
from .table import parse_table

PROSE_SIGNAL = re.compile(r"(?:CPF|CNPJ)\s*:\s*\d|\s[-–]\s*\d{2}\.\d{3}\.\d{3}/\d{4}", re.I)


def extract_document(path: str) -> DocResult:
    with open_pdf(path) as pdf:
        n = len(pdf.pages)
        texts = [page_text(p) for p in pdf.pages]
        text_layer = has_text_layer(pdf)
        if not text_layer:
            return DocResult(path, n, False, "NONE", [], [], notes=["NO_TEXT_LAYER: needs OCR"])
        prose_hits = sum(len(PROSE_SIGNAL.findall(t)) for t in texts)
        notes = [f"prose_signal={prose_hits}"]
        layout = ""
        if prose_hits >= 5:
            rows, totals = parse_prose(texts)
            strategy, layout = "PROSE", "PROSE"
            if len([r for r in rows if r.value_brl is not None]) < 3:
                r2, t2, lay = parse_table(pdf)
                if len([r for r in r2 if r.value_brl is not None]) > len(rows):
                    rows, totals, strategy, layout = r2, t2, "TABLE", lay
                    notes.append("fell back to TABLE")
        else:
            rows, totals, layout = parse_table(pdf)
            strategy = "TABLE"
            if len([r for r in rows if r.value_brl is not None]) < 3:
                r2, t2 = parse_prose(texts)
                if len([r for r in r2 if r.value_brl is not None]) > len(rows):
                    rows, totals, strategy, layout = r2, t2, "PROSE", "PROSE"
                    notes.append("fell back to PROSE")
        for r in rows:
            if is_summary_row(r.creditor_name_as_printed, bool(r.document_number)) and "NOT_A_CLAIM" not in r.flags:
                r.flags += ["NOT_A_CLAIM", "SUMMARY_ROW"]
        return DocResult(path, n, True, strategy, rows, totals, layout_id=layout, notes=notes)
