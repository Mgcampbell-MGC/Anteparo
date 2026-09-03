"""Thin pdfplumber wrapper: words with geometry per page, page text, text-layer check."""
from __future__ import annotations

import warnings
from contextlib import contextmanager

warnings.filterwarnings("ignore")
import pdfplumber  # noqa: E402


@contextmanager
def open_pdf(path: str):
    with pdfplumber.open(path) as pdf:
        yield pdf


def page_words(page):
    """Words with geometry; rotated (non-upright) text such as gazette margin stamps is dropped."""
    ws = page.extract_words(keep_blank_chars=False, use_text_flow=False, extra_attrs=["upright"])
    return [w for w in ws if w.get("upright", True)]


def page_text(page) -> str:
    return page.extract_text() or ""


def has_text_layer(pdf, sample_pages: int = 3) -> bool:
    n = 0
    for p in pdf.pages[:sample_pages]:
        n += len(page_text(p))
    return n > 80
