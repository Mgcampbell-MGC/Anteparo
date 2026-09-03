"""Scanned creditor lists: render → tesseract (por) → words with geometry → the same band parser.

Opt-in (ANTEPARO_OCR=1 or explicit call). OCR output is treated exactly like a text layer: values only
from the row's own value column, reconciled against printed totals where they exist. OCR digit damage is
caught by the CNPJ check digits and by the reconcile gate; rows are flagged OCR so the sheet says so.
"""
from __future__ import annotations

import csv
import io
import os
import shutil
import subprocess
import tempfile

from .table import TableState, _bands, _rows_from_bands, _learn_geometry

DPI = 300
SCALE = 72.0 / DPI


def ocr_available() -> bool:
    return bool(shutil.which("tesseract") and shutil.which("pdftoppm"))


def ocr_page_words(pdf_path: str, pageno: int, workdir: str) -> list[dict]:
    """Words for one page (1-based) in PDF point coordinates, pdfplumber-style dicts."""
    prefix = os.path.join(workdir, f"p{pageno}")
    subprocess.run(["pdftoppm", "-r", str(DPI), "-f", str(pageno), "-l", str(pageno), "-png", pdf_path, prefix],
                   check=True, capture_output=True, timeout=120)
    imgs = sorted(f for f in os.listdir(workdir) if f.startswith(f"p{pageno}") and f.endswith(".png"))
    if not imgs:
        return []
    img = os.path.join(workdir, imgs[-1])
    r = subprocess.run(["tesseract", img, "stdout", "-l", "por", "--psm", "6", "tsv"], capture_output=True, timeout=180)
    words = []
    for row in csv.DictReader(io.StringIO(r.stdout.decode("utf-8", "replace")), delimiter="\t", quoting=csv.QUOTE_NONE):
        if row.get("level") != "5" or not (row.get("text") or "").strip():
            continue
        try:
            conf = float(row.get("conf", "0"))
        except ValueError:
            conf = 0.0
        if conf < 30:
            continue
        x, y, w, h = (int(row["left"]), int(row["top"]), int(row["width"]), int(row["height"]))
        words.append({"text": row["text"].strip(), "x0": x * SCALE, "x1": (x + w) * SCALE, "top": y * SCALE,
                      "bottom": (y + h) * SCALE, "upright": True, "conf": conf})
    os.remove(img)
    return words


def parse_scanned(pdf_path: str, pages: int, max_pages: int = 60):
    """(rows, totals, layout) for a scanned PDF using the band parser on OCR words."""
    st = TableState()
    rows, totals = [], []
    with tempfile.TemporaryDirectory() as wd:
        for pn in range(1, min(pages, max_pages) + 1):
            words = ocr_page_words(pdf_path, pn, wd)
            if not words:
                continue
            bands = _bands(words)
            r, t = _rows_from_bands(bands, pn, st, 0)
            for row in r:
                row.flags.append("OCR")
                row.strategy = "TABLE_BANDS_OCR"
            rows.extend(r)
            totals.extend(t)
    return rows, totals, "TABLE_BANDS_OCR"
