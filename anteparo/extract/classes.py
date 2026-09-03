"""Creditor class detection (Lei 11.101 art. 41): I labour · II secured · III unsecured/trade · IV ME/EPP."""
from __future__ import annotations

import re

CLASS_HEAD_RE = re.compile(r"\bCLASSE\s+(IV|III|II|I)\b", re.I)
_WORDS = [
    ("III", ("QUIROGRAF",)),
    ("II", ("GARANTIA REAL", "GARANTIAS REAIS", "COM GARANTIA")),
    ("I", ("TRABALHIST", "DERIVADOS DA LEGISLA")),
    ("IV", ("ME/EPP", "ME E EPP", "MICROEMPRESA", "PEQUENO PORTE", "ME-EPP", "ME EPP")),
]
NOISE_RE = re.compile(
    r"(assinado eletronicamente|https?://|n[úu]mero do documento|\bp[áa]g\.?\s*\d|\bcontinua\b|"
    r"documento assinado|certifica[çc][ãa]o digital|^\s*p[áa]gina\s+\d)",
    re.I,
)
TOTAL_RE = re.compile(r"\b(TOTAL|SOMA|SUBTOTAL|MONTANTE)\b", re.I)


def class_from_text(s: str) -> str | None:
    m = CLASS_HEAD_RE.search(s)
    if m:
        return m.group(1).upper()
    u = s.upper()
    for k, words in _WORDS:
        if any(w in u for w in words):
            return k
    return None


def is_class_heading(s: str) -> bool:
    """A section heading: names a class, is short-ish, and is not a data row."""
    if not s or len(s) > 160:
        return False
    if not CLASS_HEAD_RE.search(s) and not re.search(r"\bCREDORES?\b", s, re.I):
        return False
    return class_from_text(s) is not None


def is_noise(s: str) -> bool:
    return bool(NOISE_RE.search(s or ""))
