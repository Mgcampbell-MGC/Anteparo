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
_COMPANY_SUFFIX_RE = re.compile(r"\b(LTDA|S\.?A\.?|S/A|EIRELI|EPP|ME|CIA|COMERCIO|COM[ÉE]RCIO|INDUSTRIA|IND[ÚU]STRIA|SERVI[ÇC]OS|ENGENHARIA|TRANSPORTES|LOG[ÍI]STICA)\b", re.I)


def is_total_line(s: str) -> bool:
    """A printed total, not a creditor whose name contains 'Total'."""
    if not s or not TOTAL_RE.search(s):
        return False
    if _COMPANY_SUFFIX_RE.search(s):
        return False
    return bool(re.search(r"^\s*(?:VALOR\s+)?(?:SUB)?TOTAL\b|\bTOTAL\s+(?:GERAL|CONCURSAL|D[OAE]S?\b|CLASSE|DA\s+CLASSE)|\bSOMA\b|\bMONTANTE\b|VALOR\s+TOTAL|TOTAL\s*:", s, re.I))
# Sections that are not one of the four classes: rows under them are kept but never counted as class claims.
BOUNDARY_RE = re.compile(
    r"\b(CR[ÉE]DITOS?\s+N[ÃA]O\s+SUJEITOS?|N[ÃA]O\s+SUJEITOS?\s+(?:À|A)\s+RECUPERA|N[ÃA]O\s+SUBMETID|EXTRACONCURSA\w*|"
    r"RETARDAT[ÁA]RI\w*|TOTAL\s+GERAL|CR[ÉE]DITOS?\s+IL[ÍI]QUIDOS?|QUADRO\s+RESUMO|RESUMO\s+GERAL|RESERVA\s+DE\s+CR[ÉE]DITO|"
    r"A[ÇC][ÕO]ES\s+TRABALHISTAS|GARANTIA\s+FIDUCI[ÁA]RIA|ALIENA[ÇC][ÃA]O\s+FIDUCI[ÁA]RIA|CESS[ÃA]O\s+FIDUCI[ÁA]RIA|"
    r"ADIANTAMENTO\s+(?:DE|A)\s+CONTRATO\s+DE\s+C[ÂA]MBIO|\bACC\b|ARRENDAMENTO\s+MERCANTIL|\bLEASING\b|"
    r"RESERVA\s+DE\s+DOM[ÍI]NIO|CR[ÉE]DITOS?\s+FISCA(?:L|IS)|FAZENDA\s+P[ÚU]BLICA|CREDORES\s+N[ÃA]O\s+SUJEITOS)\b",
    re.I,
)


def boundary_tag(s: str):
    m = BOUNDARY_RE.search(s or "")
    return re.sub(r"\W+", "_", m.group(1).upper()) if m else None


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
    if class_from_text(s) is None:
        return False
    if CLASS_HEAD_RE.search(s):
        return True
    return bool(re.match(r"^\s*(?:[A-D]\s*[-–).]\s*)?(?:CREDORES?|TITULARES|RELA[ÇC][ÃA]O|LISTA|QUADRO|CR[ÉE]DITOS?)\b", s, re.I))
SUMMARY_NAME_RE = re.compile(r"CLASSIFICA|QUANTIDADE|\bTOTAL\b|\bSOMA\b|SUBTOTAL|RESUMO|VALOR DA CAUSA|VALOR DE CR[ÉE]DITOS|^\s*R\$", re.I)


def is_summary_row(name: str, has_document: bool) -> bool:
    """A row that names a class or a total instead of a creditor (summary tables, running lines)."""
    if has_document:
        return False
    n = (name or "").strip()
    if not n:
        return False
    if SUMMARY_NAME_RE.search(n):
        return True
    return class_from_text(n) is not None and len(n) < 45 and not re.search(r"\d", n)


def is_noise(s: str) -> bool:
    return bool(NOISE_RE.search(s or ""))
