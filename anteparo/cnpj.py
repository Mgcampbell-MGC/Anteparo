"""CNPJ / CPF handling — normalisation, mod-11 validation, root, single-digit repair.

Spec step E: 14 digits = CNPJ (validate check digits); 11 digits = CPF (individual,
never a target); anything else = NONE. A CNPJ failing its check digits is usually one
OCR-damaged digit — try every single-digit correction and accept only if exactly one
is valid (caller must additionally confirm the RFB name matches).
"""
from __future__ import annotations

import re

_CNPJ_RE = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}")
_CPF_RE = re.compile(r"\d{3}\.?\d{3}\.?\d{3}-?\d{2}")


def digits(s: str | None) -> str:
    return re.sub(r"\D", "", s or "")


def _mod11(nums: str, weights: list[int]) -> int:
    total = sum(int(n) * w for n, w in zip(nums, weights))
    r = total % 11
    return 0 if r < 2 else 11 - r


def is_valid_cnpj(d: str) -> bool:
    d = digits(d)
    if len(d) != 14 or len(set(d)) == 1:
        return False
    w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    w2 = [6] + w1
    return _mod11(d[:12], w1) == int(d[12]) and _mod11(d[:13], w2) == int(d[13])


def is_valid_cpf(d: str) -> bool:
    d = digits(d)
    if len(d) != 11 or len(set(d)) == 1:
        return False
    w1 = list(range(10, 1, -1))
    w2 = list(range(11, 1, -1))
    return _mod11(d[:9], w1) == int(d[9]) and _mod11(d[:10], w2) == int(d[10])


def root(d: str) -> str:
    """8-digit CNPJ root (cnpj_basico)."""
    return digits(d)[:8]


def format_cnpj(d: str) -> str:
    d = digits(d)
    if len(d) != 14:
        return d
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"


def repair_single_digit(d: str) -> list[str]:
    """All single-digit substitutions of a 14-digit string that yield a valid CNPJ."""
    d = digits(d)
    if len(d) != 14:
        return []
    out = []
    for i in range(14):
        for c in "0123456789":
            if c == d[i]:
                continue
            cand = d[:i] + c + d[i + 1 :]
            if is_valid_cnpj(cand):
                out.append(cand)
    return sorted(set(out))


def classify(raw: str | None) -> tuple[str, str, list[str]]:
    """Return (document_type, normalised_digits, flags).

    document_type ∈ {CNPJ, CPF, NONE}. Flags may include CNPJ_INVALID.
    Does NOT auto-apply single-digit repair — that needs an RFB name match (step E).
    """
    d = digits(raw)
    if len(d) == 14:
        return ("CNPJ", d, [] if is_valid_cnpj(d) else ["CNPJ_INVALID"])
    if len(d) == 11:
        return ("CPF", d, ["INDIVIDUAL"] if is_valid_cpf(d) else ["INDIVIDUAL", "CPF_INVALID"])
    return ("NONE", d, ["NO_DOCUMENT"] if not d else ["DOC_LEN_" + str(len(d))])


def find_documents(text: str) -> list[tuple[str, str]]:
    """Find all CNPJ/CPF-looking strings in free text → [(kind, digits)]."""
    out = []
    for m in _CNPJ_RE.finditer(text):
        out.append(("CNPJ", digits(m.group())))
    # avoid double-counting CPF patterns that are substrings of a CNPJ match
    spans = [m.span() for m in _CNPJ_RE.finditer(text)]
    for m in _CPF_RE.finditer(text):
        if any(a <= m.start() and m.end() <= b for a, b in spans):
            continue
        out.append(("CPF", digits(m.group())))
    return out
