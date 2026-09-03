"""BRL money parsing.

Court documents print values as 31.981.695,98. pdfplumber sometimes splits a value
into fragments ('3' + '1.981.695,98', '9' + '02,44') — the "space bug". merge_fragments
re-joins fragments that are horizontally contiguous in the value column, then the
result is validated against the BRL grammar. A value that cannot be read cleanly is
returned as None and must be flagged VALUE_MISSING / VALUE_UNPARSEABLE — never guessed.
"""
from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation

# 1.234.567,89  |  1234,56  |  0,00   (thousands groups optional but if present must be 3-digit)
BRL_RE = re.compile(r"^(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}$")
# a fragment that could be part of a value: digits, dots, comma
FRAG_RE = re.compile(r"^[\d.,]+$")
CURRENCY_TOKENS = {"R$", "RS", "R$.", "R", "$"}


def is_brl(s: str | None) -> bool:
    return bool(s) and bool(BRL_RE.match(s.strip()))


def parse_brl(s: str | None) -> Decimal | None:
    if not s:
        return None
    t = s.strip().replace("R$", "").replace(" ", "")
    if t.startswith("-") or t.startswith("("):
        return None  # negatives don't occur in creditor lists; treat as unreadable
    if not BRL_RE.match(t):
        return None
    try:
        return Decimal(t.replace(".", "").replace(",", "."))
    except InvalidOperation:
        return None


def merge_fragments(tokens: list[dict], max_gap: float = 2.5, right_edge: float | None = None) -> str | None:
    """Join word tokens (pdfplumber dicts with text/x0/x1) that sit contiguously.

    tokens: candidate value-column words, left→right. Currency tokens are dropped.
    Returns the joined string if it parses as BRL, else None.
    """
    toks = []
    for t in tokens:
        txt = re.sub(r"^R\$", "", t["text"]).strip()
        if not txt or txt in CURRENCY_TOKENS or not FRAG_RE.match(txt):
            continue
        toks.append({**t, "text": txt})
    if not toks:
        return None
    toks = sorted(toks, key=lambda t: t["x0"])
    runs = [[toks[0]]]
    for prev, cur in zip(toks, toks[1:]):
        if cur["x0"] - prev["x1"] <= max_gap:
            runs[-1].append(cur)
        else:
            runs.append([cur])
    cands = []
    for run in runs:
        joined = "".join(t["text"] for t in run)
        if is_brl(joined):
            cands.append((joined, run[-1]["x1"]))
    if not cands:
        return None
    if right_edge is not None:
        return min(cands, key=lambda c: abs(c[1] - right_edge))[0]
    return cands[-1][0]   # rightmost run


def fmt_brl(d: Decimal | None) -> str:
    if d is None:
        return ""
    s = f"{d:,.2f}"  # 31,981,695.98
    return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")
