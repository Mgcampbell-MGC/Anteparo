"""Prose / inline creditor lists (editais published as running text).

Observed grammars:
  B  NAME, CPF: 000.000.000-00, R$ 1.234,56;  NAME - 00.000.000/0001-00, CNPJ: ..., R$ ...;
     header: 'QUIROGRAFÁRIO - CLASSE III - 45 (QUARENTA E CINCO) CREDORES ... VALOR DE R$ 39.944.118,88 (...)'
     a second 'CLASSE III - 09 (NOVE) CREDORES EM DÓLARES ... USD$ ...' section holds foreign-currency claims
  C  NAME - 00.000.000/0001-00|00.000.000/0002-00: R$ 1.234,56;
     header: 'CLASSE III – QUIROGRAFÁRIOS (25 CREDORES | R$ 53.682.740,64):'
Headers print per-class counts and totals, so step D can reconcile — per (class, currency).
Sections such as 'CRÉDITOS NÃO SUJEITOS' / 'EXTRACONCURSAIS' / 'TOTAL GERAL' end the class
and any rows under them are kept but flagged, never counted as class claims.
Tokens may wrap across lines/pages, so document and value patterns tolerate whitespace
between atoms; offsets are preserved for page provenance.
"""
from __future__ import annotations

import re

from ..cnpj import classify
from ..money import parse_brl
from .classes import CLASS_HEAD_RE
from .models import ClaimRow, PrintedTotal


def _ws(shape: str) -> str:
    """Wrap-tolerant pattern from a shape string: 'd' → \\d, '.' '/' '-' literal; \\s* between atoms."""
    atoms = {"d": r"\d", ".": r"\.", "/": "/", "-": "-"}
    return r"\s*".join(atoms[c] for c in shape)


CNPJ_P = _ws("dd.ddd.ddd/dddd-dd")
CPF_P = _ws("ddd.ddd.ddd-dd")
DOC_RE = re.compile(rf"(?:{CNPJ_P})|(?:{CPF_P})")
VAL_P = r"(?:\d{1,3}(?:\s*\.\s*\d{3})*|\d+)\s*,\s*\d{2}"
CUR_SYM = r"(R\s*\$|US\s*\$|USD\s*\$?|U\s*S\s*D|EUR(?:O)?S?\s*\$?|€)"
CUR_VAL_RE = re.compile(CUR_SYM + r"\s*(" + VAL_P + r")(?![\d,])", re.I)
MARKER_RE = re.compile(r",\s*(?:CPF|CNPJ)\s*:", re.I)
DOC_LABEL_RE = re.compile(r"(?:CPF|CNPJ)\s*:\s*([\w./\-]+)", re.I)
COUNT_RE = re.compile(r"\b(\d{1,5})\b\s*(?:\([^)]*\)\s*)?CREDOR", re.I)
BOUNDARY_RE = re.compile(
    r"\b(CR[ÉE]DITOS?\s+N[ÃA]O\s+SUJEITOS?|N[ÃA]O\s+SUJEITOS?\s+(?:À|A)\s+RECUPERA|EXTRACONCURSA\w*|"
    r"RETARDAT[ÁA]RI\w*|TOTAL\s+GERAL|CR[ÉE]DITOS?\s+IL[ÍI]QUIDOS?|QUADRO\s+RESUMO|RESUMO\s+GERAL)\b",
    re.I,
)


def _clean(s: str) -> str:
    return re.sub(r"\s+", "", s)


def _cur_of(sym: str) -> str:
    s = _clean(sym).upper()
    if s.startswith("R$"):
        return "BRL"
    if s.startswith("US") or s.startswith("U$"):
        return "USD"
    if s.startswith("EUR") or s == "€":
        return "EUR"
    return "BRL"


def _header_currency(head: str) -> str:
    if re.search(r"D[ÓO]LAR|USD|US\s*\$", head, re.I):
        return "USD"
    if re.search(r"\bEUROS?\b|€", head, re.I):
        return "EUR"
    return "BRL"


def _combined(page_texts):
    """Join pages; return (text, [(start_offset, pageno)])."""
    parts, offsets, pos = [], [], 0
    for i, t in enumerate(page_texts, start=1):
        t = t.replace("\n", " ")
        offsets.append((pos, i))
        parts.append(t)
        pos += len(t) + 1
    return " ".join(parts), offsets


def _page_at(offsets, pos):
    pg = offsets[0][1]
    for start, p in offsets:
        if start <= pos:
            pg = p
        else:
            break
    return pg


def _headers(text):
    """Class headers and boundary sections → sorted list of segment starts."""
    out = []
    for m in CLASS_HEAD_RE.finditer(text):
        if text[m.start():m.start() + 6] != "CLASSE":   # skip 'Classe e Assunto' style prose
            continue
        klass = m.group(1).upper()
        window = text[m.end(): m.end() + 300]
        cut = len(window)
        for pat in (MARKER_RE, DOC_RE):
            mm = pat.search(window)
            if mm and mm.start() < cut:
                cut = mm.start()
        head = window[:cut]
        cnt = COUNT_RE.search(head)
        tot = CUR_VAL_RE.search(head)
        cur = _header_currency(head)
        end = m.end()
        if tot:
            cur = _cur_of(tot.group(1)) if _cur_of(tot.group(1)) != "BRL" else cur
            end = m.end() + tot.end()
            mm = re.match(r"\s*\)?\s*(?:\([^)]*\))?\s*[.:]?", text[end:end + 300])
            if mm:
                end += mm.end()
        else:
            mm = re.search(r"[:.]", head)
            end = m.end() + (mm.end() if mm else 0)
        # a boundary word just before this header (e.g. 'CREDORES RETARDATÁRIOS - CLASSE III') is a tag
        pre = text[max(0, m.start() - 90):m.start()]
        bm = BOUNDARY_RE.search(pre)
        tag = bm.group(1).upper() if bm else None
        out.append({"start": m.start(), "end": end, "klass": klass, "currency": cur, "tag": tag,
                    "count": int(cnt.group(1)) if cnt else None,
                    "total": parse_brl(_clean(tot.group(2))) if tot else None,
                    "text": text[m.start():end][:200]})
    for bm in BOUNDARY_RE.finditer(text):
        # ignore if a class header sits within 90 chars after it (already used as a tag)
        if any(0 <= h["start"] - bm.start() <= 90 for h in out):
            continue
        tail = text[bm.end():bm.end() + 200]
        tot = CUR_VAL_RE.search(tail)
        out.append({"start": bm.start(), "end": bm.end(), "klass": None, "currency": "BRL",
                    "tag": bm.group(1).upper(), "count": None,
                    "total": parse_brl(_clean(tot.group(2))) if (tot and "TOTAL" in bm.group(1).upper()) else None,
                    "text": text[bm.start():bm.end() + 60]})
    out.sort(key=lambda h: h["start"])
    ded = []
    for h in out:
        if ded and h["klass"] == ded[-1]["klass"] and h["klass"] is not None and h["start"] - ded[-1]["end"] < 40:
            continue
        ded.append(h)
    return ded


def _entry(chunk: str, base_off: int, offsets, klass, heading, seg_cur, tag, idx):
    docs = list(DOC_RE.finditer(chunk))
    marker = MARKER_RE.search(chunk)
    label = DOC_LABEL_RE.search(chunk)
    first_doc_pos = docs[0].start() if docs else None
    cands = []
    if marker:
        cands.append(marker.start())
    if first_doc_pos is not None:
        pre = re.search(r"\s*[-–]\s*$", chunk[:first_doc_pos])
        cands.append(pre.start() if pre else first_doc_pos)
    if label and (first_doc_pos is None or label.start() < first_doc_pos):
        pre = re.search(r",?\s*$", chunk[:label.start()])
        cands.append(pre.start() if pre else label.start())
    search_from = docs[-1].end() if docs else (label.end() if label else 0)
    vm = CUR_VAL_RE.search(chunk, search_from) or CUR_VAL_RE.search(chunk)
    if not vm:
        return None
    if not cands:
        cands.append(vm.start())
    name_end = min(cands)
    name = re.sub(r"[\s,;:.\-–]+$", "", chunk[:name_end]).strip()
    name = re.sub(r"^[\s,;:.\-–)]+", "", name)
    if not name or len(name) < 2:
        return None
    value_str = _clean(vm.group(2))
    val = parse_brl(value_str)
    cur = _cur_of(vm.group(1))
    if cur == "BRL" and seg_cur != "BRL":
        cur = seg_cur
    raw_doc = _clean(docs[0].group()) if docs else (label.group(1) if label else "")
    dtype, dnum, dflags = classify(raw_doc) if raw_doc else ("NONE", "", ["NO_DOCUMENT"])
    flags = list(dflags)
    if val is None:
        flags.append("VALUE_UNPARSEABLE")
    if klass is None:
        flags.append("CLASS_UNKNOWN")
    if tag:
        flags.append("SECTION_" + re.sub(r"\W+", "_", tag))
    if cur != "BRL":
        flags.append("FOREIGN_CURRENCY")
    all_docs = [classify(_clean(d.group()))[1] for d in docs]
    if len({d[:8] for d in all_docs if len(d) == 14}) > 1:
        flags.append("MULTI_ROOT")
    if len(chunk) > 400:
        flags.append("LONG_CHUNK")
    page = _page_at(offsets, base_off + max(0, name_end))
    return ClaimRow(
        page=page, row_index=idx, seq_as_printed=None,
        creditor_name_as_printed=name, document_as_printed=raw_doc,
        document_number=dnum, document_type=dtype, all_documents=all_docs,
        klass=klass, class_set_by="INLINE_HEADER" if klass else "NONE",
        value_as_printed=value_str, value_brl=val, currency=cur, section_heading=heading,
        flags=flags, strategy="PROSE",
    )


def parse_prose(page_texts):
    text, offsets = _combined(page_texts)
    heads = _headers(text)
    rows, totals = [], []
    if not heads:
        segments = [(0, len(text), None, None, "BRL", None)]
    else:
        segments = []
        for i, h in enumerate(heads):
            nxt = heads[i + 1]["start"] if i + 1 < len(heads) else len(text)
            segments.append((h["end"], nxt, h["klass"], h["text"], h["currency"], h["tag"]))
            if h["klass"] or h["total"] is not None:
                totals.append(PrintedTotal(klass=h["klass"], total=h["total"], count=h["count"],
                                           page=_page_at(offsets, h["start"]), text=h["text"],
                                           currency=h["currency"]))
    per_page_idx = {}
    for (s, e, klass, heading, cur, tag) in segments:
        seg = text[s:e]
        pos = 0
        for chunk in re.split(r";", seg):
            base = s + pos
            pos += len(chunk) + 1
            c = chunk.strip()
            if not c or not CUR_VAL_RE.search(c):
                continue
            if not (DOC_RE.search(c) or DOC_LABEL_RE.search(c)) and len(c) > 250:
                continue
            row = _entry(c, base, offsets, klass, heading, cur, tag, 0)
            if row:
                per_page_idx[row.page] = per_page_idx.get(row.page, 0) + 1
                row.row_index = per_page_idx[row.page]
                rows.append(row)
    return rows, totals
