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
from .classes import CLASS_HEAD_RE, BOUNDARY_RE as CLASS_BOUNDARY_RE, NOISE_RE
PAGE_NOISE_RE = re.compile(r"^(Poder Judiciário|Tribunal de Justiça|Diário de Justiça|Certidão de publicação|Número do processo|Classe:|Tribunal:|Órgão:|Tipo de documento|Disponibilizado em|Inteiro teor|Destinatários|Advogado|Teor da Comunicação|Processo:|Comarca|Usuário:|ANO\s+[XVI]+\s*-\s*EDIÇÃO)", re.I)
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
BOUNDARY_RE = CLASS_BOUNDARY_RE
# a generic '(N CREDORES | R$ total):' section header that names no class, e.g. 'RESERVA DE CRÉDITO – AÇÕES TRABALHISTAS (20 CREDORES | R$ …):'
SECTION_RE = re.compile(r"\(\s*\d{1,4}\s*CREDOR(?:ES)?\s*\|\s*R\s*\$", re.I)
TOTAL_CONTEXT_RE = re.compile(r"(VALOR|TOTAL|SOMA|MONTANTE|IMPORT\w*|PERFAZ\w*|TOTALIZ\w*|\||\(|:|=)\s*(?:DE|EM|GERAL|DA\s+CLASSE|TOTAL|:)?\s*$", re.I)


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
        lines = t.split("\n")
        # blank out (keep length) header/footer lines so entries that straddle a page break stay intact
        lines = [(" " * len(ln)) if (NOISE_RE.search(ln) or (i > 1 and PAGE_NOISE_RE.match(ln.strip()))) else ln for ln in lines]
        t = " ".join(lines)
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
        for pat in (MARKER_RE, DOC_RE, CLASS_HEAD_RE, BOUNDARY_RE):
            mm = pat.search(window)
            if mm and mm.start() < cut:
                cut = mm.start()
        head = window[:cut]
        cnt = COUNT_RE.search(head)
        tot = None
        for cand in CUR_VAL_RE.finditer(head):
            if TOTAL_CONTEXT_RE.search(head[max(0, cand.start() - 40):cand.start()]):
                tot = cand
                break
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
    generic = []
    for gm in SECTION_RE.finditer(text):
        if any(h["start"] - 5 <= gm.start() <= h["end"] + 5 for h in out):
            continue   # this '(N CREDORES | R$ …)' belongs to a class header
        # walk back to the start of the section title (previous '.', ';' boundary)
        st0 = max(text.rfind(". ", 0, gm.start()), text.rfind("; ", 0, gm.start()), gm.start() - 90)
        title = text[st0 + 2:gm.start()].strip(" :–-")
        if len(title) < 8 or CLASS_HEAD_RE.search(title) or BOUNDARY_RE.search(title):
            continue   # handled by the class / boundary passes
        if CLASS_HEAD_RE.search(text[max(0, gm.start() - 160):gm.start()]):
            continue   # a class header sits just before: its own count/total
        generic.append((st0 + 2, gm, title))
    for bm in list(BOUNDARY_RE.finditer(text)) + [g[1] for g in generic]:
        # ignore if a class header sits within 90 chars after it (already used as a tag)
        if any(0 <= h["start"] - bm.start() <= 90 for h in out):
            continue
        tail = text[bm.end():bm.end() + 200]
        tot = CUR_VAL_RE.search(tail)
        is_generic = bm.re is SECTION_RE
        label = next((g[2] for g in generic if g[1] is bm), None) if is_generic else bm.group(1)
        end = bm.end()
        if is_generic:
            mm = re.match(r"[^)]*\)\s*:?", tail)
            end = bm.end() + (mm.end() if mm else 0)
        out.append({"start": bm.start(), "end": end, "klass": None, "currency": "BRL",
                    "tag": re.sub(r"\W+", "_", (label or "SECTION").upper())[:40], "count": None,
                    "total": parse_brl(_clean(tot.group(2))) if (tot and "TOTAL" in (label or "").upper()) else None,
                    "text": text[bm.start():end + 60]})
    out.sort(key=lambda h: h["start"])
    ded = []
    for h in out:
        if ded and h["klass"] == ded[-1]["klass"] and h["klass"] is not None and h["start"] - ded[-1]["end"] < 40:
            continue
        ded.append(h)
    return ded


def _entry(chunk: str, base_off: int, offsets, klass, heading, seg_cur, tag, idx, section=0):
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
        flags=flags, strategy="PROSE", section=section,
    )


def parse_prose(page_texts):
    text, offsets = _combined(page_texts)
    heads = _headers(text)
    rows, totals = [], []
    if not heads:
        segments = [(0, len(text), None, None, "BRL", None, 0)]
    else:
        segments = []
        for i, h in enumerate(heads, start=1):
            nxt = heads[i]["start"] if i < len(heads) else len(text)
            segments.append((h["end"], nxt, h["klass"], h["text"], h["currency"], h["tag"], i))
            if h["klass"] or h["total"] is not None:
                totals.append(PrintedTotal(klass=h["klass"], total=h["total"], count=h["count"],
                                           page=_page_at(offsets, h["start"]), text=h["text"],
                                           currency=h["currency"], section=i))
    per_page_idx = {}
    for (s, e, klass, heading, cur, tag, section) in segments:
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
            row = _entry(c, base, offsets, klass, heading, cur, tag, 0, section)
            if row:
                per_page_idx[row.page] = per_page_idx.get(row.page, 0) + 1
                row.row_index = per_page_idx[row.page]
                rows.append(row)
    return rows, totals
