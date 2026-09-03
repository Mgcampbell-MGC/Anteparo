"""Column-aware table extraction (spec step C).

Two paths, chosen per page:
  RULED  — pdfplumber finds the table from its ruling lines/rects. Cells give exact row
           segmentation (multi-line names stay in their cell). Column roles are learned
           from the header row or from cell content. The value cell's whitespace is
           stripped before parsing, which neutralises the fragment bug ('R$ 3 1.981.695,98').
  BANDS  — no rulings: words are clustered into bands, rows anchored on a value/document,
           and name-only bands attached to the nearest anchor.
Class comes from a CLASSE column when the layout has one, otherwise from the section
heading in force (carried across pages). Values come only from the row's own value cell.
"""
from __future__ import annotations

import re
from statistics import median

from ..cnpj import classify
from ..money import merge_fragments, parse_brl, is_brl
from .classes import class_from_text, is_class_heading, is_noise, TOTAL_RE
from .models import ClaimRow, PrintedTotal

DOC_TOKEN_RE = re.compile(r"^\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}$|^\d{3}\.\d{3}\.\d{3}-\d{2}$|^\d{11}$|^\d{14}$")
DOC_IN_RE = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}|\d{3}\.?\d{3}\.?\d{3}-?\d{2}")
DOCISH_RE = re.compile(r"^[\d./-]{8,}$")
SEQ_RE = re.compile(r"^\d{1,5}$")
CUR_RE = re.compile(r"^R\$?$")
CLASS_CELL_RE = re.compile(r"^(IV|III|II|I)$", re.I)
CUR_PREFIX_RE = re.compile(r"^(R\$|US\$|USD\$?|U\$|EUR|€)", re.I)


# ----------------------------------------------------------------------------- shared
def _bands(words, tol=2.6):
    ws = sorted(words, key=lambda w: (w["top"], w["x0"]))
    bands = []
    for w in ws:
        if bands and abs(w["top"] - bands[-1]["top"]) <= tol:
            bands[-1]["words"].append(w)
        else:
            bands.append({"top": w["top"], "words": [w]})
    for b in bands:
        b["words"].sort(key=lambda w: w["x0"])
        b["text"] = " ".join(w["text"] for w in b["words"])
        b["bottom"] = max(w["bottom"] for w in b["words"])
    return bands


class TableState:
    def __init__(self):
        self.klass = None
        self.heading = None
        self.doc_x = None
        self.val_left = None
        self.seq_split = None
        self.dev_split = None
        self.roles = None          # learned column roles for the ruled path
        self.layout = None


def _cell_value(text: str):
    """Parse a value cell: strip whitespace and currency, validate BRL grammar."""
    if not text:
        return None, "BRL"
    t = re.sub(r"\s+", "", text)
    cur = "BRL"
    m = CUR_PREFIX_RE.match(t)
    if m:
        sym = m.group(1).upper()
        cur = "USD" if sym.startswith("U") else "EUR" if sym.startswith("E") or sym == "€" else "BRL"
        t = t[m.end():]
    t = t.strip("*")
    return (t if is_brl(t) else None), cur


# ----------------------------------------------------------------------------- ruled
def _learn_roles(grid):
    """Return (roles: dict col->role, header_row_index or None)."""
    ncol = max(len(r) for r in grid)
    roles, hdr = {}, None
    for i, row in enumerate(grid[:3]):
        u = " ".join((c or "") for c in row).upper()
        if ("CREDOR" in u or "NOME" in u or "RAZ" in u) and ("CNPJ" in u or "CPF" in u or "DOCUMENTO" in u):
            hdr = i
            for j, c in enumerate(row):
                t = (c or "").upper().replace("\n", " ")
                if not t:
                    continue
                if "DEVEDOR" in t or "RECUPERAND" in t:
                    roles[j] = "debtor"
                elif "CNPJ" in t or "CPF" in t or "DOCUMENTO" in t:
                    roles[j] = "doc"
                elif "VALOR" in t or "CRÉDITO" in t or "CREDITO" in t or "IMPORT" in t:
                    roles.setdefault(j, "value")
                elif t.startswith("CLASSE") or t == "CLASSE":
                    roles[j] = "class"
                elif "CREDOR" in t or "NOME" in t or "RAZ" in t:
                    roles[j] = "name"
                elif t.startswith("N") and len(t) <= 4 or "ITEM" in t or "SEQ" in t or "ORD" in t:
                    roles[j] = "seq"
                elif "MOEDA" in t:
                    roles[j] = "currency"
            break
    data = grid[hdr + 1:] if hdr is not None else grid
    n = max(1, len(data))
    for j in range(ncol):
        if j in roles:
            continue
        cells = [(r[j] or "").strip() for r in data if j < len(r)]
        nonempty = [c for c in cells if c]
        if not nonempty:
            continue
        docs = sum(1 for c in nonempty if DOC_IN_RE.search(re.sub(r"\s+", "", c)))
        vals = sum(1 for c in nonempty if _cell_value(c)[0] is not None)
        seqs = sum(1 for c in nonempty if SEQ_RE.match(c))
        cls = sum(1 for c in nonempty if CLASS_CELL_RE.match(c) or (class_from_text(c) and len(c) < 40))
        if vals / len(nonempty) > 0.6 and "value" not in roles.values():
            roles[j] = "value"
        elif docs / len(nonempty) > 0.6 and "doc" not in roles.values():
            roles[j] = "doc"
        elif seqs / len(nonempty) > 0.8 and "seq" not in roles.values():
            roles[j] = "seq"
        elif cls / len(nonempty) > 0.8 and "class" not in roles.values():
            roles[j] = "class"
    # remaining text columns: the widest is the name; a low-cardinality one is the debtor
    text_cols = [j for j in range(ncol) if j not in roles]
    if text_cols:
        stats = []
        for j in text_cols:
            cells = [(r[j] or "").strip() for r in data if j < len(r) and (r[j] or "").strip()]
            stats.append((j, sum(len(c) for c in cells) / max(1, len(cells)), len(set(cells)), len(cells)))
        name_col = max(stats, key=lambda s: s[1])[0]
        roles[name_col] = "name"
        for j, avg, distinct, cnt in stats:
            if j != name_col and cnt >= 3 and distinct / cnt <= 0.34 and "debtor" not in roles.values():
                roles[j] = "debtor"
    return roles, hdr


def _rows_from_table(table, pageno, st: TableState, start_idx):
    grid = table.extract()
    if not grid:
        return [], []
    roles, hdr = _learn_roles(grid)
    if "value" not in roles.values():
        return None, None       # not a creditor table — let the caller fall back
    st.roles = roles
    inv = {v: k for k, v in roles.items()}
    rows, totals = [], []
    idx = start_idx
    for ri, cells in enumerate(grid):
        if hdr is not None and ri == hdr:
            continue
        get = lambda role: (cells[inv[role]] if role in inv and inv[role] < len(cells) else None) or ""
        joined = " ".join((c or "").replace("\n", " ") for c in cells)
        if not joined.strip() or is_noise(joined):
            continue
        val_str, cur = _cell_value(get("value"))
        doc_txt = re.sub(r"\s+", "", get("doc"))
        doc_m = DOC_IN_RE.search(doc_txt) or DOC_IN_RE.search(re.sub(r"\s+", "", joined))
        raw_doc = doc_m.group() if doc_m else ""
        name = re.sub(r"\s+", " ", get("name").replace("\n", " ")).strip()
        if TOTAL_RE.search(joined) and not raw_doc and val_str:
            totals.append(PrintedTotal(klass=class_from_text(joined) or st.klass, total=parse_brl(val_str),
                                       count=None, page=pageno, text=joined[:120], currency=cur))
            continue
        if not val_str and not raw_doc:
            if is_class_heading(joined):
                st.klass, st.heading = class_from_text(joined), joined[:120]
            continue
        klass, set_by = st.klass, "SECTION_HEADING"
        if "class" in inv:
            ct = get("class").strip()
            k = class_from_text(ct) if ct else None
            if k:
                klass, set_by = k, "COLUMN"
        if "currency" in inv and get("currency").strip():
            cu = get("currency").strip().upper()
            cur = "USD" if "US" in cu or "DÓLAR" in cu or "DOLAR" in cu else "EUR" if "EUR" in cu else cur
        seq = get("seq").strip() or None
        if seq and not SEQ_RE.match(seq):
            seq = None
        dtype, dnum, dflags = classify(raw_doc) if raw_doc else ("NONE", "", ["NO_DOCUMENT"])
        flags = list(dflags)
        val = parse_brl(val_str) if val_str else None
        if val is None:
            flags.append("VALUE_MISSING" if not get("value").strip() else "VALUE_UNPARSEABLE")
        if klass is None:
            flags.append("CLASS_UNKNOWN")
        if not name:
            flags.append("NAME_MISSING")
        if klass is None and val is None:
            flags.append("NOT_A_CLAIM")
        if cur != "BRL":
            flags.append("FOREIGN_CURRENCY")
        idx += 1
        rows.append(ClaimRow(
            page=pageno, row_index=idx, seq_as_printed=seq,
            creditor_name_as_printed=name, document_as_printed=raw_doc,
            document_number=dnum, document_type=dtype, all_documents=[dnum] if dnum else [],
            klass=klass, class_set_by=set_by if klass else "NONE",
            value_as_printed=val_str, value_brl=val, currency=cur,
            debtor_as_printed=(re.sub(r"\s+", " ", get("debtor")).strip() or None),
            section_heading=st.heading, flags=flags, strategy="TABLE_RULED",
        ))
    return rows, totals


# ----------------------------------------------------------------------------- bands
def _learn_geometry(bands, st: TableState):
    doc_xs, cur_xs, brl_x0 = [], [], []
    for b in bands:
        for w in b["words"]:
            if DOC_TOKEN_RE.match(w["text"]):
                doc_xs.append(w["x0"])
            elif CUR_RE.match(w["text"]):
                cur_xs.append(w["x0"])
            elif is_brl(w["text"]):
                brl_x0.append(w["x0"])
    if doc_xs:
        st.doc_x = median(doc_xs)
    if cur_xs:
        st.val_left = median(cur_xs) - 3
    elif brl_x0:
        st.val_left = min(brl_x0) - 30


def _learn_header_band(band, st: TableState):
    u = band["text"].upper()
    if not (("CREDOR" in u or "NOME" in u or "RAZ" in u) and ("CNPJ" in u or "CPF" in u or "DOCUMENTO" in u)):
        return False
    xs = {}
    for w in band["words"]:
        t = w["text"].upper()
        if t.startswith("N") and ("º" in t or "°" in t or t in ("N.", "NO", "Nº", "N°", "ITEM", "SEQ")):
            xs.setdefault("seq", w["x0"])
        elif "DEVEDOR" in t or "RECUPERAND" in t:
            xs.setdefault("dev", w["x0"])
        elif t.startswith("CREDOR") or t in ("NOME", "RAZÃO", "RAZAO"):
            xs.setdefault("credor", w["x0"])
    if "seq" in xs and "dev" in xs:
        st.seq_split = (xs["seq"] + xs["dev"]) / 2
    if "dev" in xs and "credor" in xs:
        st.dev_split = (xs["dev"] + xs["credor"]) / 2
    return True


def _rows_from_bands(bands, pageno, st: TableState, start_idx):
    _learn_geometry(bands, st)
    heights = [b["bottom"] - b["top"] for b in bands] or [10]
    line_h = median(heights)
    max_orphan_dist = line_h * 4.5 + 2
    anchors, orphans, totals = [], [], []
    for b in bands:
        txt = b["text"]
        if _learn_header_band(b, st) or is_noise(txt):
            continue
        val_toks = [w for w in b["words"] if st.val_left is not None and w["x0"] >= st.val_left]
        value_str = merge_fragments(val_toks) if val_toks else None
        doc_toks = [w for w in b["words"] if DOC_TOKEN_RE.match(w["text"]) or
                    (st.doc_x is not None and DOCISH_RE.match(w["text"]) and abs(w["x0"] - st.doc_x) < 30)]
        docs = [w["text"] for w in doc_toks]
        if TOTAL_RE.search(txt) and value_str and not docs:
            totals.append(PrintedTotal(klass=class_from_text(txt) or st.klass, total=parse_brl(value_str),
                                       count=None, page=pageno, text=txt))
            continue
        if not value_str and not docs:
            if is_class_heading(txt):
                st.klass, st.heading = class_from_text(txt), txt
            else:
                orphans.append(b)
            continue
        anchors.append({"band": b, "value_str": value_str, "docs": docs, "doc_toks": doc_toks,
                        "val_toks": val_toks, "klass": st.klass, "heading": st.heading})
    for o in orphans:
        if not anchors:
            break
        best = min(anchors, key=lambda a: abs(a["band"]["top"] - o["top"]))
        if abs(best["band"]["top"] - o["top"]) <= max_orphan_dist:
            best.setdefault("extra", []).append(o)
    rows = []
    idx = start_idx
    for a in anchors:
        b = a["band"]
        seq, name_toks, dev_toks = None, [], []
        doc_ids = {id(w) for w in a["doc_toks"]}
        val_ids = {id(w) for w in a["val_toks"]}
        name_right = (st.doc_x - 10) if st.doc_x is not None else (st.val_left - 3 if st.val_left else 9e9)
        for bb in sorted([b] + a.get("extra", []), key=lambda x: x["top"]):
            for w in bb["words"]:
                if id(w) in doc_ids or id(w) in val_ids:
                    continue
                t = w["text"]
                if bb is b and seq is None and SEQ_RE.match(t) and w["x0"] < 80 and (st.seq_split is None or w["x0"] < st.seq_split):
                    seq = t
                    continue
                if st.dev_split is not None and w["x0"] < st.dev_split and (st.seq_split is None or w["x0"] >= st.seq_split):
                    dev_toks.append(t)
                    continue
                if w["x0"] < name_right and not CUR_RE.match(t):
                    name_toks.append(t)
        name = " ".join(name_toks).strip()
        raw_doc = a["docs"][0] if a["docs"] else ""
        dtype, dnum, dflags = classify(raw_doc) if raw_doc else ("NONE", "", ["NO_DOCUMENT"])
        flags = list(dflags)
        val = parse_brl(a["value_str"]) if a["value_str"] else None
        if val is None:
            flags.append("VALUE_MISSING" if not a["val_toks"] else "VALUE_UNPARSEABLE")
        if a["klass"] is None:
            flags.append("CLASS_UNKNOWN")
        if not name:
            flags.append("NAME_MISSING")
        if a["klass"] is None and val is None:
            flags.append("NOT_A_CLAIM")
        idx += 1
        rows.append(ClaimRow(
            page=pageno, row_index=idx, seq_as_printed=seq,
            creditor_name_as_printed=name, document_as_printed=raw_doc,
            document_number=dnum, document_type=dtype,
            all_documents=[classify(d)[1] for d in a["docs"]],
            klass=a["klass"], class_set_by="SECTION_HEADING" if a["klass"] else "NONE",
            value_as_printed=a["value_str"], value_brl=val,
            debtor_as_printed=" ".join(dev_toks) or None, section_heading=a["heading"],
            flags=flags, strategy="TABLE_BANDS",
        ))
    return rows, totals


# ----------------------------------------------------------------------------- page
def parse_table_page(page, pageno: int, st: TableState):
    from .pdfio import page_words

    words = page_words(page)
    if not words:
        return [], []
    tables = []
    try:
        tables = page.find_tables()
    except Exception:
        tables = []
    tables = [t for t in tables if len(t.rows) >= 1 and t.rows and len(t.rows[0].cells) >= 3]
    if not tables:
        st.layout = st.layout or "TABLE_BANDS"
        return _rows_from_bands(_bands(words), pageno, st, 0)

    # words outside every table bbox → headings / totals / noise, processed in y-order with the tables
    def inside(w, bbox):
        x0, top, x1, bottom = bbox
        cx, cy = (w["x0"] + w["x1"]) / 2, (w["top"] + w["bottom"]) / 2
        return x0 - 1 <= cx <= x1 + 1 and top - 1 <= cy <= bottom + 1

    outside = [w for w in words if not any(inside(w, t.bbox) for t in tables)]
    items = [("band", b["top"], b) for b in _bands(outside)] + [("table", t.bbox[1], t) for t in tables]
    items.sort(key=lambda it: it[1])
    rows, totals = [], []
    idx = 0
    for kind, _, obj in items:
        if kind == "band":
            txt = obj["text"]
            if is_noise(txt):
                continue
            if is_class_heading(txt):
                st.klass, st.heading = class_from_text(txt), txt
            elif TOTAL_RE.search(txt):
                v = merge_fragments(obj["words"])
                if v:
                    totals.append(PrintedTotal(klass=class_from_text(txt) or st.klass, total=parse_brl(v),
                                               count=None, page=pageno, text=txt))
        else:
            r, t = _rows_from_table(obj, pageno, st, idx)
            if r is None:                      # table without a value column → try bands for the page
                st.layout = st.layout or "TABLE_BANDS"
                return _rows_from_bands(_bands(words), pageno, st, 0)
            st.layout = "TABLE_RULED"
            idx += len(r)
            rows.extend(r)
            totals.extend(t)
    return rows, totals


def parse_table(pdf):
    st = TableState()
    rows, totals = [], []
    for i, page in enumerate(pdf.pages, start=1):
        r, t = parse_table_page(page, i, st)
        rows.extend(r)
        totals.extend(t)
    return rows, totals, st.layout or "TABLE_BANDS"
