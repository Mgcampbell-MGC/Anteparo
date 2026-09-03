"""Extract + reconcile every harvested PDF into documents/claims (steps C+D, append-only)."""
from __future__ import annotations

import json
import re
from collections import Counter

from ..db import now
from ..extract.engine import extract_document
from ..extract.pdfio import open_pdf, page_text
from ..steps.reconcile import reconcile

CNJ_RE = re.compile(r"\b\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}\b")
DATE_RE = re.compile(r"\b(\d{2})/(\d{2})/(20\d{2})\b")
DEBTOR_RE = re.compile(
    r"RECUPERA[ÇC][ÃA]O JUDICIAL\s+(?:D[EAO]S?\s+|REQUERIDA PELA(?: EMPRESA)?\s+|AJUIZAD[AO] P(?:ELA|OR)\s+)?"
    r"([A-ZÀ-Ú0-9&.,'\-\s]{4,90}?(?:LTDA\.?|S\.?A\.?|S/A|EIRELI|\bME\b|\bEPP\b|CIA\.?|E OUTR[AO]S?))", re.I)


def doc_type_from_text(t: str) -> str:
    u = t.upper()
    if "QUADRO GERAL" in u or "QUADRO-GERAL" in u:
        return "QGC"
    if re.search(r"ART\.?\s*7\s*[º°o]?\s*,?\s*(?:§|PAR[ÁA]GRAFO)\s*2", u) or "2ª LISTA" in u or "SEGUNDA LISTA" in u or "2A LISTA" in u:
        return "AJ_LIST"
    if "PLANO DE RECUPERA" in u[:600] and "RELA" not in u[:600]:
        return "PLAN"
    if re.search(r"ART\.?\s*5[12]", u):
        return "DEBTOR_LIST"
    if "ADMINISTRADOR" in u and ("RELA" in u or "LISTA" in u):
        return "AJ_LIST"
    if "HOMOLOG" in u:
        return "HOMOLOG"
    return "UNKNOWN"


def first_pages_text(path: str, n: int = 3) -> str:
    with open_pdf(path) as pdf:
        return "\n".join(page_text(p) for p in pdf.pages[:n])


def extract_for_ingest(args):
    """Worker: pure extraction, no DB. Returns everything ingest_document needs to write."""
    sha1, path, url, page_cases, page_title, domain = args
    try:
        head = first_pages_text(path)
        doc = extract_document(path)
        rec = reconcile(doc)
        if doc.has_text_layer and not [r for r in doc.rows if r.value_brl is not None]:
            rec["status"] = "NO_ROWS"
        return {"ok": True, "sha1": sha1, "path": path, "url": url, "page_cases": page_cases, "page_title": page_title,
                "domain": domain, "head": head, "doc": doc, "rec": rec}
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "sha1": sha1, "path": path, "url": url, "domain": domain, "error": f"{type(e).__name__}: {str(e)[:300]}"}


def write_ingested(db, run_id, w, by="ingest"):
    head, doc, rec = w["head"], w["doc"], w["rec"]
    sha1, path, url, page_cases, page_title, domain = w["sha1"], w["path"], w["url"], w["page_cases"], w["page_title"], w["domain"]
    cnj = Counter(CNJ_RE.findall(head))
    case_number = cnj.most_common(1)[0][0] if cnj else ((page_cases or "").split(",")[0] or None)
    case_flag = "" if cnj else ("CASE_FROM_PAGE" if case_number else "CASE_NUMBER_MISSING")
    dm = DATE_RE.search(head)
    pub = f"{dm.group(3)}-{dm.group(2)}-{dm.group(1)}" if dm else None
    dtype = doc_type_from_text(head)
    dh = DEBTOR_RE.search(head.replace("\n", " "))
    debtor_hint = re.sub(r"\s+", " ", dh.group(1)).strip(" ,.") if dh else (page_title or None)
    return _write(db, run_id, sha1, path, url, case_number, case_flag, pub, dtype, debtor_hint, domain, doc, rec, by)


def ingest_document(db, run_id: str, sha1: str, path: str, url: str, page_cases: str, page_title: str, domain: str, by="ingest"):
    w = extract_for_ingest((sha1, path, url, page_cases, page_title, domain))
    if not w["ok"]:
        raise RuntimeError(w["error"])
    return write_ingested(db, run_id, w, by)


def _write(db, run_id, sha1, path, url, case_number, case_flag, pub, dtype, debtor_hint, domain, doc, rec, by):
    db.execute("UPDATE documents SET superseded_by=? WHERE doc_id=? AND superseded_by IS NULL", (run_id, sha1))
    db.execute("UPDATE claims SET superseded_by=? WHERE doc_id=? AND superseded_by IS NULL", (run_id, sha1))
    db.execute("""INSERT OR REPLACE INTO documents VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
               (sha1, run_id, case_number, dtype, pub, url, path, doc.pages, doc.strategy, doc.layout_id,
                json.dumps([{"class": t.klass, "currency": t.currency, "total": str(t.total) if t.total is not None else None,
                             "count": t.count, "page": t.page, "section": t.section} for t in doc.printed_totals], ensure_ascii=False),
                rec["status"] if doc.has_text_layer else "NO_TEXT_LAYER",
                json.dumps(rec, default=str, ensure_ascii=False),
                "; ".join(doc.notes + ([case_flag] if case_flag else []) + ["classIII=" + rec.get("per_class", {}).get("III/BRL", "NO_ROWS")]),
                debtor_hint, domain, now(), by, None))
    rows = [(sha1, run_id, case_number, r.page, r.row_index, r.seq_as_printed, r.creditor_name_as_printed,
             r.document_as_printed, r.document_number, r.document_type, "|".join(r.all_documents), r.klass, r.class_set_by,
             r.value_as_printed, str(r.value_brl) if r.value_brl is not None else None, r.currency, r.debtor_as_printed,
             r.section_heading, "|".join(r.flags), r.strategy, now(), by, None) for r in doc.rows]
    db.executemany("""INSERT INTO claims(doc_id,run_id,case_number,page,row_index,seq_as_printed,creditor_name_as_printed,
        document_as_printed,document_number,document_type,all_documents,class,class_set_by,value_as_printed,value_brl,currency,
        debtor_as_printed,section_heading,flags,strategy,extracted_at,extracted_by,superseded_by) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", rows)
    db.commit()
    return {"doc_id": sha1, "case_number": case_number, "doc_type": dtype, "status": rec["status"] if doc.has_text_layer else "NO_TEXT_LAYER",
            "rows": len(doc.rows), "strategy": doc.strategy, "layout": doc.layout_id, "pages": doc.pages}


def ingest_all(db, run_id: str, log=print, kinds=("LIST", "PLAN", "HOMOLOG"), limit=None, workers: int = 1):
    q = db.execute("""SELECT sha1, path, url, case_numbers, page_title, domain FROM raw_documents
                      WHERE http_status=200 AND path IS NOT NULL AND doc_kind IN (%s)
                      AND sha1 NOT IN (SELECT doc_id FROM documents WHERE run_id=?)""" % ",".join("?" * len(kinds)),
                   (*kinds, run_id)).fetchall()
    if limit:
        q = q[:limit]
    stats = Counter()

    def handle(i, w):
        if not w["ok"]:
            stats["ERROR"] += 1
            log(f"  [{i}/{len(q)}] {w['domain']} ERROR {w['error'][:120]}")
            db.execute("INSERT OR REPLACE INTO documents(doc_id,run_id,source_url,file_path,status,notes,extracted_at,extracted_by) VALUES(?,?,?,?,?,?,?,?)",
                       (w["sha1"], run_id, w["url"], w["path"], "ERROR", w["error"], now(), "ingest"))
            db.commit()
            return
        r = write_ingested(db, run_id, w)
        stats[r["status"]] += 1
        log(f"  [{i}/{len(q)}] {w['domain']} {r['doc_type']:<11} {r['status']:<13} rows={r['rows']:<4} {r['strategy']}/{r['layout']} case={r['case_number']} p={r['pages']}")

    if workers > 1:
        from concurrent.futures import ProcessPoolExecutor
        with ProcessPoolExecutor(workers) as ex:
            for i, w in enumerate(ex.map(extract_for_ingest, q, chunksize=2), 1):
                handle(i, w)
    else:
        for i, args in enumerate(q, 1):
            handle(i, extract_for_ingest(args))
    return stats
