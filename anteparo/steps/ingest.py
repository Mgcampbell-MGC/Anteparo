"""Extract + reconcile every harvested PDF into documents/claims (steps C+D, append-only)."""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path

ROOT = str(Path(__file__).resolve().parents[2])
DOC_TIMEOUT = int(os.environ.get("ANTEPARO_DOC_TIMEOUT", "240"))
DOC_MEM = int(os.environ.get("ANTEPARO_DOC_MEM", str(2 * 1024 ** 3)))

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
    """Run the extractor for one document in a separate process (memory-capped, time-boxed)."""
    sha1, path, url, page_cases, page_title, domain = args
    base = {"sha1": sha1, "path": path, "url": url, "page_cases": page_cases, "page_title": page_title, "domain": domain}
    cmd = [sys.executable, "-m", "anteparo.steps.worker", path, str(DOC_MEM)]
    try:
        p = subprocess.run(cmd, capture_output=True, timeout=DOC_TIMEOUT, cwd=ROOT)
    except subprocess.TimeoutExpired:
        return {**base, "ok": False, "error": f"TIMEOUT: extraction exceeded {DOC_TIMEOUT}s"}
    if p.returncode != 0:
        tail = p.stderr.decode("utf-8", "replace").strip().splitlines()[-1:] or [""]
        return {**base, "ok": False, "error": f"WORKER_EXIT_{p.returncode}: {tail[0][:250]}"}
    try:
        w = json.loads(p.stdout.decode("utf-8"))
    except ValueError as e:
        return {**base, "ok": False, "error": f"BAD_WORKER_OUTPUT: {e}"}
    return {**base, "ok": True, **w}


def write_ingested(db, run_id, w, by="ingest"):
    head, rec = w["head"], w["rec"]
    sha1, path, url, page_cases, page_title, domain = w["sha1"], w["path"], w["url"], w["page_cases"], w["page_title"], w["domain"]
    cnj = Counter(CNJ_RE.findall(head))
    case_number = cnj.most_common(1)[0][0] if cnj else ((page_cases or "").split(",")[0] or None)
    case_flag = "" if cnj else ("CASE_FROM_PAGE" if case_number else "CASE_NUMBER_MISSING")
    dm = DATE_RE.search(head)
    pub = f"{dm.group(3)}-{dm.group(2)}-{dm.group(1)}" if dm else None
    dtype = doc_type_from_text(head)
    dh = DEBTOR_RE.search(head.replace("\n", " "))
    debtor_hint = re.sub(r"\s+", " ", dh.group(1)).strip(" ,.") if dh else (page_title or None)
    return _write(db, run_id, sha1, path, url, case_number, case_flag, pub, dtype, debtor_hint, domain, w, rec, by)


def ingest_document(db, run_id: str, sha1: str, path: str, url: str, page_cases: str, page_title: str, domain: str, by="ingest"):
    w = extract_for_ingest((sha1, path, url, page_cases, page_title, domain))
    if not w["ok"]:
        raise RuntimeError(w["error"])
    return write_ingested(db, run_id, w, by)


def _write(db, run_id, sha1, path, url, case_number, case_flag, pub, dtype, debtor_hint, domain, w, rec, by):
    db.execute("UPDATE documents SET superseded_by=? WHERE doc_id=? AND superseded_by IS NULL", (run_id, sha1))
    db.execute("UPDATE claims SET superseded_by=? WHERE doc_id=? AND superseded_by IS NULL", (run_id, sha1))
    status = rec["status"] if w["has_text_layer"] else "NO_TEXT_LAYER"
    db.execute("""INSERT OR REPLACE INTO documents VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
               (sha1, run_id, case_number, dtype, pub, url, path, w["pages"], w["strategy"], w["layout_id"],
                json.dumps(w["printed_totals"], ensure_ascii=False), status,
                json.dumps(rec, default=str, ensure_ascii=False),
                "; ".join(w["notes"] + ([case_flag] if case_flag else []) + ["classIII=" + rec.get("per_class", {}).get("III/BRL", "NO_ROWS")]),
                debtor_hint, domain, now(), by, None))
    rows = [(sha1, run_id, case_number, r["page"], r["row_index"], r["seq_as_printed"], r["creditor_name_as_printed"],
             r["document_as_printed"], r["document_number"], r["document_type"], r["all_documents"], r["klass"], r["class_set_by"],
             r["value_as_printed"], r["value_brl"] or None, r["currency"], r["debtor_as_printed"],
             r["section_heading"], r["flags"], r["strategy"], now(), by, None) for r in w["rows"]]
    db.executemany("""INSERT INTO claims(doc_id,run_id,case_number,page,row_index,seq_as_printed,creditor_name_as_printed,
        document_as_printed,document_number,document_type,all_documents,class,class_set_by,value_as_printed,value_brl,currency,
        debtor_as_printed,section_heading,flags,strategy,extracted_at,extracted_by,superseded_by) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", rows)
    db.commit()
    return {"doc_id": sha1, "case_number": case_number, "doc_type": dtype, "status": status,
            "rows": len(w["rows"]), "strategy": w["strategy"], "layout": w["layout_id"], "pages": w["pages"]}


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
        from concurrent.futures import ThreadPoolExecutor      # threads only wait on the per-document subprocesses
        with ThreadPoolExecutor(workers) as ex:
            for i, w in enumerate(ex.map(extract_for_ingest, q), 1):
                handle(i, w)
    else:
        for i, args in enumerate(q, 1):
            handle(i, extract_for_ingest(args))
    return stats
