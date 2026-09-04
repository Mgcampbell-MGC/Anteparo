"""report.md — what the run produced, what it could not, and why."""
from __future__ import annotations
import time
from pathlib import Path


def write_report(db, run_id, out_dir, counts):
    docs = db.execute("SELECT status, COUNT(*) FROM documents WHERE run_id=? GROUP BY status", (run_id,)).fetchall()
    layouts = db.execute("SELECT strategy||'/'||COALESCE(layout_id,''), status, COUNT(*) FROM documents WHERE run_id=? GROUP BY 1,2", (run_id,)).fetchall()
    partial = db.execute("SELECT COUNT(*) FROM documents WHERE run_id=? AND status='QUARANTINED' AND notes LIKE '%classIII=PASS%'", (run_id,)).fetchone()[0]
    quar = db.execute("SELECT doc_id, source_url FROM documents WHERE run_id=? AND status='QUARANTINED' AND notes LIKE '%classIII=FAIL%' ORDER BY doc_id", (run_id,)).fetchall()
    cov = db.execute("""SELECT c.court, COUNT(DISTINCT c.case_number), COUNT(DISTINCT d.case_number)
                        FROM cases c LEFT JOIN documents d ON d.case_number=c.case_number AND d.run_id=? AND d.status IN ('OK','OK_NO_TOTALS')
                        GROUP BY c.court ORDER BY 2 DESC""", (run_id,)).fetchall()
    matched = db.execute("SELECT COUNT(DISTINCT d.case_number) FROM documents d JOIN cases c ON c.case_number=d.case_number WHERE d.run_id=?", (run_id,)).fetchone()[0]
    doc_cases = db.execute("SELECT COUNT(DISTINCT case_number) FROM documents WHERE run_id=? AND case_number IS NOT NULL", (run_id,)).fetchone()[0]
    rfb = db.execute("SELECT rfb_source, MIN(rfb_fetched_at), MAX(rfb_fetched_at), COUNT(*) FROM companies GROUP BY rfb_source").fetchall()
    blocked = db.execute("SELECT domain, COUNT(*) FROM raw_documents WHERE http_status<>200 GROUP BY domain ORDER BY 2 DESC LIMIT 15").fetchall()
    zero = db.execute("SELECT domain FROM raw_documents GROUP BY domain HAVING SUM(CASE WHEN http_status=200 THEN 1 ELSE 0 END)=0").fetchall()
    unres = db.execute("SELECT COUNT(*) FROM claims WHERE run_id=? AND class='III' AND document_type='NONE' AND value_brl IS NOT NULL", (run_id,)).fetchone()[0]
    excl = db.execute("""SELECT SUM(c.is_bank), SUM(c.is_public), SUM(c.is_inactive) FROM (SELECT DISTINCT cnpj_basico FROM targets WHERE run_id=? AND band='FLOOR') t
                         JOIN companies c USING(cnpj_basico)""", (run_id,)).fetchone()
    byname = db.execute("SELECT COUNT(DISTINCT cnpj_basico) FROM targets WHERE run_id=? AND band='FLOOR' AND (flags LIKE '%LIKELY_FINANCIAL%' OR flags LIKE '%LIKELY_PUBLIC%')", (run_id,)).fetchone()[0]
    lines = [f"# ANTEPARO index — Phase 1 report", f"run `{run_id}` · {time.strftime('%Y-%m-%d %H:%MZ', time.gmtime())}", "",
             "## Call sheet",
             f"- **{counts['floor']} leads ≥ R$200k** on the Call List · {counts['pool']} pooling (R$100–200k) · {counts.get('needs_cnpj', 0)} name-only rows ≥ R$200k on *Needs CNPJ*",
             f"- excluded from the Call List by registry data: banks {excl[0] or 0} · public {excl[1] or 0} · inactive {excl[2] or 0}; by name pre-filter: {byname}",
             f"- {counts['docs']} documents ingested · {counts['cases']} cases with a document · {matched} of {doc_cases} document cases matched to DataJud (rest are pre-2020 or unmatched numbers)", "",
             "## Documents by status", *[f"- {s}: {n}" for s, n in docs],
             f"- of the QUARANTINED, {partial} reconciled on class III and are used (flag PARTIAL_DOC_CLASS_III_RECONCILED)", "",
             "## Layouts", *[f"- {l} · {s}: {n}" for l, s, n in layouts], "",
             f"## Name-resolution queue", f"- class III rows with a value but no printed document number: {unres} (Phase 2: name → CNPJ with an RFB name match)", "",
             "## RFB releases used (via mirror)", *[f"- {s}: {n} companies, fetched {a} … {b}" for s, a, b, n in rfb], "",
             "## Coverage by court (DataJud RJ cases since 2020 → cases with ≥1 usable creditor list)",
             *[f"- {c}: {n} cases · {m} with a usable list ({(100*m//n) if n else 0}%)" for c, n, m in cov], "",
             "## Quarantined documents where class III did not reconcile", *([f"- {d[:12]} {u}" for d, u in quar] or ["- none"]), "",
             "## Sources that failed or were blocked", *([f"- {d}: {n} failed downloads" for d, n in blocked] or ["- none"]),
             *[f"- {d[0]}: no documents retrievable (form/JS-gated or blocked)" for d in zero], ""]
    Path(out_dir, "report.md").write_text("\n".join(lines), encoding="utf-8")
