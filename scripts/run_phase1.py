"""Phase 1 driver: ingest harvested PDFs → targets → enrich → six CSVs + call sheet + report."""
from __future__ import annotations
import argparse, csv, json, sys, time
from collections import Counter
from decimal import Decimal
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(ROOT))
from anteparo.db import connect, new_run
from anteparo.steps.ingest import ingest_all
from anteparo.steps.build import build_targets, enrich
from anteparo.export.csv_out import export_csvs
from anteparo.export.xlsx_out import export_xlsx

ap = argparse.ArgumentParser()
ap.add_argument("--run", default="new", help="'new' or an existing run_id to continue")
ap.add_argument("--no-enrich", action="store_true")
ap.add_argument("--enrich-limit", type=int, default=None)
ap.add_argument("--out", default=str(ROOT / "data" / "out"))
ap.add_argument("--workers", type=int, default=4)
ap.add_argument("--limit", type=int, default=None, help="ingest at most N documents (smoke test)")
args = ap.parse_args()

DB = str(ROOT / "data" / "anteparo.sqlite")
db = connect(DB)
run = new_run(db, "phase1") if args.run == "new" else args.run
log = lambda m: print(time.strftime("%H:%M:%S"), m, flush=True)
log(f"run {run}")

stats = ingest_all(db, run, log=log, workers=args.workers, limit=args.limit)
log(f"ingest: {dict(stats)}")
build_targets(db, run, log=log)
if not args.no_enrich:
    enrich(db, run, DB, log=log, limit=args.enrich_limit)
out = export_csvs(db, run, args.out)
xlsx = Path(args.out) / "ANTEPARO_call_sheet.xlsx"
counts = export_xlsx(db, run, str(xlsx))
log(f"xlsx: {counts} → {xlsx}")

# ---- diff vs the old Drive sheet (by CNPJ root × case) ----
old = Path(ROOT / "data/seeds/old_sheet_targets.csv")
diff_rows = []
if old.exists():
    new = {}
    for root, case, face, name in db.execute("SELECT cnpj_basico, case_number, class_iii_face_sum, creditor_name_as_printed FROM targets WHERE run_id=?", (run,)):
        new[(root, case or "")] = (Decimal(face), name)
    with open(old, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            key = (r["cnpj"][:8], r["case_number"])
            try:
                old_amt = Decimal(r["amount_old"]) if r["amount_old"] else None
            except Exception:  # noqa: BLE001
                old_amt = None
            if key in new:
                new_amt, nm = new[key]
                delta = (new_amt - old_amt) if old_amt is not None else None
                diff_rows.append({"creditor": r["creditor_name"], "cnpj": r["cnpj"], "case": r["case_number"],
                                  "old_amount": str(old_amt or ""), "new_face_root_sum": str(new_amt), "delta": str(delta or ""),
                                  "verdict": "SAME" if delta is not None and abs(delta) < 1 else ("CHANGED" if delta is not None else "NEW_ONLY"),
                                  "old_status": r["status_old"]})
            else:
                diff_rows.append({"creditor": r["creditor_name"], "cnpj": r["cnpj"], "case": r["case_number"], "old_amount": r["amount_old"],
                                  "new_face_root_sum": "", "delta": "", "verdict": "NOT_IN_NEW_RUN", "old_status": r["status_old"]})
    with open(Path(args.out) / "diff_vs_drive_sheet.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(diff_rows[0].keys()) if diff_rows else ["empty"]); w.writeheader(); w.writerows(diff_rows)
    vc = Counter(r["verdict"] for r in diff_rows)
    log(f"diff vs Drive sheet: {dict(vc)}")

# ---- report.md ----
docs = db.execute("SELECT status, COUNT(*) FROM documents WHERE run_id=? GROUP BY status", (run,)).fetchall()
layouts = db.execute("SELECT strategy||'/'||COALESCE(layout_id,''), status, COUNT(*) FROM documents WHERE run_id=? GROUP BY 1,2", (run,)).fetchall()
quar = db.execute("SELECT doc_id, source_url, reconcile_json FROM documents WHERE run_id=? AND status='QUARANTINED'", (run,)).fetchall()
cov = db.execute("""SELECT c.court, COUNT(DISTINCT c.case_number), COUNT(DISTINCT d.case_number)
                    FROM cases c LEFT JOIN documents d ON d.case_number=c.case_number AND d.run_id=? AND d.status IN ('OK','OK_NO_TOTALS')
                    GROUP BY c.court ORDER BY 2 DESC""", (run,)).fetchall()
rfb = db.execute("SELECT rfb_source, MIN(rfb_fetched_at), MAX(rfb_fetched_at), COUNT(*) FROM companies GROUP BY rfb_source").fetchall()
blocked = db.execute("SELECT domain, COUNT(*) FROM raw_documents WHERE http_status<>200 GROUP BY domain ORDER BY 2 DESC LIMIT 15").fetchall()
unres = db.execute("SELECT COUNT(*) FROM claims WHERE run_id=? AND class='III' AND document_type='NONE' AND value_brl IS NOT NULL", (run,)).fetchone()[0]
lines = [f"# ANTEPARO index — Phase 1 report", f"run `{run}` · {time.strftime('%Y-%m-%d %H:%MZ', time.gmtime())}", "",
         "## Documents by status", *[f"- {s}: {n}" for s, n in docs], "",
         "## Layouts", *[f"- {l} · {s}: {n}" for l, s, n in layouts], "",
         f"## Targets", f"- ≥ R$200k: {counts['floor']} · pooling 100–200k: {counts['pool']} · name-only (needs CNPJ) ≥ 200k: {counts.get('needs_cnpj', 0)}",
         f"- class III rows without a printed document number (name-resolution queue): {unres}", "",
         "## RFB releases used (via mirror)", *[f"- {s}: {n} companies, fetched {a} … {b}" for s, a, b, n in rfb], "",
         "## Coverage by court (cases in DataJud universe since 2020 → cases with ≥1 usable creditor list)",
         *[f"- {c}: {n} cases · {m} with a usable list ({(100*m//n) if n else 0}%)" for c, n, m in cov], "",
         "## Quarantined documents", *([f"- {d[:12]} {u}" for d, u, _ in quar] or ["- none"]), "",
         "## Sources that failed or were blocked", *([f"- {d}: {n}" for d, n in blocked] or ["- none"]), ""]
Path(args.out, "report.md").write_text("\n".join(lines), encoding="utf-8")
log("report.md written")
log("PHASE 1 DONE")
