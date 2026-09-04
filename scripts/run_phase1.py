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
ap.add_argument("--reingest", default=None, help="'all' or a comma list of statuses (e.g. QUARANTINED,ERROR,NO_ROWS): drop those documents from this run and extract them again")
ap.add_argument("--reingest-before", default=None, help="ISO time: also re-extract documents of this run extracted before that moment (older extractor code)")
args = ap.parse_args()

DB = str(ROOT / "data" / "anteparo.sqlite")
db = connect(DB)
run = new_run(db, "phase1") if args.run == "new" else args.run
log = lambda m: print(time.strftime("%H:%M:%S"), m, flush=True)
log(f"run {run}")

if (args.reingest or args.reingest_before) and args.run != "new":
    ids = set()
    if args.reingest == "all":
        ids |= {r[0] for r in db.execute("SELECT doc_id FROM documents WHERE run_id=?", (run,))}
    elif args.reingest:
        sts = [x.strip() for x in args.reingest.split(",")]
        ids |= {r[0] for r in db.execute(f"SELECT doc_id FROM documents WHERE run_id=? AND status IN ({','.join('?'*len(sts))})", (run, *sts))}
    if args.reingest_before:
        ids |= {r[0] for r in db.execute("SELECT doc_id FROM documents WHERE run_id=? AND extracted_at < ?", (run, args.reingest_before))}
    ids = sorted(ids)
    for i in range(0, len(ids), 500):
        chunk = ids[i:i + 500]
        db.execute(f"DELETE FROM claims WHERE run_id=? AND doc_id IN ({','.join('?'*len(chunk))})", (run, *chunk))
        db.execute(f"DELETE FROM documents WHERE run_id=? AND doc_id IN ({','.join('?'*len(chunk))})", (run, *chunk))
    db.commit()
    log(f"reingest: dropped {len(ids)} documents from run {run} for re-extraction")
stats = ingest_all(db, run, log=log, workers=args.workers, limit=args.limit)
log(f"ingest: {dict(stats)}")
build_targets(db, run, log=log)
if not args.no_enrich:
    enrich(db, run, DB, log=log, limit=args.enrich_limit)
out = export_csvs(db, run, args.out)
xlsx = Path(args.out) / "ANTEPARO_call_sheet.xlsx"
counts = export_xlsx(db, run, str(xlsx), old_csv=str(ROOT / "data/seeds/old_sheet_targets.csv"))
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
from anteparo.export.report import write_report
write_report(db, run, args.out, counts)
log("report.md written")
log("PHASE 1 DONE")
