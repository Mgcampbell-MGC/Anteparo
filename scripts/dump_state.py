"""Dump the hand-curated / enrichment tables to data/state/*.csv so the repo carries what the SQLite file (git-ignored) holds.
Usage: dump_state.py"""
from __future__ import annotations
import csv, sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(ROOT))
from anteparo.db import connect
db = connect(str(ROOT / "data/anteparo.sqlite"))
out = ROOT / "data/state"; out.mkdir(parents=True, exist_ok=True)
for table in ("debtors", "doc_debtor", "apollo_contacts", "plan_terms", "case_notes"):
    if not db.execute("SELECT name FROM sqlite_master WHERE name=?", (table,)).fetchone(): continue
    cols = [r[1] for r in db.execute(f"PRAGMA table_info({table})")]
    rows = db.execute(f"SELECT * FROM {table} ORDER BY 1" + (", 2" if table == "apollo_contacts" else "")).fetchall()
    with open(out / f"{table}.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(cols); w.writerows(rows)
    print(f"{table}: {len(rows)} rows → data/state/{table}.csv")
