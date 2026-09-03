"""Re-export CSVs + workbook for an existing run (no re-ingest). Usage: reexport.py [run_id|latest] [out_dir]"""
from __future__ import annotations
import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(ROOT))
from anteparo.db import connect
from anteparo.export.csv_out import export_csvs
from anteparo.export.xlsx_out import export_xlsx
db = connect(str(ROOT / "data/anteparo.sqlite"))
run = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] != "latest" else db.execute("SELECT run_id FROM runs WHERE note='phase1' ORDER BY started_at DESC LIMIT 1").fetchone()[0]
out = sys.argv[2] if len(sys.argv) > 2 else str(ROOT / "data/out")
export_csvs(db, run, out)
counts = export_xlsx(db, run, str(Path(out) / "ANTEPARO_call_sheet.xlsx"), old_csv=str(ROOT / "data/seeds/old_sheet_targets.csv"))
print(run, counts)
