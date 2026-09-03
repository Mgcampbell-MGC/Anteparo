"""Step A driver — every RJ case since 2020 in all 27 state courts → cases table."""
from __future__ import annotations
import sys, time
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(ROOT))
from anteparo.db import connect, now
from anteparo.sources.datajud import COURTS, iter_cases, derive

db = connect(str(ROOT / "data/anteparo.sqlite"))
courts = sys.argv[1:] or COURTS
total = 0
for c in courts:
    n = 0
    try:
        for h in iter_cases(c, since="2020-01-01", size=1000):
            d = derive(h)
            if not d["case_number"]:
                continue
            db.execute("""INSERT INTO cases(case_number,court,chamber,chamber_code,grau,filing_date,stage,rj_granted_signal,
                last_movement,n_movements,movements_json,system,source_url,datajud_updated,extracted_at,extracted_by)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(case_number) DO UPDATE SET stage=excluded.stage, last_movement=excluded.last_movement,
                n_movements=excluded.n_movements, movements_json=excluded.movements_json, datajud_updated=excluded.datajud_updated,
                rj_granted_signal=excluded.rj_granted_signal, extracted_at=excluded.extracted_at""",
                (d["case_number"], d["court"], d["chamber"], d["chamber_code"], d["grau"], d["filing_date"], d["stage"],
                 int(d["rj_granted_signal"]), d["last_movement"], d["n_movements"], d["movements_json"], d["system"],
                 d["source_url"], d["datajud_updated"], now(), "datajud-loader"))
            n += 1
            if n % 500 == 0:
                db.commit()
        db.commit()
    except Exception as e:  # noqa: BLE001
        print(time.strftime("%H:%M:%S"), f"{c}: ERROR {type(e).__name__}: {e}", flush=True)
    total += n
    print(time.strftime("%H:%M:%S"), f"{c}: {n} cases since 2020", flush=True)
print(time.strftime("%H:%M:%S"), f"CASES DONE: {total}", flush=True)
