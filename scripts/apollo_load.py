"""Load Apollo picks (data/apollo/results.jsonl) and revealed people (data/apollo/matched.jsonl) into apollo_contacts.

rank 1 = decision maker, rank 2 = backup. Only people whose full record was revealed by bulk_match are loaded;
a pick that was never matched has only a masked surname and is skipped.
"""
import json, sys, time
BAD_DOMAINS = {"stonebrewing.com": "Apollo matched a US brewery, not Stone SCD", "ecopneus.it": "Italian namesake, not Ecopneus PA",
               "casaraomc.com.br": "side job domain, not Laponia", "fazenda.no": "personal domain", "sicoobtranscredi.com.br": "previous employer domain"}
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(ROOT))
from anteparo.db import connect
db = connect(str(ROOT / "data/anteparo.sqlite"))
db.executescript("""CREATE TABLE IF NOT EXISTS apollo_contacts(
  cnpj_basico TEXT, rank INTEGER, apollo_id TEXT, person_name TEXT, title TEXT, email TEXT, email_status TEXT,
  mobile TEXT, direct_phone TEXT, org_phone TEXT, linkedin TEXT, org_name TEXT, city TEXT, state TEXT, pick_note TEXT,
  source TEXT, fetched_at TEXT, PRIMARY KEY(cnpj_basico, rank));""")
matched = {}
for line in open(ROOT / "data/apollo/matched.jsonl", encoding="utf-8"):
    r = json.loads(line); matched[r["id"]] = r
n = 0; skipped = 0
for line in open(ROOT / "data/apollo/results.jsonl", encoding="utf-8"):
    row = json.loads(line)
    roots = [row.get("root")] + list(row.get("also_roots") or [])
    picks = row.get("picks") or []
    for root in roots:
        if not root: continue
        rank = 0
        for p in picks:
            m = matched.get(p["id"])
            if not m:
                skipped += 1; continue
            rank += 1
            phones = dict((t, num) for t, num in (m.get("phones") or []) if num)
            mobile = phones.get("mobile"); direct = phones.get("work_direct") or phones.get("direct") or phones.get("other")
            email, status, note = m.get("email"), m.get("email_status"), p.get("note")
            # emails the enrichment workers identified as sitting on a stranger's domain (namesake company, personal domain, ex-employer)
            dom = (email or "").rsplit("@", 1)[-1].lower()
            if email and dom in BAD_DOMAINS:
                note = (note + "; " if note else "") + f"email dropped: {BAD_DOMAINS[dom]}"; email = None; status = "dropped"
            db.execute("INSERT OR REPLACE INTO apollo_contacts VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                       (root, rank, m["id"], m.get("name"), m.get("title") or p.get("title"), email, status,
                        mobile, direct, m.get("org_phone"), m.get("linkedin"), m.get("org"), m.get("city"), m.get("state"),
                        note, "apollo", time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())))
            n += 1
db.commit()
print(f"apollo_contacts: {n} rows loaded, {skipped} picks not yet revealed; companies with a contact: "
      f"{db.execute('SELECT COUNT(DISTINCT cnpj_basico) FROM apollo_contacts').fetchone()[0]}; with email: "
      f"{db.execute('SELECT COUNT(DISTINCT cnpj_basico) FROM apollo_contacts WHERE email IS NOT NULL').fetchone()[0]}")
