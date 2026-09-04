"""Parse a saved apollo_people_bulk_match tool result (raw JSON, or the persisted [{"type":"text","text":...}] wrapper)
and append the revealed people to data/apollo/matched.jsonl."""
import json, sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
s = open(sys.argv[1], encoding="utf-8").read().strip()
d = json.loads(s[s.find("{"):]) if s.startswith("{") or not s.startswith("[") else json.loads(s)
if isinstance(d, list):   # persisted wrapper
    d = json.loads(next(x["text"] for x in d if x.get("type") == "text"))
print("credits_consumed:", d.get("credits_consumed"), "| unique:", d.get("unique_enriched_records"), "| missing:", d.get("missing_records"))
seen = {json.loads(l)["id"] for l in open(ROOT / "data/apollo/matched.jsonl", encoding="utf-8")} if (ROOT / "data/apollo/matched.jsonl").exists() else set()
n = 0
with open(ROOT / "data/apollo/matched.jsonl", "a", encoding="utf-8") as f:
    for m in d.get("matches") or []:
        if not m or m.get("id") in seen: continue
        org = m.get("organization") or {}
        rec = {"id": m.get("id"), "name": m.get("name"), "first": m.get("first_name"), "last": m.get("last_name"), "title": m.get("title"),
               "email": m.get("email"), "email_status": m.get("email_status"), "linkedin": m.get("linkedin_url"),
               "org": org.get("name"), "org_domain": org.get("primary_domain"), "org_phone": org.get("phone") or org.get("sanitized_phone"),
               "city": m.get("city"), "state": m.get("state"),
               "phones": [(x.get("type"), x.get("sanitized_number") or x.get("raw_number")) for x in (m.get("phone_numbers") or [])],
               "headline": m.get("headline"), "seniority": m.get("seniority"), "departments": m.get("departments")}
        f.write(json.dumps(rec, ensure_ascii=False) + "\n"); n += 1
        print(json.dumps({k: rec[k] for k in ("name", "title", "email", "email_status", "org")}, ensure_ascii=False))
print("appended:", n)
