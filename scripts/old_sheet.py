"""Parse the TARGETS tab of the existing Drive sheet ('SELLERS list TEST') into a CSV for the diff."""
from __future__ import annotations
import csv, re, sys
from pathlib import Path
SRC = Path("/tmp/claude-0/-home-user-Anteparo/f6b38504-968f-5e67-8071-9786cc41bb5a/scratchpad/sellers_sheet_raw.txt")
OUT = Path(__file__).resolve().parents[1] / "data" / "seeds" / "old_sheet_targets.csv"
t = SRC.read_text(encoding="utf-8")
rows = [ln for ln in t.splitlines() if ln.startswith("|")]
hdr = None; out = []; seen = set()
CNPJ = re.compile(r"\b\d{14}\b")
CASE = re.compile(r"\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}")
AMT = re.compile(r"R\$\s*([\d.,]+)")
for ln in rows:
    cells = [c.strip().replace("\\_", "_") for c in ln.strip("|").split("|")]
    if any("CREDITOR" in c.upper() and "OWED" in c.upper() for c in cells):
        hdr = cells; continue
    if set("".join(cells)) <= set(":- "):
        continue
    if hdr is None:
        continue
    joined = " | ".join(cells)
    cn = CNPJ.search(joined); cs = CASE.search(joined); am = AMT.search(joined)
    url = re.search(r"https?://\S+", joined)
    if not cn:
        continue
    amt = am.group(1).replace(",", "") if am else ""
    key = (cn.group(), cs.group() if cs else "")
    if key in seen:
        continue
    seen.add(key)
    out.append({"creditor_name": cells[0], "cnpj": cn.group(), "case_number": cs.group() if cs else "",
                "amount_old": amt, "status_old": cells[1] if len(cells) > 1 else "", "source_url": url.group() if url else "",
                "raw": joined[:300]})
OUT.parent.mkdir(parents=True, exist_ok=True)
with open(OUT, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=list(out[0].keys())); w.writeheader(); w.writerows(out)
print(f"old sheet rows parsed: {len(out)} → {OUT}")
for r in out[:3]:
    print("  ", r["creditor_name"][:40], r["cnpj"], r["case_number"], r["amount_old"], r["status_old"])
