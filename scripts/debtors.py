"""Debtor identification + enrichment + banding.

1. Debtor CNPJ from each document's first pages, context-scored so the administrator's own CNPJ is never taken.
2. RFB enrichment of the debtor; accepted only when the registry name matches the printed debtor name.
3. Band A–D per case from stage, plan document, registry status, age and dormancy.
"""
from __future__ import annotations
import re, sys, json, time
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(ROOT))
from anteparo.db import connect, now
from anteparo.cnpj import find_documents, is_valid_cnpj, root as cnpj_root
from anteparo.sources.rfb import RFB, normalise
from anteparo.extract.pdfio import open_pdf, page_text
from rapidfuzz import fuzz

RUN = sys.argv[1] if len(sys.argv) > 1 else None
db = connect(str(ROOT / "data/anteparo.sqlite"))
run = RUN or db.execute("SELECT run_id FROM runs WHERE note='phase1' ORDER BY started_at DESC LIMIT 1").fetchone()[0]
db.executescript("""
CREATE TABLE IF NOT EXISTS debtors(case_number TEXT PRIMARY KEY, debtor_name TEXT, debtor_cnpj TEXT, cnpj_source TEXT, name_match REAL,
  razao_social TEXT, situacao TEXT, situacao_desc TEXT, porte TEXT, capital_social TEXT, uf TEXT, municipio TEXT, cnae_desc TEXT,
  rfb_source TEXT, band TEXT, band_reasons TEXT, plan_status TEXT, stage TEXT, filing_year INTEGER, last_movement TEXT, updated_at TEXT);
""")
POS = re.compile(r"RECUPERANDA|REQUERENTE|AUTOR|DEVEDOR|EM RECUPERA|GRUPO|RECUPERA[ÇC][ÃA]O JUDICIAL D", re.I)
NEG = re.compile(r"ADMINISTRADOR|ADMINISTRA[ÇC][ÃA]O JUDICIAL|\bAJ\b|CONSULTORIA|PER[ÍI]CIA|ADVOGAD|OAB|ESCRIT[ÓO]RIO|SOCIEDADE INDIVIDUAL", re.I)
log = lambda m: print(time.strftime("%H:%M:%S"), m, flush=True)

# --- 1. debtor name + CNPJ per case (best document per case) ---
cases = {}
for case, doc_id, path, hint in db.execute("""SELECT d.case_number, d.doc_id, d.file_path, COALESCE(dd.debtor, d.debtor_name_hint)
        FROM documents d LEFT JOIN doc_debtor dd ON dd.doc_id=d.doc_id
        WHERE d.run_id=? AND d.case_number IS NOT NULL AND d.status IN ('OK','OK_NO_TOTALS','QUARANTINED') ORDER BY d.case_number""", (run,)):
    cases.setdefault(case, []).append((doc_id, path, hint))
log(f"{len(cases)} cases with documents")
already = {c for (c,) in db.execute("SELECT case_number FROM debtors")}   # idempotent re-runs after a crash
found = 0
for case, docs in cases.items():
    if case in already:
        continue
    best = None
    dname = next((h for _, _, h in docs if h), None)
    for doc_id, path, hint in docs[:3]:
        try:
            with open_pdf(path) as pdf:
                head = "\n".join(page_text(p) for p in pdf.pages[:2])
        except Exception:
            continue
        for kind, digits in find_documents(head):
            if kind != "CNPJ" or not is_valid_cnpj(digits):
                continue
            i = head.replace(".", "").replace("/", "").replace("-", "").find(digits)   # rough position in stripped text
            # window around the CNPJ in the original text
            m = re.search(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}", head)
            pos = None
            for mm in re.finditer(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}", head):
                if re.sub(r"\D", "", mm.group()) == digits:
                    pos = mm.start(); break
            if pos is None:
                continue
            win = head[max(0, pos - 160):pos + 40]
            score = 0
            if POS.search(win): score += 2
            if NEG.search(win): score -= 3
            if dname and fuzz.partial_ratio(dname[:30].upper(), win.upper()) >= 80: score += 3
            if score > 0 and (best is None or score > best[0]):
                best = (score, digits, doc_id)
    if best:
        found += 1
        db.execute("INSERT OR REPLACE INTO debtors(case_number, debtor_name, debtor_cnpj, cnpj_source, updated_at) VALUES(?,?,?,?,?)",
                   (case, dname, best[1], f"doc:{best[2][:8]} score={best[0]}", now()))
    else:
        db.execute("INSERT OR IGNORE INTO debtors(case_number, debtor_name, updated_at) VALUES(?,?,?)", (case, dname, now()))
db.commit()
log(f"debtor CNPJ found for {found}/{len(cases)} cases")

# --- 2. RFB enrichment with name match ---
rfb = RFB(str(ROOT / "data/anteparo.sqlite"))
rows = db.execute("SELECT case_number, debtor_name, debtor_cnpj FROM debtors WHERE debtor_cnpj IS NOT NULL AND rfb_source IS NULL").fetchall()
log(f"enriching {len(rows)} debtors")
ok = rej = miss = 0
for i, (case, dname, cnpj) in enumerate(rows, 1):
    db.commit()   # release our write lock: rfb.get() writes rfb_cache through its own connection
    d = rfb.get(cnpj)
    if not d:
        miss += 1; db.execute("UPDATE debtors SET rfb_source='NOT_IN_RFB', updated_at=? WHERE case_number=?", (now(), case)); continue
    n = normalise(d)
    match = fuzz.token_set_ratio((dname or "").upper(), (n["razao_social"] or "").upper()) / 100 if dname else None
    if dname and match is not None and match < 0.55:
        rej += 1
        db.execute("UPDATE debtors SET rfb_source='NAME_MISMATCH', razao_social=?, name_match=?, updated_at=? WHERE case_number=?", (n["razao_social"], match, now(), case)); continue
    ok += 1
    db.execute("""UPDATE debtors SET razao_social=?, situacao=?, situacao_desc=?, porte=?, capital_social=?, uf=?, municipio=?, cnae_desc=?, rfb_source=?, name_match=?, updated_at=? WHERE case_number=?""",
               (n["razao_social"], n["situacao_cadastral"], n["situacao_desc"], n["porte"], str(n["capital_social"] or ""), n["uf"], n["municipio"], n["cnae_desc"], n["source"], match, now(), case))
    if i % 50 == 0: db.commit(); log(f"  {i}/{len(rows)} ok={ok} rej={rej} miss={miss}")
db.commit()
log(f"debtor enrichment: accepted={ok} name-mismatch={rej} not-in-rfb={miss}")

# --- 3. bands ---
def band_for(case):
    reasons = []
    stage, last_mv, filing = (db.execute("SELECT stage, last_movement, filing_date FROM cases WHERE case_number=?", (case,)).fetchone() or (None, None, None))
    m = re.match(r"\d{7}-\d{2}\.(\d{4})\.", case or "")
    year = int(filing[:4]) if filing else (int(m.group(1)) if m else None)
    plan = db.execute("SELECT COUNT(*) FROM documents WHERE run_id=? AND case_number=? AND doc_type IN ('PLAN','HOMOLOG')", (run, case)).fetchone()[0]
    sit = db.execute("SELECT situacao, rfb_source FROM debtors WHERE case_number=?", (case,)).fetchone() or (None, None)
    age = (2026 - year) if year else None
    dormant = bool(last_mv) and last_mv < "2025-09-01"
    if stage == "CONVERTED_TO_BANKRUPTCY":
        reasons.append("falência decreed"); band = "D"
    elif sit[1] and sit[1] not in ("NOT_IN_RFB", "NAME_MISMATCH") and sit[0] and sit[0] != "02":
        reasons.append(f"debtor not ATIVA at RFB ({sit[0]})"); band = "D"
    elif plan and stage != "CLOSED":
        reasons.append("plan/homologation document found"); band = "A"
        if age is not None and age > 8: reasons.append(f"but case is {age}y old"); band = "B"
    elif (age is not None and age > 4 and not plan) or dormant or stage == "CLOSED":
        band = "C"
        if age is not None and age > 4: reasons.append(f"{age}y since filing, no plan document")
        if dormant: reasons.append(f"no movement since {last_mv}")
        if stage == "CLOSED": reasons.append("case closed/archived")
    else:
        band = "B"; reasons.append("live case, no plan document yet")
    if stage is None and not m: reasons.append("case not matched")
    plan_status = "CONFIRMED" if plan else ("ESTIMATED" if stage else "UNKNOWN")
    return band, "; ".join(reasons), plan_status, stage, year, last_mv

n = 0
for (case,) in db.execute("SELECT case_number FROM debtors").fetchall():
    b, r, ps, st, yr, lm = band_for(case)
    db.execute("UPDATE debtors SET band=?, band_reasons=?, plan_status=?, stage=?, filing_year=?, last_movement=?, updated_at=? WHERE case_number=?", (b, r, ps, st, yr, lm, now(), case)); n += 1
db.commit()
log(f"bands: {db.execute('SELECT band, COUNT(*) FROM debtors GROUP BY band').fetchall()}")
q_band = "SELECT COALESCE(d.band, '?'), COUNT(*) FROM targets t LEFT JOIN debtors d ON d.case_number=t.case_number WHERE t.run_id=? AND t.band=? GROUP BY 1"
log(f"targets by debtor band: {db.execute(q_band, (run, 'FLOOR')).fetchall()}")
log("DEBTORS DONE")
