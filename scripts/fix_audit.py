"""Data corrections driven by the audit + PDF readers. Idempotent; run before scripts/debtors.py.

- no-case documents whose portal record carries exactly one valid state-court case number get that case
  (EGESA: 5088952-81.2025.8.13.0024, TJMG) and a debtor name from the portal slug;
- doc_debtor '' → NULL (an empty string used to mask the page-title hint through COALESCE);
- an extracted debtor CNPJ that is itself a creditor row on the same case is rejected (rfb_source='CREDITOR_CNPJ');
- plan_terms: per-case class III terms read from the plan PDF (Grupo 3E) override the band defaults;
- recuperação extrajudicial documents are marked (proceeding 'RE').
"""
from __future__ import annotations
import re, sys, json, time
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(ROOT))
from anteparo.db import connect

db = connect(str(ROOT / "data/anteparo.sqlite"))
run = sys.argv[1] if len(sys.argv) > 1 else db.execute("SELECT run_id FROM runs WHERE note='phase1' ORDER BY started_at DESC LIMIT 1").fetchone()[0]
now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
CNJ = re.compile(r"\d{7}-\d{2}\.\d{4}\.8\.\d{2}\.\d{4}")
log = lambda m: print(time.strftime("%H:%M:%S"), m, flush=True)

db.executescript("""
CREATE TABLE IF NOT EXISTS plan_terms(case_number TEXT PRIMARY KEY, recovery REAL, grace REAL, term REAL, source TEXT, note TEXT, updated_at TEXT);
""")
for col, typ in (("display_name", "TEXT"), ("display_source", "TEXT"), ("verified_name", "TEXT"), ("verified_cnpj", "TEXT"), ("group_members", "TEXT"),
                 ("administrator", "TEXT"), ("proceeding", "TEXT"), ("verified_confidence", "REAL"), ("verified_evidence", "TEXT"), ("verified_at", "TEXT")):
    if col not in [r[1] for r in db.execute("PRAGMA table_info(debtors)")]:
        db.execute(f"ALTER TABLE debtors ADD COLUMN {col} {typ}")

# 1. doc_debtor '' masks the hint
n = db.execute("UPDATE doc_debtor SET debtor=NULL WHERE debtor=''").rowcount
log(f"doc_debtor '' → NULL: {n}")

# 2. no-case documents with exactly one valid state-court number on the portal record
def slug_name(page_url):
    m = re.search(r"recuperacao-judicial_([^/]+?)__\d+$", page_url or "")
    if not m: return None
    s = m.group(1).replace("+", ".").replace("-", " ")
    s = re.sub(r"\s*\(\s*", " (", s); s = re.sub(r"\s+", " ", s).strip()
    return s.upper()
fixed = 0
for doc_id, url, page_url, nums in db.execute("""SELECT d.doc_id, d.source_url, r.page_url, r.case_numbers FROM documents d JOIN raw_documents r ON r.url=d.source_url
        WHERE d.run_id=? AND d.case_number IS NULL""", (run,)).fetchall():
    cands = sorted({c for c in CNJ.findall(nums or "") if int(c[11:15]) >= 2005})
    name = slug_name(page_url)
    if len(cands) == 1:
        case = cands[0]
        db.execute("UPDATE documents SET case_number=?, notes=COALESCE(notes,'')||? WHERE doc_id=? AND run_id=?", (case, f"; case from portal record {case} ({now})", doc_id, run))
        db.execute("UPDATE targets SET case_number=?, flags=COALESCE(flags,'')||' CASE_FROM_PORTAL' WHERE doc_id=? AND run_id=?", (case, doc_id, run))
        db.execute("UPDATE claims SET case_number=? WHERE doc_id=? AND run_id=?", (case, doc_id, run))
        db.execute("INSERT OR IGNORE INTO debtors(case_number, debtor_name, updated_at) VALUES(?,?,?)", (case, name, now))
        if name:
            db.execute("UPDATE debtors SET debtor_name=COALESCE(debtor_name, ?), display_name=COALESCE(verified_name, ?), display_source=COALESCE(display_source,'hint') WHERE case_number=?", (name, name, case))
        fixed += 1; log(f"  {doc_id[:8]} → {case} | {name}")
    elif name:
        db.execute("INSERT OR REPLACE INTO doc_debtor VALUES(?,?)", (doc_id, name)); log(f"  {doc_id[:8]} no case; debtor from slug: {name}")
log(f"documents given a portal case number: {fixed}")
# the EGESA slug lists ten group companies; the closer needs the group name, the members go to group_members
egesa = "5088952-81.2025.8.13.0024"
if db.execute("SELECT 1 FROM debtors WHERE case_number=?", (egesa,)).fetchone():
    members = ["EGESA ENGENHARIA S.A.", "EGESUR PARTICIPAÇÕES E EMPREENDIMENTOS S.A.", "MATRIX INFRA LTDA", "BEMVIVER ENGENHARIA AMBIENTAL E SERVIÇOS LTDA", "DKF CONSTRUÇÕES E EMPREENDIMENTOS LTDA",
               "EGEPEL LTDA", "ETR ASSESSORIA EMPRESARIAL E PARTICIPAÇÕES LTDA", "MVT ENGENHARIA E SERVIÇOS LTDA", "PARQUES DO VALE LTDA", "ELMO RIBEIRO E ANA RIBEIRO"]
    db.execute("UPDATE debtors SET display_name=COALESCE(verified_name, ?), display_source=COALESCE(NULLIF(display_source,'hint'),'hint'), group_members=COALESCE(group_members, ?) WHERE case_number=?",
               ("GRUPO EGESA (Egesa Engenharia S.A. + 9 group companies; see group members)", json.dumps(members, ensure_ascii=False), egesa))

# 3. debtor CNPJ that is a creditor on the same case → rejected
rej = 0
for case, cnpj in db.execute("SELECT case_number, debtor_cnpj FROM debtors WHERE debtor_cnpj IS NOT NULL AND verified_cnpj IS NULL").fetchall():
    hit = db.execute("SELECT 1 FROM claims WHERE run_id=? AND case_number=? AND substr(document_number,1,8)=? LIMIT 1", (run, case, cnpj[:8])).fetchone()
    if hit:
        db.execute("UPDATE debtors SET rfb_source='CREDITOR_CNPJ', updated_at=? WHERE case_number=? AND rfb_source NOT IN ('CREDITOR_CNPJ')", (now, case)); rej += 1
log(f"debtor CNPJs rejected because they are creditor rows on the same case: {rej}")

# 4. plan terms read from plan PDFs (class III general condition — what a fund buyer gets)
plans = [
    ("5341375-33.2026.8.09.0049", 0.12, 2.0, 13.0, "Plano de Recuperação Judicial Grupo 3E v2.0 (Métis), mov. 321, pp.24-25",
     "Subclasse III-A: deságio 88%, carência 24 meses da homologação, 13 parcelas anuais SAC anos 3-15, TR + 1% a.a. Subclasse III-B (fornecedores parceiros que assinam termo de adesão e mantêm fornecimento): deságio 70%, carência 12 meses, 48 parcelas mensais — not available to a fund buyer."),
]
for case, rec, g, t, src, note in plans:
    db.execute("INSERT OR REPLACE INTO plan_terms VALUES(?,?,?,?,?,?,?)", (case, rec, g, t, src, note, now))
log(f"plan_terms rows: {db.execute('SELECT COUNT(*) FROM plan_terms').fetchone()[0]}")

# 5. recuperação extrajudicial
for (case,) in db.execute("SELECT DISTINCT case_number FROM documents WHERE run_id=? AND case_number IS NOT NULL AND (source_url LIKE '%extrajudicial%' OR debtor_name_hint LIKE '%extrajudicial%')", (run,)).fetchall():
    db.execute("INSERT OR IGNORE INTO debtors(case_number, updated_at) VALUES(?,?)", (case, now))
    db.execute("UPDATE debtors SET proceeding='RE' WHERE case_number=? AND proceeding IS NULL", (case,)); log(f"  extrajudicial: {case}")
db.commit()
log("FIX_AUDIT DONE")
