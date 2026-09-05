"""Data corrections driven by the audit, the PDF readers and the verification round. Idempotent; run before scripts/debtors.py.

- no-case documents (or documents carrying a non-state-court number) whose portal record carries exactly one valid
  state-court case number get that case (EGESA: 5088952-81.2025.8.13.0024, TJMG) and a debtor name from the portal slug;
- doc_debtor '' → NULL (an empty string used to mask the page-title hint through COALESCE);
- an extracted debtor CNPJ that is itself a creditor row on the same case is rejected (rfb_source='CREDITOR_CNPJ');
- plan_terms: per-case class III terms read from the plan PDF (Grupo 3E) override the band defaults;
- recuperação extrajudicial documents are marked (proceeding 'RE');
- grant decisions (sentença de homologação / concessão da RJ) filed as AJ_LIST/UNKNOWN by the text classifier become HOMOLOG,
  which is what band A keys on (art. 58);
- proceedings read from the AJ portal: JOB Fertilizantes (RJ extinguished by desistência, 04/08/2026), Agimax (falência decreed
  13/01/2026), Visão (falência ENCERRADA 01/08/2025);
- case_notes: caveats the closer must see in 'Band reason';
- RFB enrichment for every above-floor creditor root that still has no companies row.
"""
from __future__ import annotations
import re, sys, json, time
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(ROOT))
from anteparo.db import connect

DB_PATH = str(ROOT / "data/anteparo.sqlite")
db = connect(DB_PATH)
run = sys.argv[1] if len(sys.argv) > 1 else db.execute("SELECT run_id FROM runs WHERE note='phase1' ORDER BY started_at DESC LIMIT 1").fetchone()[0]
now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
CNJ = re.compile(r"\d{7}-\d{2}\.\d{4}\.8\.\d{2}\.\d{4}")
log = lambda m: print(time.strftime("%H:%M:%S"), m, flush=True)

db.executescript("""
CREATE TABLE IF NOT EXISTS plan_terms(case_number TEXT PRIMARY KEY, recovery REAL, grace REAL, term REAL, source TEXT, note TEXT, updated_at TEXT);
CREATE TABLE IF NOT EXISTS case_notes(case_number TEXT PRIMARY KEY, note TEXT, source TEXT, updated_at TEXT);
""")
for col, typ in (("display_name", "TEXT"), ("display_source", "TEXT"), ("verified_name", "TEXT"), ("verified_cnpj", "TEXT"), ("group_members", "TEXT"),
                 ("administrator", "TEXT"), ("proceeding", "TEXT"), ("verified_confidence", "REAL"), ("verified_evidence", "TEXT"), ("verified_at", "TEXT")):
    if col not in [r[1] for r in db.execute("PRAGMA table_info(debtors)")]:
        db.execute(f"ALTER TABLE debtors ADD COLUMN {col} {typ}")


def set_case(doc_id, case, why):
    db.execute("UPDATE documents SET case_number=?, notes=COALESCE(notes,'')||? WHERE doc_id=? AND run_id=?", (case, f"; case {why} {case} ({now})", doc_id, run))
    db.execute("UPDATE targets SET case_number=?, flags=COALESCE(flags,'')||' CASE_FROM_PORTAL' WHERE doc_id=? AND run_id=? AND COALESCE(flags,'') NOT LIKE '%CASE_FROM_PORTAL%'", (case, doc_id, run))
    db.execute("UPDATE targets SET case_number=? WHERE doc_id=? AND run_id=?", (case, doc_id, run))
    db.execute("UPDATE claims SET case_number=? WHERE doc_id=? AND run_id=?", (case, doc_id, run))
    db.execute("INSERT OR IGNORE INTO debtors(case_number, updated_at) VALUES(?,?)", (case, now))


# 1. doc_debtor '' masks the hint
n = db.execute("UPDATE doc_debtor SET debtor=NULL WHERE debtor=''").rowcount
log(f"doc_debtor '' → NULL: {n}")

# 2. documents without a state-court case number, whose portal record carries exactly one
def slug_name(page_url):
    m = re.search(r"/(?:recuperacao-judicial|falencia|recuperacao-extrajudicial)_([^/]+?)__\d+$", page_url or "")
    if not m: return None
    s = m.group(1).replace("+", ".").replace("-", " ")
    s = re.sub(r"\s*\(\s*", " (", s); s = re.sub(r"\s+", " ", s).strip()
    return s.upper()
fixed = 0
for doc_id, url, page_url, nums, cur in db.execute("""SELECT d.doc_id, d.source_url, r.page_url, r.case_numbers, d.case_number FROM documents d JOIN raw_documents r ON r.url=d.source_url
        WHERE d.run_id=? AND (d.case_number IS NULL OR substr(d.case_number,17,1)<>'8')""", (run,)).fetchall():
    cands = sorted({c for c in CNJ.findall(nums or "") if int(c[11:15]) >= 2005})
    name = slug_name(page_url)
    if len(cands) == 1 and cands[0] != cur:
        case = cands[0]
        set_case(doc_id, case, "from portal record" if cur is None else f"replaced non-state-court number {cur} by portal record")
        if name:
            db.execute("UPDATE debtors SET debtor_name=COALESCE(debtor_name, ?) WHERE case_number=?", (name, case))
        fixed += 1; log(f"  {doc_id[:8]} {cur or '(none)'} → {case} | {name}")
    elif name and cur is None:
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

# 6. grant decisions mis-typed by the text classifier (the ADMINISTRADOR + RELAÇÃO rule fires before HOMOLOG)
GRANTS = {"295719aa": "Homologação do plano (decisão de 11/11/2024, condição resolutiva) — Frigorífico Alfa / CTX",
          "aa45ffaa": "Homologação definitiva (sentença de 10/03/2025) — Frigorífico Alfa / CTX",
          "39382ddb": "Sentença de homologação — WWS Services / Worldwide", "11c1d739": "Sentença homologatória do plano — Magazine S.A.",
          "41be4de3": "Homologação do plano de recuperação judicial", "c126258d": "Sentença de homologação — Meqso Distribuição S.A. e outros",
          "f3652e36": "Edital: aviso aos credores da homologação do PRJ (DJE TJRJ)", "d8a79937": "Edital de intimação: concessão da RJ (HSM, publ. 29/06/26)"}
for pfx, why in GRANTS.items():
    n = db.execute("UPDATE documents SET doc_type='HOMOLOG', notes=COALESCE(notes,'')||? WHERE run_id=? AND doc_id LIKE ? AND doc_type<>'HOMOLOG'",
                   (f"; doc_type→HOMOLOG: {why} ({now})", run, pfx + "%")).rowcount
    if n: log(f"  HOMOLOG: {pfx} {why[:60]}")

# 7. proceedings read from the AJ portal (not visible to DataJud's RJ-class pull)
job = "5000561-63.2025.8.24.0536"
for pfx in ("8b4c169d", "2fc83c38"):
    for (doc_id, cur) in db.execute("SELECT doc_id, case_number FROM documents WHERE run_id=? AND doc_id LIKE ?", (run, pfx + "%")).fetchall():
        if cur != job: set_case(doc_id, job, "from AJ portal page (scalzilliaj.com.br, JOB Fertilizantes)"); log(f"  {pfx} → {job}")
db.execute("INSERT OR IGNORE INTO debtors(case_number, updated_at) VALUES(?,?)", (job, now))
db.execute("""UPDATE debtors SET debtor_name=COALESCE(NULLIF(debtor_name,'SCZ Scalzilli Administração Judicial'), 'JOB FERTILIZANTES LTDA'), proceeding='EXTINCT',
              verified_evidence=COALESCE(verified_evidence, ?) WHERE case_number=?""",
           ("AJ portal (scalzilliaj.com.br, fetched 2026-09-04): RJ petition 24/07/2025, deferred 21/08/2025; docs 11-12 'Sentença – Extinção por desistência' 04/08/2026 and 'Edital – Homologação do pedido de desistência'", job))
agimax = "5009529-81.2025.8.24.0019"
db.execute("INSERT OR IGNORE INTO debtors(case_number, updated_at) VALUES(?,?)", (agimax, now))
db.execute("""UPDATE debtors SET proceeding='FALENCIA', debtor_name=COALESCE(NULLIF(debtor_name,'SCZ Scalzilli Administração Judicial'), 'MASSA FALIDA DE AGIMAX ESQUADRIAS METALICAS LTDA'),
              verified_evidence=COALESCE(verified_evidence, ?) WHERE case_number=?""",
           ("AJ portal category 'falencia' (scalzilliaj.com.br, fetched 2026-09-04): 'Massa Falida de Agimax Esquadrias Metálicas LTDA', falência decreed 13/01/2026, Vara Regional de Falências de Concórdia/SC; list headings use art. 83 classes", agimax))
db.execute("UPDATE debtors SET proceeding='FALENCIA' WHERE case_number='0003232-89.2024.8.16.0185'")
log("proceedings set from the AJ portal: JOB (EXTINCT), Agimax (FALENCIA), Visão (FALENCIA, encerrada)")

# 8. case notes the closer must see
NOTES = [("5014104-85.2021.8.21.0010", "eproc header prints 'RECUPERAÇÃO EXTRAJUDICIAL', the edital body invokes arts. 7 §2 and 55 (RJ-only) — confirm the case class in eproc before quoting", "verification reader, p1 of 84c58636"),
         ("0003232-89.2024.8.16.0185", "falência ENCERRADA by sentença of 01/08/2025 (art. 156 edital on file); debtor BAIXADA at RFB — nothing left to habilitate", "bbsadvogados.com.br edital, 7b043ed69346e1b5.pdf"),
         ("0005010-50.2013.8.24.0026", "falência label rests on the AJ portal category and the file name 'relacao-de-credores-da-falida.pdf'; the list itself is headed 'recuperanda' — confirm on the docket", "verification critic"),
         ("5003439-20.2020.8.24.0282", "falência label rests on the AJ portal category — confirm on the docket", "verification critic"),
         ("0873061-47.2023.8.19.0001", "list prints two 'Valor' columns per creditor (updates of 05/06/2023 and 18/07/2025); the face is the 18/07/2025 figure and the list uses art. 83 class labels", "verification spot-check, 8eb5d0d64353bf5f.pdf p1")]
for case, note, src in NOTES:
    db.execute("INSERT OR REPLACE INTO case_notes VALUES(?,?,?,?)", (case, note, src, now))
log(f"case_notes rows: {db.execute('SELECT COUNT(*) FROM case_notes').fetchone()[0]}")
db.commit()

# 9. RFB enrichment for above-floor creditor roots with no companies row (registry phone/e-mail, porte, QSA partners)
if "--no-enrich" not in sys.argv:
    from anteparo.steps.build import enrich
    missing = db.execute("SELECT COUNT(DISTINCT cnpj_basico) FROM targets WHERE run_id=? AND band='FLOOR' AND cnpj_basico NOT IN (SELECT cnpj_basico FROM companies)", (run,)).fetchone()[0]
    log(f"creditor roots above floor without a companies row: {missing}")
    if missing:
        db.commit()
        enrich(db, run, DB_PATH, log, include_by_name=True)
        db.commit()
log("FIX_AUDIT DONE")
