"""Debtor identification + enrichment + banding.

1. Debtor CNPJ from each document's first pages, context-scored so the administrator's own CNPJ is never taken,
   and never a CNPJ that is itself a creditor row on the same document.
2. RFB enrichment of the debtor; accepted only when the registry name matches the printed debtor name, or when the
   registry name itself carries the "EM RECUPERAÇÃO JUDICIAL" suffix and the page-title hint was not a company name.
3. Band A–D per case from stage, plan document, registry status, age and dormancy.
4. Verified names: ingest data/state/debtor_read/final.json (two independent readers + adjudicator recorded the
   recuperanda exactly as printed on each PDF head) → display_name / verified_cnpj / group_members / administrator,
   re-enrich verified CNPJs at RFB, then re-band.

Idempotent: re-runs skip cases already resolved in step 1 and only re-enrich what is new.
Usage: debtors.py [run_id] [--reset-display]
"""
from __future__ import annotations
import re, sys, json, time
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(ROOT))
from anteparo.db import connect, now
from anteparo.cnpj import find_documents, is_valid_cnpj
from anteparo.sources.rfb import RFB, normalise
from anteparo.extract.pdfio import open_pdf, page_text
from anteparo.export.callsheet import clean_hint, debtor_display
from rapidfuzz import fuzz

args = [a for a in sys.argv[1:] if not a.startswith("--")]
RUN = args[0] if args else None
db = connect(str(ROOT / "data/anteparo.sqlite"))
run = RUN or db.execute("SELECT run_id FROM runs WHERE note='phase1' ORDER BY started_at DESC LIMIT 1").fetchone()[0]
db.executescript("""
CREATE TABLE IF NOT EXISTS debtors(case_number TEXT PRIMARY KEY, debtor_name TEXT, debtor_cnpj TEXT, cnpj_source TEXT, name_match REAL,
  razao_social TEXT, situacao TEXT, situacao_desc TEXT, porte TEXT, capital_social TEXT, uf TEXT, municipio TEXT, cnae_desc TEXT,
  rfb_source TEXT, band TEXT, band_reasons TEXT, plan_status TEXT, stage TEXT, filing_year INTEGER, last_movement TEXT, updated_at TEXT);
""")
for col, typ in (("display_name", "TEXT"), ("display_source", "TEXT"), ("verified_name", "TEXT"), ("verified_cnpj", "TEXT"), ("group_members", "TEXT"),
                 ("administrator", "TEXT"), ("proceeding", "TEXT"), ("verified_confidence", "REAL"), ("verified_evidence", "TEXT"), ("verified_at", "TEXT")):
    if col not in [r[1] for r in db.execute("PRAGMA table_info(debtors)")]:
        db.execute(f"ALTER TABLE debtors ADD COLUMN {col} {typ}")
db.commit()
POS = re.compile(r"RECUPERANDA|REQUERENTE|AUTOR|DEVEDOR|EM RECUPERA|GRUPO|RECUPERA[ÇC][ÃA]O JUDICIAL D", re.I)
NEG = re.compile(r"ADMINISTRADOR|ADMINISTRA[ÇC][ÃA]O JUDICIAL|\bAJ\b|CONSULTORIA|PER[ÍI]CIA|ADVOGAD|OAB|ESCRIT[ÓO]RIO|SOCIEDADE INDIVIDUAL", re.I)
RJ_SUFFIX = re.compile(r"RECUPERA[CÇ][AÃ]O JUDICIAL", re.I)
log = lambda m: print(time.strftime("%H:%M:%S"), m, flush=True)


# --- 1. debtor name + CNPJ per case (best document per case) ---
def step1():
    cases = {}
    for case, doc_id, path, hint in db.execute("""SELECT d.case_number, d.doc_id, d.file_path, COALESCE(NULLIF(dd.debtor,''), d.debtor_name_hint)
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
        creditor_cnpjs = {r[0] for r in db.execute("SELECT DISTINCT document_number FROM claims WHERE run_id=? AND case_number=?", (run, case))}
        for doc_id, path, hint in docs[:3]:
            try:
                with open_pdf(path) as pdf:
                    head = "\n".join(page_text(p) for p in pdf.pages[:2])
            except Exception:
                continue
            for kind, digits in find_documents(head):
                if kind != "CNPJ" or not is_valid_cnpj(digits) or digits in creditor_cnpjs:
                    continue
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
    log(f"debtor CNPJ found for {found} new cases")


# --- 2. RFB enrichment with name match ---
def enrich(rows, target_col="debtor_cnpj"):
    rfb = RFB(str(ROOT / "data/anteparo.sqlite"))
    ok = rej = miss = 0
    for i, (case, dname, cnpj) in enumerate(rows, 1):
        db.commit()   # release our write lock: rfb.get() writes rfb_cache through its own connection
        d = rfb.get(cnpj)
        if not d:
            miss += 1; db.execute("UPDATE debtors SET rfb_source='NOT_IN_RFB', updated_at=? WHERE case_number=?", (now(), case)); continue
        n = normalise(d)
        hint = clean_hint(dname)
        match = fuzz.token_set_ratio((hint or dname or "").upper(), (n["razao_social"] or "").upper()) / 100 if (hint or dname) else None
        accept = match is None or match >= 0.55 or (not hint and RJ_SUFFIX.search(n["razao_social"] or ""))
        if not accept:
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
    granted = db.execute("SELECT COUNT(*) FROM documents WHERE run_id=? AND case_number=? AND doc_type='HOMOLOG'", (run, case)).fetchone()[0]
    plan_filed = db.execute("SELECT COUNT(*) FROM documents WHERE run_id=? AND case_number=? AND doc_type='PLAN'", (run, case)).fetchone()[0]
    rj_granted = (db.execute("SELECT rj_granted_signal FROM cases WHERE case_number=?", (case,)).fetchone() or (0,))[0]
    plan = granted or rj_granted   # band A needs the GRANT (art. 58); a filed plan (art. 53) is band B
    sit = db.execute("SELECT situacao, rfb_source, proceeding, COALESCE(display_name, razao_social, debtor_name) FROM debtors WHERE case_number=?", (case,)).fetchone() or (None, None, None, None)
    age = (2026 - year) if year else None
    dormant = bool(last_mv) and last_mv < "2025-09-01"
    accepted = sit[1] and sit[1] not in ("NOT_IN_RFB", "NAME_MISMATCH", "CREDITOR_CNPJ")
    falida = bool(re.search(r"MASSA FALIDA|\bFALID[AO]\b", sit[3] or "", re.I))
    if stage == "CONVERTED_TO_BANKRUPTCY" or sit[2] == "FALENCIA" or falida:
        reasons.append("falência decreed (DataJud)" if stage == "CONVERTED_TO_BANKRUPTCY" else "falência proceeding (list/name says falida)"); band = "D"
        if year and year < 2005: reasons.append("pre-Lei 11.101 proceeding (DL 7.661/45)")
    elif accepted and sit[0] == "08":
        reasons.append("debtor closed at RFB (baixada)"); band = "D"
    elif plan and stage != "CLOSED":
        reasons.append("recovery granted (art. 58): grant document or DataJud grant movement"); band = "A"
        if age is not None and age > 8: reasons.append(f"but case is {age}y old"); band = "B"
    elif (age is not None and age > 4) or dormant or stage == "CLOSED" or (accepted and sit[0] in ("03", "04")):
        band = "C"
        if age is not None and age > 4: reasons.append(f"{age}y since filing; plan not in our documents")
        if dormant: reasons.append(f"no movement since {last_mv}")
        if stage == "CLOSED": reasons.append("case closed/archived (DataJud)")
        if accepted and sit[0] in ("03", "04"): reasons.append(f"debtor {'SUSPENSA' if sit[0]=='03' else 'INAPTA'} at RFB")
    else:
        band = "B"; reasons.append("plan filed (art. 53), not yet granted" if plan_filed else "live case, no plan document yet")
    if stage is None and not m: reasons.append("case not matched")
    plan_status = "CONFIRMED" if plan else ("PLAN_FILED" if plan_filed else ("ESTIMATED" if stage else "UNKNOWN"))
    return band, "; ".join(reasons), plan_status, stage, year, last_mv


def step3():
    n = 0
    for (case,) in db.execute("SELECT case_number FROM debtors").fetchall():
        b, r, ps, st, yr, lm = band_for(case)
        db.execute("UPDATE debtors SET band=?, band_reasons=?, plan_status=?, stage=?, filing_year=?, last_movement=?, updated_at=? WHERE case_number=?", (b, r, ps, st, yr, lm, now(), case)); n += 1
    db.commit()
    log(f"bands: {db.execute('SELECT band, COUNT(*) FROM debtors GROUP BY band').fetchall()}")


# --- 4. verified names from the PDF readers + display name for every case ---
def step4():
    fin = ROOT / "data/state/debtor_read/final.json"
    if fin.exists():
        data = json.load(open(fin, encoding="utf-8"))
        rows = data["final"] if isinstance(data, dict) else data
        n_ok = 0
        for c in rows:
            name = (c.get("debtor_name_as_printed") or "").strip()
            if not name or (c.get("confidence") or 0) < 0.6:
                continue
            cnpj = re.sub(r"\D", "", c.get("debtor_cnpj") or "")
            db.execute("""UPDATE debtors SET verified_name=?, verified_cnpj=?, group_members=?, administrator=?, proceeding=?, verified_confidence=?, verified_evidence=?, verified_at=?
                          WHERE case_number=?""", (name, cnpj if len(cnpj) == 14 else None, json.dumps(c.get("group_members") or [], ensure_ascii=False),
                                                   c.get("administrator") or None, c.get("proceeding") or None, c.get("confidence"), (c.get("evidence_quote") or "")[:400], now(), c["case_number"]))
            db.execute("INSERT OR IGNORE INTO debtors(case_number, debtor_name, updated_at) VALUES(?,?,?)", (c["case_number"], name, now()))
            n_ok += 1
        db.commit()
        log(f"verified names ingested: {n_ok}")
        # a verified CNPJ that differs from the extracted one replaces it and is enriched
        todo = db.execute("""SELECT case_number, verified_name, verified_cnpj FROM debtors WHERE verified_cnpj IS NOT NULL AND (debtor_cnpj IS NULL OR debtor_cnpj<>verified_cnpj)""").fetchall()
        for case, vname, vcnpj in todo:
            db.execute("UPDATE debtors SET debtor_cnpj=?, cnpj_source='reader', rfb_source=NULL, razao_social=NULL, situacao=NULL, name_match=NULL WHERE case_number=?", (vcnpj, case))
        db.commit()
        if todo:
            enrich([(c, v, k) for c, v, k in todo])
    # display name for every case
    n = 0
    for case, vname, razao, src, hint, gm in db.execute("SELECT case_number, verified_name, razao_social, rfb_source, debtor_name, group_members FROM debtors").fetchall():
        hints = [h for (h,) in db.execute("SELECT COALESCE(NULLIF(dd.debtor,''), d.debtor_name_hint) FROM documents d LEFT JOIN doc_debtor dd ON dd.doc_id=d.doc_id WHERE d.case_number=? AND d.run_id=?", (case, run))]
        hints += [t for (t,) in db.execute("SELECT r.page_title FROM raw_documents r JOIN documents d ON d.source_url=r.url WHERE d.case_number=? AND d.run_id=?", (case, run))]
        # the portal slug ('recuperacao-judicial_<name>__id') names the debtor more reliably than the page title
        for (pu,) in db.execute("SELECT r.page_url FROM raw_documents r JOIN documents d ON d.source_url=r.url WHERE d.case_number=? AND d.run_id=?", (case, run)):
            mm = re.search(r"recuperacao-judicial_([^/]+?)__\d+$", pu or "")
            if mm: hints.insert(0, re.sub(r"\s+", " ", mm.group(1).replace("+", ".").replace("-", " ")).strip().upper())
        if not db.execute("SELECT proceeding FROM debtors WHERE case_number=?", (case,)).fetchone()[0]:
            blob = " ".join(h for h in hints if h) + " " + (razao or "")
            if re.search(r"MASSA FALIDA|\bFALID[AO]\b|FAL[ÊE]NCIA D[EAO]", blob, re.I): db.execute("UPDATE debtors SET proceeding='FALENCIA' WHERE case_number=?", (case,))
            elif re.search(r"EXTRAJUDICIAL", blob, re.I): db.execute("UPDATE debtors SET proceeding='RE' WHERE case_number=?", (case,))
        if vname:
            disp, srcname = vname, "pdf-read"
            if " / " in vname and len(vname) > 60:   # reader listed the members; the portal's group name is what a caller says
                g = next((clean_hint(h) for h in [hint] + hints if clean_hint(h) and clean_hint(h).upper().startswith("GRUPO")), None)
                if g: disp = f"{g} ({vname[:70]}…)"
        else:
            disp = debtor_display(None, razao, src, hint, *hints)
            srcname = "rfb" if disp and disp == razao else ("hint" if disp else "")
        db.execute("UPDATE debtors SET display_name=?, display_source=?, updated_at=? WHERE case_number=?", (disp or None, srcname or None, now(), case)); n += 1
    db.commit()
    log(f"display names: {db.execute('SELECT display_source, COUNT(*) FROM debtors GROUP BY 1').fetchall()}")


if __name__ == "__main__":
    if "--reset-display" in sys.argv:
        db.execute("UPDATE debtors SET display_name=NULL, display_source=NULL"); db.commit()
    step1()
    rows = db.execute("SELECT case_number, debtor_name, debtor_cnpj FROM debtors WHERE debtor_cnpj IS NOT NULL AND rfb_source IS NULL").fetchall()
    log(f"enriching {len(rows)} debtors")
    if rows: enrich(rows)
    step4()
    step3()
    q_band = "SELECT COALESCE(d.band, '?'), COUNT(*) FROM targets t LEFT JOIN debtors d ON d.case_number=t.case_number WHERE t.run_id=? AND t.band=? GROUP BY 1"
    log(f"targets by debtor band: {db.execute(q_band, (run, 'FLOOR')).fetchall()}")
    log("DEBTORS DONE")
