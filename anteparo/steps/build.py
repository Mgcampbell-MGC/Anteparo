"""Steps E–H: clean, aggregate to targets, enrich from RFB, find decision-makers."""
from __future__ import annotations

import json
from collections import defaultdict
from decimal import Decimal

from ..cnpj import is_valid_cnpj, matriz, root as cnpj_root
from ..db import now
from ..sources.rfb import RFB, normalise

FLOOR = Decimal(200000)
POOL_MIN = Decimal(100000)
TYPE_RANK = {"QGC": 4, "AJ_LIST": 3, "DEBTOR_LIST": 2, "UNKNOWN": 1, "HOMOLOG": 0, "PLAN": 0}
STATUS_RANK = {"OK": 3, "PARTIAL_DOC": 2, "OK_NO_TOTALS": 1}
DM_CODES = {49: "Sócio-Administrador", 5: "Administrador", 10: "Diretor", 16: "Presidente"}


def best_document_per_case(db, run_id):
    """Rank usable creditor-list documents per case: QGC > AJ_LIST > DEBTOR_LIST, latest, most rows."""
    docs = db.execute("""SELECT d.doc_id, d.case_number, d.doc_type, d.status, d.publication_date,
                                (SELECT COUNT(*) FROM claims c WHERE c.doc_id=d.doc_id AND c.run_id=d.run_id) AS n, d.notes
                         FROM documents d WHERE d.run_id=? AND (d.status IN ('OK','OK_NO_TOTALS')
                              OR (d.status='QUARANTINED' AND d.notes LIKE '%classIII=PASS%'))
                         AND d.doc_type IN ('QGC','AJ_LIST','DEBTOR_LIST','UNKNOWN')""", (run_id,)).fetchall()
    best = {}
    for doc_id, case, dtype, status, pub, n, notes in docs:
        key = case or ("UNKNOWN:" + doc_id[:10])
        if status == "QUARANTINED":
            status = "PARTIAL_DOC"
        score = (TYPE_RANK.get(dtype, 0), STATUS_RANK.get(status, 0), pub or "", n)
        if key not in best or score > best[key][0]:
            best[key] = (score, doc_id, dtype, status)
    return {k: v[1:] for k, v in best.items()}


def build_targets(db, run_id, log=print):
    best = best_document_per_case(db, run_id)
    db.execute("UPDATE targets SET superseded_by=? WHERE superseded_by IS NULL AND run_id<>?", (run_id, run_id))
    db.execute("DELETE FROM targets WHERE run_id=?", (run_id,))   # idempotent rebuild within the same run
    n_targets, n_pool, n_below, seen_dupes = 0, 0, 0, 0
    debtor_by_case = dict(db.execute("SELECT case_number, debtor_name_hint FROM documents WHERE run_id=? AND case_number IS NOT NULL", (run_id,)).fetchall())
    stage_by_case = dict(db.execute("SELECT case_number, stage FROM cases").fetchall())
    for case_key, (doc_id, dtype, status) in best.items():
        rows = db.execute("""SELECT id, document_number, all_documents, value_brl, class, currency, flags, creditor_name_as_printed, page, row_index
                             FROM claims WHERE doc_id=? AND run_id=?""", (doc_id, run_id)).fetchall()
        agg = defaultdict(lambda: {"sum": Decimal(0), "ests": set(), "n": 0, "names": [], "pages": [], "flags": set(), "ids": []})
        dedupe = set()
        for cid, dnum, alldocs, val, klass, cur, flags, name, page, ridx in rows:
            fl = set((flags or "").split("|")) - {""}
            if klass != "III" or cur != "BRL" or val is None:
                continue
            if fl & {"NOT_A_CLAIM", "INDIVIDUAL", "CNPJ_INVALID", "VALUE_UNPARSEABLE"} or any(f.startswith("SECTION_") for f in fl):
                continue
            if not dnum or len(dnum) != 14 or not is_valid_cnpj(dnum):
                continue
            key = (dnum, val, klass)
            if key in dedupe:
                seen_dupes += 1
                continue
            dedupe.add(key)
            r = cnpj_root(dnum)
            a = agg[r]
            a["sum"] += Decimal(val)
            a["ests"].update(d for d in (alldocs or dnum).split("|") if len(d) == 14)
            a["n"] += 1
            a["names"].append(name or "")
            a["pages"].append(f"p{page}r{ridx}")
            a["flags"] |= fl
            a["ids"].append(cid)
        for r, a in agg.items():
            if a["sum"] >= FLOOR:
                band, above = "FLOOR", 1
                n_targets += 1
            elif a["sum"] >= POOL_MIN:
                band, above = "POOL", 0
                n_pool += 1
            else:
                n_below += 1
                continue
            fl = set(a["flags"])
            if status == "OK_NO_TOTALS":
                fl.add("UNRECONCILED_SOURCE")
            if status == "PARTIAL_DOC":
                fl.add("PARTIAL_DOC_CLASS_III_RECONCILED")
            if case_key.startswith("UNKNOWN:"):
                fl.add("CASE_NUMBER_MISSING")
            case_number = None if case_key.startswith("UNKNOWN:") else case_key
            db.execute("""INSERT INTO targets(run_id,cnpj_basico,case_number,doc_id,class_iii_face_sum,establishment_cnpjs,claim_count,
                          creditor_name_as_printed,debtor_name,stage,is_related_party,above_floor,band,flags,extracted_at,extracted_by)
                          VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                       (run_id, r, case_number, doc_id, str(a["sum"]), "|".join(sorted(a["ests"])), a["n"],
                        max(a["names"], key=len), debtor_by_case.get(case_number), stage_by_case.get(case_number),
                        None, above, band, "|".join(sorted(fl | {"ROWS:" + ",".join(a["pages"])})), now(), "build"))
    db.commit()
    log(f"targets: {n_targets} at/above R$200k · {n_pool} pooling (100–200k) · {n_below} below · {seen_dupes} exact dupes skipped · {len(best)} cases")
    return n_targets, n_pool


def enrich(db, run_id, db_path, log=print, limit=None):
    rfb = RFB(db_path)
    roots = [r[0] for r in db.execute("""SELECT DISTINCT cnpj_basico FROM targets WHERE run_id=? AND cnpj_basico NOT IN
                                         (SELECT cnpj_basico FROM companies) ORDER BY CAST(class_iii_face_sum AS REAL) DESC""", (run_id,)).fetchall()]
    if limit:
        roots = roots[:limit]
    log(f"enrich: {len(roots)} companies to look up (throttled)")
    hit, miss = 0, 0
    for i, r in enumerate(roots, 1):
        ests = db.execute("SELECT establishment_cnpjs FROM targets WHERE run_id=? AND cnpj_basico=? LIMIT 1", (run_id, r)).fetchone()[0]
        tried = [matriz(r)] + [e for e in (ests or "").split("|") if e and e != matriz(r)]
        rec = None
        for c in tried:
            d = rfb.get(c)
            if d:
                rec = normalise(d)
                break
        if not rec:
            miss += 1
            db.execute("INSERT OR REPLACE INTO companies(cnpj_basico, cnpj_matriz, rfb_source, extracted_at, extracted_by, is_bank, is_public, is_inactive) VALUES(?,?,?,?,?,?,?,?)",
                       (r, matriz(r), "NOT_IN_RFB", now(), "enrich", None, None, None))
            db.commit()
            continue
        hit += 1
        nat = rec["natureza_juridica"] or ""
        cnae = rec["cnae_principal"] or ""
        is_bank = int(cnae[:2] in ("64", "65", "66"))
        is_public = int(nat.startswith("1"))
        is_inactive = int(rec["situacao_cadastral"] != "02")
        db.execute("""INSERT OR REPLACE INTO companies VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                   (r, rec["cnpj"], rec["razao_social"], rec["nome_fantasia"], nat, rec["natureza_juridica_desc"], rec["porte"],
                    rec["situacao_cadastral"], rec["situacao_desc"], rec["data_situacao"], cnae, rec["cnae_desc"], rec["uf"],
                    rec["municipio"], rec["phone"], rec["phone2"], rec["email"], str(rec["capital_social"] or ""), rec["inicio_atividade"],
                    is_bank, is_public, is_inactive, rec["source"], rec["fetched_at"], json.dumps(rec["qsa"], ensure_ascii=False), now(), "enrich"))
        # contacts (step H) — registry only, confidence LOW; LinkedIn/Apollo is a later pass
        db.execute("DELETE FROM contacts WHERE cnpj_basico=? AND source LIKE 'RFB%'", (r,))
        dms = [q for q in rec["qsa"] if q.get("qual_code") in DM_CODES]
        if not dms:
            dms = [q for q in rec["qsa"] if q.get("qual_code") == 22][:1]
        for q in dms:
            db.execute("INSERT INTO contacts(cnpj_basico,person_name,role,role_code,cpf_masked,phone,email,source,confidence,extracted_at,extracted_by) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                       (r, q.get("nome"), q.get("qual") or DM_CODES.get(q.get("qual_code")), q.get("qual_code"), q.get("cpf_masked"),
                        rec["phone"], rec["email"], f"RFB via {rec['source']} ({rec['fetched_at'][:10]})", "LOW", now(), "enrich"))
        db.commit()
        if i % 25 == 0:
            log(f"  enrich {i}/{len(roots)}  hit={hit} miss={miss}")
    log(f"enrich done: hit={hit} miss={miss}")
