"""Load Apollo picks (data/apollo/results.jsonl) and revealed people (data/apollo/matched.jsonl) into apollo_contacts.

rank 1 = decision maker, rank 2 = backup. Only people whose full record was revealed by bulk_match are loaded;
a pick that was never matched has only a masked surname and is skipped.
"""
import json, re, sys, time, unicodedata
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
    line = line.strip()
    if line:
        rec = json.loads(line); matched[rec["id"]] = rec
# merge every row for a root (a company may have been searched twice: keyword pass + finance-title pass); dedupe picks by id
byroot, rownote = {}, {}
for line in open(ROOT / "data/apollo/results.jsonl", encoding="utf-8"):
    line = line.strip()
    if not line: continue
    row = json.loads(line)
    roots = [row.get("root")] + list(row.get("also_roots") or [])
    for root in roots:
        if not root: continue
        seen = {p["id"] for p in byroot.get(root, [])}
        for p in row.get("picks") or []:
            if p.get("id") and p["id"] not in seen:
                byroot.setdefault(root, []).append(dict(p)); seen.add(p["id"])
        if row.get("note") and ("VERIFY" in row["note"].upper() or "LARGE CORP" in row["note"].upper()):
            rownote[root] = (rownote.get(root, "") + "; " if rownote.get(root) else "") + row["note"]
company = {r[0]: (r[1] or "", r[2] or "", r[3] or "") for r in db.execute("SELECT cnpj_basico, razao_social, nome_fantasia, uf FROM companies")}
STOP = {"LTDA", "SA", "S", "A", "ME", "EPP", "EIRELI", "DE", "DO", "DA", "DOS", "DAS", "E", "EM", "RECUPERACAO", "JUDICIAL", "INDUSTRIA", "COMERCIO", "SERVICOS", "BRASIL", "DO", "GRUPO", "CIA", "COMPANHIA", "BANCO", "FUNDO", "INVESTIMENTO", "DIREITOS", "CREDITORIOS", "FIDC", "SOCIEDADE", "CREDITO", "COOPERATIVA", "PARTICIPACOES", "ADMINISTRACAO", "DISTRIBUIDORA", "IMPORTACAO", "EXPORTACAO", "PRODUTOS", "ALIMENTOS", "AGRICOLA", "AGROPECUARIA", "CONSTRUTORA", "ENGENHARIA", "TRANSPORTES", "TEXTIL", "THE", "OF", "AND", "GROUP", "INC", "LLC", "CO", "CORP", "BR"}
def toks(s):
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().upper()
    return {t for t in re.findall(r"[A-Z0-9]{3,}", s) if t not in STOP}
n = skipped = 0
for root, picks in byroot.items():
    db.execute("DELETE FROM apollo_contacts WHERE cnpj_basico=?", (root,))
    razao, fantasia, uf = company.get(root, ("", "", ""))
    ctoks = toks(razao) | toks(fantasia)
    rank = 0
    for p in picks:
        m = matched.get(p["id"])
        if not m:
            skipped += 1; continue
        rank += 1
        phones = dict((t, num) for t, num in (m.get("phones") or []) if num)
        mobile = phones.get("mobile"); direct = phones.get("work_direct") or phones.get("direct") or phones.get("other")
        email, status, note = m.get("email"), m.get("email_status"), p.get("note") or ""
        if rank == 1 and rownote.get(root): note = (note + "; " if note else "") + rownote[root]
        # org sanity: no shared token between the Apollo org (name/domain) and the creditor's registry names → say so
        otoks = toks(m.get("org")) | toks((m.get("org_domain") or "").split(".")[0])
        if ctoks and otoks and not (ctoks & otoks) and "VERIFY" not in note.upper():
            note = (note + "; " if note else "") + f"VERIFY org: Apollo has this person at '{m.get('org')}'"
        dom = (email or "").rsplit("@", 1)[-1].lower()
        if email and dom in BAD_DOMAINS:
            note = (note + "; " if note else "") + f"email dropped: {BAD_DOMAINS[dom]}"; email = None; status = "dropped"
        db.execute("INSERT OR REPLACE INTO apollo_contacts VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                   (root, rank, m["id"], m.get("name"), m.get("title") or p.get("title"), email, status,
                    mobile, direct, m.get("org_phone"), m.get("linkedin"), m.get("org"), m.get("city"), m.get("state"),
                    note or None, "apollo", time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())))
        n += 1
db.commit()
print(f"apollo_contacts: {n} rows loaded, {skipped} picks not yet revealed; companies with a contact: "
      f"{db.execute('SELECT COUNT(DISTINCT cnpj_basico) FROM apollo_contacts').fetchone()[0]}; with email: "
      f"{db.execute('SELECT COUNT(DISTINCT cnpj_basico) FROM apollo_contacts WHERE email IS NOT NULL').fetchone()[0]}")
