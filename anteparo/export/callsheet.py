"""The closer's workbook: DASHBOARD (live CRM formulas over both call tabs) + CALL SHEET + FINANCIAL CREDITORS.

Ranking = fund price (∝ our ágio), not face value. Every quote is an Excel formula off the Assumptions
tab keyed by DEBTOR BAND, so re-pricing a band re-prices every lead in it. Proof-of-claim columns travel
with the row. Assumptions is hidden in the closer file and visible in the ops file.

Rules that protect the reader:
- the DEBTOR column never prints a registry name whose match was rejected (that is how BANCO BRADESCO once
  appeared as a debtor); it prints the verified name, else an accepted registry name, else a cleaned page title;
- rows where the creditor is the debtor itself or a group member are excluded (intercompany paper);
- years-to-payment runs from the plan, not from the filing: an old case with no plan prices LOWER, not higher;
- PROOF QUALITY says whether the printed value was reconciled against the list's own totals.
"""
from __future__ import annotations

import json
import re
from datetime import date
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from rapidfuzz import fuzz

STATUSES = ["NEW", "QUEUED", "EMAILED", "CALLED", "VOICEMAIL", "CONNECTED", "INTERESTED", "NOT INTERESTED", "CALL BACK", "DEAD", "DEAL"]
# band, class III recovery, grace (y), term (y), plan lag (y), meaning
#   plan lag: A = typical filing→plan confirmation delay used to estimate how far into the plan the case is;
#             B/C/?/D = years from TODAY until a plan is confirmed (or, for D, until a liquidation payout)
BANDS = [("A", 0.40, 2, 8, 1.5, "Plan confirmed (art. 58) or plan document found; case live; debtor ATIVA at RFB"),
         ("B", 0.30, 2, 8, 1.5, "Live case, no confirmed plan found yet — the payment clock has not started"),
         ("C", 0.15, 3, 10, 2.5, "Stale: >4 years since filing with no plan, or no movement in 12 months, or case closed"),
         ("D", 0.03, 0, 0, 6.0, "Avoid: falência decreed, or debtor no longer ATIVA at RFB — liquidation payout years away"),
         ("?", 0.25, 3, 8, 2.5, "Debtor or case not resolved — conservative default")]
BAND_MAP = {b: (rec, g, t, lag) for b, rec, g, t, lag, _ in BANDS}
DEFAULT_IRR, DEFAULT_SHARE = 0.25, 0.60
NAVY = "1F3A5F"; HEAD = PatternFill("solid", fgColor=NAVY); WHITE = Font(bold=True, color="FFFFFF")
FILL = {"crm": "FFF4D6", "deal": "E6EEF8", "proof": "E8F1E8", "contact": "F6F0FA", "ctx": "F3F3F3"}
BAND_FILL = {"A": "C6EFCE", "B": "E2EFDA", "C": "FFEB9C", "D": "FFC7CE", "?": "EDEDED"}
PROOF_FILL = {"NO TOTALS": "FFEB9C", "PARTIAL": "F8CBAD"}
TR = {"26": "SP", "19": "RJ", "16": "PR", "21": "RS", "24": "SC", "13": "MG", "09": "GO", "17": "PE", "05": "BA", "11": "MT", "12": "MS",
      "14": "PA", "15": "PB", "08": "ES", "10": "MA", "07": "DF", "06": "CE", "02": "AL", "25": "SE", "27": "TO", "20": "RN", "18": "PI",
      "22": "RO", "01": "AC", "04": "AM", "23": "RR", "03": "AP"}
LIST_KIND = {"AJ_LIST": "AJ list (art. 7 §2)", "DEBTOR_LIST": "Debtor list (art. 51)", "QGC": "QGC (art. 18)", "PLAN": "Plan / homologation",
             "EDITAL": "Edital (art. 7 §2)", "UNKNOWN": "Unclassified list"}

COLS = [  # (header, width, group)
    ("CREDITOR — company owed money", 40, "lead"), ("AMOUNT OWED (R$)", 17, "lead"), ("DEBTOR — company in RJ", 36, "lead"),
    ("STATUS", 15, "crm"), ("LAST TOUCH", 11, "crm"), ("NEXT ACTION", 22, "crm"), ("NEXT DATE", 11, "crm"), ("NOTES", 38, "crm"),
    ("DECISION MAKER", 30, "contact"), ("Role", 24, "contact"), ("Company phone 1", 16, "contact"), ("Company phone 2", 16, "contact"),
    ("Mobile (Apollo)", 16, "contact"), ("Email", 30, "contact"), ("Email 2", 26, "contact"),
    ("BACKUP CONTACT", 28, "contact"), ("Backup role", 22, "contact"), ("Backup phone", 16, "contact"), ("Backup email", 28, "contact"),
    ("Contact note", 34, "contact"),
    ("FACE VALUE (as printed)", 17, "deal"), ("DEBTOR BAND", 8, "deal"), ("Exp. recovery %", 9, "deal"), ("Years to payment", 8, "deal"),
    ("FUND PRICE @ IRR (R$)", 17, "deal"), ("cents", 7, "deal"), ("OUR QUOTE (R$)", 17, "deal"), ("cents", 7, "deal"),
    ("OUR ÁGIO (R$)", 16, "deal"), ("Ágio pts of face", 8, "deal"),
    ("PROOF QUALITY", 26, "proof"), ("PROOF: source document", 50, "proof"), ("Page / row", 16, "proof"), ("Value as printed", 30, "proof"),
    ("Which list", 20, "proof"), ("List date", 11, "proof"), ("Plan evidence", 16, "proof"),
    ("Case number", 26, "ctx"), ("Court", 7, "ctx"), ("Stage", 12, "ctx"), ("Filed", 7, "ctx"), ("Band reason", 40, "ctx"), ("Debtor CNPJ", 19, "ctx"),
    ("City", 18, "ctx"), ("UF", 5, "ctx"), ("Sector", 34, "ctx"), ("Size", 10, "ctx"), ("Capital (R$)", 15, "ctx"), ("Seller fit", 13, "ctx"),
    ("CNPJ root", 10, "ctx"), ("Claims", 6, "ctx"), ("Flags", 30, "ctx"),
]
C = {h: i + 1 for i, (h, _, _) in enumerate(COLS)}          # header → 1-based column (the two 'cents' share a key; resolved by position below)
L = lambda h: get_column_letter(C[h])                          # header → letter


def _phone(p):
    d = "".join(ch for ch in str(p or "") if ch.isdigit())
    if len(d) == 10: return f"({d[:2]}) {d[2:6]}-{d[6:]}"
    if len(d) == 11: return f"({d[:2]}) {d[2:7]}-{d[7:]}"
    if len(d) == 8: return f"{d[:4]}-{d[4:]}"
    if len(d) == 9: return f"{d[:5]}-{d[5:]}"
    return str(p or "")


def _cnpj(d):
    d = "".join(ch for ch in str(d or "") if ch.isdigit())
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}" if len(d) == 14 else (d or "")


def _seller_fit(razao, capital, porte):
    """LARGE CORP = the Petrobras/Hypera tier that sells through a desk or not at all; capital is RFB registered capital in R$."""
    try: cap = float(capital or 0)
    except ValueError: cap = 0
    if cap >= 1_000_000_000 or (re.search(r"\bS\.?A\.?\s*$|\bS/A\s*$|\bS\.A\.\b", razao or "") and cap >= 300_000_000):
        return "LARGE CORP"
    if porte and ("MICRO" in porte.upper() or "PEQUENO" in porte.upper()):
        return "SMALL"
    return "MID-MARKET"


_ADMIN = re.compile(r"administra[çc][ãa]o judicial|administradora judicial|administrador judicial|administra[çc][õo]es judiciais|\bAJ\b|advogad|consultoria|"
                    r"escrit[óo]rio|per[íi]cia|l[íi]deres em recupera|recupera[çc][õo]es judiciais|p[áa]gina inicial|\bbase\s*\||^ativos$|^base$", re.I)
_PREFIX = re.compile(r"^\s*(?:recupera[çc][ãa]o judicial (?:de|da|do|das|dos)?|rj (?:de|da|do)?|massa falida (?:de|da|do)?|fal[êe]ncia (?:de|da|do)?|"
                     r"processo(?: n[º°.]*)?\s*[\d.\-/]+\s*[-–—:]?|e fal[êe]ncia|aludida empresa|empresa|promovida por)\s*", re.I)
_SUFFIX_RJ = re.compile(r"\s*[-–—(]?\s*em recupera[çc][ãa]o judicial\)?\s*$", re.I)
_CASE = re.compile(r"\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}")
_SPLIT = re.compile(r"\s+[–—|]\s+|\s+-\s+")


def clean_hint(h):
    """Portal page title / debtor hint → a company name, or '' when it is not one (administrator's site name, case number, prose)."""
    if not h: return ""
    h = re.sub(r"\s+", " ", str(h)).strip()
    keep = [p for p in _SPLIT.split(h) if p and not _ADMIN.search(p) and not _CASE.fullmatch(p.strip())]
    if not keep: return ""
    h = _SUFFIX_RJ.sub("", _PREFIX.sub("", keep[0])).strip(" -–—:|,")
    if len(h) < 4 or _ADMIN.search(h) or (_CASE.search(h) and len(h) < 40): return ""
    prose = len(re.findall(r"\b(litiscons[óo]rcio|conforme|entendimento|requerimento|foi|apresentad[oa]|autos|conson[âa]ncia|movimento|supramencionad[oa]s?|"
                           r"referid[oa]|aludid[oa]|mediante|sendo|seja|cuj[oa]|qual|pelo|pela|nos|nas)\b", h, re.I))
    if prose >= 1 and not re.search(r"\b(LTDA|S\.?A\.?|S/A|EIRELI|EPP|ME|GRUPO)\b", h, re.I): return ""
    if prose >= 2: return ""
    if not re.search(r"\b[A-ZÀ-Ý][A-ZÀ-Ýa-zà-ÿ&'.]{2,}", h): return ""   # no capitalised token → not a company name
    return h


def debtor_display(display, razao, src, *hints):
    """What the closer sees as the debtor. Never a registry name whose match was rejected."""
    if display: return display
    if razao and src not in (None, "NAME_MISMATCH", "NOT_IN_RFB"): return razao
    if razao and src == "NAME_MISMATCH" and re.search(r"RECUPERA[CÇ][AÃ]O JUDICIAL", razao, re.I):
        return razao   # the registry itself carries the RJ suffix: it is the recuperanda even though the page title did not match
    for h in hints:
        c = clean_hint(h)
        if c: return c
    return ""


def _years(band, ysf):
    rec, g, term, lag = BAND_MAP.get(band, BAND_MAP["?"])
    if band == "A":
        return max(1.0, g + term / 2 - max(0.0, ysf - lag))
    return lag + g + term / 2


def _related(root, razao, name_printed, debtor_root, debtor_names):
    if debtor_root and root == debtor_root: return "same CNPJ as debtor"
    for nm in debtor_names:
        for cand in (razao, name_printed):
            if nm and cand and fuzz.token_set_ratio(nm.upper(), cand.upper()) >= 90 and len(nm) > 8:
                return f"creditor name = debtor/group member ({nm[:40]})"
    return ""


def _leads(db, run_id):
    q = db.execute("""
        SELECT t.cnpj_basico, t.case_number, t.doc_id, t.class_iii_face_sum, t.establishment_cnpjs, t.claim_count, t.creditor_name_as_printed, t.flags,
               c.razao_social, c.phone, c.phone2, c.email, c.uf, c.municipio, c.cnae_desc, c.porte, c.capital_social, c.is_bank, c.is_public, c.is_inactive,
               d.source_url, d.doc_type, d.publication_date, d.status, d.file_path, d.debtor_name_hint,
               cs.court, cs.stage, cs.filing_date,
               db2.debtor_name, db2.razao_social, db2.band, db2.band_reasons, db2.plan_status, db2.filing_year, db2.stage, db2.rfb_source, db2.debtor_cnpj,
               db2.display_name, db2.verified_cnpj, db2.group_members, dd.debtor
        FROM targets t
        LEFT JOIN companies c ON c.cnpj_basico=t.cnpj_basico
        LEFT JOIN documents d ON d.doc_id=t.doc_id AND d.run_id=t.run_id
        LEFT JOIN cases cs ON cs.case_number=t.case_number
        LEFT JOIN debtors db2 ON db2.case_number=t.case_number
        LEFT JOIN doc_debtor dd ON dd.doc_id=t.doc_id
        WHERE t.run_id=? AND t.band='FLOOR'""", (run_id,)).fetchall()
    out, related = [], []
    for r in q:
        (root, case, doc_id, face, ests, ncl, name_printed, flags, razao, phone, phone2, email, uf, city, cnae, porte, capital,
         is_bank, is_public, is_inactive, url, dtype, pub, dstatus, fpath, dhint, court, stage, filed,
         dname, drazao, band, breason, pstatus, fyear, dstage, dsrc, dcnpj, ddisplay, vcnpj, gmembers, dd_debtor) = r
        if is_public or is_inactive or "LIKELY_PUBLIC" in (flags or ""):
            continue
        fin = bool(is_bank) or "LIKELY_FINANCIAL" in (flags or "")   # banks / FIDCs / securitizers / credit co-ops → their own tab
        if not case or case[16:17] != "8":   # RJ cases live in state courts (J=8); anything else is a creditor-side number picked off the page
            case = ""; band = "?"; breason = "RJ case number not identified on the document"; stage = stage or ""
        band = band if band in BAND_MAP else "?"
        debtor = debtor_display(ddisplay, drazao, dsrc, dd_debtor, dname, dhint)
        accepted_cnpj = dcnpj if dsrc not in (None, "NAME_MISMATCH", "NOT_IN_RFB") else ""   # a rejected registry match is not the debtor's CNPJ
        debtor_root = (vcnpj or accepted_cnpj or "")[:8]
        members = [m for m in (json.loads(gmembers) if gmembers else [])] if gmembers else []
        rel = _related(root, razao, name_printed, debtor_root, [debtor] + members)
        face_d = Decimal(face)
        year = fyear or (int(filed[:4]) if filed else None)
        if not year:
            m = re.match(r"\d{7}-\d{2}\.(\d{4})\.", case or ""); year = int(m.group(1)) if m else None
        ysf = (date.today().year - year) if year else 0
        rec = BAND_MAP[band][0]
        fund = float(face_d) * rec / (1 + DEFAULT_IRR) ** _years(band, ysf)
        # contacts: Apollo first (ranked), then RFB partners (QSA)
        ap = db.execute("SELECT person_name, title, email, mobile, direct_phone, email_status, pick_note FROM apollo_contacts WHERE cnpj_basico=? ORDER BY rank", (root,)).fetchall()
        rf = db.execute("SELECT person_name, role FROM contacts WHERE cnpj_basico=? ORDER BY CASE role_code WHEN 49 THEN 0 WHEN 5 THEN 1 WHEN 16 THEN 2 WHEN 10 THEN 3 ELSE 9 END", (root,)).fetchall()
        primary = {"name": rf[0][0] if rf else "", "role": (rf[0][1] + " (RFB partner)") if rf else "", "mobile": "", "email": "", "email2": ""}
        notes = []
        if ap:
            a = ap[0]; primary.update({"name": a[0] or primary["name"], "role": a[1] or primary["role"], "mobile": a[3] or "", "email": a[2] or ""})
            if len(ap) > 1 and ap[1][2]: primary["email2"] = ap[1][2]
            if a[6] and "VERIFY" in a[6].upper(): notes.append(a[6])
            if a[2] and a[5] == "extrapolated": notes.append("email is pattern-guessed, not verified")
            elif a[2] and a[5] == "verified": notes.append("email verified by Apollo")
            elif not a[2]: notes.append("no work email found in Apollo")
        elif rf:
            notes.append("RFB partner (QSA); no Apollo match")
        else:
            notes.append("no named contact found")
        backup = {"name": "", "role": "", "phone": "", "email": ""}
        if len(ap) > 1:
            backup = {"name": ap[1][0] or "", "role": ap[1][1] or "", "phone": ap[1][3] or ap[1][4] or "", "email": ap[1][2] or ""}
        elif len(rf) > 1:
            backup = {"name": rf[1][0], "role": rf[1][1] + " (RFB partner)", "phone": "", "email": ""}
        if backup["phone"] and _phone(backup["phone"]) in (_phone(phone), _phone(phone2)):
            backup["phone"] = ""   # same switchboard as the company phones: do not pretend it is a second route in
        claims = db.execute("SELECT page, row_index, value_as_printed FROM claims WHERE doc_id=? AND run_id=? AND class='III' AND currency='BRL' AND document_number IN (%s) ORDER BY page,row_index"
                            % ",".join("?" * len((ests or "").split("|"))), (doc_id, run_id, *((ests or "").split("|")))).fetchall()
        proof = {"OK": "OK — totals reconciled", "OK_NO_TOTALS": "NO TOTALS — list prints none; check page", "QUARANTINED": "PARTIAL — class III reconciled, other classes not"}.get(dstatus, dstatus or "")
        court_txt = court if court and court != "UNKNOWN" else (("TJ" + TR.get(case[18:20], "")) if case else "")
        lead = {"root": root, "case": case, "face": face_d, "n": ncl, "name_printed": name_printed, "razao": razao, "phone": phone, "phone2": phone2,
                "uf": uf, "city": city, "cnae": cnae, "porte": porte, "capital": capital, "url": url, "dtype": LIST_KIND.get(dtype or "UNKNOWN", dtype or ""),
                "pub": pub, "proof": proof, "court": court_txt, "stage": stage or dstage or "", "year": year, "ysf": ysf, "debtor": debtor,
                "debtor_cnpj": vcnpj or (dcnpj if dsrc not in (None, "NAME_MISMATCH", "NOT_IN_RFB") else "") or "",
                "band": band, "breason": breason or "", "pstatus": {"CONFIRMED": "PLAN DOC FOUND", "ESTIMATED": "NONE FOUND", "UNKNOWN": "NONE FOUND"}.get(pstatus or "UNKNOWN", pstatus),
                "fund": fund, "primary": primary, "backup": backup, "note": "; ".join(notes), "claims": claims, "flags": flags or "",
                "fit": _seller_fit(razao, capital, porte), "financial": fin, "related": rel}
        (related if rel else out).append(lead)
    out.sort(key=lambda x: -x["fund"])
    return out, related


def _assumptions(ws, hidden=False):
    ws.title = "Assumptions"
    ws.column_dimensions["A"].width = 14; ws.column_dimensions["F"].width = 95
    ws["A1"] = "PRICING ASSUMPTIONS — yellow cells are editable; every quote on both call tabs recalculates"; ws["A1"].font = Font(bold=True, size=12)
    ws["A3"] = "Target investor IRR (annual)"; ws["B3"] = DEFAULT_IRR; ws["B3"].number_format = "0%"
    ws["A4"] = "Our buy price as share of fund price"; ws["B4"] = DEFAULT_SHARE; ws["B4"].number_format = "0%"
    for c in ("B3", "B4"): ws[c].fill = PatternFill("solid", fgColor="FFF2CC")
    for c, h in zip("ABCDEF", ("Debtor band", "Class III recovery %", "Grace (y)", "Term (y)", "Plan lag (y)", "What the band means")):
        ws[f"{c}6"] = h; ws[f"{c}6"].font = WHITE; ws[f"{c}6"].fill = HEAD
    for i, (b, rec, g, t, lag, note) in enumerate(BANDS, start=7):
        ws[f"A{i}"], ws[f"B{i}"], ws[f"C{i}"], ws[f"D{i}"], ws[f"E{i}"], ws[f"F{i}"] = b, rec, g, t, lag, note
        ws[f"B{i}"].number_format = "0%"
        for c in "BCDE": ws[f"{c}{i}"].fill = PatternFill("solid", fgColor="FFF2CC")
    lines = [
        "Fund price = face × recovery% ÷ (1+IRR)^(years to payment). Our quote = fund price × share. Our ágio = fund price − quote.",
        "Years to payment — band A (plan confirmed): max(1, grace + term/2 − years already elapsed under the plan), where the plan is assumed confirmed 'plan lag' years after filing.",
        "Years to payment — bands B, C, ?: plan lag (years from today until a plan is confirmed) + grace + term/2. Band D: plan lag = years until a liquidation payout.",
        "So an old case with no confirmed plan prices LOWER than a fresh one, never higher. Recovery % is an assumption per band, not a term read from any plan.",
        "Financial creditors (banks, FIDCs, co-ops) usually get worse class III terms than trade suppliers (longer terms, larger haircuts, no small-creditor preference); the same band recovery is an UPPER bound for them.",
        "Everything on this tab is a MODEL. The face value, the page/row and the source document are court data; the quote is ours.",
    ]
    for i, t in enumerate(lines, start=13): ws[f"A{i}"] = t
    if hidden: ws.sheet_state = "hidden"


def _call_sheet(ws, leads, title="CALL SHEET"):
    ws.title = title
    ws.freeze_panes = "D2"
    for j, (h, w, g) in enumerate(COLS, 1):
        c = ws.cell(row=1, column=j, value=h); c.font = WHITE; c.fill = HEAD; c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[1].height = 34
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True); ws.add_data_validation(dv)
    dvb = DataValidation(type="list", formula1='"A,B,C,D,?"', allow_blank=True); ws.add_data_validation(dvb)
    thin = Side(style="thin", color="D9D9D9")
    VL = lambda band, k, dflt: f"IFERROR(VLOOKUP({band},Assumptions!$A$7:$F$11,{k},FALSE),{dflt})"
    fund_col = C["FUND PRICE @ IRR (R$)"]; quote_col = C["OUR QUOTE (R$)"]
    for i, Ld in enumerate(leads, start=2):
        r = i; face = f"{L('FACE VALUE (as printed)')}{r}"; band = f"{L('DEBTOR BAND')}{r}"; filed = f"{L('Filed')}{r}"
        rec = f"{L('Exp. recovery %')}{r}"; yrs = f"{L('Years to payment')}{r}"; fund = f"{L('FUND PRICE @ IRR (R$)')}{r}"; quote = f"{L('OUR QUOTE (R$)')}{r}"
        elapsed = f'IF({filed}="",0,YEAR(TODAY())-{filed})'
        p, b = Ld["primary"], Ld["backup"]
        vals = {
            "CREDITOR — company owed money": Ld["razao"] or Ld["name_printed"], "AMOUNT OWED (R$)": float(Ld["face"]), "DEBTOR — company in RJ": Ld["debtor"] or "(debtor not identified — see proof)",
            "STATUS": "NEW", "LAST TOUCH": "", "NEXT ACTION": "", "NEXT DATE": "", "NOTES": "",
            "DECISION MAKER": p["name"], "Role": p["role"], "Company phone 1": _phone(Ld["phone"]), "Company phone 2": _phone(Ld["phone2"]), "Mobile (Apollo)": _phone(p["mobile"]),
            "Email": p["email"], "Email 2": p["email2"],
            "BACKUP CONTACT": b["name"], "Backup role": b["role"], "Backup phone": _phone(b["phone"]) if b["phone"] else "", "Backup email": b["email"], "Contact note": Ld["note"],
            "FACE VALUE (as printed)": float(Ld["face"]), "DEBTOR BAND": Ld["band"],
            "Exp. recovery %": f"={VL(band, 2, 0)}",
            "Years to payment": f'=IF({band}="A",MAX(1,{VL(band, 3, 3)}+{VL(band, 4, 8)}/2-MAX(0,{elapsed}-{VL(band, 5, 1.5)})),{VL(band, 5, 2.5)}+{VL(band, 3, 3)}+{VL(band, 4, 8)}/2)',
            "FUND PRICE @ IRR (R$)": f"={face}*{rec}/(1+Assumptions!$B$3)^{yrs}",
            "OUR QUOTE (R$)": f"={fund}*Assumptions!$B$4",
            "OUR ÁGIO (R$)": f"={fund}-{quote}", "Ágio pts of face": f'=IF({face}=0,"",ROUND(({fund}-{quote})/{face}*100,1))',
            "PROOF QUALITY": Ld["proof"], "PROOF: source document": Ld["url"] or "", "Page / row": "; ".join(f"p{pg} r{ri}" for pg, ri, _ in Ld["claims"]),
            "Value as printed": "; ".join(f"p{pg}: R$ {v}" for pg, ri, v in Ld["claims"]), "Which list": Ld["dtype"], "List date": Ld["pub"] or "", "Plan evidence": Ld["pstatus"],
            "Case number": Ld["case"] or "(not printed)", "Court": Ld["court"], "Stage": Ld["stage"], "Filed": Ld["year"] or "", "Band reason": Ld["breason"], "Debtor CNPJ": _cnpj(Ld["debtor_cnpj"]),
            "City": Ld["city"] or "", "UF": Ld["uf"] or "", "Sector": Ld["cnae"] or "", "Size": Ld["porte"] or "",
            "Capital (R$)": (float(Ld["capital"]) if Ld["capital"] not in (None, "") else ""), "Seller fit": Ld["fit"],
            "CNPJ root": Ld["root"], "Claims": Ld["n"], "Flags": Ld["flags"].replace("|", " "),
        }
        for j, (h, w, g) in enumerate(COLS, 1):
            v = vals.get(h)
            if h == "cents":
                base = fund if j == fund_col + 1 else quote
                v = f'=IF({face}=0,"",ROUND({base}/{face}*100,1))'
            c = ws.cell(row=r, column=j, value=v)
            if g in FILL: c.fill = PatternFill("solid", fgColor=FILL[g])
            c.border = Border(bottom=thin)
        for h in ("AMOUNT OWED (R$)", "FACE VALUE (as printed)", "FUND PRICE @ IRR (R$)", "OUR QUOTE (R$)", "OUR ÁGIO (R$)", "Capital (R$)"):
            ws.cell(row=r, column=C[h]).number_format = "#,##0"
        ws.cell(row=r, column=C["Exp. recovery %"]).number_format = "0%"; ws.cell(row=r, column=C["Years to payment"]).number_format = "0.0"
        ws.cell(row=r, column=C["OUR QUOTE (R$)"]).font = Font(bold=True); ws.cell(row=r, column=C["OUR ÁGIO (R$)"]).font = Font(bold=True)
        ws.cell(row=r, column=C["CREDITOR — company owed money"]).font = Font(bold=True)
        bc = ws.cell(row=r, column=C["DEBTOR BAND"]); bc.fill = PatternFill("solid", fgColor=BAND_FILL.get(Ld["band"], "EDEDED")); bc.alignment = Alignment(horizontal="center"); bc.font = Font(bold=True)
        pq = ws.cell(row=r, column=C["PROOF QUALITY"])
        for k, colr in PROOF_FILL.items():
            if Ld["proof"].startswith(k): pq.fill = PatternFill("solid", fgColor=colr)
        if not Ld["debtor"]: ws.cell(row=r, column=C["DEBTOR — company in RJ"]).font = Font(italic=True, color="9C0006")
        if Ld["url"]:
            c = ws.cell(row=r, column=C["PROOF: source document"]); c.hyperlink = Ld["url"]; c.font = Font(color="0563C1", underline="single")
        for h in ("LAST TOUCH", "NEXT DATE"): ws.cell(row=r, column=C[h]).number_format = "yyyy-mm-dd"
        dv.add(f"{L('STATUS')}{r}"); dvb.add(band)
    n = len(leads) + 1
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{max(2, n)}"
    S = L("STATUS")
    for s, colr in (("DEAL", "C6EFCE"), ("INTERESTED", "C6EFCE"), ("DEAD", "FFC7CE"), ("NOT INTERESTED", "FFC7CE")):
        ws.conditional_formatting.add(f"{S}2:{S}{n}", CellIsRule(operator="equal", formula=[f'"{s}"'], fill=PatternFill("solid", fgColor=colr)))
    ND = L("NEXT DATE")
    ws.conditional_formatting.add(f"{ND}2:{ND}{n}", FormulaRule(formula=[f'AND({ND}2<>"",{ND}2<TODAY(),{S}2<>"DEAD",{S}2<>"DEAL")'], fill=PatternFill("solid", fgColor="FFC7CE")))
    ws.conditional_formatting.add(f"{ND}2:{ND}{n}", FormulaRule(formula=[f'{ND}2=TODAY()'], fill=PatternFill("solid", fgColor="FFEB9C")))
    return n


STACK = [("CREDITOR — company owed money", "text"), ("DEBTOR — company in RJ", "text"), ("FACE VALUE (as printed)", "num"),
         ("FUND PRICE @ IRR (R$)", "num"), ("OUR ÁGIO (R$)", "num"), ("DEBTOR BAND", "text"), ("STATUS", "text"),
         ("NEXT DATE", "date"), ("UF", "text"), ("Seller fit", "text"), ("Email", "text"), ("PROOF QUALITY", "text")]
SC = {h: get_column_letter(i + 1) for i, (h, _) in enumerate(STACK)}
SC["BOOK"] = get_column_letter(len(STACK) + 1)
SC["Where"] = get_column_letter(len(STACK) + 2)


def _book(ws, books):
    """Hidden sheet stacking every call tab by direct cell reference, so the dashboard measures the whole book, not one tab."""
    ws.title = "_BOOK"
    for j, h in enumerate([h for h, _ in STACK] + ["BOOK", "Where"], 1):
        ws.cell(row=1, column=j, value=h).font = Font(bold=True)
    r = 1
    for title, last in books:
        for i in range(2, last + 1):
            r += 1
            for j, (h, kind) in enumerate(STACK, 1):
                ref = f"'{title}'!${L(h)}${i}"
                ws.cell(row=r, column=j, value=f"={ref}" if kind == "num" else f'=IF({ref}="","",{ref})')
            ws.cell(row=r, column=len(STACK) + 1, value="FINANCIAL" if "FINANCIAL" in title else "TRADE")
            ws.cell(row=r, column=len(STACK) + 2, value=f"{title} · row {i}")
    ws.sheet_state = "hidden"
    return r


def _dashboard(ws, leads, nrows, books):
    ws.title = "DASHBOARD"
    rng = lambda h: f"_BOOK!${SC[h]}$2:${SC[h]}${nrows}"
    A = rng("OUR ÁGIO (R$)"); F = rng("FACE VALUE (as printed)"); S = rng("STATUS"); BK = rng("BOOK")
    ws.column_dimensions["A"].width = 34; ws.column_dimensions["B"].width = 16; ws.column_dimensions["C"].width = 16; ws.column_dimensions["D"].width = 4
    for col, w in zip("EFGHIJK", (34, 30, 15, 7, 14, 12, 26)): ws.column_dimensions[col].width = w
    ws["A1"] = "ANTEPARO — CALL DASHBOARD"; ws["A1"].font = Font(bold=True, size=16, color=NAVY)
    ws["A2"] = f"Both call tabs — CALL SHEET (trade) and FINANCIAL CREDITORS — counted together · built {date.today().isoformat()}"
    ws["A2"].font = Font(italic=True, color="666666")

    def block(row, title, cols):
        ws[f"A{row}"] = title
        for c, lab in zip("BC", cols): ws[f"{c}{row}"] = lab
        for c in "ABC": ws[f"{c}{row}"].font = WHITE; ws[f"{c}{row}"].fill = HEAD
        return row + 1

    kpis = [("Leads in book (both tabs)", f'=SUMPRODUCT(--({rng("CREDITOR — company owed money")}<>""))', "0"),
            ("Total face (R$)", f"=SUM({F})", "#,##0"),
            ("Fund price pipeline (R$)", f'=SUM({rng("FUND PRICE @ IRR (R$)")})', "#,##0"),
            ("OUR ÁGIO pipeline (R$)", f"=SUM({A})", "#,##0"),
            ("Untouched (NEW)", f'=COUNTIF({S},"NEW")', "0"),
            ("Calls due today", f'=COUNTIF({rng("NEXT DATE")},TODAY())', "0"),
            ("Overdue follow-ups", f'=COUNTIFS({rng("NEXT DATE")},"<"&TODAY(),{rng("NEXT DATE")},"<>",{S},"<>DEAD",{S},"<>DEAL")', "0"),
            ("Interested + Deal — ágio (R$)", f'=SUMIF({S},"INTERESTED",{A})+SUMIF({S},"DEAL",{A})', "#,##0"),
            ("Rows with a decision-maker email", f'=SUMPRODUCT(--({rng("Email")}<>""))', "0"),
            ("Rows needing a proof check", f'=COUNTIF({rng("PROOF QUALITY")},"NO TOTALS*")+COUNTIF({rng("PROOF QUALITY")},"PARTIAL*")', "0")]
    r = block(4, "PIPELINE", ("", ""))
    for lab, f, fmt in kpis:
        ws[f"A{r}"] = lab; ws[f"B{r}"] = f; ws[f"B{r}"].number_format = fmt; ws[f"B{r}"].font = Font(bold=True, size=12); r += 1

    r = block(r + 1, "BOOK", ("leads", "ágio (R$)"))
    for lab, key in (("TRADE — CALL SHEET tab", "TRADE"), ("FINANCIAL — banks, FIDCs, co-ops", "FINANCIAL")):
        ws[f"A{r}"] = lab; ws[f"B{r}"] = f'=COUNTIF({BK},"{key}")'; ws[f"C{r}"] = f'=SUMIF({BK},"{key}",{A})'; ws[f"C{r}"].number_format = "#,##0"; r += 1
    ws[f"A{r}"] = "Both books"; ws[f"B{r}"] = f'=SUMPRODUCT(--({BK}<>""))'; ws[f"C{r}"] = f"=SUM({A})"; ws[f"C{r}"].number_format = "#,##0"
    for c in "ABC": ws[f"{c}{r}"].font = Font(bold=True)
    r += 1

    r = block(r + 1, "FUNNEL", ("leads", "ágio (R$)"))
    for s in STATUSES:
        ws[f"A{r}"] = s; ws[f"B{r}"] = f'=COUNTIF({S},"{s}")'; ws[f"C{r}"] = f'=SUMIF({S},"{s}",{A})'; ws[f"C{r}"].number_format = "#,##0"; r += 1

    r = block(r + 1, "DEBTOR BAND", ("leads", "ágio (R$)"))
    B = rng("DEBTOR BAND")
    for b, _, _, _, _, note in BANDS:
        ws[f"A{r}"] = f"{b} — {note[:30]}"; ws[f"B{r}"] = f'=COUNTIF({B},"{b}")'; ws[f"C{r}"] = f'=SUMIF({B},"{b}",{A})'
        ws[f"C{r}"].number_format = "#,##0"; ws[f"A{r}"].fill = PatternFill("solid", fgColor=BAND_FILL[b]); r += 1

    r = block(r + 1, "SELLER FIT", ("leads", "ágio (R$)"))
    F2 = rng("Seller fit")
    for fit in ("MID-MARKET", "SMALL", "LARGE CORP"):
        ws[f"A{r}"] = fit + (" — sells through a desk, if at all" if fit == "LARGE CORP" else ""); ws[f"B{r}"] = f'=COUNTIF({F2},"{fit}")'
        ws[f"C{r}"] = f'=SUMIF({F2},"{fit}",{A})'; ws[f"C{r}"].number_format = "#,##0"; r += 1

    r = block(r + 1, "PROOF QUALITY", ("leads", "ágio (R$)"))
    PQ = rng("PROOF QUALITY")
    for lab, pat in (("OK — reconciled against printed totals", "OK*"), ("NO TOTALS — list prints none", "NO TOTALS*"), ("PARTIAL — class III reconciled only", "PARTIAL*")):
        ws[f"A{r}"] = lab; ws[f"B{r}"] = f'=COUNTIF({PQ},"{pat}")'; ws[f"C{r}"] = f'=SUMIF({PQ},"{pat}",{A})'; ws[f"C{r}"].number_format = "#,##0"; r += 1

    notes = ["How to use: work each call tab top-down (both are sorted by our ágio). Set STATUS and NEXT DATE on every touch; this tab updates itself.",
             "Bands: A plan confirmed · B live, no plan yet · C stale · D avoid · ? unresolved. Quotes are a model off the Assumptions tab; FACE VALUE, page/row and the PROOF link are court data.",
             "PROOF QUALITY: OK = the list's own printed totals reconcile; NO TOTALS = the list prints none, so the value stands on its row alone; PARTIAL = class III reconciled but another class on the same list did not.",
             "FINANCIAL CREDITORS are banks, FIDCs, securitizers and credit co-operatives: same class III paper, but they sell through a desk, and plans usually give them worse terms than trade suppliers.",
             "Do not delete rows on the call tabs (the dashboard reads them by position); mark DEAD instead. Sorting and filtering are fine."]
    for i, t in enumerate(notes, start=r + 1): ws[f"A{i}"] = t

    ws["E4"] = "TOP 15 BY OUR ÁGIO"; ws["F4"] = "Debtor"; ws["G4"] = "Ágio (R$)"; ws["H4"] = "Band"; ws["I4"] = "Status"; ws["J4"] = "Next date"; ws["K4"] = "Where to find it"
    for c in "EFGHIJK": ws[f"{c}4"].font = WHITE; ws[f"{c}4"].fill = HEAD
    for k in range(1, 16):
        rr = 4 + k
        ws[f"G{rr}"] = f'=IFERROR(LARGE({A},{k}),"")'; ws[f"G{rr}"].number_format = "#,##0"
        # k-th occurrence among equal values, so ties never show the same row twice
        mrow = f"MATCH(G{rr},{A},0)+COUNTIF(G$5:G{rr},G{rr})-1"
        for col, h in (("E", "CREDITOR — company owed money"), ("F", "DEBTOR — company in RJ"), ("H", "DEBTOR BAND"), ("I", "STATUS"), ("J", "NEXT DATE"), ("K", "Where")):
            ws[f"{col}{rr}"] = f'=IFERROR(INDEX({rng(h)},{mrow}),"")'
        ws[f"J{rr}"].number_format = "yyyy-mm-dd"

    from collections import Counter, defaultdict
    cnt = Counter(Ld["debtor"] for Ld in leads if Ld["debtor"]); agio = defaultdict(float)
    for Ld in leads: agio[Ld["debtor"]] += Ld["fund"] * (1 - DEFAULT_SHARE)
    top = sorted(cnt.items(), key=lambda kv: (-kv[1], -agio[kv[0]]))[:12]
    r2 = 21
    ws[f"E{r2}"] = "CASE OWNERSHIP — deepest debtors"; ws[f"F{r2}"] = "leads"; ws[f"G{r2}"] = "ágio (R$)"; ws[f"H{r2}"] = "called"; ws[f"I{r2}"] = "interested"
    for c in "EFGHI": ws[f"{c}{r2}"].font = WHITE; ws[f"{c}{r2}"].fill = HEAD
    D = rng("DEBTOR — company in RJ")
    for i, (deb, _) in enumerate(top, start=r2 + 1):
        d = deb.replace('"', '""')
        ws[f"E{i}"] = deb; ws[f"F{i}"] = f'=COUNTIF({D},"{d}")'; ws[f"G{i}"] = f'=SUMIF({D},"{d}",{A})'; ws[f"G{i}"].number_format = "#,##0"
        ws[f"H{i}"] = f'=COUNTIFS({D},"{d}",{S},"<>NEW")'; ws[f"I{i}"] = f'=COUNTIFS({D},"{d}",{S},"INTERESTED")+COUNTIFS({D},"{d}",{S},"DEAL")'
    r3 = r2 + len(top) + 2
    ws[f"E{r3}"] = "BY STATE"; ws[f"F{r3}"] = "leads"; ws[f"G{r3}"] = "ágio (R$)"
    for c in "EFG": ws[f"{c}{r3}"].font = WHITE; ws[f"{c}{r3}"].fill = HEAD
    U = rng("UF")
    for i, (uf, _) in enumerate(Counter(Ld["uf"] for Ld in leads if Ld["uf"]).most_common(10), start=r3 + 1):
        ws[f"E{i}"] = uf; ws[f"F{i}"] = f'=COUNTIF({U},"{uf}")'; ws[f"G{i}"] = f'=SUMIF({U},"{uf}",{A})'; ws[f"G{i}"].number_format = "#,##0"


def export_callsheet(db, run_id, path, hide_assumptions=True):
    every, related = _leads(db, run_id)
    leads = [l for l in every if not l["financial"]]
    fin = [l for l in every if l["financial"]]
    wb = Workbook()
    dash = wb.active
    books = [("CALL SHEET", _call_sheet(wb.create_sheet(), leads))]
    if fin:   # banks, FIDCs, securitizers, credit co-ops: same layout, own tab, same dashboard
        books.append(("FINANCIAL CREDITORS", _call_sheet(wb.create_sheet(), fin, title="FINANCIAL CREDITORS")))
    nrows = _book(wb.create_sheet(), books)
    _dashboard(dash, every, nrows, books)
    _assumptions(wb.create_sheet(), hidden=hide_assumptions)
    wb.save(path)
    from collections import Counter
    tally = lambda rows: {"n": len(rows), "bands": dict(Counter(l["band"] for l in rows)), "with_email": sum(1 for l in rows if l["primary"]["email"]),
                          "with_mobile": sum(1 for l in rows if l["primary"]["mobile"]), "with_backup": sum(1 for l in rows if l["backup"]["name"]),
                          "debtor_named": sum(1 for l in rows if l["debtor"]), "face": round(sum(float(l["face"]) for l in rows)),
                          "fund": round(sum(l["fund"] for l in rows)), "agio": round(sum(l["fund"] for l in rows) * (1 - DEFAULT_SHARE)),
                          "proof": dict(Counter(l["proof"].split(" ")[0] for l in rows))}
    return {"trade": tally(leads), "financial": tally(fin), "book": tally(every), "fit": dict(Counter(l["fit"] for l in every)),
            "related_party_excluded": [(l["razao"] or l["name_printed"], l["debtor"], float(l["face"]), l["related"]) for l in related]}
