"""The closer's workbook: DASHBOARD (live CRM formulas) + CALL SHEET in the agreed column order.

Ranking = fund price (∝ our ágio), not face value. Every quote is an Excel formula off the Assumptions
tab keyed by DEBTOR BAND, so re-pricing a band re-prices every lead in it. Proof-of-claim columns travel
with the row. Assumptions is hidden in the closer file and visible in the ops file.
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

STATUSES = ["NEW", "QUEUED", "EMAILED", "CALLED", "VOICEMAIL", "CONNECTED", "INTERESTED", "NOT INTERESTED", "CALL BACK", "DEAD", "DEAL"]
BANDS = [("A", 0.40, 2, 8, "Plan approved / homologation document found, case live, debtor active at RFB"),
         ("B", 0.30, 2, 8, "Live case, no plan document found yet"),
         ("C", 0.15, 3, 10, "Stale: >4 years since filing with no plan, or no movement in 12 months, or case closed"),
         ("D", 0.03, 4, 4, "Avoid: falência decreed, or debtor no longer ATIVA"),
         ("?", 0.25, 3, 8, "Debtor could not be resolved; conservative default")]
DEFAULT_IRR, DEFAULT_SHARE = 0.25, 0.60
NAVY = "1F3A5F"; HEAD = PatternFill("solid", fgColor=NAVY); WHITE = Font(bold=True, color="FFFFFF")
FILL = {"crm": "FFF4D6", "deal": "E6EEF8", "proof": "E8F1E8", "contact": "F6F0FA", "ctx": "F3F3F3"}
BAND_FILL = {"A": "C6EFCE", "B": "E2EFDA", "C": "FFEB9C", "D": "FFC7CE", "?": "EDEDED"}

COLS = [  # (header, width, group)
    ("CREDITOR — company owed money", 40, "lead"), ("AMOUNT OWED (R$)", 17, "lead"), ("DEBTOR — company in RJ", 36, "lead"),
    ("STATUS", 15, "crm"), ("LAST TOUCH", 11, "crm"), ("NEXT ACTION", 22, "crm"), ("NEXT DATE", 11, "crm"), ("NOTES", 38, "crm"),
    ("DECISION MAKER", 30, "contact"), ("Role", 18, "contact"), ("Phone 1", 16, "contact"), ("Phone 2", 16, "contact"),
    ("Mobile (Apollo)", 16, "contact"), ("Email", 30, "contact"), ("Email 2", 26, "contact"),
    ("BACKUP CONTACT", 28, "contact"), ("Backup role", 18, "contact"), ("Backup phone", 16, "contact"), ("Backup email", 28, "contact"),
    ("FACE VALUE (as printed)", 17, "deal"), ("DEBTOR BAND", 8, "deal"), ("Exp. recovery %", 9, "deal"), ("Years to payment", 8, "deal"),
    ("FUND PRICE @ IRR (R$)", 17, "deal"), ("cents", 7, "deal"), ("OUR QUOTE (R$)", 17, "deal"), ("cents", 7, "deal"),
    ("OUR ÁGIO (R$)", 16, "deal"), ("Ágio pts of face", 8, "deal"),
    ("Plan status", 12, "proof"), ("PROOF: source document", 50, "proof"), ("Page / row", 16, "proof"), ("Value as printed", 30, "proof"),
    ("Which list", 11, "proof"), ("Published", 11, "proof"), ("Reconciled", 13, "proof"),
    ("Case number", 26, "ctx"), ("Court", 7, "ctx"), ("Stage", 11, "ctx"), ("Filed", 7, "ctx"), ("Band reason", 40, "ctx"),
    ("City", 18, "ctx"), ("UF", 5, "ctx"), ("Sector", 34, "ctx"), ("Size", 10, "ctx"), ("Seller fit", 13, "ctx"),
    ("CNPJ root", 10, "ctx"), ("Claims", 6, "ctx"), ("Flags", 30, "ctx"),
]
C = {h: i + 1 for i, (h, _, _) in enumerate(COLS)}          # header → 1-based column
L = lambda h: get_column_letter(C[h])                          # header → letter


def _phone(p):
    d = "".join(ch for ch in str(p or "") if ch.isdigit())
    if len(d) == 10: return f"({d[:2]}) {d[2:6]}-{d[6:]}"
    if len(d) == 11: return f"({d[:2]}) {d[2:7]}-{d[7:]}"
    return str(p or "")


def _seller_fit(razao, capital, porte):
    try: cap = float(capital or 0)
    except ValueError: cap = 0
    if cap >= 100_000_000 or (re.search(r"\bS\.?A\.?$|\bS/A$", razao or "") and cap >= 30_000_000):
        return "LARGE CORP"
    if porte and "MICRO" in porte.upper() or (porte and "PEQUENO" in porte.upper()):
        return "SMALL"
    return "MID-MARKET"


_AJ_JUNK = re.compile(r"administra[çc][ãa]o judicial|administradora judicial|\bAJ\b|advogad|consultoria|escrit[óo]rio|\bbase\s*\|", re.I)


def _debtor_name(*cands):
    """First candidate that looks like a company, never the administrator's own name or a portal title."""
    for c in cands:
        c = (c or "").strip()
        if c and not _AJ_JUNK.search(c):
            return c
    return ""


def _leads(db, run_id):
    q = db.execute("""
        SELECT t.cnpj_basico, t.case_number, t.doc_id, t.class_iii_face_sum, t.establishment_cnpjs, t.claim_count, t.creditor_name_as_printed, t.flags,
               c.razao_social, c.phone, c.phone2, c.email, c.uf, c.municipio, c.cnae_desc, c.porte, c.capital_social, c.is_bank, c.is_public, c.is_inactive,
               d.source_url, d.doc_type, d.publication_date, d.status, d.file_path,
               cs.court, cs.stage, cs.filing_date,
               db2.debtor_name, db2.razao_social, db2.band, db2.band_reasons, db2.plan_status, db2.filing_year, db2.stage, dd.debtor
        FROM targets t
        LEFT JOIN companies c ON c.cnpj_basico=t.cnpj_basico
        LEFT JOIN documents d ON d.doc_id=t.doc_id AND d.run_id=t.run_id
        LEFT JOIN cases cs ON cs.case_number=t.case_number
        LEFT JOIN debtors db2 ON db2.case_number=t.case_number
        LEFT JOIN doc_debtor dd ON dd.doc_id=t.doc_id
        WHERE t.run_id=? AND t.band='FLOOR'""", (run_id,)).fetchall()
    band_map = {b: (rec, g, term) for b, rec, g, term, _ in BANDS}
    out = []
    for r in q:
        (root, case, doc_id, face, ests, ncl, name_printed, flags, razao, phone, phone2, email, uf, city, cnae, porte, capital,
         is_bank, is_public, is_inactive, url, dtype, pub, dstatus, fpath, court, stage, filed, dname, drazao, band, breason, pstatus, fyear, dstage, dd_debtor) = r
        if is_public or is_inactive or "LIKELY_PUBLIC" in (flags or ""):
            continue
        fin = bool(is_bank) or "LIKELY_FINANCIAL" in (flags or "")   # banks / FIDCs / securitizers / credit co-ops → their own tab
        if not case or case[16:17] != "8":   # RJ cases live in state courts (J=8); anything else is a creditor-side number picked off the page
            case = ""; band = "?"; breason = "RJ case number not identified on the document"; stage = stage or ""
        band = band or "?"
        face_d = Decimal(face)
        year = fyear or (int(filed[:4]) if filed else None)
        if not year:
            m = re.match(r"\d{7}-\d{2}\.(\d{4})\.", case or ""); year = int(m.group(1)) if m else None
        ysf = (date.today().year - year) if year else 0
        rec, grace, term = band_map[band]
        t = max(0.5, grace + term / 2 - ysf)
        fund = float(face_d) * rec / (1 + DEFAULT_IRR) ** t
        # contacts: Apollo first (ranked), then RFB partners
        ap = db.execute("SELECT person_name, title, email, mobile, direct_phone FROM apollo_contacts WHERE cnpj_basico=? ORDER BY rank", (root,)).fetchall()
        rf = db.execute("SELECT person_name, role FROM contacts WHERE cnpj_basico=? ORDER BY CASE role_code WHEN 49 THEN 0 WHEN 5 THEN 1 WHEN 16 THEN 2 WHEN 10 THEN 3 ELSE 9 END", (root,)).fetchall()
        primary = {"name": rf[0][0] if rf else "", "role": rf[0][1] if rf else "", "mobile": "", "email": "", "email2": ""}
        if ap:
            a = ap[0]; primary.update({"name": a[0] or primary["name"], "role": a[1] or primary["role"], "mobile": a[3] or "", "email": a[2] or ""})
            if len(ap) > 1 and ap[1][2]: primary["email2"] = ap[1][2]
        backup = {"name": "", "role": "", "phone": "", "email": ""}
        if len(ap) > 1:
            backup = {"name": ap[1][0] or "", "role": ap[1][1] or "", "phone": ap[1][3] or ap[1][4] or "", "email": ap[1][2] or ""}
        elif len(rf) > 1:
            backup = {"name": rf[1][0], "role": rf[1][1], "phone": _phone(phone2 or phone), "email": ""}
        claims = db.execute("SELECT page, row_index, value_as_printed FROM claims WHERE doc_id=? AND run_id=? AND class='III' AND currency='BRL' AND document_number IN (%s) ORDER BY page,row_index"
                            % ",".join("?" * len((ests or "").split("|"))), (doc_id, run_id, *((ests or "").split("|")))).fetchall()
        out.append({"root": root, "case": case, "face": face_d, "n": ncl, "name_printed": name_printed, "razao": razao, "phone": phone, "phone2": phone2,
                    "uf": uf, "city": city, "cnae": cnae, "porte": porte, "capital": capital, "url": url, "dtype": dtype, "pub": pub, "dstatus": dstatus,
                    "court": court or (("TJ" + {"26": "SP", "19": "RJ", "16": "PR", "21": "RS", "24": "SC", "13": "MG", "09": "GO", "17": "PE", "05": "BA", "11": "MT", "12": "MS", "14": "PA", "15": "PB", "08": "ES", "10": "MA", "07": "DF", "06": "CE", "02": "AL", "25": "SE", "27": "TO", "20": "RN", "18": "PI", "22": "RO", "01": "AC", "04": "AM", "23": "RR", "03": "AP"}.get((case or "")[16:18], "")) if case else ""),
                    "stage": stage or dstage or "", "year": year, "ysf": ysf, "debtor": _debtor_name(drazao, dname, dd_debtor), "band": band, "breason": breason or "",
                    "pstatus": pstatus or ("CONFIRMED" if band == "A" else "UNKNOWN"), "fund": fund, "primary": primary, "backup": backup,
                    "claims": claims, "flags": flags or "", "fit": _seller_fit(razao, capital, porte), "financial": fin})
    out.sort(key=lambda x: -x["fund"])
    return out


def _assumptions(ws, hidden=False):
    ws.title = "Assumptions"
    ws.column_dimensions["A"].width = 14; ws.column_dimensions["E"].width = 90
    ws["A1"] = "PRICING ASSUMPTIONS — yellow cells are editable; every quote on the CALL SHEET recalculates"; ws["A1"].font = Font(bold=True, size=12)
    ws["A3"] = "Target investor IRR (annual)"; ws["B3"] = DEFAULT_IRR; ws["B3"].number_format = "0%"
    ws["A4"] = "Our buy price as share of fund price"; ws["B4"] = DEFAULT_SHARE; ws["B4"].number_format = "0%"
    for c in ("B3", "B4"): ws[c].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["A6"], ws["B6"], ws["C6"], ws["D6"], ws["E6"] = "Debtor band", "Class III recovery %", "Grace (y)", "Term (y)", "What the band means"
    for c in "ABCDE": ws[f"{c}6"].font = WHITE; ws[f"{c}6"].fill = HEAD
    for i, (b, rec, g, t, note) in enumerate(BANDS, start=7):
        ws[f"A{i}"], ws[f"B{i}"], ws[f"C{i}"], ws[f"D{i}"], ws[f"E{i}"] = b, rec, g, t, note
        ws[f"B{i}"].number_format = "0%"
        for c in "BCD": ws[f"{c}{i}"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["A13"] = "Fund price = face × recovery% ÷ (1+IRR)^(years to payment). Years = grace + term/2 − years since filing, floor 0.5."
    ws["A14"] = "Our quote = fund price × share. Our ágio = fund price − quote. All of this is a MODEL; the face value is court data."
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
    for i, Ld in enumerate(leads, start=2):
        r = i; face = f"{L('FACE VALUE (as printed)')}{r}"; band = f"{L('DEBTOR BAND')}{r}"; filed = f"{L('Filed')}{r}"
        rec = f"{L('Exp. recovery %')}{r}"; yrs = f"{L('Years to payment')}{r}"; fund = f"{L('FUND PRICE @ IRR (R$)')}{r}"; quote = f"{L('OUR QUOTE (R$)')}{r}"
        p, b = Ld["primary"], Ld["backup"]
        vals = {
            "CREDITOR — company owed money": Ld["razao"] or Ld["name_printed"], "AMOUNT OWED (R$)": float(Ld["face"]), "DEBTOR — company in RJ": Ld["debtor"],
            "STATUS": "NEW", "LAST TOUCH": "", "NEXT ACTION": "", "NEXT DATE": "", "NOTES": "",
            "DECISION MAKER": p["name"], "Role": p["role"], "Phone 1": _phone(Ld["phone"]), "Phone 2": _phone(Ld["phone2"]), "Mobile (Apollo)": _phone(p["mobile"]),
            "Email": p["email"], "Email 2": p["email2"],
            "BACKUP CONTACT": b["name"], "Backup role": b["role"], "Backup phone": _phone(b["phone"]) if b["phone"] else "", "Backup email": b["email"],
            "FACE VALUE (as printed)": float(Ld["face"]), "DEBTOR BAND": Ld["band"],
            "Exp. recovery %": f"=IFERROR(VLOOKUP({band},Assumptions!$A$7:$D$11,2,FALSE),0)",
            "Years to payment": f'=MAX(0.5,IFERROR(VLOOKUP({band},Assumptions!$A$7:$D$11,3,FALSE),3)+IFERROR(VLOOKUP({band},Assumptions!$A$7:$D$11,4,FALSE),8)/2-IF({filed}="",0,YEAR(TODAY())-{filed}))',
            "FUND PRICE @ IRR (R$)": f"={face}*{rec}/(1+Assumptions!$B$3)^{yrs}",
            "cents": None,  # filled below (two 'cents' columns)
            "OUR QUOTE (R$)": f"={fund}*Assumptions!$B$4",
            "OUR ÁGIO (R$)": f"={fund}-{quote}", "Ágio pts of face": f'=IF({face}=0,"",ROUND(({fund}-{quote})/{face}*100,1))',
            "Plan status": Ld["pstatus"], "PROOF: source document": Ld["url"] or "", "Page / row": "; ".join(f"p{pg} r{ri}" for pg, ri, _ in Ld["claims"]),
            "Value as printed": "; ".join(f"p{pg}: R$ {v}" for pg, ri, v in Ld["claims"]), "Which list": Ld["dtype"] or "", "Published": Ld["pub"] or "", "Reconciled": Ld["dstatus"] or "",
            "Case number": Ld["case"] or "(not printed)", "Court": Ld["court"] or "", "Stage": Ld["stage"], "Filed": Ld["year"] or "", "Band reason": Ld["breason"],
            "City": Ld["city"] or "", "UF": Ld["uf"] or "", "Sector": Ld["cnae"] or "", "Size": Ld["porte"] or "", "Seller fit": Ld["fit"],
            "CNPJ root": Ld["root"], "Claims": Ld["n"], "Flags": Ld["flags"].replace("|", " "),
        }
        for j, (h, w, g) in enumerate(COLS, 1):
            v = vals.get(h)
            if h == "cents":
                base = fund if j == C["FUND PRICE @ IRR (R$)"] + 1 else quote
                v = f'=IF({face}=0,"",ROUND({base}/{face}*100,1))'
            c = ws.cell(row=r, column=j, value=v)
            if g in FILL: c.fill = PatternFill("solid", fgColor=FILL[g])
            c.border = Border(bottom=thin)
        for h in ("AMOUNT OWED (R$)", "FACE VALUE (as printed)", "FUND PRICE @ IRR (R$)", "OUR QUOTE (R$)", "OUR ÁGIO (R$)"):
            ws.cell(row=r, column=C[h]).number_format = "#,##0"
        ws.cell(row=r, column=C["Exp. recovery %"]).number_format = "0%"; ws.cell(row=r, column=C["Years to payment"]).number_format = "0.0"
        ws.cell(row=r, column=C["OUR QUOTE (R$)"]).font = Font(bold=True); ws.cell(row=r, column=C["OUR ÁGIO (R$)"]).font = Font(bold=True)
        ws.cell(row=r, column=C["CREDITOR — company owed money"]).font = Font(bold=True)
        bc = ws.cell(row=r, column=C["DEBTOR BAND"]); bc.fill = PatternFill("solid", fgColor=BAND_FILL.get(Ld["band"], "EDEDED")); bc.alignment = Alignment(horizontal="center"); bc.font = Font(bold=True)
        if Ld["url"]:
            c = ws.cell(row=r, column=C["PROOF: source document"]); c.hyperlink = Ld["url"]; c.font = Font(color="0563C1", underline="single")
        for h in ("LAST TOUCH", "NEXT DATE"): ws.cell(row=r, column=C[h]).number_format = "yyyy-mm-dd"
        dv.add(f"{L('STATUS')}{r}"); dvb.add(band)
    n = len(leads) + 1
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{max(2, n)}"
    # status colouring
    S = L("STATUS")
    ws.conditional_formatting.add(f"{S}2:{S}{n}", CellIsRule(operator="equal", formula=['"DEAL"'], fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add(f"{S}2:{S}{n}", CellIsRule(operator="equal", formula=['"INTERESTED"'], fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add(f"{S}2:{S}{n}", CellIsRule(operator="equal", formula=['"DEAD"'], fill=PatternFill("solid", fgColor="FFC7CE")))
    ws.conditional_formatting.add(f"{S}2:{S}{n}", CellIsRule(operator="equal", formula=['"NOT INTERESTED"'], fill=PatternFill("solid", fgColor="FFC7CE")))
    ND = L("NEXT DATE")
    ws.conditional_formatting.add(f"{ND}2:{ND}{n}", FormulaRule(formula=[f'AND({ND}2<>"",{ND}2<TODAY(),{S}2<>"DEAD",{S}2<>"DEAL")'], fill=PatternFill("solid", fgColor="FFC7CE")))
    ws.conditional_formatting.add(f"{ND}2:{ND}{n}", FormulaRule(formula=[f'{ND}2=TODAY()'], fill=PatternFill("solid", fgColor="FFEB9C")))
    return n


def _dashboard(ws, leads, n):
    ws.title = "DASHBOARD"
    cs = "'CALL SHEET'!"
    rng = lambda h: f"{cs}${L(h)}$2:${L(h)}${n}"
    ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 16; ws.column_dimensions["C"].width = 16; ws.column_dimensions["D"].width = 4
    for col, w in zip("EFGHIJ", (36, 32, 16, 8, 14, 12)): ws.column_dimensions[col].width = w
    ws["A1"] = "ANTEPARO — CALL DASHBOARD"; ws["A1"].font = Font(bold=True, size=16, color=NAVY)
    ws["A2"] = f"Live off the CALL SHEET tab · sorted by fund price (∝ our ágio) · {date.today().isoformat()}"; ws["A2"].font = Font(italic=True, color="666666")
    tile = lambda cell, label, formula, fmt=None: (ws.__setitem__(cell, label), ws.__setitem__(cell.replace("A", "B").replace("C", "D") if False else cell[0] + str(int(cell[1:]) + 1), formula))
    # KPI block
    kpis = [("Leads on sheet", f"=COUNTA({rng('CREDITOR — company owed money')})", "0"),
            ("Total face (R$)", f"=SUM({rng('FACE VALUE (as printed)')})", "#,##0"),
            ("Fund price pipeline (R$)", f"=SUM({rng('FUND PRICE @ IRR (R$)')})", "#,##0"),
            ("OUR ÁGIO pipeline (R$)", f"=SUM({rng('OUR ÁGIO (R$)')})", "#,##0"),
            ("Untouched (NEW)", f'=COUNTIF({rng("STATUS")},"NEW")', "0"),
            ("Calls due today", f'=COUNTIF({rng("NEXT DATE")},TODAY())', "0"),
            ("Overdue follow-ups", f'=COUNTIFS({rng("NEXT DATE")},"<"&TODAY(),{rng("NEXT DATE")},"<>",{rng("STATUS")},"<>DEAD",{rng("STATUS")},"<>DEAL")', "0"),
            ("Interested + Deal — ágio (R$)", f'=SUMIF({rng("STATUS")},"INTERESTED",{rng("OUR ÁGIO (R$)")})+SUMIF({rng("STATUS")},"DEAL",{rng("OUR ÁGIO (R$)")})', "#,##0")]
    ws["A4"] = "PIPELINE"; ws["A4"].font = WHITE; ws["A4"].fill = HEAD; ws["B4"].fill = HEAD
    for i, (lab, f, fmt) in enumerate(kpis, start=5):
        ws[f"A{i}"] = lab; ws[f"B{i}"] = f; ws[f"B{i}"].number_format = fmt; ws[f"B{i}"].font = Font(bold=True, size=12)
    # funnel by status
    r0 = 5 + len(kpis) + 1
    ws[f"A{r0}"] = "FUNNEL"; ws[f"B{r0}"] = "leads"; ws[f"C{r0}"] = "ágio (R$)"
    for c in "ABC": ws[f"{c}{r0}"].font = WHITE; ws[f"{c}{r0}"].fill = HEAD
    for i, s in enumerate(STATUSES, start=r0 + 1):
        ws[f"A{i}"] = s; ws[f"B{i}"] = f'=COUNTIF({rng("STATUS")},"{s}")'; ws[f"C{i}"] = f'=SUMIF({rng("STATUS")},"{s}",{rng("OUR ÁGIO (R$)")})'; ws[f"C{i}"].number_format = "#,##0"
    # by band
    r1 = r0 + len(STATUSES) + 2
    ws[f"A{r1}"] = "DEBTOR BAND"; ws[f"B{r1}"] = "leads"; ws[f"C{r1}"] = "ágio (R$)"
    for c in "ABC": ws[f"{c}{r1}"].font = WHITE; ws[f"{c}{r1}"].fill = HEAD
    for i, (b, _, _, _, note) in enumerate(BANDS, start=r1 + 1):
        ws[f"A{i}"] = f"{b} — {note[:28]}"; ws[f"B{i}"] = f'=COUNTIF({rng("DEBTOR BAND")},"{b}")'; ws[f"C{i}"] = f'=SUMIF({rng("DEBTOR BAND")},"{b}",{rng("OUR ÁGIO (R$)")})'; ws[f"C{i}"].number_format = "#,##0"
        ws[f"A{i}"].fill = PatternFill("solid", fgColor=BAND_FILL[b])
    # top 15 by ágio (live: LARGE + INDEX/MATCH)
    ws["E4"] = "TOP 15 BY OUR ÁGIO"; ws["F4"] = "Debtor"; ws["G4"] = "Ágio (R$)"; ws["H4"] = "Band"; ws["I4"] = "Status"; ws["J4"] = "Next date"
    for c in "EFGHIJ": ws[f"{c}4"].font = WHITE; ws[f"{c}4"].fill = HEAD
    A = rng("OUR ÁGIO (R$)")
    for k in range(1, 16):
        r = 4 + k
        ws[f"G{r}"] = f"=IFERROR(LARGE({A},{k}),\"\")"; ws[f"G{r}"].number_format = "#,##0"
        mrow = f"MATCH(G{r},{A},0)"
        for col, h in (("E", "CREDITOR — company owed money"), ("F", "DEBTOR — company in RJ"), ("H", "DEBTOR BAND"), ("I", "STATUS"), ("J", "NEXT DATE")):
            ws[f"{col}{r}"] = f'=IFERROR(INDEX({rng(h)},{mrow}),"")'
        ws[f"J{r}"].number_format = "yyyy-mm-dd"
    # case ownership: top debtors by lead count (names precomputed, counts live)
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
        ws[f"H{i}"] = f'=COUNTIFS({D},"{d}",{rng("STATUS")},"<>NEW")'; ws[f"I{i}"] = f'=COUNTIFS({D},"{d}",{rng("STATUS")},"INTERESTED")+COUNTIFS({D},"{d}",{rng("STATUS")},"DEAL")'
    # by state
    r3 = r2 + len(top) + 2
    ws[f"E{r3}"] = "BY STATE"; ws[f"F{r3}"] = "leads"; ws[f"G{r3}"] = "ágio (R$)"
    for c in "EFG": ws[f"{c}{r3}"].font = WHITE; ws[f"{c}{r3}"].fill = HEAD
    ufs = Counter(Ld["uf"] for Ld in leads if Ld["uf"]).most_common(10)
    for i, (uf, _) in enumerate(ufs, start=r3 + 1):
        ws[f"E{i}"] = uf; ws[f"F{i}"] = f'=COUNTIF({rng("UF")},"{uf}")'; ws[f"G{i}"] = f'=SUMIF({rng("UF")},"{uf}",{A})'; ws[f"G{i}"].number_format = "#,##0"
    ws[f"A{r1 + len(BANDS) + 2}"] = "How to use: work the CALL SHEET top-down (it is sorted by our ágio). Set STATUS and NEXT DATE on every touch; this tab updates itself."
    ws[f"A{r1 + len(BANDS) + 3}"] = "Bands: A plan approved · B live · C stale · D avoid. Quotes are a model off the Assumptions tab; the FACE VALUE and the PROOF link are court data."


def export_callsheet(db, run_id, path, hide_assumptions=True):
    every = _leads(db, run_id)
    leads = [l for l in every if not l["financial"]]
    fin = [l for l in every if l["financial"]]
    wb = Workbook()
    dash = wb.active
    cs = wb.create_sheet(); n = _call_sheet(cs, leads)
    _dashboard(dash, leads, n)
    if fin:   # banks, FIDCs, securitizers, credit co-ops: same layout, separate tab, not counted on the dashboard
        _call_sheet(wb.create_sheet(), fin, title="FINANCIAL CREDITORS")
    _assumptions(wb.create_sheet(), hidden=hide_assumptions)
    wb.save(path)
    from collections import Counter
    return {"leads": len(leads), "financial": len(fin), "bands": dict(Counter(l["band"] for l in leads)), "with_email": sum(1 for l in leads if l["primary"]["email"]),
            "with_mobile": sum(1 for l in leads if l["primary"]["mobile"]), "with_backup": sum(1 for l in leads if l["backup"]["name"]),
            "fit": dict(Counter(l["fit"] for l in leads))}
