"""Phase 0 — run the extractor on known PDFs and prove it against hand-verified facts."""
from __future__ import annotations

import csv
import sys
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from anteparo.extract.engine import extract_document
from anteparo.steps.reconcile import reconcile
from anteparo.money import fmt_brl
from anteparo.cnpj import root

SP = Path("/tmp/claude-0/-home-user-Anteparo/f6b38504-968f-5e67-8071-9786cc41bb5a/scratchpad/pdfs")
OUT = Path(__file__).resolve().parents[1] / "data" / "out"
OUT.mkdir(parents=True, exist_ok=True)

FACTS = {
    "assistjud_conduscabos.pdf": [("33233778000108", Decimal("31981695.98"), "TAG INDUSTRIA E LAMINACAO LTDA p19 — char-verified"),
                                  ("33233778000108", Decimal("32731819.08"), "TAG (CONDUSCABOS section) p10 — cell 'R$ 3 2.731.819,08'"),
                                  ("10842374000108", Decimal("160996.26"), "MULTIPLO FIDC p19"),
                                  ("58805466000144", Decimal("902.44"), "PERFIL COMERCIAL p19 (space bug '9 02,44')")],
    "diligence_edital.pdf": [("98521909000190", Decimal("4421862.65"), "VINICOLA est. 0001"),
                             ("98521909000270", Decimal("9558482.85"), "VINICOLA est. 0002")],
    "lrf_lideres.pdf": [("48430290000130", Decimal("7135516.97"), "ADDIANTE S.A"),
                        ("01789121001107", Decimal("11525550.26"), "ALBAUGH (3 establishments, one value)")],
}

all_ok = True
for pdf in sorted(SP.glob("*.pdf")):
    doc = extract_document(str(pdf))
    rec = reconcile(doc)
    print("=" * 78)
    print(f"{pdf.name}  pages={doc.pages}  strategy={doc.strategy}/{doc.layout_id}  rows={len(doc.rows)}  notes={doc.notes}")
    print(f"  status: {rec['status']}")
    for c in rec["checks"]:
        print(f"    class {c['class']}: printed {fmt_brl(c['printed_total'])} / n={c['printed_count']}   "
              f"extracted {fmt_brl(c['extracted_sum'])} / n={c['extracted_count']}   "
              f"delta={c.get('delta')}  missing={c['value_missing']}  {'PASS' if c['pass'] else 'FAIL'}")
    if not rec["checks"]:
        for k, v in sorted(rec["extracted_sums"].items()):
            print(f"    class {k}: extracted {fmt_brl(v)} / n={rec['extracted_counts'][k]}  missing={rec['value_missing'].get(k,0)}  (no printed totals)")
    # facts
    by_doc = {}
    for r in doc.rows:
        for d in r.all_documents or [r.document_number]:
            by_doc.setdefault(d, []).append(r)
    for dnum, expect, label in FACTS.get(pdf.name, []):
        rows = by_doc.get(dnum, [])
        got = [r.value_brl for r in rows]
        ok = expect in got
        all_ok &= ok
        print(f"  {'✓' if ok else '✗'} {label}: expect {fmt_brl(expect)}  got {[fmt_brl(g) for g in got]}"
              + (f"  name={rows[0].creditor_name_as_printed!r} class={rows[0].klass} p{rows[0].page}" if rows else "  ROW NOT FOUND"))
    # VINICOLA root aggregation
    if pdf.name == "diligence_edital.pdf":
        tot = sum((r.value_brl or 0) for r in doc.rows if root(r.document_number) == "98521909")
        ok = tot == Decimal("13980345.50")
        all_ok &= ok
        print(f"  {'✓' if ok else '✗'} VINICOLA root 98521909 aggregated: {fmt_brl(tot)} (expect R$ 13.980.345,50)")
    # flag summary
    from collections import Counter
    fc = Counter(f for r in doc.rows for f in r.flags)
    print(f"  flags: {dict(fc)}")
    # dump
    with open(OUT / f"calib_{pdf.stem}.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(doc.rows[0].to_dict().keys()) if doc.rows else ["empty"])
        w.writeheader()
        for r in doc.rows:
            w.writerow(r.to_dict())
print("=" * 78)
print("ALL FACTS PASS" if all_ok else "SOME FACTS FAILED")
