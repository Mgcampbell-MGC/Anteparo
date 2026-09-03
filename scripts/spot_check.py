"""Step I end-to-end sample: random targets → the PDF page named on the row → the printed value must be there.

Usage: spot_check.py [run_id|latest] [n]
A row passes when the value string (digits only) of every claim behind the target is found in the
text of the cited page. This is the automated half of the 'stranger test'; the other half is a human
opening the link.
"""
from __future__ import annotations
import random, re, sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(ROOT))
from anteparo.db import connect
from anteparo.extract.pdfio import open_pdf, page_text

db = connect(str(ROOT / "data/anteparo.sqlite"))
run = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] != "latest" else db.execute("SELECT run_id FROM runs WHERE note='phase1' ORDER BY started_at DESC LIMIT 1").fetchone()[0]
n = int(sys.argv[2]) if len(sys.argv) > 2 else 50
random.seed(7)
targets = db.execute("""SELECT t.id, t.cnpj_basico, t.doc_id, t.class_iii_face_sum, t.creditor_name_as_printed, d.file_path, t.establishment_cnpjs
                        FROM targets t JOIN documents d ON d.doc_id=t.doc_id AND d.run_id=t.run_id
                        WHERE t.run_id=? AND t.band='FLOOR' AND t.flags NOT LIKE '%LIKELY_%'""", (run,)).fetchall()
sample = random.sample(targets, min(n, len(targets)))
ok = fail = 0
cache = {}
for tid, root, doc_id, face, name, path, ests in sample:
    claims = db.execute("SELECT page, value_as_printed, creditor_name_as_printed FROM claims WHERE doc_id=? AND run_id=? AND class='III' AND currency='BRL' AND document_number IN (%s)"
                        % ",".join("?" * len(ests.split("|"))), (doc_id, run, *ests.split("|"))).fetchall()
    row_ok = True
    detail = []
    for page, val, cname in claims:
        key = (path, page)
        if key not in cache:
            try:
                with open_pdf(path) as pdf:
                    t = page_text(pdf.pages[page - 1])
                cache[key] = re.sub(r"\s+", "", t)
            except Exception as e:  # noqa: BLE001
                cache[key] = ""
        found = re.sub(r"\s+", "", val or "") in cache[key]
        # doubled-glyph pages: the value appears with every char twice
        if not found and val:
            found = "".join(ch * 2 for ch in re.sub(r"\s+", "", val)) in cache[key]
        row_ok &= found
        detail.append(f"p{page} {val} {'✓' if found else '✗'}")
    ok += row_ok; fail += (not row_ok)
    print(f"  {'PASS' if row_ok else 'FAIL'} {name[:38]:<38} root={root} face={float(face):>16,.2f}  {' · '.join(detail)}")
print(f"\nspot check: {ok} PASS / {fail} FAIL of {len(sample)} sampled targets (run {run})")
