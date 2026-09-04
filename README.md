# ANTEPARO — creditor index

The standing list of class III (trade) creditors in Brazilian *recuperação judicial* cases — every value
traceable to a page of a public court document — and the call sheet closers dial from.
Spec: *202 Building the index*. Business context: *101 The Forward Flow Claims Platform*.

## What it produces

| output | what it is |
|---|---|
| `data/out/ANTEPARO_call_sheet.xlsx` | **The call list.** CRM columns · lead · IRR-driven quote · debtor/case · decision-maker · **proof of claim** (link, page, row, value as printed) |
| `data/out/{cases,documents,claims,companies,targets,contacts}.csv` | The six tables, exact spec schema |
| `data/out/diff_vs_drive_sheet.csv` (+ workbook tab) | Every number that changed vs. the previous Drive sheet, and why |
| `data/out/report.md` | Documents by status/layout, coverage by court, quarantines, blocked sources, RFB release used |
| `data/anteparo.sqlite` | Append-only store: nothing deleted, re-extractions carry `superseded_by` |

## Pipeline (spec steps A–I)

```
A  DataJud (CNJ)          classe 129 since 2020, all 27 courts        → cases
B  AJ portals             crawl → creditor lists, plans, homologations → raw_documents (sha1-addressed PDFs)
C  extract                ruled tables · unruled tables (bands) · prose editais
D  reconcile              per (section, class, currency) against the document's own printed totals
E  clean                  mod-11 CNPJ, CPF out, exact dupes, aggregate on the 8-digit root
F  enrich                 RFB registry via mirror (status, CNAE, UF, phone, email), throttled + cached
G  targets                class III · active · non-bank · non-public · ≥ R$200k (pooling band 100–200k)
H  contacts               sócio-administrador / administrador / diretor / presidente from the partner register
I  QC                     statuses, per-class verdicts, coverage by court, flags on every row
```

Sources reachable from outside Brazil: DataJud (public key), administrator portals, state court portals,
RFB via `minhareceita.org` / BrasilAPI. **Geo-blocked** (need a Brazilian egress): DJEN, the RFB bulk files.
Both have adapter seams in `anteparo/sources/`.

## Extraction — what it defends against

Every rule below came from a real document in the harvest.

- **Values only from the row's own value cell/column.** Never proximity. The rule that turned
  `R$ 1.981.695,98` back into the printed **R$ 31.981.695,98** (TAG INDUSTRIA, assistjud p19).
- Fragment repair for pdfplumber splits (`R$ 3 1.981.695,98`, `R$ 9 02,44`).
- Fake-bold doubled glyphs (`GGuussttaavvoo` → `Gustavo`), rotated margin stamps dropped (gazette *valor da causa*).
- Multi-line names kept whole (ruled cells; nearest-anchor bands); running headers/footers ignored.
- Class from section headings (`CLASSE III`, `C – Titulares de créditos quirografários…`) or a CLASSE column;
  **sections** for group filings (one total per debtor); header-style vs footer-style totals; running totals.
- Non-class sections reset the class: *não sujeitos, extraconcursais, reserva de crédito, fiduciária, ACC…*
- Two value columns (`Valor Nominal | Valor Atualizado`, `1ª Lista | 2ª Lista`): the admitted one wins, the other is kept as `ALT_VALUE`.
- Foreign-currency sections (`… EM DÓLARES`) reconciled separately, never summed into BRL.
- Summary rows (`CLASSIFICAÇÃO`, `TOTAL`, bare class words) are never claims; creditors *named* "Total …" are.
- A document that does not reconcile is **QUARANTINED**. A class that reconciled on its own (class III, the one
  we sell against) is still usable, flagged `PARTIAL_DOC_CLASS_III_RECONCILED`. No printed totals → `OK_NO_TOTALS`,
  rows flagged `UNRECONCILED_SOURCE`.
- Each document extracts in its own process (memory cap, timeout): one bad PDF is an `ERROR` row, not a dead run.

## Run

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python scripts/calibrate.py            # Phase 0: proves the extractor on hand-verified facts
.venv/bin/python scripts/load_cases.py           # A: DataJud universe (all courts) → cases
.venv/bin/python scripts/harvest.py              # B: direct URLs + crawl data/seeds/domains.txt
.venv/bin/python scripts/run_phase1.py           # C–I + exports (new run)
.venv/bin/python scripts/run_phase1.py --run <id> --reingest QUARANTINED,ERROR,NO_ROWS   # repair pass
.venv/bin/python scripts/reexport.py latest      # regenerate CSVs + workbook without re-extracting
```

## Reading the workbook

- **Class III face** is the admitted value *as printed* — what the company is **owed**, updated only to the
  filing date (art. 9 II). It is not what will be paid.
- **OUR QUOTE** is a model: face → assumed plan terms (by *Plan status*) → discounted at the investor IRR →
  fund price → our share. Edit the **Assumptions** tab; every quote recalculates. Nothing on that tab is court data.
- **Plan status**: `CONFIRMED` a plan/homologation document exists for the case · `ESTIMATED` case live in
  DataJud, terms assumed · `UNKNOWN` · `CONVERTED_TO_BANKRUPTCY`.
- **Proof columns (green)**: click the source, go to the page/row, read the value. `Reconciled = OK` means the
  document's class totals matched its own printed totals within 0.5%.
- **Needs CNPJ** tab: class III rows above the floor whose list prints no document number — real leads that
  need name resolution before enrichment.
- **Diff vs Drive sheet** tab: old amount vs new root-summed face, with the reason.

## Status — Phase 1 (2026-09-04)
- 44 administrator portals crawled → 1,224 PDFs · 1,199 ingested · 9,770 RJ cases since 2020 in the DataJud universe
- **559 leads ≥ R$200k on the Call List**, 236 pooling, 1,063 name-only leads ≥ R$200k awaiting CNPJ resolution
- 943 companies enriched from the RFB registry (100% hit) · 2,800+ decision-maker contacts
- End-to-end spot check: 50/50 sampled leads → the printed value is on the cited page
- Details: `data/out/report.md`. Phase 2 levers: OCR for 74 scanned lists, name→CNPJ resolution, TJRJ/TJSP court portals.
