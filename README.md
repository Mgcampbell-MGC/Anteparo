# ANTEPARO — creditor index

Builds the standing list of class III (trade) creditors in Brazilian recuperação judicial cases,
with every value traceable to a page of a public court document. Spec: *202 Building the index*.

## Layout
- `anteparo/cnpj.py`, `anteparo/money.py` — CNPJ/CPF mod-11, BRL parsing, fragment repair
- `anteparo/extract/` — column-aware PDF extraction: ruled tables, unruled tables (bands), prose editais
- `anteparo/steps/` — D reconcile · E clean · F enrich · G targets · H contacts · I QC
- `anteparo/sources/` — DataJud (cases), administrator portals (documents), RFB via mirror (companies)
- `scripts/calibrate.py` — Phase 0: proves the extractor against hand-verified facts

## Phase 0 — calibration (3 layouts)
| document | layout | status | evidence |
|---|---|---|---|
| assistjud (TJBA) | TABLE_RULED | OK_NO_TOTALS | TAG INDUSTRIA = R$ 31.981.695,98 (p19) and R$ 32.731.819,08 (p10), char-verified |
| diligence (TJPE, PJe) | PROSE | **OK** | I, III/BRL, III/USD, IV all reconcile to printed totals (max delta R$0,02) |
| lrf lideres (TJPA, DJEN) | PROSE | **OK** | II, III, IV reconcile exactly; 25/25 creditors |

Run: `.venv/bin/python scripts/calibrate.py`
