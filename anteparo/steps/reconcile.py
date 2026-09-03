"""Step D — prove the reading against the document's own printed totals.

Reconcile unit = (section, class, currency). A section is one class heading occurrence, so a
group filing with two debtors' class III lists is checked twice, each against its own total.
Totals printed outside any section of that class (cover summaries) are document-level: a
class with no section total is checked against them, summing across its sections.
Tolerance 0.5% or R$1.000; a printed count must match exactly. Any failure quarantines the
document (spec) but per-class results are kept so a class that reconciled — class III, the one
we sell against — can still be used, flagged PARTIAL_DOC.
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

from ..extract.models import DocResult


def _tol(t):
    return max(t * Decimal("0.005"), Decimal(1000))


def reconcile(doc: DocResult) -> dict:
    sums, counts, missing = defaultdict(Decimal), defaultdict(int), defaultdict(int)   # keyed (section, class/cur)
    for r in doc.rows:
        if "NOT_A_CLAIM" in r.flags:
            continue
        k = (r.section, f"{r.klass or '?'}/{r.currency or 'BRL'}")
        counts[k] += 1
        if r.value_brl is not None:
            sums[k] += r.value_brl
        else:
            missing[k] += 1
    sec_tot, doc_tot = defaultdict(list), defaultdict(list)
    for t in doc.printed_totals:
        if not t.klass or (t.total is None and t.count is None):
            continue
        ck = f"{t.klass}/{t.currency or 'BRL'}"
        if (t.section, ck) in counts:
            sec_tot[(t.section, ck)].append(t)
        else:
            doc_tot[ck].append(t)
    checks, per_class = [], {}
    status = "OK"

    def check(label, ext_sum, ext_n, miss, cands):
        best = None
        for t in cands:
            ok, delta = True, None
            if t.total is not None:
                delta = ext_sum - t.total
                ok = ok and abs(delta) <= _tol(t.total)
            if t.count is not None:
                ok = ok and ext_n == t.count
            score = (ok, -(abs(delta) if delta is not None else Decimal(0)))
            if best is None or score > best[0]:
                best = (score, t, delta, ok)
        _, t, delta, ok = best
        checks.append({"class": label, "printed_total": t.total, "extracted_sum": ext_sum, "printed_count": t.count,
                       "extracted_count": ext_n, "value_missing": miss, "delta": delta,
                       "tolerance": _tol(t.total) if t.total is not None else None,
                       "count_delta": (ext_n - t.count) if t.count is not None else None, "pass": ok,
                       "n_printed_candidates": len(cands), "page": t.page})
        return ok

    covered = set()
    for (sec, ck), cands in sec_tot.items():
        ok = check(f"{ck}#s{sec}", sums.get((sec, ck), Decimal(0)), counts.get((sec, ck), 0), missing.get((sec, ck), 0), cands)
        covered.add(ck)
        per_class[ck] = "FAIL" if (not ok or per_class.get(ck) == "FAIL") else "PASS"
        if not ok:
            status = "QUARANTINED"
    for ck, cands in doc_tot.items():
        if ck in covered:
            continue
        secs = [k for k in counts if k[1] == ck]
        ext_sum = sum((sums[k] for k in secs), Decimal(0))
        ext_n = sum(counts[k] for k in secs)
        miss = sum(missing[k] for k in secs)
        ok = check(ck, ext_sum, ext_n, miss, cands)
        per_class[ck] = "PASS" if ok else "FAIL"
        if not ok:
            status = "QUARANTINED"
    if not sec_tot and not doc_tot:
        status = "OK_NO_TOTALS"
    for (_, ck) in counts:
        per_class.setdefault(ck, "NO_TOTAL")
    agg_sums = defaultdict(Decimal); agg_counts = defaultdict(int); agg_missing = defaultdict(int)
    for (sec, ck) in counts:
        agg_sums[ck] += sums[(sec, ck)]; agg_counts[ck] += counts[(sec, ck)]; agg_missing[ck] += missing[(sec, ck)]
    return {"status": status, "checks": checks, "per_class": per_class,
            "extracted_sums": dict(agg_sums), "extracted_counts": dict(agg_counts), "value_missing": dict(agg_missing)}
