"""Step D — prove the reading against the document's own printed totals."""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

from ..extract.models import DocResult


def reconcile(doc: DocResult) -> dict:
    sums, counts = defaultdict(Decimal), defaultdict(int)
    missing = defaultdict(int)
    for r in doc.rows:
        if "NOT_A_CLAIM" in r.flags:
            continue
        k = f"{r.klass or '?'}/{r.currency or 'BRL'}"
        counts[k] += 1
        if r.value_brl is not None:
            sums[k] += r.value_brl
        else:
            missing[k] += 1
    printed = {}
    for t in doc.printed_totals:
        if t.klass and (t.total is not None or t.count is not None):
            printed[f"{t.klass}/{t.currency or 'BRL'}"] = t
    checks = []
    status = "OK"
    for k, t in printed.items():
        ext_sum, ext_n = sums.get(k, Decimal(0)), counts.get(k, 0)
        row = {"class": k, "printed_total": t.total, "extracted_sum": ext_sum,
               "printed_count": t.count, "extracted_count": ext_n, "value_missing": missing.get(k, 0)}
        ok = True
        if t.total is not None:
            delta = ext_sum - t.total
            tol = max(t.total * Decimal("0.005"), Decimal(1000))
            row["delta"] = delta
            row["tolerance"] = tol
            ok = ok and abs(delta) <= tol
        if t.count is not None:
            row["count_delta"] = ext_n - t.count
            ok = ok and ext_n == t.count
        row["pass"] = ok
        checks.append(row)
        if not ok:
            status = "QUARANTINED"
    if not printed:
        status = "OK_NO_TOTALS"
    return {"status": status, "checks": checks,
            "extracted_sums": dict(sums), "extracted_counts": dict(counts), "value_missing": dict(missing)}
