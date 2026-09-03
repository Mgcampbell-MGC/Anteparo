"""One document, one process: extract + reconcile, print JSON. Crashes and memory blow-ups stay contained."""
from __future__ import annotations

import json
import resource
import sys


def main(path: str):
    from ..extract.engine import extract_document
    from .reconcile import reconcile
    from .ingest import first_pages_text

    head = first_pages_text(path)
    doc = extract_document(path)
    rec = reconcile(doc)
    if doc.has_text_layer and not [r for r in doc.rows if r.value_brl is not None]:
        rec["status"] = "NO_ROWS"
    out = {
        "head": head, "pages": doc.pages, "has_text_layer": doc.has_text_layer, "strategy": doc.strategy,
        "layout_id": doc.layout_id, "notes": doc.notes,
        "printed_totals": [{"class": t.klass, "currency": t.currency, "total": str(t.total) if t.total is not None else None,
                            "count": t.count, "page": t.page, "section": t.section} for t in doc.printed_totals],
        "rows": [r.to_dict() for r in doc.rows],
        "rec": json.loads(json.dumps(rec, default=str)),
    }
    json.dump(out, sys.stdout, ensure_ascii=False)


if __name__ == "__main__":
    try:
        lim = int(sys.argv[2]) if len(sys.argv) > 2 else 2 * 1024 ** 3
        resource.setrlimit(resource.RLIMIT_AS, (lim, lim))
    except Exception:  # noqa: BLE001
        pass
    main(sys.argv[1])
