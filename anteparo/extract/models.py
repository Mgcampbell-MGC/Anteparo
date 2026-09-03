"""Row-level extraction results with full provenance (the 'stranger test')."""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from decimal import Decimal


@dataclass
class ClaimRow:
    page: int                       # 1-based page in the source PDF
    row_index: int                  # ordinal of this row on that page (1-based)
    seq_as_printed: str | None      # the list's own row number, if printed
    creditor_name_as_printed: str
    document_as_printed: str        # exactly as printed (may be malformed)
    document_number: str            # digits only
    document_type: str              # CNPJ / CPF / NONE
    all_documents: list[str]        # every doc found on the row (multi-establishment entries)
    klass: str | None               # I / II / III / IV
    class_set_by: str               # SECTION_HEADING / COLUMN / INLINE_HEADER / NONE
    value_as_printed: str | None    # e.g. "31.981.695,98"
    value_brl: Decimal | None
    currency: str = "BRL"
    debtor_as_printed: str | None = None   # for group filings with a DEVEDORA column
    section_heading: str | None = None     # the heading text in force for this row
    flags: list[str] = field(default_factory=list)
    strategy: str = ""              # TABLE / PROSE

    def to_dict(self):
        d = asdict(self)
        d["value_brl"] = str(self.value_brl) if self.value_brl is not None else ""
        d["all_documents"] = "|".join(self.all_documents)
        d["flags"] = "|".join(self.flags)
        return d


@dataclass
class PrintedTotal:
    klass: str | None       # None = grand total
    total: Decimal | None
    count: int | None
    page: int
    text: str
    currency: str = "BRL"


@dataclass
class DocResult:
    file_path: str
    pages: int
    has_text_layer: bool
    strategy: str
    rows: list[ClaimRow]
    printed_totals: list[PrintedTotal]
    layout_id: str = ""
    notes: list[str] = field(default_factory=list)
