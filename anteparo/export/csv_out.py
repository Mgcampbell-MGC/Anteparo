"""The six CSVs, exact spec schema."""
from __future__ import annotations

import csv
from pathlib import Path


def _w(path, header, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def export_csvs(db, run_id, out_dir: str):
    out = Path(out_dir); out.mkdir(parents=True, exist_ok=True)
    _w(out / "cases.csv",
       ["case_number", "court", "chamber", "debtor_name", "debtor_cnpj", "filing_date", "stage", "administrator_name", "administrator_cnpj", "source_url", "extracted_at", "extracted_by"],
       db.execute("""SELECT c.case_number, c.court, c.chamber, COALESCE(c.debtor_name, d.debtor_name_hint), c.debtor_cnpj, c.filing_date, c.stage,
                     COALESCE(c.administrator_name, d.administrator_hint), c.administrator_cnpj, c.source_url, c.extracted_at, c.extracted_by
                     FROM cases c LEFT JOIN documents d ON d.case_number=c.case_number AND d.run_id=?
                     WHERE c.case_number IN (SELECT case_number FROM documents WHERE run_id=?) GROUP BY c.case_number""", (run_id, run_id)).fetchall())
    _w(out / "documents.csv",
       ["doc_id", "case_number", "doc_type", "publication_date", "source_url", "file_path", "pages", "printed_total_by_class", "status", "strategy", "layout_id", "notes", "extracted_at", "extracted_by"],
       db.execute("SELECT doc_id, case_number, doc_type, publication_date, source_url, file_path, pages, printed_totals_json, status, strategy, layout_id, notes, extracted_at, extracted_by FROM documents WHERE run_id=?", (run_id,)).fetchall())
    _w(out / "claims.csv",
       ["case_number", "doc_id", "page", "row_index", "creditor_name_as_printed", "document_number", "document_type", "class", "value_brl", "value_as_printed", "currency", "flags", "seq_as_printed", "debtor_as_printed", "extracted_at", "extracted_by"],
       db.execute("SELECT case_number, doc_id, page, row_index, creditor_name_as_printed, document_number, document_type, class, value_brl, value_as_printed, currency, flags, seq_as_printed, debtor_as_printed, extracted_at, extracted_by FROM claims WHERE run_id=? ORDER BY doc_id, page, row_index", (run_id,)).fetchall())
    _w(out / "companies.csv",
       ["cnpj_basico", "razao_social", "natureza_juridica", "porte", "situacao_cadastral", "cnae_principal", "uf", "municipio", "phone", "email", "is_bank", "is_public", "is_inactive", "rfb_source", "rfb_fetched_at", "extracted_at", "extracted_by"],
       db.execute("SELECT cnpj_basico, razao_social, natureza_juridica, porte, situacao_cadastral, cnae_principal, uf, municipio, phone, email, is_bank, is_public, is_inactive, rfb_source, rfb_fetched_at, extracted_at, extracted_by FROM companies WHERE cnpj_basico IN (SELECT cnpj_basico FROM targets WHERE run_id=?)", (run_id,)).fetchall())
    _w(out / "targets.csv",
       ["cnpj_basico", "case_number", "class_iii_face_sum", "establishment_cnpjs", "claim_count", "debtor_name", "stage", "is_related_party", "above_floor", "band", "doc_id", "creditor_name_as_printed", "flags", "extracted_at", "extracted_by"],
       db.execute("SELECT cnpj_basico, case_number, class_iii_face_sum, establishment_cnpjs, claim_count, debtor_name, stage, is_related_party, above_floor, band, doc_id, creditor_name_as_printed, flags, extracted_at, extracted_by FROM targets WHERE run_id=? ORDER BY CAST(class_iii_face_sum AS REAL) DESC", (run_id,)).fetchall())
    _w(out / "contacts.csv",
       ["cnpj_basico", "person_name", "role", "cpf_masked", "linkedin_url", "phone", "email", "source", "confidence", "extracted_at", "extracted_by"],
       db.execute("SELECT cnpj_basico, person_name, role, cpf_masked, linkedin_url, phone, email, source, confidence, extracted_at, extracted_by FROM contacts WHERE cnpj_basico IN (SELECT cnpj_basico FROM targets WHERE run_id=?)", (run_id,)).fetchall())
    return out
