"""SQLite store for the six index tables plus raw-document manifest and RFB cache.

Append-only: nothing is deleted. A re-extraction inserts new rows under a new run_id and
marks the previous rows for the same document with superseded_by = <new run_id>.
"""
from __future__ import annotations

import sqlite3
import time
import uuid

SCHEMA = """
CREATE TABLE IF NOT EXISTS runs(run_id TEXT PRIMARY KEY, started_at TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS cases(
  case_number TEXT PRIMARY KEY, court TEXT, chamber TEXT, chamber_code INTEGER, grau TEXT,
  filing_date TEXT, stage TEXT, rj_granted_signal INTEGER, last_movement TEXT, n_movements INTEGER,
  movements_json TEXT, system TEXT, source_url TEXT, datajud_updated TEXT,
  debtor_name TEXT, debtor_cnpj TEXT, administrator_name TEXT, administrator_cnpj TEXT,
  extracted_at TEXT, extracted_by TEXT);
CREATE TABLE IF NOT EXISTS raw_documents(
  sha1 TEXT PRIMARY KEY, url TEXT UNIQUE, domain TEXT, page_url TEXT, page_title TEXT, link_text TEXT,
  case_numbers TEXT, doc_kind TEXT, path TEXT, size INTEGER, content_type TEXT,
  fetched_at TEXT, http_status INTEGER, error TEXT);
CREATE TABLE IF NOT EXISTS documents(
  doc_id TEXT, run_id TEXT, case_number TEXT, doc_type TEXT, publication_date TEXT, source_url TEXT,
  file_path TEXT, pages INTEGER, strategy TEXT, layout_id TEXT, printed_totals_json TEXT, status TEXT,
  reconcile_json TEXT, notes TEXT, debtor_name_hint TEXT, administrator_hint TEXT,
  extracted_at TEXT, extracted_by TEXT, superseded_by TEXT, PRIMARY KEY(doc_id, run_id));
CREATE TABLE IF NOT EXISTS claims(
  id INTEGER PRIMARY KEY AUTOINCREMENT, doc_id TEXT, run_id TEXT, case_number TEXT, page INTEGER, row_index INTEGER,
  seq_as_printed TEXT, creditor_name_as_printed TEXT, document_as_printed TEXT, document_number TEXT,
  document_type TEXT, all_documents TEXT, class TEXT, class_set_by TEXT, value_as_printed TEXT,
  value_brl TEXT, currency TEXT, debtor_as_printed TEXT, section_heading TEXT, flags TEXT, strategy TEXT,
  extracted_at TEXT, extracted_by TEXT, superseded_by TEXT);
CREATE INDEX IF NOT EXISTS ix_claims_doc ON claims(doc_id, run_id);
CREATE INDEX IF NOT EXISTS ix_claims_docnum ON claims(document_number);
CREATE TABLE IF NOT EXISTS companies(
  cnpj_basico TEXT PRIMARY KEY, cnpj_matriz TEXT, razao_social TEXT, nome_fantasia TEXT,
  natureza_juridica TEXT, natureza_desc TEXT, porte TEXT, situacao_cadastral TEXT, situacao_desc TEXT,
  data_situacao TEXT, cnae_principal TEXT, cnae_desc TEXT, uf TEXT, municipio TEXT, phone TEXT, phone2 TEXT,
  email TEXT, capital_social TEXT, inicio_atividade TEXT, is_bank INTEGER, is_public INTEGER, is_inactive INTEGER,
  rfb_source TEXT, rfb_fetched_at TEXT, qsa_json TEXT, extracted_at TEXT, extracted_by TEXT);
CREATE TABLE IF NOT EXISTS targets(
  id INTEGER PRIMARY KEY AUTOINCREMENT, run_id TEXT, cnpj_basico TEXT, case_number TEXT, doc_id TEXT,
  class_iii_face_sum TEXT, establishment_cnpjs TEXT, claim_count INTEGER, creditor_name_as_printed TEXT,
  debtor_name TEXT, stage TEXT, is_related_party INTEGER, above_floor INTEGER, band TEXT, flags TEXT,
  extracted_at TEXT, extracted_by TEXT, superseded_by TEXT);
CREATE TABLE IF NOT EXISTS contacts(
  id INTEGER PRIMARY KEY AUTOINCREMENT, cnpj_basico TEXT, person_name TEXT, role TEXT, role_code INTEGER,
  cpf_masked TEXT, linkedin_url TEXT, phone TEXT, email TEXT, source TEXT, confidence TEXT,
  extracted_at TEXT, extracted_by TEXT);
CREATE INDEX IF NOT EXISTS ix_contacts_root ON contacts(cnpj_basico);
"""


def now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def connect(path: str) -> sqlite3.Connection:
    db = sqlite3.connect(path, timeout=60)
    db.execute("PRAGMA journal_mode=WAL")
    db.executescript(SCHEMA)
    return db


def new_run(db, note: str) -> str:
    rid = time.strftime("%Y%m%dT%H%M%SZ", time.gmtime()) + "-" + uuid.uuid4().hex[:6]
    db.execute("INSERT INTO runs VALUES(?,?,?)", (rid, now(), note))
    db.commit()
    return rid
