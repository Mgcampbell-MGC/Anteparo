"""Steps F/H — company registry data from the Receita Federal open data, via a mirror.

The RFB bulk files are geo-blocked from this host, so we read the same RFB release through
public per-CNPJ mirrors (minhareceita.org primary, BrasilAPI fallback). Every record carries
its source and fetch time. Calls are throttled and cached on disk; nothing is ever filled
from memory — a CNPJ the mirror does not know stays NOT_IN_RFB.
"""
from __future__ import annotations

import json
import sqlite3
import time

import requests

SOURCES = [
    ("minhareceita", "https://minhareceita.org/{cnpj}"),
    ("brasilapi", "https://brasilapi.com.br/api/cnpj/v1/{cnpj}"),
]
MIN_INTERVAL = 0.7  # seconds between calls, per host — these are free public services


class RFB:
    def __init__(self, db_path: str):
        self.db = sqlite3.connect(db_path, timeout=120)
        self.db.execute("""CREATE TABLE IF NOT EXISTS rfb_cache(
            cnpj TEXT PRIMARY KEY, source TEXT, fetched_at TEXT, status INTEGER, json TEXT)""")
        self.db.commit()
        self._last = {}
        self.s = requests.Session()
        self.s.headers["User-Agent"] = "anteparo-index/0.1 (+registry lookups, throttled)"

    def _throttle(self, host):
        dt = time.time() - self._last.get(host, 0)
        if dt < MIN_INTERVAL:
            time.sleep(MIN_INTERVAL - dt)
        self._last[host] = time.time()

    def get(self, cnpj: str) -> dict | None:
        cnpj = "".join(ch for ch in cnpj if ch.isdigit())
        if len(cnpj) != 14:
            return None
        row = self.db.execute("SELECT source, fetched_at, status, json FROM rfb_cache WHERE cnpj=?", (cnpj,)).fetchone()
        if row and row[2] == 200:
            d = json.loads(row[3]); d["_source"], d["_fetched_at"] = row[0], row[1]
            return d
        if row and row[2] == 404:
            return None
        for name, url in SOURCES:
            self._throttle(name)
            try:
                r = self.s.get(url.format(cnpj=cnpj), timeout=30)
            except requests.RequestException:
                continue
            if r.status_code == 200:
                try:
                    d = r.json()
                except ValueError:
                    continue
                if not d.get("razao_social") and not d.get("cnpj"):
                    continue
                now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
                self.db.execute("INSERT OR REPLACE INTO rfb_cache VALUES(?,?,?,?,?)", (cnpj, name, now, 200, json.dumps(d, ensure_ascii=False)))
                self.db.commit()
                d["_source"], d["_fetched_at"] = name, now
                return d
            if r.status_code == 404:
                now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
                self.db.execute("INSERT OR REPLACE INTO rfb_cache VALUES(?,?,?,?,?)", (cnpj, name, now, 404, "{}"))
                self.db.commit()
                return None
            if r.status_code == 429:
                time.sleep(5)
        return None


def normalise(d: dict) -> dict:
    """Flatten a mirror record to the companies schema (both mirrors share the RFB field names)."""
    sit = d.get("situacao_cadastral")
    sit_code = str(sit).zfill(2) if sit is not None else ""
    cnae = d.get("cnae_fiscal")
    return {
        "cnpj": "".join(ch for ch in str(d.get("cnpj", "")) if ch.isdigit()),
        "razao_social": d.get("razao_social"),
        "nome_fantasia": d.get("nome_fantasia"),
        "natureza_juridica": str(d.get("codigo_natureza_juridica") or ""),
        "natureza_juridica_desc": d.get("natureza_juridica"),
        "porte": d.get("porte") or d.get("descricao_porte"),
        "situacao_cadastral": sit_code,
        "situacao_desc": d.get("descricao_situacao_cadastral"),
        "data_situacao": d.get("data_situacao_cadastral"),
        "cnae_principal": str(cnae or ""),
        "cnae_desc": d.get("cnae_fiscal_descricao"),
        "uf": d.get("uf"),
        "municipio": d.get("municipio"),
        "phone": d.get("ddd_telefone_1"),
        "phone2": d.get("ddd_telefone_2"),
        "email": d.get("email"),
        "capital_social": d.get("capital_social"),
        "inicio_atividade": d.get("data_inicio_atividade"),
        "qsa": [{"nome": q.get("nome_socio"), "qual_code": q.get("codigo_qualificacao_socio"),
                 "qual": q.get("qualificacao_socio"), "cpf_masked": q.get("cnpj_cpf_do_socio"),
                 "entrada": q.get("data_entrada_sociedade")} for q in (d.get("qsa") or [])],
        "source": d.get("_source"),
        "fetched_at": d.get("_fetched_at"),
    }
