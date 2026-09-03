"""Step A — the case universe from the CNJ DataJud public API.

Every state court, classe 129 (Recuperação Judicial), paged with search_after.
DataJud carries no party names; debtor name/CNPJ come from the first edital (step B).
Movements are generic procedural codes — plan homologation is NOT reliably visible here,
so stage is a coarse derivation and plan status stays for the document harvest.
"""
from __future__ import annotations

import json
import os
import time

import requests

COURTS = ["tjac", "tjal", "tjam", "tjap", "tjba", "tjce", "tjdft", "tjes", "tjgo", "tjma", "tjmg", "tjms",
          "tjmt", "tjpa", "tjpb", "tjpe", "tjpi", "tjpr", "tjrj", "tjrn", "tjro", "tjrr", "tjrs", "tjsc",
          "tjse", "tjsp", "tjto"]
URL = "https://api-publica.datajud.cnj.jus.br/api_publica_{court}/_search"
# The public key the CNJ publishes on the DataJud page (override with DATAJUD_API_KEY).
KEY = os.environ.get("DATAJUD_API_KEY", "cDZHYzlZa0JadVREZDJCendQbXY6SkJlTzNjLV9TRENyQk1RdnFKZGRQdw==")
CLASSE_RJ = 129

BANKRUPTCY_WORDS = ("decretação de falência", "convolação em falência", "decreta a falência")
CLOSED_WORDS = ("extinção", "baixa definitiva", "arquivamento definitivo", "encerramento")
RJ_GRANT_WORDS = ("concessão da recuperação", "homologação do plano", "homologação de plano")


def _session():
    s = requests.Session()
    s.headers.update({"Authorization": f"APIKey {KEY}", "Content-Type": "application/json"})
    return s


def iter_cases(court: str, since: str = "2020-01-01", size: int = 1000, sleep: float = 0.4):
    """Yield raw DataJud hits for classe 129 in one court, filed on/after `since`."""
    s = _session()
    since_ts = since.replace("-", "") + "000000"
    body = {
        "size": size,
        "query": {"bool": {"must": [{"match": {"classe.codigo": CLASSE_RJ}}],
                           "filter": [{"range": {"dataAjuizamento": {"gte": since_ts}}}]}},
        "sort": [{"@timestamp": {"order": "asc"}}, {"numeroProcesso.keyword": {"order": "asc"}}],
    }
    search_after = None
    while True:
        if search_after:
            body["search_after"] = search_after
        for attempt in range(4):
            try:
                r = s.post(URL.format(court=court), data=json.dumps(body), timeout=90)
                if r.status_code == 200:
                    break
                time.sleep(2 * (attempt + 1))
            except requests.RequestException:
                time.sleep(2 * (attempt + 1))
        else:
            raise RuntimeError(f"DataJud {court}: gave up")
        hits = r.json().get("hits", {}).get("hits", [])
        if not hits:
            return
        for h in hits:
            yield h
        search_after = hits[-1].get("sort")
        if len(hits) < size:
            return
        time.sleep(sleep)


def derive(hit: dict) -> dict:
    src = hit.get("_source", {})
    movs = src.get("movimentos") or []
    names = [(m.get("nome") or "").lower() for m in movs]
    dates = sorted([m.get("dataHora") for m in movs if m.get("dataHora")])
    stage = "ACTIVE"
    if any(any(w in n for w in BANKRUPTCY_WORDS) for n in names):
        stage = "CONVERTED_TO_BANKRUPTCY"
    elif any(any(w in n for w in CLOSED_WORDS) for n in names):
        stage = "CLOSED"
    rj_grant = any(any(w in n for w in RJ_GRANT_WORDS) for n in names)
    da = src.get("dataAjuizamento") or ""
    filing = f"{da[:4]}-{da[4:6]}-{da[6:8]}" if len(da) >= 8 else None
    oj = src.get("orgaoJulgador") or {}
    return {
        "case_number": src.get("numeroProcesso"),
        "court": (src.get("tribunal") or "").upper(),
        "chamber": oj.get("nome"),
        "chamber_code": oj.get("codigo"),
        "grau": src.get("grau"),
        "filing_date": filing,
        "stage": stage,
        "rj_granted_signal": rj_grant,
        "last_movement": dates[-1][:10] if dates else None,
        "n_movements": len(movs),
        "movements_json": json.dumps([{"c": m.get("codigo"), "n": m.get("nome"), "d": (m.get("dataHora") or "")[:10]}
                                      for m in movs[-40:]], ensure_ascii=False),
        "system": (src.get("sistema") or {}).get("nome"),
        "source_url": URL.format(court=(src.get("tribunal") or "").lower()),
        "datajud_updated": src.get("dataHoraUltimaAtualizacao"),
    }
