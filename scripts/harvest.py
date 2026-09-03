"""Step B driver — download the direct URLs, then crawl every seed domain."""
from __future__ import annotations

import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from anteparo.db import connect
from anteparo.sources.portals import Crawler

import argparse, threading, time

ap = argparse.ArgumentParser()
ap.add_argument("--max-pages", type=int, default=120)
ap.add_argument("--max-docs", type=int, default=60)
ap.add_argument("--workers", type=int, default=6)
ap.add_argument("--only", default=None, help="substring filter on seed domains")
ap.add_argument("--skip-direct", action="store_true")
args = ap.parse_args()

lock = threading.Lock()
def log(msg):
    with lock:
        print(time.strftime("%H:%M:%S"), msg, flush=True)

db_path = str(ROOT / "data" / "anteparo.sqlite")
out = str(ROOT / "data" / "pdfs")

if not args.skip_direct:
    db = connect(db_path)
    c = Crawler(db, out, log=log)
    n = 0
    for u in (ROOT / "data/seeds/direct_urls.txt").read_text().split():
        if "comunicaapi.pje.jus.br" in u:
            continue
        dom = urlparse(u).netloc.lower().replace("www.", "")
        if c.download(u, dom, "", "", "direct-from-sheet", "", "LIST"):
            n += 1
    log(f"direct URLs downloaded: {n}")

seeds = [s.strip() for s in (ROOT / "data/seeds/domains.txt").read_text().splitlines() if s.strip() and not s.startswith("#")]
if args.only:
    seeds = [s for s in seeds if args.only in s]

def work(seed):
    db = connect(db_path)          # one connection per thread
    c = Crawler(db, out, max_pages=args.max_pages, max_docs=args.max_docs, log=log)
    try:
        return c.crawl_domain(seed)
    except Exception as e:  # noqa: BLE001
        log(f"[{seed}] FAILED {type(e).__name__}: {e}")
        return 0

with ThreadPoolExecutor(args.workers) as ex:
    total = sum(ex.map(work, seeds))
log(f"HARVEST DONE: {total} documents from {len(seeds)} domains")
