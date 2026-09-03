"""Step B — creditor-list and plan documents from judicial-administrator portals.

Every AJ site has the same shape: an index of cases → a page per case → PDFs. We crawl each
domain breadth-first, following only links that look like case navigation, collect PDF links,
classify them by filename/link text (creditor list · plan · homologation · skip), and download
the useful ones with a per-domain cap. Files are content-addressed (sha1) and recorded in
raw_documents with the page they came from and any CNJ case numbers printed on that page.
"""
from __future__ import annotations

import hashlib
import html
import os
import re
import time
import unicodedata
from urllib.parse import urljoin, urlparse, urldefrag

import requests

from ..db import now

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) anteparo-index/0.1 (creditor-list harvest; contact via repo)"}
CNJ_RE = re.compile(r"\b\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}\b")
HREF_RE = re.compile(r"""<a\b[^>]*?href\s*=\s*["']([^"'#]+)["'][^>]*>(.*?)</a>""", re.I | re.S)
TAG_RE = re.compile(r"<[^>]+>")
TITLE_RE = re.compile(r"<title>(.*?)</title>", re.I | re.S)

NAV_RE = re.compile(r"processo|caso|recupera|falenc|cliente|lista|credor|\brj\b|ativo|edital|andamento|documento|page/|pagina|/p/\d", re.I)
LIST_RE = re.compile(r"credor|qgc|quadro.?geral|relac|lista|edital|art\.?\s?7|art\.?\s?52|art\.?\s?51", re.I)
LIST_STRONG_RE = re.compile(r"qgc|quadro.?geral|(?:lista|relac|rela[cç][aã]o|edital)[^|]{0,40}credor|credor[^|]{0,30}(?:lista|relac)", re.I)
HARD_SKIP_RE = re.compile(r"\bata\b|assembleia|\bagc\b|relat[oó]rio|\brma\b|honor|parecer|balan|boleto|leil|impugna|objec", re.I)
PLAN_RE = re.compile(r"plano|\bprj\b|aditivo|modificativ", re.I)
HOMOLOG_RE = re.compile(r"homolog|concess|senten[cç]a", re.I)
SKIP_RE = re.compile(r"relat[oó]rio|\brma\b|mensal|honor|parecer|\bata\b|balan|demonstra|boleto|manual|faq|curr[ií]cul|"
                     r"convoca|assembleia|agc|laudo|avalia|leil|alien|impugna|objec|habilita[cç][aã]o\b|contesta|"
                     r"peti[cç][aã]o|manifesta|decis[aã]o|despacho|of[ií]cio|procura|certid|cronograma", re.I)


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in s if not unicodedata.combining(c)).lower()


def classify_link(href: str, text: str) -> tuple[str, int]:
    """→ (kind, score). kind ∈ LIST / PLAN / HOMOLOG / SKIP / OTHER."""
    s = _norm(href + " " + text)
    if HARD_SKIP_RE.search(s):
        return "SKIP", 0
    if SKIP_RE.search(s) and not LIST_STRONG_RE.search(s):
        return "SKIP", 0
    if LIST_STRONG_RE.search(s):
        return "LIST", 3
    if PLAN_RE.search(s):
        return "PLAN", 2
    if HOMOLOG_RE.search(s):
        return "HOMOLOG", 2
    if LIST_RE.search(s):
        return "LIST", 1
    return "OTHER", 0


class Crawler:
    def __init__(self, db, out_dir: str, max_pages=120, max_docs=60, delay=0.6, log=print):
        self.db, self.out_dir, self.max_pages, self.max_docs, self.delay, self.log = db, out_dir, max_pages, max_docs, delay, log
        self.s = requests.Session()
        self.s.headers.update(UA)

    # ---------------------------------------------------------------- fetch
    def _get(self, url, stream=False):
        try:
            r = self.s.get(url, timeout=40, allow_redirects=True, stream=stream)
            return r
        except requests.RequestException as e:
            return e

    def _robots_disallow(self, base):
        try:
            r = self.s.get(urljoin(base, "/robots.txt"), timeout=15)
            if r.status_code != 200:
                return []
            dis, ua_all = [], False
            for ln in r.text.splitlines():
                ln = ln.strip()
                if ln.lower().startswith("user-agent:"):
                    ua_all = ln.split(":", 1)[1].strip() == "*"
                elif ua_all and ln.lower().startswith("disallow:"):
                    p = ln.split(":", 1)[1].strip()
                    if p and p != "/":
                        dis.append(p)
            return dis
        except requests.RequestException:
            return []

    # ---------------------------------------------------------------- download
    def download(self, url, domain, page_url="", page_title="", link_text="", case_numbers="", kind="LIST"):
        row = self.db.execute("SELECT sha1, path FROM raw_documents WHERE url=?", (url,)).fetchone()
        if row:
            return row[0]
        r = self._get(url, stream=True)
        if not isinstance(r, requests.Response):
            self._record(None, url, domain, page_url, page_title, link_text, case_numbers, kind, None, 0, "", 0, str(r)[:200])
            return None
        ct = r.headers.get("content-type", "")
        try:
            data = r.raw.read(40_000_000, decode_content=True)
        except Exception as e:  # noqa: BLE001
            self._record(None, url, domain, page_url, page_title, link_text, case_numbers, kind, None, 0, ct, r.status_code, str(e)[:200])
            return None
        if r.status_code != 200 or not data.startswith(b"%PDF"):
            self._record(None, url, domain, page_url, page_title, link_text, case_numbers, kind, None, len(data), ct, r.status_code, "not a PDF")
            return None
        sha1 = hashlib.sha1(data).hexdigest()
        d = os.path.join(self.out_dir, re.sub(r"[^\w.-]", "_", domain))
        os.makedirs(d, exist_ok=True)
        path = os.path.join(d, sha1[:16] + ".pdf")
        if not os.path.exists(path):
            with open(path, "wb") as f:
                f.write(data)
        existing = self.db.execute("SELECT sha1 FROM raw_documents WHERE sha1=?", (sha1,)).fetchone()
        if existing:   # same file under another URL — keep the first record, note the alias
            self.db.execute("UPDATE raw_documents SET error=COALESCE(error,'')||' alias:'||? WHERE sha1=?", (url[:200], sha1))
            self.db.commit()
            return sha1
        self._record(sha1, url, domain, page_url, page_title, link_text, case_numbers, kind, path, len(data), ct, 200, None)
        return sha1

    def _record(self, sha1, url, domain, page_url, page_title, link_text, case_numbers, kind, path, size, ct, status, err):
        key = sha1 or ("ERR:" + hashlib.sha1(url.encode()).hexdigest())
        self.db.execute("INSERT OR REPLACE INTO raw_documents VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (key, url, domain, page_url, page_title[:200], link_text[:200], case_numbers, kind, path, size, ct[:80], now(), status, err))
        self.db.commit()

    # ---------------------------------------------------------------- crawl one domain
    def crawl_domain(self, start_url: str):
        domain = urlparse(start_url).netloc.lower().replace("www.", "")
        if "s3.amazonaws.com" in domain or "digitaloceanspaces.com" in domain:
            return self._crawl_bucket(start_url, domain)
        disallow = self._robots_disallow(start_url)
        seen, queue = set(), [(start_url, 0)]
        pdf_candidates = {}  # url -> (score, kind, page_url, title, text, cases)
        pages = 0
        while queue and pages < self.max_pages:
            url, depth = queue.pop(0)
            url = urldefrag(url)[0]
            if url in seen:
                continue
            seen.add(url)
            path = urlparse(url).path or "/"
            if any(path.startswith(p) for p in disallow):
                continue
            r = self._get(url)
            time.sleep(self.delay)
            if not isinstance(r, requests.Response) or r.status_code != 200:
                continue
            ct = r.headers.get("content-type", "")
            if "pdf" in ct:
                pdf_candidates.setdefault(url, (1, "LIST", url, "", "", ""))
                continue
            if "html" not in ct and "xml" not in ct:
                continue
            pages += 1
            body = r.text
            title = html.unescape(TAG_RE.sub("", (TITLE_RE.search(body) or [None, ""])[1] if TITLE_RE.search(body) else "")).strip()
            text = TAG_RE.sub(" ", body)
            cases = ",".join(sorted(set(CNJ_RE.findall(text)))[:6])
            for href, inner in HREF_RE.findall(body):
                href = html.unescape(href.strip())
                if href.startswith(("mailto:", "tel:", "javascript:")):
                    continue
                absu = urljoin(url, href)
                pu = urlparse(absu)
                if pu.scheme not in ("http", "https"):
                    continue
                ltext = html.unescape(TAG_RE.sub(" ", inner)).strip()[:200]
                is_pdf = pu.path.lower().endswith(".pdf") or "download" in pu.path.lower() or "arquivo" in pu.path.lower() or ".pdf" in absu.lower()
                if is_pdf:
                    kind, score = classify_link(absu, ltext)
                    if kind in ("LIST", "PLAN", "HOMOLOG"):
                        prev = pdf_candidates.get(absu)
                        if not prev or prev[0] < score:
                            pdf_candidates[absu] = (score, kind, url, title, ltext, cases)
                    continue
                same = pu.netloc.lower().replace("www.", "") == domain
                if same and depth < 3 and absu not in seen and (NAV_RE.search(absu) or NAV_RE.search(_norm(ltext))):
                    queue.append((absu, depth + 1))
        ranked = sorted(pdf_candidates.items(), key=lambda kv: -kv[1][0])
        got = 0
        for absu, (score, kind, page_url, title, ltext, cases) in ranked:
            if got >= self.max_docs:
                break
            sha = self.download(absu, domain, page_url, title, ltext, cases, kind)
            if sha:
                got += 1
            time.sleep(self.delay)
        self.log(f"[{domain}] pages={pages} pdf_links={len(pdf_candidates)} downloaded={got}")
        return got

    def _crawl_bucket(self, start_url, domain):
        base = f"{urlparse(start_url).scheme}://{urlparse(start_url).netloc}/"
        r = self._get(base)
        if not isinstance(r, requests.Response) or r.status_code != 200:
            self.log(f"[{domain}] bucket listing failed")
            return 0
        keys = re.findall(r"<Key>(.*?)</Key>", r.text)
        cands = []
        for k in keys:
            if not k.lower().endswith(".pdf"):
                continue
            kind, score = classify_link(k, "")
            if kind in ("LIST", "PLAN", "HOMOLOG"):
                cands.append((score, kind, k))
        cands.sort(reverse=True)
        got = 0
        for score, kind, k in cands[: self.max_docs]:
            if self.download(base + k, domain, base, "", k, "", kind):
                got += 1
            time.sleep(self.delay)
        self.log(f"[{domain}] bucket keys={len(keys)} pdf_cands={len(cands)} downloaded={got}")
        return got
