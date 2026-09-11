#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
schema_validator.py — Google yapılandırılmış veri / Rich Results doğrulayıcı.

Kullanım:
  python schema_validator.py <URL | yerel.html> [--json]

Denetler:
 - Tüm JSON-LD blokları geçerli JSON mu, @context schema.org mu
 - @graph yapısı, @id benzersiz mi, dahili @id referansları çözülüyor mu
 - Tip bazında Google zorunlu/önerilen alanlar (Organization, WebSite, WebPage,
   BreadcrumbList, Product, Offer/AggregateOffer, SoftwareApplication, Service,
   Article/BlogPosting, FAQPage, LocalBusiness, ImageObject)
 - Google politika riski: uydurma aggregateRating/review, boş/placeholder sameAs,
   fiyatsız Offer, geçmiş priceValidUntil
"""

import sys
import os
import json
import re
import urllib.request
import urllib.error
from datetime import datetime, timezone

REQUIRED = {
    "Organization": {"must": ["name", "url"], "should": ["logo", "sameAs"]},
    "WebSite": {"must": ["name", "url"], "should": ["potentialAction"]},
    "WebPage": {"must": ["name"], "should": ["url", "isPartOf"]},
    "BreadcrumbList": {"must": ["itemListElement"], "should": []},
    "ListItem": {"must": ["position", "name"], "should": ["item"]},
    "Product": {"must": ["name"], "should": ["image", "description", "offers", "brand"]},
    "Offer": {"must": ["price", "priceCurrency"], "should": ["availability", "url", "priceValidUntil"]},
    "AggregateOffer": {"must": ["lowPrice", "priceCurrency"], "should": ["highPrice", "offerCount"]},
    "SoftwareApplication": {"must": ["name", "applicationCategory"], "should": ["offers", "operatingSystem", "aggregateRating"]},
    "Service": {"must": ["name"], "should": ["provider", "areaServed", "serviceType"]},
    "Article": {"must": ["headline"], "should": ["image", "datePublished", "author", "publisher"]},
    "BlogPosting": {"must": ["headline"], "should": ["image", "datePublished", "author", "publisher"]},
    "NewsArticle": {"must": ["headline"], "should": ["image", "datePublished", "author", "publisher"]},
    "FAQPage": {"must": ["mainEntity"], "should": []},
    "Question": {"must": ["name", "acceptedAnswer"], "should": []},
    "Answer": {"must": ["text"], "should": []},
    "LocalBusiness": {"must": ["name", "address"], "should": ["telephone", "openingHoursSpecification", "geo"]},
    "ImageObject": {"must": ["url"], "should": ["width", "height"]},
    "AggregateRating": {"must": ["ratingValue", "ratingCount"], "should": ["bestRating"]},
    "Review": {"must": ["reviewRating", "author"], "should": []},
}

PLACEHOLDER_RX = re.compile(r"(example\.com|yourdomain|placeholder|xxx+|lorem|test\.test|\{\{)", re.IGNORECASE)


class Report:
    """Doğrulama sonuçlarını toplayan ve biçimlendiren rapor sınıfı."""
    def __init__(self, title: str):
        self.title = title
        self.rows = []
        self.pass_count = 0
        self.fail_count = 0
        self.warn_count = 0

    def pass_(self, test_id: str, msg: str):
        self.rows.append(("PASS", test_id, msg))
        self.pass_count += 1

    def fail(self, test_id: str, msg: str):
        self.rows.append(("FAIL", test_id, msg))
        self.fail_count += 1

    def warn(self, test_id: str, msg: str):
        self.rows.append(("WARN", test_id, msg))
        self.warn_count += 1

    def info(self, test_id: str, msg: str):
        self.rows.append(("INFO", test_id, msg))

    def summary(self):
        return {
            "pass": self.pass_count,
            "fail": self.fail_count,
            "warn": self.warn_count,
            "total": len(self.rows)
        }

    def to_text(self) -> str:
        lines = [f"# {self.title}", ""]
        for status, test_id, msg in self.rows:
            lines.append(f"- [{status}] {test_id} — {msg}")
        lines.append("")
        lines.append(f"SONUÇ: PASS={self.pass_count}  FAIL={self.fail_count}  WARN={self.warn_count}")
        return "\n".join(lines)

    def print_report(self) -> int:
        print(self.to_text())
        return 0 if self.fail_count == 0 else 1


def type_list(node) -> list:
    if not isinstance(node, dict):
        return []
    t = node.get("@type")
    if isinstance(t, list):
        return [item for item in t if item]
    return [t] if t else []


def has_field(node, key: str) -> bool:
    if not isinstance(node, dict) or key not in node:
        return False
    val = node[key]
    if val is None or val == "":
        return False
    if isinstance(val, list) and len(val) == 0:
        return False
    return True


def is_date_in_past(date_str: str) -> bool:
    if not isinstance(date_str, str) or not date_str.strip():
        return False
    cleaned = date_str.strip().replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(cleaned)
        if dt.tzinfo is None:
            return dt < datetime.now()
        return dt < datetime.now(timezone.utc)
    except Exception:
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", cleaned)
        if m:
            try:
                y, mo, d = map(int, m.groups())
                return datetime(y, mo, d) < datetime.now()
            except Exception:
                pass
    return False


def fetch_url(target: str, timeout: int = 20):
    req = urllib.request.Request(
        target,
        headers={
            "User-Agent": "super-seo-motor/1.0 (+audit)",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        }
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            content = resp.read()
            charset = resp.headers.get_content_charset() or "utf-8"
            html = content.decode(charset, errors="replace")
            return {"html": html, "status": resp.status, "finalUrl": resp.geturl(), "headers": dict(resp.headers)}
    except urllib.error.HTTPError as e:
        content = e.read()
        charset = e.headers.get_content_charset() or "utf-8" if e.headers else "utf-8"
        html = content.decode(charset, errors="replace")
        return {"html": html, "status": e.code, "finalUrl": e.geturl() if hasattr(e, "geturl") else target, "headers": dict(e.headers) if e.headers else {}}


def load_html(input_target: str):
    if re.match(r"^https?://", input_target, re.IGNORECASE):
        return fetch_url(input_target)
    path_resolved = os.path.abspath(input_target)
    if not os.path.exists(path_resolved):
        raise FileNotFoundError(f"Dosya bulunamadı: {path_resolved}")
    with open(path_resolved, "r", encoding="utf-8-sig", errors="replace") as f:
        html = f.read()
    return {"html": html, "status": 200, "finalUrl": "file://" + path_resolved, "headers": {}}


def extract_json_ld(html: str) -> list:
    pattern = re.compile(r'<script[^>]*type=["\']application/ld\+json["\'][^>]*>([\s\S]*?)</script>', re.IGNORECASE)
    blocks = []
    for m in pattern.finditer(html):
        raw = m.group(1).strip()
        try:
            data = json.loads(raw)
            blocks.append({"ok": True, "raw": raw, "data": data})
        except Exception as e:
            blocks.append({"ok": False, "raw": raw, "error": str(e)})
    return blocks


def flatten_nodes(parsed_blocks: list) -> list:
    nodes = []
    for b in parsed_blocks:
        if not b.get("ok"):
            continue
        d = b.get("data")
        if isinstance(d, list):
            for n in d:
                nodes.append(n)
        elif isinstance(d, dict) and isinstance(d.get("@graph"), list):
            for n in d["@graph"]:
                nodes.append(n)
        elif isinstance(d, dict):
            nodes.append(d)
    return nodes


def finish(rep: Report, json_mode: bool):
    if json_mode:
        print(json.dumps({"summary": rep.summary(), "text": rep.to_text()}, indent=2, ensure_ascii=False))
        sys.exit(1 if rep.summary()["fail"] > 0 else 0)
    sys.exit(rep.print_report())


def main():
    args = sys.argv[1:]
    json_mode = False
    cleaned_args = []
    for a in args:
        if a == "--json":
            json_mode = True
        elif a.startswith("--json="):
            json_mode = a.split("=", 1)[1].lower() in ("true", "1")
        else:
            cleaned_args.append(a)

    if not cleaned_args:
        print("Kullanım: python schema_validator.py <URL|dosya> [--json]", file=sys.stderr)
        sys.exit(2)

    input_target = cleaned_args[0]
    loaded = load_html(input_target)
    html = loaded["html"]
    status = loaded["status"]

    rep = Report(f"SCHEMA / RICH RESULTS DOĞRULAMA — {input_target}")
    if status and status >= 400:
        rep.warn("HTTP", f"Sayfa HTTP {status} döndü")

    blocks = extract_json_ld(html)
    if len(blocks) == 0:
        rep.fail("LD-00", "Sayfada hiç JSON-LD bloğu yok")
        finish(rep, json_mode)
        return

    rep.pass_("LD-00", f"{len(blocks)} JSON-LD bloğu bulundu")

    for i, b in enumerate(blocks):
        if not b["ok"]:
            rep.fail("LD-PARSE", f"Blok #{i + 1} JSON hatası: {b['error']}")
            continue
        data = b["data"]
        ctx = data.get("@context") if isinstance(data, dict) else None
        if ctx is None and isinstance(data, dict) and data.get("@graph"):
            ctx = "(graph)"
        ctx_dump = json.dumps(data.get("@context") if isinstance(data, dict) and data.get("@context") is not None else "")
        if not re.search(r"schema\.org", ctx_dump, re.IGNORECASE):
            rep.warn("LD-CTX", f"Blok #{i + 1} @context schema.org değil: {json.dumps(ctx, ensure_ascii=False)}")

    nodes = flatten_nodes(blocks)
    rep.info("LD-CNT", f"{len(nodes)} node (graph düzleştirildi)")

    # @id benzersizlik + referans çözümü
    ids = {}
    for n in nodes:
        if isinstance(n, dict) and n.get("@id"):
            node_id = n["@id"]
            if node_id in ids:
                rep.warn("LD-ID", f"Tekrar eden @id: {node_id}")
            ids[node_id] = n

    ref_rx = re.compile(r'"@id"\s*:\s*"([^"]+)"')
    flat_json = json.dumps(nodes, ensure_ascii=False)
    seen_ref = set(ref_rx.findall(flat_json))
    for ref in seen_ref:
        if ref not in ids and "#" in ref:
            rep.warn("LD-REF", f"Çözülemeyen dahili @id referansı: {ref}")

    # Tip bazında alan denetimi
    found_types = set()
    for n in nodes:
        if not isinstance(n, dict):
            continue
        for t in type_list(n):
            found_types.add(t)
            spec = REQUIRED.get(t)
            if not spec:
                continue
            for k in spec["must"]:
                if not has_field(n, k):
                    rep.fail("REQ", f"{t}: zorunlu alan eksik → {k}")
            for k in spec["should"]:
                if not has_field(n, k):
                    rep.warn("REC", f"{t}: önerilen alan eksik → {k}")

        # Politika riskleri
        if has_field(n, "sameAs"):
            val = n["sameAs"]
            arr = val if isinstance(val, list) else [val]
            for s in arr:
                if not isinstance(s, str) or not re.match(r"^https?://.+\..+", s) or PLACEHOLDER_RX.search(s):
                    rep.warn("POL-SAMEAS", f"Geçersiz/placeholder sameAs: {json.dumps(s, ensure_ascii=False)}")

        types_for_node = type_list(n)
        if "Offer" in types_for_node:
            p = None
            try:
                p = float(n.get("price"))
            except (TypeError, ValueError):
                pass
            if p is None or not (p > 0):
                rep.fail("POL-OFFER", "Offer.price ≤ 0 / geçersiz — Google geçersiz Offer sayar")
            pvu = n.get("priceValidUntil")
            if pvu and is_date_in_past(str(pvu)):
                rep.warn("POL-PVU", f"priceValidUntil geçmişte: {pvu}")

        if "AggregateRating" in types_for_node:
            rc_val = n.get("ratingCount") if n.get("ratingCount") is not None else n.get("reviewCount")
            rc = None
            try:
                rc = float(rc_val)
            except (TypeError, ValueError):
                pass
            if rc is None or not (rc > 0):
                rep.fail("POL-RATING", "aggregateRating var ama ratingCount/reviewCount yok — uydurma veri riski")

    # Taban blok beklentisi
    if "Organization" not in found_types:
        rep.warn("BASE", "Organization node yok (taban blok beklenir)")
    if "WebSite" not in found_types:
        rep.warn("BASE", "WebSite node yok (taban blok beklenir)")

    rep.info("TYPES", f"Bulunan tipler: {', '.join(sorted(found_types))}")

    finish(rep, json_mode)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"HATA: {e}", file=sys.stderr)
        sys.exit(2)
