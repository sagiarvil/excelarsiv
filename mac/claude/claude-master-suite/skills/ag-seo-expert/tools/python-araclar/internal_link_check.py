#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
internal_link_check.py — Sığ tarama ile kırık iç link + zayıf çapa metni denetimi.
Kullanım: python internal_link_check.py <başlangıç-URL> [--max=60] [--depth=1] [--json]
"""

import sys
import re
import json
import ssl
from collections import deque
import urllib.request
import urllib.parse
import urllib.error

GENERIC = re.compile(
    r'^(t[ıi]kla(y[ıi]n)?|buraya?|devam[ıi]?|read more|daha fazla|link|here|more|>>|»)$',
    re.IGNORECASE
)

SKIP_HREF = re.compile(r'^(mailto:|tel:|javascript:|#)', re.IGNORECASE)
A_HREF_RX = re.compile(r'<a\b[^>]*\bhref=["\']([^"\']+)["\'][^>]*>(.*?)</a>', re.IGNORECASE | re.DOTALL)


def strip_tags(s: str) -> str:
    if not s:
        return ""
    text = re.sub(r'<[^>]+>', ' ', str(s))
    return re.sub(r'\s+', ' ', text).strip()


def parse_args(argv=None):
    if argv is None:
        argv = sys.argv[1:]
    out = {"_": []}
    for a in argv:
        if a.startswith("--"):
            parts = a[2:].split("=", 1)
            key = parts[0]
            val = parts[1] if len(parts) > 1 else True
            out[key] = val
        else:
            out["_"].append(a)
    return out


class NoRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def fetch_url(target: str, timeout: float = 20.0, redirects: int = 0):
    if redirects >= 6:
        raise RuntimeError(f"Çok fazla yönlendirme: {target}")

    parsed = urllib.parse.urlsplit(target)
    if parsed.scheme not in ("http", "https"):
        raise ValueError(f"Geçersiz URL şeması: {target}")

    ssl_context = ssl.create_default_context()
    ssl_context.check_hostname = False
    ssl_context.verify_mode = ssl.CERT_NONE

    req = urllib.request.Request(
        target,
        headers={
            "User-Agent": "super-seo-motor/1.0 (+audit)",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        },
    )

    opener = urllib.request.build_opener(
        urllib.request.HTTPSHandler(context=ssl_context),
        NoRedirectHandler
    )

    try:
        with opener.open(req, timeout=timeout) as res:
            status = getattr(res, 'status', res.code)
            headers = dict(res.headers)
            body = res.read().decode('utf-8', errors='replace')
            loc = headers.get('Location') or headers.get('location')

            if 300 <= status < 400 and loc and redirects < 6:
                next_url = urllib.parse.urljoin(target, loc)
                return fetch_url(next_url, timeout=timeout, redirects=redirects + 1)

            return {"status": status, "headers": headers, "body": body, "url": target}
    except urllib.error.HTTPError as e:
        headers = dict(e.headers)
        loc = headers.get('Location') or headers.get('location')

        if 300 <= e.code < 400 and loc and redirects < 6:
            next_url = urllib.parse.urljoin(target, loc)
            return fetch_url(next_url, timeout=timeout, redirects=redirects + 1)

        body = ""
        try:
            body = e.read().decode('utf-8', errors='replace')
        except Exception:
            pass

        return {"status": e.code, "headers": headers, "body": body, "url": target}


class Report:
    def __init__(self, title: str):
        self.title = title
        self.rows = []
        self.pass_count = 0
        self.fail_count = 0
        self.warn_count = 0

    def pass_(self, id_: str, msg: str):
        self.rows.append(('PASS', id_, msg))
        self.pass_count += 1

    def fail(self, id_: str, msg: str):
        self.rows.append(('FAIL', id_, msg))
        self.fail_count += 1

    def warn(self, id_: str, msg: str):
        self.rows.append(('WARN', id_, msg))
        self.warn_count += 1

    def info(self, id_: str, msg: str):
        self.rows.append(('INFO', id_, msg))

    def summary(self):
        return {
            "pass": self.pass_count,
            "fail": self.fail_count,
            "warn": self.warn_count,
            "total": len(self.rows),
        }

    def to_text(self):
        lines = [f"# {self.title}", ""]
        for s, id_, msg in self.rows:
            lines.append(f"- [{s}] {id_} — {msg}")
        lines.append("")
        lines.append(f"SONUÇ: PASS={self.pass_count}  FAIL={self.fail_count}  WARN={self.warn_count}")
        return "\n".join(lines)

    def print_report(self):
        print(self.to_text())
        return 0 if self.fail_count == 0 else 1


def main():
    args = parse_args()
    if not args["_"]:
        print("Kullanım: python internal_link_check.py <URL> [--max=60] [--depth=1] [--json]", file=sys.stderr)
        sys.exit(2)

    start_raw = args["_"][0]
    parsed_start = urllib.parse.urlsplit(start_raw)
    if not parsed_start.scheme or not parsed_start.netloc:
        print("HATA: Geçersiz başlangıç URL'si: " + start_raw, file=sys.stderr)
        sys.exit(2)

    origin = f"{parsed_start.scheme}://{parsed_start.netloc}"
    start_url = urllib.parse.urlunsplit((parsed_start.scheme, parsed_start.netloc, parsed_start.path or '/', parsed_start.query, ''))

    try:
        max_pages = int(args.get("max", 60))
    except ValueError:
        max_pages = 60

    try:
        max_depth = int(args.get("depth", 1))
    except ValueError:
        max_depth = 1

    is_json = bool(args.get("json", False))

    rep = Report('İÇ LİNK DENETİMİ — ' + origin)

    seen = set()
    queue = deque([(start_url, 0)])
    status_cache = {}
    checked = 0

    def head(u: str) -> int:
        if u in status_cache:
            return status_cache[u]
        try:
            r = fetch_url(u)
            st = r.get("status", 0)
        except Exception:
            st = 0
        status_cache[u] = st
        return st

    while queue and checked < max_pages:
        url, depth = queue.popleft()
        if url in seen:
            continue
        seen.add(url)
        checked += 1

        try:
            r = fetch_url(url)
            html = r.get("body", "")
            if r.get("status", 0) >= 400:
                rep.fail('PAGE', f"{r.get('status')} ← {url}")
                continue
        except Exception:
            rep.fail('PAGE', 'erişilemedi: ' + url)
            continue

        links = []
        for m in A_HREF_RX.finditer(html):
            links.append({"href": m.group(1), "text": strip_tags(m.group(2))})

        for item in links:
            href = item["href"]
            text = item["text"]

            if SKIP_HREF.search(href):
                continue

            try:
                abs_url = urllib.parse.urljoin(url, href)
                abs_parsed = urllib.parse.urlsplit(abs_url)
                if not abs_parsed.scheme or not abs_parsed.netloc:
                    raise ValueError("Geçersiz URL")
            except Exception:
                rep.warn('HREF', f"geçersiz href: {href} @ {url}")
                continue

            abs_origin = f"{abs_parsed.scheme}://{abs_parsed.netloc}"
            if abs_origin != origin:
                continue

            clean = urllib.parse.urlunsplit((abs_parsed.scheme, abs_parsed.netloc, abs_parsed.path, abs_parsed.query, ''))

            st = head(clean)
            if st == 0:
                rep.fail('LINK', f"erişilemedi: {clean}  (kaynak: {url})")
            elif st >= 400:
                rep.fail('LINK', f"{st}: {clean}  (kaynak: {url})")
            elif st >= 300:
                rep.warn('LINK', f"{st} redirect: {clean}")

            if not text:
                rep.warn('ANCHOR', 'boş çapa metni → ' + clean)
            elif GENERIC.match(text):
                rep.warn('ANCHOR', f'jenerik çapa "{text}" → {clean}')

            if depth < max_depth and clean not in seen and st and st < 300:
                queue.append((clean, depth + 1))

    rep.info('CRAWL', f"{checked} sayfa tarandı, {len(status_cache)} benzersiz iç link kontrol edildi")

    if is_json:
        result = {"summary": rep.summary(), "text": rep.to_text()}
        print(json.dumps(result, indent=2, ensure_ascii=False))
        sys.exit(1 if rep.summary()["fail"] else 0)

    sys.exit(rep.print_report())


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f"HATA: {e}", file=sys.stderr)
        sys.exit(2)
