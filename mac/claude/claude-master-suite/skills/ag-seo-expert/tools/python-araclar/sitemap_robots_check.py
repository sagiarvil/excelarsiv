#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sitemap_robots_check.py — robots.txt, sitemap.xml, llms.txt, agent-card.json denetimi.
Kullanım: python sitemap_robots_check.py <origin-URL> [--max=40] [--json]
  <origin-URL>: https://alan.tld  (yol verilirse origin'e indirilir)
"""

import sys
import re
import json
import ssl
import urllib.request
import urllib.parse
import urllib.error
from typing import List, Dict


class FetchResult:
    def __init__(self, status: int, headers, body: str, url: str):
        self.status = status
        self.headers = headers
        self.body = body
        self.url = url


def fetch_url(target: str, method: str = 'GET', timeout: int = 20) -> FetchResult:
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    req = urllib.request.Request(
        target,
        headers={
            'User-Agent': 'super-seo-motor/1.0 (+audit)',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        },
        method=method
    )

    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
            body = resp.read().decode('utf-8', errors='replace')
            return FetchResult(status=resp.status, headers=resp.headers, body=body, url=resp.geturl())
    except urllib.error.HTTPError as e:
        body = e.read().decode('utf-8', errors='replace')
        return FetchResult(status=e.code, headers=e.headers, body=body, url=target)
    except Exception as e:
        raise e


class Report:
    def __init__(self, title: str):
        self.title = title
        self.rows: List[tuple] = []
        self.pass_count = 0
        self.fail_count = 0
        self.warn_count = 0

    def pass_(self, category: str, msg: str):
        self.rows.append(('PASS', category, msg))
        self.pass_count += 1

    def fail(self, category: str, msg: str):
        self.rows.append(('FAIL', category, msg))
        self.fail_count += 1

    def warn(self, category: str, msg: str):
        self.rows.append(('WARN', category, msg))
        self.warn_count += 1

    def info(self, category: str, msg: str):
        self.rows.append(('INFO', category, msg))

    def summary(self) -> Dict[str, int]:
        return {'pass': self.pass_count, 'fail': self.fail_count, 'warn': self.warn_count, 'total': len(self.rows)}

    def to_text(self) -> str:
        lines = [f"# {self.title}", ""]
        for status, category, msg in self.rows:
            lines.append(f"- [{status}] {category} — {msg}")
        lines.append("")
        lines.append(f"SONUÇ: PASS={self.pass_count}  FAIL={self.fail_count}  WARN={self.warn_count}")
        return "\n".join(lines)

    def print_report(self) -> int:
        print(self.to_text())
        return 0 if self.fail_count == 0 else 1


def parse_args():
    out = {'_': [], 'max': '40', 'json': False}
    for a in sys.argv[1:]:
        if a.startswith('--'):
            parts = a[2:].split('=', 1)
            k = parts[0]
            v = parts[1] if len(parts) > 1 else True
            out[k] = v
        else:
            out['_'].append(a)
    return out


def main():
    args = parse_args()

    if not args['_']:
        sys.stderr.write("Kullanım: python sitemap_robots_check.py <https://alan.tld> [--max=40] [--json]\n")
        sys.exit(2)

    raw_url = args['_'][0]
    if not re.match(r'^https?://', raw_url, re.IGNORECASE):
        raw_url = 'https://' + raw_url

    parsed = urllib.parse.urlparse(raw_url)
    origin = f"{parsed.scheme}://{parsed.netloc}"

    try:
        max_count = int(args['max'])
    except (ValueError, TypeError):
        max_count = 40

    is_json = bool(args.get('json', False))

    rep = Report(f"ROBOTS / SITEMAP / LLM YÜZEYLERİ — {origin}")

    # 1. robots.txt kontrolü
    sitemaps = []
    try:
        r = fetch_url(origin + '/robots.txt')
        if r.status == 200 and r.body.strip():
            byte_len = len(r.body.encode('utf-8'))
            rep.pass_('ROBOTS', f"robots.txt bulundu ({byte_len} bayt)")
            if re.search(r'Disallow:\s*/\s*$', r.body, re.IGNORECASE | re.MULTILINE):
                rep.fail('ROBOTS', 'Tüm site Disallow: / ile kapalı')
            sitemaps = re.findall(r'^\s*Sitemap:\s*(\S+)', r.body, re.IGNORECASE | re.MULTILINE)
            if not sitemaps:
                rep.warn('ROBOTS', 'robots.txt içinde Sitemap: satırı yok')
            else:
                rep.pass_('ROBOTS', f"{len(sitemaps)} sitemap bildirimi")
        else:
            rep.fail('ROBOTS', f"robots.txt yok / HTTP {r.status}")
    except Exception as e:
        rep.fail('ROBOTS', f"robots.txt alınamadı: {e}")

    if not sitemaps:
        sitemaps = [origin + '/sitemap.xml']

    # 2. sitemap(ler) kontrolü
    urls = []
    for sm in sitemaps[:5]:
        try:
            r = fetch_url(sm)
            if r.status != 200:
                rep.fail('SITEMAP', f"{sm} → HTTP {r.status}")
                continue
            if not re.search(r'<(urlset|sitemapindex)\b', r.body, re.IGNORECASE):
                rep.fail('SITEMAP', f"{sm} geçerli XML değil")
                continue
            is_index = bool(re.search(r'<sitemapindex\b', r.body, re.IGNORECASE))
            locs = re.findall(r'<loc>\s*([^<\s]+)\s*</loc>', r.body, re.IGNORECASE)
            rep.pass_('SITEMAP', f"{sm} → {len(locs)} {'alt-sitemap' if is_index else 'URL'}")

            if is_index:
                for sub in locs[:5]:
                    try:
                        rs = fetch_url(sub)
                        sub_locs = re.findall(r'<loc>\s*([^<\s]+)\s*</loc>', rs.body, re.IGNORECASE)
                        urls.extend(sub_locs)
                    except Exception:
                        rep.warn('SITEMAP', f"alt sitemap alınamadı: {sub}")
            else:
                urls.extend(locs)

            if re.search(r'<lastmod>\s*</lastmod>', r.body, re.IGNORECASE):
                rep.warn('SITEMAP', 'boş <lastmod> var')
        except Exception as e:
            rep.fail('SITEMAP', f"{sm} alınamadı: {e}")

    # 3. URL örnekleme — HTTP durum kontrolü
    sample = list(dict.fromkeys(urls))[:max_count]
    bad = 0
    for u in sample:
        try:
            r = fetch_url(u, method='GET')
            if r.status >= 400:
                rep.fail('URL', f"{r.status} ← {u}")
                bad += 1
            elif r.url.rstrip('/') != u.rstrip('/'):
                rep.warn('URL', f"redirect: {u} → {r.url}")
        except Exception:
            rep.fail('URL', f"erişilemedi: {u}")
            bad += 1

    if sample:
        if bad:
            rep.warn('URL', f"{len(sample)} URL örneklendi, {bad} hatalı")
        else:
            rep.pass_('URL', f"{len(sample)} URL örneklendi, {bad} hatalı")

    # 4. LLM / makine yüzeyleri
    for p in ['/llms.txt', '/llms-full.txt', '/.well-known/agent-card.json', '/.well-known/ai-plugin.json']:
        try:
            r = fetch_url(origin + p)
            if r.status == 200 and r.body.strip():
                rep.pass_('LLM', f"{p} mevcut")
            else:
                rep.warn('LLM', f"{p} yok (HTTP {r.status})")
        except Exception:
            rep.warn('LLM', f"{p} alınamadı")

    # 5. Çıktı
    if is_json:
        out = {"summary": rep.summary(), "text": rep.to_text()}
        print(json.dumps(out, indent=2, ensure_ascii=False))
        sys.exit(1 if rep.summary()['fail'] else 0)

    sys.exit(rep.print_report())


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        sys.stderr.write(f"HATA: {e}\n")
        sys.exit(2)
