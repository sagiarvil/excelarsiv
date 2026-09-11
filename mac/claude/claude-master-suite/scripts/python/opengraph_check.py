#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
opengraph_check.py — <head> SEO + Open Graph + Twitter Card denetimi.
Kullanım: python opengraph_check.py <URL|dosya> [--json]
"""

import sys
import os
import re
import json
import urllib.request
import urllib.error
from typing import Optional, Dict, Any, List, Tuple


class Report:
    def __init__(self, title: str):
        self.title = title
        self.rows: List[Tuple[str, str, str]] = []
        self.pass_count = 0
        self.fail_count = 0
        self.warn_count = 0

    def pass_(self, id_: str, msg: str):
        self.rows.append(('PASS', id_, str(msg)))
        self.pass_count += 1

    def fail(self, id_: str, msg: str):
        self.rows.append(('FAIL', id_, str(msg)))
        self.fail_count += 1

    def warn(self, id_: str, msg: str):
        self.rows.append(('WARN', id_, str(msg)))
        self.warn_count += 1

    def info(self, id_: str, msg: str):
        self.rows.append(('INFO', id_, str(msg)))

    def summary(self) -> Dict[str, int]:
        return {'pass': self.pass_count, 'fail': self.fail_count, 'warn': self.warn_count, 'total': len(self.rows)}

    def to_text(self) -> str:
        lines = [f"# {self.title}", ""]
        for s, id_, msg in self.rows:
            lines.append(f"- [{s}] {id_} — {msg}")
        lines.append("")
        lines.append(f"SONUÇ: PASS={self.pass_count}  FAIL={self.fail_count}  WARN={self.warn_count}")
        return "\n".join(lines)

    def print_report(self) -> int:
        print(self.to_text())
        return 0 if self.fail_count == 0 else 1


def parse_args() -> Dict[str, Any]:
    out: Dict[str, Any] = {'_': []}
    for a in sys.argv[1:]:
        if a.startswith('--'):
            parts = a[2:].split('=', 1)
            k = parts[0]
            v = parts[1] if len(parts) > 1 else True
            out[k] = v
        else:
            out['_'].append(a)
    return out


def load_html(input_str: str) -> Dict[str, Any]:
    if re.match(r'^https?://', input_str, re.IGNORECASE):
        req = urllib.request.Request(
            input_str,
            headers={
                'User-Agent': 'super-seo-motor/1.0 (+audit)',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            }
        )
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                status = resp.status if hasattr(resp, 'status') else resp.getcode()
                final_url = resp.geturl()
                content = resp.read()
                charset = resp.headers.get_content_charset() or 'utf-8'
                html = content.decode(charset, errors='replace')
                return {'html': html, 'status': status, 'finalUrl': final_url, 'headers': dict(resp.headers)}
        except urllib.error.HTTPError as e:
            content = e.read() if hasattr(e, 'read') else b''
            charset = e.headers.get_content_charset() if hasattr(e, 'headers') else 'utf-8'
            html = content.decode(charset or 'utf-8', errors='replace')
            return {'html': html, 'status': e.code, 'finalUrl': input_str, 'headers': dict(e.headers) if hasattr(e, 'headers') else {}}
    else:
        abs_path = os.path.abspath(input_str)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Dosya bulunamadı: {abs_path}")
        with open(abs_path, 'r', encoding='utf-8', errors='replace') as f:
            html = f.read()
        return {'html': html, 'status': 200, 'finalUrl': 'file://' + abs_path, 'headers': {}}


def first_title(html: str) -> Optional[str]:
    m = re.search(r'<title[^>]*>(.*?)</title>', html, re.IGNORECASE | re.DOTALL)
    return m.group(1).strip() if m else None


def first_meta_name(name: str, html: str) -> Optional[str]:
    pattern = rf'<meta\b[^>]*(?:name=["\'{re.escape(name)}["\'][^>]*content=["\']([^"\']*)["\']|content=["\']([^"\']*)["\'][^>]*name=["\']{re.escape(name)}["\'])'
    m = re.search(pattern, html, re.IGNORECASE | re.DOTALL)
    if m:
        val = m.group(1) if m.group(1) is not None else m.group(2)
        return val.strip() if val else None
    return None


def first_meta_prop(prop: str, html: str) -> Optional[str]:
    pattern = rf'<meta\b[^>]*(?:property=["\'{re.escape(prop)}["\'][^>]*content=["\']([^"\']*)["\']|content=["\']([^"\']*)["\'][^>]*property=["\']{re.escape(prop)}["\'])'
    m = re.search(pattern, html, re.IGNORECASE | re.DOTALL)
    if m:
        val = m.group(1) if m.group(1) is not None else m.group(2)
        return val.strip() if val else None
    return None


def first_link_rel(rel: str, html: str) -> Optional[str]:
    pattern = rf'<link\b[^>]*(?:rel=["\'{re.escape(rel)}["\'][^>]*href=["\']([^"\']*)["\']|href=["\']([^"\']*)["\'][^>]*rel=["\']{re.escape(rel)}["\'])'
    m = re.search(pattern, html, re.IGNORECASE | re.DOTALL)
    if m:
        val = m.group(1) if m.group(1) is not None else m.group(2)
        return val.strip() if val else None
    return None


def main():
    args = parse_args()
    input_val = args['_'][0] if args['_'] else None
    if not input_val:
        sys.stderr.write('Kullanım: python opengraph_check.py <URL|dosya> [--json]\n')
        sys.exit(2)

    data = load_html(input_val)
    html = data['html']
    status = data['status']
    final_url = data['finalUrl']

    rep = Report('OPEN GRAPH / META DENETİMİ — ' + input_val)
    if status >= 400:
        rep.warn('HTTP', 'HTTP ' + str(status))

    # Title
    title = first_title(html)
    if not title:
        rep.fail('TITLE', '<title> yok')
    else:
        rep.pass_('TITLE', f'"{title}" ({len(title)} kr)')
        if len(title) < 15:
            rep.warn('TITLE', 'Title çok kısa (<15)')
        if len(title) > 65:
            rep.warn('TITLE', 'Title uzun (>65) — SERP kırpılır')

    # Meta Description
    desc = first_meta_name('description', html)
    if not desc:
        rep.fail('DESC', 'meta description yok')
    else:
        rep.pass_('DESC', f'{len(desc)} kr')
        if len(desc) < 70:
            rep.warn('DESC', 'Description kısa (<70)')
        if len(desc) > 165:
            rep.warn('DESC', 'Description uzun (>165) — kırpılır')

    # Canonical
    canon = first_link_rel('canonical', html)
    if not canon:
        rep.fail('CANON', 'rel=canonical yok')
    else:
        rep.pass_('CANON', canon)
        if not re.match(r'^https?://', canon, re.IGNORECASE):
            rep.warn('CANON', 'Canonical mutlak URL değil')

    # Robots
    robots = first_meta_name('robots', html)
    if robots and re.search(r'noindex', robots, re.IGNORECASE):
        rep.warn('ROBOTS', 'Sayfa noindex: ' + robots)
    else:
        rep.info('ROBOTS', robots if robots else '(yok — indexlenir)')

    # Viewport
    viewport = first_meta_name('viewport', html)
    if not viewport:
        rep.fail('VIEWPORT', 'meta viewport yok (mobil uyumsuz)')

    # Open Graph
    og_required = ['og:title', 'og:type', 'og:url', 'og:image', 'og:description']
    for p in og_required:
        v = first_meta_prop(p, html)
        if not v:
            rep.fail('OG', f'{p} yok')
        else:
            rep.pass_('OG', f'{p} = {v[:80]}')

    og_img = first_meta_prop('og:image', html)
    if og_img and not re.match(r'^https?://', og_img, re.IGNORECASE):
        rep.warn('OG', 'og:image mutlak URL değil')

    if not first_meta_prop('og:site_name', html):
        rep.warn('OG', 'og:site_name yok')

    if not first_meta_prop('og:locale', html):
        rep.warn('OG', 'og:locale yok (ör. tr_TR)')

    # Twitter
    t_card = first_meta_name('twitter:card', html)
    if not t_card:
        rep.fail('TW', 'twitter:card yok')
    else:
        rep.pass_('TW', 'twitter:card = ' + t_card)

    for n in ['twitter:title', 'twitter:description', 'twitter:image']:
        if not first_meta_name(n, html):
            rep.warn('TW', n + ' yok')

    # og:type=article ise zaman damgaları
    og_type = first_meta_prop('og:type', html)
    if og_type == 'article':
        for p in ['article:published_time', 'article:modified_time']:
            if not first_meta_prop(p, html):
                rep.warn('OG-ART', p + ' yok')

    rep.info('URL', final_url)

    if args.get('json'):
        print(json.dumps({'summary': rep.summary(), 'text': rep.to_text()}, indent=2, ensure_ascii=False))
        sys.exit(1 if rep.summary()['fail'] > 0 else 0)

    sys.exit(rep.print_report())


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        sys.stderr.write(f'HATA: {e}\n')
        sys.exit(2)
