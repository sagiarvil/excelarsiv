#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
web_vitals_hints.py — Core Web Vitals statik ipuçları (lab değil, kaynak taraması).
Kullanım: python web_vitals_hints.py <URL|dosya> [--json]
"""

import sys
import os
import re
import json
import urllib.request
import urllib.parse
import urllib.error


class Report:
    def __init__(self, title):
        self.title = title
        self.rows = []
        self.pass_count = 0
        self.fail_count = 0
        self.warn_count = 0

    def pass_(self, id_, msg):
        self.rows.append(('PASS', id_, msg))
        self.pass_count += 1

    def fail(self, id_, msg):
        self.rows.append(('FAIL', id_, msg))
        self.fail_count += 1

    def warn(self, id_, msg):
        self.rows.append(('WARN', id_, msg))
        self.warn_count += 1

    def info(self, id_, msg):
        self.rows.append(('INFO', id_, msg))

    def summary(self):
        return {
            'pass': self.pass_count,
            'fail': self.fail_count,
            'warn': self.warn_count,
            'total': len(self.rows)
        }

    def to_text(self):
        lines = [f"# {self.title}", ""]
        for status, id_, msg in self.rows:
            lines.append(f"- [{status}] {id_} — {msg}")
        lines.append("")
        lines.append(f"SONUÇ: PASS={self.pass_count}  FAIL={self.fail_count}  WARN={self.warn_count}")
        return "\n".join(lines)

    def print_report(self):
        print(self.to_text())
        return 0 if self.fail_count == 0 else 1


def parse_args(argv=None):
    if argv is None:
        argv = sys.argv[1:]
    out = {'_': []}
    for a in argv:
        if a.startswith('--'):
            part = a[2:]
            if '=' in part:
                k, v = part.split('=', 1)
                out[k] = v
            else:
                out[part] = True
        else:
            out['_'].append(a)
    return out


def fetch_url(target, timeout=20):
    req = urllib.request.Request(
        target,
        headers={
            'User-Agent': 'super-seo-motor/1.0 (+audit)',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        }
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        body = resp.read().decode('utf-8', errors='replace')
        status = resp.status
        final_url = resp.geturl()
        # Başlık anahtarlarını küçük harfe normalize et
        headers = {k.lower(): v for k, v in resp.headers.items()}
        return {
            'html': body,
            'status': status,
            'finalUrl': final_url,
            'headers': headers
        }


def load_html(input_target):
    if re.match(r'^https?://', input_target, re.IGNORECASE):
        r = fetch_url(input_target)
        return r['html'], r['headers']
    
    file_path = os.path.abspath(input_target)
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        html = f.read()
    return html, {}


def get_attr(tag, name):
    m = re.search(rf'\b{re.escape(name)}=["\']([^"\']*)["\']', tag, re.IGNORECASE)
    return m.group(1) if m else None


def main():
    args = parse_args()
    if not args['_']:
        print('Kullanım: python web_vitals_hints.py <URL|dosya> [--json]', file=sys.stderr)
        sys.exit(2)

    target = args['_'][0]
    html, headers = load_html(target)
    rep = Report(f"CORE WEB VITALS — STATİK İPUÇLARI — {target}")

    byte_len = len(html.encode('utf-8'))
    rep.info('HTML', f"{byte_len} bayt")
    if byte_len > 102400:
        rep.warn('HTML', 'HTML > 100KB — TTFB/LCP riski')
    rep.info('AST-14KB', 'İlk 14KB bütçesi içinde' if byte_len <= 14336 else 'İlk 14KB bütçesi aşıldı (kritik içerik geç)')

    # Render-blocking
    head_match = re.search(r'<head[^>]*>([\s\S]*?)</head>', html, re.IGNORECASE)
    head = head_match.group(1) if head_match else html

    all_css_links = [m.group(0) for m in re.finditer(r'<link[^>]*rel=["\']stylesheet["\'][^>]*>', head, re.IGNORECASE)]
    blk_css = [
        m for m in all_css_links
        if not re.search(r'\bmedia=["\'](print|[^"\']*\bprint)["\']', m, re.IGNORECASE)
        and not re.search(r'\bonload=', m, re.IGNORECASE)
    ]
    if len(blk_css) > 2:
        rep.warn('CSS', f"{len(blk_css)} render-blocking stylesheet (<=2 hedefle)")
    else:
        rep.pass_('CSS', f"{len(blk_css)} render-blocking stylesheet")

    head_scripts = [
        m.group(0) for m in re.finditer(r'<script\b[^>]*\bsrc=["\'][^"\']+["\'][^>]*>', head, re.IGNORECASE)
        if not re.search(r'\b(async|defer|type=["\']module["\'])\b', m.group(0), re.IGNORECASE)
    ]
    if head_scripts:
        rep.fail('JS', f"{len(head_scripts)} <head> içinde async/defer olmayan script")
    else:
        rep.pass_('JS', '<head> scriptleri async/defer/module')

    # 3. parti
    rx_script_src = re.compile(r'<script\b[^>]*\bsrc=["\']([^"\']+)["\'][^>]*>', re.IGNORECASE)
    ext = [m.group(1) for m in rx_script_src.finditer(html) if re.match(r'^https?://', m.group(1), re.IGNORECASE)]
    third_party = [s for s in ext if not re.search(r'(^/|localhost)', s)]
    if third_party:
        hosts = []
        for u in third_party:
            try:
                hosts.append(urllib.parse.urlparse(u).netloc or u)
            except Exception:
                hosts.append(u)
        unique_hosts = list(dict.fromkeys(hosts))
        rep.warn('3P', f"{len(third_party)} harici script: {', '.join(unique_hosts)}")

    # Görseller — boyut / lazy / modern format
    rx_img_tag = re.compile(r'<img\b[^>]*>', re.IGNORECASE)
    imgs = [m.group(0) for m in rx_img_tag.finditer(html)]
    no_dim = 0
    no_lazy = 0
    legacy = 0
    for t in imgs:
        if get_attr(t, 'width') is None or get_attr(t, 'height') is None:
            no_dim += 1
        if not re.search(r'\bloading=["\']lazy["\']', t, re.IGNORECASE) and not re.search(r'\bfetchpriority=["\']high["\']', t, re.IGNORECASE):
            no_lazy += 1
        src = get_attr(t, 'src') or ''
        if re.search(r'\.(jpe?g|png|gif)(\?|$)', src, re.IGNORECASE):
            legacy += 1

    if imgs:
        if no_dim:
            rep.warn('IMG-CLS', f"{no_dim}/{len(imgs)} görselde width/height yok (CLS riski)")
        else:
            rep.pass_('IMG-CLS', f"{no_dim}/{len(imgs)} görselde width/height yok (CLS riski)")

        if no_lazy > 1:
            rep.warn('IMG-LAZY', f"{no_lazy}/{len(imgs)} görselde loading=lazy yok")
        else:
            rep.pass_('IMG-LAZY', f"{no_lazy}/{len(imgs)} görselde loading=lazy yok")

        if legacy:
            rep.warn('IMG-FMT', f"{legacy} görsel jpg/png/gif — webp/avif düşün")

    # preconnect / dns-prefetch
    if third_party and not re.search(r'rel=["\'](preconnect|dns-prefetch)["\']', head, re.IGNORECASE):
        rep.warn('PRECONN', 'Harici köken var ama preconnect/dns-prefetch yok')

    # Font
    if re.search(r'<link[^>]*fonts?\.(googleapis|gstatic)', head, re.IGNORECASE) and not re.search(r'rel=["\']preconnect["\'][^>]*gstatic', head, re.IGNORECASE):
        rep.warn('FONT', 'Google Fonts var, gstatic preconnect yok')
    if re.search(r'@font-face', html, re.IGNORECASE) and not re.search(r'font-display\s*:', html, re.IGNORECASE):
        rep.warn('FONT', '@font-face var, font-display yok (FOIT)')

    # Sıkıştırma / cache başlıkları (URL ise)
    if headers and headers.get('content-encoding'):
        rep.pass_('GZIP', 'content-encoding: ' + headers['content-encoding'])
    elif headers and len(headers) > 0:
        rep.warn('GZIP', 'content-encoding başlığı yok (br/gzip?)')
    if headers and not headers.get('cache-control'):
        rep.warn('CACHE', 'cache-control başlığı yok')

    if args.get('json'):
        print(json.dumps({'summary': rep.summary(), 'text': rep.to_text()}, indent=2, ensure_ascii=False))
        sys.exit(1 if rep.summary()['fail'] else 0)

    sys.exit(rep.print_report())


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'HATA: {e}', file=sys.stderr)
        sys.exit(2)
