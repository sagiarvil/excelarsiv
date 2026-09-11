#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
super-seo-motor ortak kütüphane — bağımlılıksız (yalnız Python standart çekirdeği).
Tüm araçlar buradan beslenir. SABİT YOL YOK: kök otomatik bulunur / argümandan alınır.
"""

import os
import sys
import re
import json
import urllib.request
import urllib.parse
import urllib.error

ROOT_MARKERS = [
    'composer.json',
    'package.json',
    '.git',
    'artisan',
    os.path.join('public', 'index.php')
]

def find_project_root(start_dir=None):
    """Verilen dizinden yukarı çıkarak proje kökünü bul. Bulunamazsa start_dir döner."""
    current = os.path.abspath(start_dir or os.getcwd())
    dir_path = current
    while True:
        for m in ROOT_MARKERS:
            if os.path.exists(os.path.join(dir_path, m)):
                return dir_path
        parent = os.path.dirname(dir_path)
        if parent == dir_path:
            return current
        dir_path = parent

findProjectRoot = find_project_root

def seotest_dir(root=None):
    """[PROJE_KÖKÜ]/docs/seotest yolu."""
    return os.path.join(root or find_project_root(), 'docs', 'seotest')

seotestDir = seotest_dir

def ensure_dir(d):
    """Dizinin var olduğundan emin ol (recursive)."""
    os.makedirs(d, exist_ok=True)

ensureDir = ensure_dir

def parse_args(argv=None):
    """Basit argv ayrıştırıcı: --key=val, --flag, konumsal."""
    out = {'_': []}
    args = sys.argv[1:] if argv is None else argv
    for a in args:
        if a.startswith('--'):
            parts = a[2:].split('=', 1)
            k = parts[0]
            out[k] = parts[1] if len(parts) > 1 else True
        else:
            out['_'].append(a)
    return out

parseArgs = parse_args

def fetch_url(target, opts=None, redirects=0):
    """URL getir (yönlendirme takipli). {status, headers, body, url} döner."""
    if opts is None:
        opts = {}

    try:
        parsed = urllib.parse.urlparse(target)
        if not parsed.scheme or not parsed.netloc:
            raise ValueError(f"Geçersiz URL: {target}")
    except Exception as e:
        raise ValueError(f"Geçersiz URL: {target}") from e

    headers = {
        'User-Agent': 'super-seo-motor/1.0 (+audit)',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    }
    if opts.get('headers'):
        headers.update(opts['headers'])

    method = opts.get('method', 'GET').upper()
    body_data = opts.get('body')
    if isinstance(body_data, str):
        body_data = body_data.encode('utf-8')

    timeout_sec = (opts.get('timeout', 20000)) / 1000.0

    class NoRedirect(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, req, fp, code, msg, hdrs, newurl):
            return None

    opener = urllib.request.build_opener(NoRedirect)
    req = urllib.request.Request(target, data=body_data, headers=headers, method=method)

    try:
        with opener.open(req, timeout=timeout_sec) as res:
            status = res.status
            res_headers = {k.lower(): v for k, v in res.headers.items()}
            body_bytes = res.read()
            body_str = body_bytes.decode('utf-8', errors='replace')
            loc = res.headers.get('Location')
            if 300 <= status < 400 and loc and redirects < 6:
                next_url = urllib.parse.urljoin(target, loc)
                return fetch_url(next_url, opts, redirects + 1)
            return {
                'status': status,
                'headers': res_headers,
                'body': body_str,
                'url': target
            }
    except urllib.error.HTTPError as e:
        status = e.code
        res_headers = {k.lower(): v for k, v in e.headers.items()}
        loc = e.headers.get('Location')
        if 300 <= status < 400 and loc and redirects < 6:
            next_url = urllib.parse.urljoin(target, loc)
            return fetch_url(next_url, opts, redirects + 1)
        body_bytes = e.read()
        body_str = body_bytes.decode('utf-8', errors='replace')
        return {
            'status': status,
            'headers': res_headers,
            'body': body_str,
            'url': target
        }

fetchUrl = fetch_url

def load_html(input_target):
    """Girdi URL de olabilir yerel dosya da. HTML metni döndürür."""
    if re.match(r'^https?://', input_target, re.IGNORECASE):
        r = fetch_url(input_target)
        return {
            'html': r['body'],
            'status': r['status'],
            'finalUrl': r['url'],
            'final_url': r['url'],
            'headers': r['headers']
        }
    p = os.path.abspath(input_target)
    with open(p, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    return {
        'html': content,
        'status': 200,
        'finalUrl': 'file://' + p,
        'final_url': 'file://' + p,
        'headers': {}
    }

loadHtml = load_html

# ---------- HTML yardımcıları (regex tabanlı, DOM'suz) ----------
class Rx:
    title = re.compile(r'<title[^>]*>([\s\S]*?)</title>', re.IGNORECASE)
    json_ld = re.compile(r'<script[^>]*type=["\']application/ld\+json["\'][^>]*>([\s\S]*?)</script>', re.IGNORECASE)
    jsonLd = json_ld
    headings = re.compile(r'<(h[1-6])[^>]*>([\s\S]*?)</\1>', re.IGNORECASE)
    img_tag = re.compile(r'<img\b[^>]*>', re.IGNORECASE)
    imgTag = img_tag
    script_src = re.compile(r'<script\b[^>]*\bsrc=["\']([^"\']+)["\'][^>]*>', re.IGNORECASE)
    scriptSrc = script_src
    a_href = re.compile(r'<a\b[^>]*\bhref=["\']([^"\']+)["\'][^>]*>([\s\S]*?)</a>', re.IGNORECASE)
    aHref = a_href

    @staticmethod
    def meta_name(n):
        return re.compile(r'<meta[^>]*name=["\']' + re.escape(n) + r'["\'][^>]*content=["\']([^"\']*)["\']', re.IGNORECASE)
    metaName = meta_name

    @staticmethod
    def meta_prop(p):
        return re.compile(r'<meta[^>]*property=["\']' + re.escape(p) + r'["\'][^>]*content=["\']([^"\']*)["\']', re.IGNORECASE)
    metaProp = meta_prop

    @staticmethod
    def link_rel(r):
        return re.compile(r'<link[^>]*rel=["\']' + re.escape(r) + r'["\'][^>]*href=["\']([^"\']*)["\']', re.IGNORECASE)
    linkRel = link_rel

rx = Rx()

def first(regex, s):
    """Regex ile ilk yakalanan grubu getirir."""
    if not s:
        return None
    if isinstance(regex, str):
        regex = re.compile(regex, re.IGNORECASE)
    m = regex.search(s)
    return m.group(1).strip() if m else None

def strip_tags(s):
    """HTML etiketlerini temizler."""
    text = str(s or '')
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

stripTags = strip_tags

def attr(tag, name):
    """Etiket içinden belirtilen öznitelik (attribute) değerini çeker."""
    pattern = re.compile(r'\b' + re.escape(name) + r'=["\']([^"\']*)["\']', re.IGNORECASE)
    m = pattern.search(tag)
    return m.group(1) if m else None

def extract_json_ld(html):
    """Tüm JSON-LD bloklarını çıkar ve parse et. [{'ok': bool, 'raw': str, 'data': dict/list, 'error': str}]"""
    out = []
    if not html:
        return out
    for m in Rx.json_ld.finditer(html):
        raw = m.group(1).strip()
        try:
            data = json.loads(raw)
            out.append({'ok': True, 'raw': raw, 'data': data})
        except Exception as e:
            out.append({'ok': False, 'raw': raw, 'error': str(e)})
    return out

extractJsonLd = extract_json_ld

def flatten_nodes(parsed_blocks):
    """@graph'ı düzleştirip tüm node'ları tek listeye indir."""
    nodes = []
    for b in parsed_blocks:
        if not b.get('ok'):
            continue
        d = b.get('data')
        if isinstance(d, list):
            nodes.extend(d)
        elif isinstance(d, dict) and isinstance(d.get('@graph'), list):
            nodes.extend(d['@graph'])
        elif d is not None:
            nodes.append(d)
    return nodes

flattenNodes = flatten_nodes

# ---------- Raporlama ----------
class Report:
    def __init__(self, title):
        self.title = title
        self.rows = []
        self._pass = 0
        self._fail = 0
        self._warn = 0

    def pass_(self, id_, msg):
        self.rows.append(('PASS', id_, msg))
        self._pass += 1

    def fail(self, id_, msg):
        self.rows.append(('FAIL', id_, msg))
        self._fail += 1

    def warn(self, id_, msg):
        self.rows.append(('WARN', id_, msg))
        self._warn += 1

    def info(self, id_, msg):
        self.rows.append(('INFO', id_, msg))

    def summary(self):
        return {
            'pass': self._pass,
            'fail': self._fail,
            'warn': self._warn,
            'total': len(self.rows)
        }

    def to_text(self):
        lines = [f"# {self.title}", '']
        for s, id_, msg in self.rows:
            lines.append(f"- [{s}] {id_} — {msg}")
        lines.append('')
        lines.append(f"SONUÇ: PASS={self._pass}  FAIL={self._fail}  WARN={self._warn}")
        return '\n'.join(lines)

    toText = to_text

    def print(self):
        print(self.to_text())
        return 0 if self._fail == 0 else 1

def make_report(title):
    rep = Report(title)
    setattr(rep, 'pass', rep.pass_)
    return rep

makeReport = make_report

__all__ = [
    'ROOT_MARKERS', 'find_project_root', 'findProjectRoot',
    'seotest_dir', 'seotestDir', 'ensure_dir', 'ensureDir',
    'parse_args', 'parseArgs', 'fetch_url', 'fetchUrl',
    'load_html', 'loadHtml', 'rx', 'first', 'strip_tags', 'stripTags',
    'attr', 'extract_json_ld', 'extractJsonLd', 'flatten_nodes', 'flattenNodes',
    'make_report', 'makeReport', 'Report'
]
