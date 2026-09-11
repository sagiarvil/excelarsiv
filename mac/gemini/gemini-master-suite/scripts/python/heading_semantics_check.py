#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
heading_semantics_check.py — başlık hiyerarşisi + anlamsal/landmark yapı denetleyicisi.
JavaScript sürümü (heading_semantics_check.js + _lib.js) ile %100 uyumlu, harici bağımlılıksız (yalnız Python standart kütüphanesi).

Kullanım:
    python heading_semantics_check.py <URL|dosya> [--json]
"""

import sys
import os
import re
import json
import urllib.request
import urllib.error
import urllib.parse
from typing import Dict, List, Any, Optional, Tuple


class Report:
    """_lib.js makeReport eşdeğeri raporlama sınıfı."""

    def __init__(self, title: str):
        self.title = title
        self.rows: List[Tuple[str, str, str]] = []
        self.pass_count = 0
        self.fail_count = 0
        self.warn_count = 0

    def pass_item(self, id_: str, msg: str):
        self.rows.append(('PASS', id_, msg))
        self.pass_count += 1

    def fail_item(self, id_: str, msg: str):
        self.rows.append(('FAIL', id_, msg))
        self.fail_count += 1

    def warn_item(self, id_: str, msg: str):
        self.rows.append(('WARN', id_, msg))
        self.warn_count += 1

    def info_item(self, id_: str, msg: str):
        self.rows.append(('INFO', id_, msg))

    def summary(self) -> Dict[str, int]:
        return {
            'pass': self.pass_count,
            'fail': self.fail_count,
            'warn': self.warn_count,
            'total': len(self.rows),
        }

    def to_text(self) -> str:
        lines = [f'# {self.title}', '']
        for s, id_, msg in self.rows:
            lines.append(f'- [{s}] {id_} — {msg}')
        lines.append('')
        lines.append(f'SONUÇ: PASS={self.pass_count}  FAIL={self.fail_count}  WARN={self.warn_count}')
        return '\n'.join(lines)

    def print_report(self) -> int:
        print(self.to_text())
        return 0 if self.fail_count == 0 else 1


def parse_args(argv: Optional[List[str]] = None) -> Dict[str, Any]:
    """Basit argv ayrıştırıcı: --key=val, --flag, konumsal argümanlar."""
    raw_args = argv if argv is not None else sys.argv[1:]
    out: Dict[str, Any] = {'_': []}
    for a in raw_args:
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


def strip_tags(s: Optional[str]) -> str:
    """HTML etiketlerini temizler ve boşlukları sadeleştirir."""
    if not s:
        return ''
    cleaned = re.sub(r'<[^>]+>', ' ', s)
    return re.sub(r'\s+', ' ', cleaned).strip()


def get_attr(tag: str, name: str) -> Optional[str]:
    """Bir HTML etiketinden belirtilen özniteliğin (attribute) değerini döndürür."""
    match = re.search(r'\b' + re.escape(name) + r'=["\']([^"\']*)["\']', tag, re.IGNORECASE)
    return match.group(1) if match else None


def load_html(target: str) -> Dict[str, Any]:
    """Girdi URL veya yerel dosya yolundan HTML içeriğini yükler."""
    if re.match(r'^https?://', target, re.IGNORECASE):
        req = urllib.request.Request(
            target,
            headers={
                'User-Agent': 'super-seo-motor/1.0 (+audit)',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            }
        )
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                charset = resp.headers.get_content_charset() or 'utf-8'
                body = resp.read().decode(charset, errors='replace')
                return {
                    'html': body,
                    'status': resp.status,
                    'final_url': resp.url,
                    'headers': dict(resp.headers),
                }
        except urllib.error.HTTPError as e:
            body = e.read().decode('utf-8', errors='replace')
            return {
                'html': body,
                'status': e.code,
                'final_url': target,
                'headers': dict(e.headers),
            }
    else:
        abs_path = os.path.abspath(target)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f'Dosya bulunamadı: {abs_path}')
        with open(abs_path, 'r', encoding='utf-8', errors='replace') as f:
            return {
                'html': f.read(),
                'status': 200,
                'final_url': 'file://' + abs_path,
                'headers': {},
            }


def main():
    args = parse_args()
    if not args['_']:
        sys.stderr.write('Kullanım: python heading_semantics_check.py <URL|dosya> [--json]\n')
        sys.exit(2)

    input_target = args['_'][0]
    try:
        data = load_html(input_target)
        html = data['html']
    except Exception as e:
        sys.stderr.write(f'HATA: {e}\n')
        sys.exit(2)

    rep = Report('BAŞLIK & ANLAMSAL YAPI — ' + input_target)

    # Başlıkları yakala (H1 - H6)
    heading_pattern = re.compile(r'<(h[1-6])[^>]*>([\s\S]*?)</\1>', re.IGNORECASE)
    hs = []
    for match in heading_pattern.finditer(html):
        level = int(match.group(1)[1])
        text = strip_tags(match.group(2))
        hs.append({'level': level, 'text': text})

    # H1 kontrolü
    h1 = [h for h in hs if h['level'] == 1]
    if len(h1) == 0:
        rep.fail_item('H1', 'H1 yok')
    elif len(h1) > 1:
        h1_texts = ', '.join([f'"{h["text"]}"' for h in h1])
        rep.fail_item('H1', f'{len(h1)} adet H1 (tek olmalı): {h1_texts}')
    else:
        rep.pass_item('H1', f'Tek H1: "{h1[0]["text"]}"')

    # Başlık seviyesi atlamaları (örneğin H2 -> H4)
    prev = 0
    skips = 0
    for h in hs:
        if prev and h['level'] > prev + 1:
            rep.fail_item('SKIP', f'Seviye atlama: H{prev} → H{h["level"]} ("{h["text"]}")')
            skips += 1
        prev = h['level']

    if not skips and len(hs) > 0:
        rep.pass_item('SKIP', 'Başlık seviyesi atlaması yok')

    # Boş başlık etiketleri
    for h in hs:
        if not h['text']:
            rep.warn_item('EMPTY', f'Boş H{h["level"]} etiketi')

    # OUTLINE
    if hs:
        outline_str = '\n'.join(['  ' * (h['level'] - 1) + f'H{h["level"]} {h["text"]}' for h in hs])
    else:
        outline_str = '(başlık yok)'
    rep.info_item('OUTLINE', outline_str)

    # Landmark / Anlamsal etiketler
    need = {
        '<header': 'header',
        '<nav': 'nav',
        '<main': 'main',
        '<footer': 'footer',
    }
    for tag_prefix, name in need.items():
        pattern = re.compile(tag_prefix + r'\b', re.IGNORECASE)
        count = len(pattern.findall(html))
        if name == 'main' and count != 1:
            rep.fail_item('LANDMARK', f'<main> {count} adet (tam 1 olmalı)')
        elif count == 0:
            rep.warn_item('LANDMARK', f'<{name}> yok')
        else:
            rep.pass_item('LANDMARK', f'<{name}> ×{count}')

    # HTML Lang kontrolü
    lang_match = re.search(r'<html[^>]*\blang=["\']([^"\']+)["\']', html, re.IGNORECASE)
    if not re.search(r'<html[^>]*\blang=', html, re.IGNORECASE):
        rep.fail_item('LANG', '<html lang> yok')
    else:
        lang_val = lang_match.group(1).strip() if lang_match else ''
        rep.pass_item('LANG', lang_val)

    # Görsel alt metni (IMG-ALT)
    img_tags = re.findall(r'<img\b[^>]*>', html, re.IGNORECASE)
    no_alt = [t for t in img_tags if get_attr(t, 'alt') is None]
    if img_tags:
        status_func = rep.warn_item if len(no_alt) > 0 else rep.pass_item
        status_func('IMG-ALT', f'{len(img_tags)} <img>, {len(no_alt)} alt eksik')

    # JSON veya metin çıktısı
    if args.get('json'):
        output_data = {
            'summary': rep.summary(),
            'text': rep.to_text(),
        }
        print(json.dumps(output_data, indent=2, ensure_ascii=False))
        sys.exit(1 if rep.summary()['fail'] > 0 else 0)

    exit_code = rep.print_report()
    sys.exit(exit_code)


if __name__ == '__main__':
    main()
