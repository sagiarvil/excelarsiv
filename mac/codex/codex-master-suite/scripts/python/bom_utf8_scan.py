#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bom_utf8_scan.py — BOM'suz UTF-8 taraması (+ isteğe bağlı onarım, + cache temizliği).
Kullanım:
  python bom_utf8_scan.py [kök-dizin] [--fix] [--ext=php,json,html,js,css,xml,txt,md] [--cache=storage/cache] [--json]
  kök-dizin verilmezse otomatik proje kökü bulunur.
"""

import sys
import os
import json
from pathlib import Path

# Ortak atlanacak dizinler
SKIP = {'vendor', 'node_modules', '.git', 'dist', 'build', '.venv', 'storage'}

ROOT_MARKERS = ['composer.json', 'package.json', '.git', 'artisan', os.path.join('public', 'index.php')]

def find_project_root(start_dir=None):
    """Verilen dizinden yukarı çıkarak proje kökünü bul. Bulunamazsa start_dir döner."""
    current = os.path.abspath(start_dir if start_dir else os.getcwd())
    while True:
        for marker in ROOT_MARKERS:
            if os.path.exists(os.path.join(current, marker)):
                return current
        parent = os.path.dirname(current)
        if parent == current:
            return os.path.abspath(start_dir if start_dir else os.getcwd())
        current = parent

def parse_args(argv=None):
    """Basit argv ayrıştırıcı: --key=val, --flag, konumsal."""
    args_list = argv if argv is not None else sys.argv[1:]
    out = {'_': []}
    for a in args_list:
        if a.startswith('--'):
            parts = a[2:].split('=', 1)
            k = parts[0]
            val = parts[1] if len(parts) > 1 else True
            out[k] = val
        else:
            out['_'].append(a)
    return out

class Report:
    """Raporlama yardımcı sınıfı (makeReport muadili)."""
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
        lines = [f"# {self.title}", '']
        for status, id_, msg in self.rows:
            lines.append(f"- [{status}] {id_} — {msg}")
        lines.append('')
        lines.append(f"SONUÇ: PASS={self.pass_count}  FAIL={self.fail_count}  WARN={self.warn_count}")
        return '\n'.join(lines)

    def print_report(self):
        print(self.to_text())
        return 0 if self.fail_count == 0 else 1

def walk(directory, cb):
    """Dizinleri yinelemeli olarak gezer; SKIP dizinlerini atlar."""
    try:
        entries = os.listdir(directory)
    except OSError:
        return

    for name in entries:
        p = os.path.join(directory, name)
        try:
            is_dir = os.path.isdir(p)
        except OSError:
            continue

        if is_dir:
            if name not in SKIP:
                walk(p, cb)
        else:
            cb(p)

def main():
    args = parse_args()
    root = os.path.abspath(args['_'][0]) if args['_'] else find_project_root()
    ext_arg = args.get('ext') or 'php,json,html,js,css,xml,txt,md'
    exts = set('.' + e.strip().lstrip('.').lower() for e in ext_arg.split(','))
    
    rep = Report('BOM / UTF-8 TARAMASI — ' + root)
    hits = []

    def inspect_file(fp):
        _, ext = os.path.splitext(fp)
        if ext.lower() not in exts:
            return

        try:
            with open(fp, 'rb') as f:
                buf = f.read()
        except OSError:
            return

        rel = os.path.relpath(fp, root)

        # UTF-8 BOM kontrolü (0xEF, 0xBB, 0xBF)
        if len(buf) >= 3 and buf[0] == 0xEF and buf[1] == 0xBB and buf[2] == 0xBF:
            hits.append(rel)
            if args.get('fix'):
                try:
                    with open(fp, 'wb') as f:
                        f.write(buf[3:])
                    rep.pass_('BOM-FIX', f"{rel} → BOM silindi")
                except OSError as e:
                    rep.fail('BOM-FIX', f"{rel} yazılırken hata oluştu: {e}")
            else:
                rep.fail('BOM', f"{rel} BOM içeriyor")

        # Tek satıra minify şüphesi (kaynak dosyada)
        if ext.lower() in {'.php', '.js', '.css'} and len(buf) > 3000 and (0x0A not in buf):
            rep.warn('MINIFY', f"{rel} newline içermiyor ({len(buf)} bayt) — bozuk/minify olabilir")

    walk(root, inspect_file)

    if not hits:
        rep.pass_('BOM', 'BOM içeren dosya yok')
    elif not args.get('fix'):
        rep.info('BOM', f"{len(hits)} dosyada BOM var — --fix ile onarılır")

    # Cache temizliği
    if 'cache' in args:
        cache_arg = args.get('cache')
        cache_rel = cache_arg if isinstance(cache_arg, str) and cache_arg.strip() else 'storage/cache'
        cd = os.path.join(root, cache_rel)
        if os.path.exists(cd):
            n = 0
            try:
                for f in os.listdir(cd):
                    fp = os.path.join(cd, f)
                    try:
                        if os.path.isfile(fp):
                            os.unlink(fp)
                            n += 1
                    except OSError:
                        pass
                rep.info('CACHE', f"{cd} → {n} dosya silindi")
            except OSError as e:
                rep.warn('CACHE', f"Cache dizini okunurken hata oluştu: {cd} ({e})")
        else:
            rep.warn('CACHE', f"cache dizini yok: {cd}")

    if args.get('json'):
        print(json.dumps({
            'summary': rep.summary(),
            'text': rep.to_text(),
            'files': hits
        }, indent=2, ensure_ascii=False))
        sys.exit(1 if rep.summary()['fail'] > 0 else 0)

    sys.exit(rep.print_report())

if __name__ == '__main__':
    main()
