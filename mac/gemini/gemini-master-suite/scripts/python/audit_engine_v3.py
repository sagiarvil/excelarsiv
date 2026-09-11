#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BETASOFT SATIS (MACOS SUITE) - SUPER-MANDATE ENGINE V3.0 DENETİMİ (PYTHON)
audit_engine_v3.js dosyasının birebir Python eşdeğeridir.
_lib.py modülü varsa yükler ve onun parse_args / load_html yeteneklerini kullanır.
"""

import os
import sys
import re
import json
from pathlib import Path

# _lib.py import ayarı: Önce bulunulan dizini, sonra üst dizinleri arar
_lib = None
try:
    import _lib
except ImportError:
    current_dir = Path(__file__).resolve().parent
    for search_path in [current_dir, current_dir.parent]:
        if str(search_path) not in sys.path:
            sys.path.insert(0, str(search_path))
    try:
        import _lib
    except ImportError:
        _lib = None


def get_html_content(default_path: str = r"$HOME/Sites/satis/docs/satis_test_live.html") -> str:
    """Hedef HTML içeriğini yerel dosyadan veya _lib üzerinden yükler."""
    target = None

    # 1. _lib.parseArgs / parse_args kontrolü
    if _lib:
        parse_fn = getattr(_lib, 'parse_args', None) or getattr(_lib, 'parseArgs', None)
        if callable(parse_fn):
            args = parse_fn()
            if isinstance(args, dict) and args.get('_'):
                target = args['_'][0]
            elif hasattr(args, '_') and getattr(args, '_'):
                target = getattr(args, '_')[0]
    
    # 2. Standart sys.argv fallback
    if not target and len(sys.argv) > 1 and not sys.argv[1].startswith('--'):
        target = sys.argv[1]

    target = target or default_path

    # 3. _lib.loadHtml / load_html desteği
    if _lib:
        load_fn = getattr(_lib, 'load_html', None) or getattr(_lib, 'loadHtml', None)
        if callable(load_fn):
            res = load_fn(target)
            if isinstance(res, dict) and 'html' in res:
                return res['html']
            elif isinstance(res, str):
                return res

    # 4. Doğrudan dosya okuma fallback
    if not os.path.exists(target):
        # Eğer verilen dosya bulunamazsa varsayılan yolu dene
        if os.path.exists(default_path):
            target = default_path
        else:
            raise FileNotFoundError(f"HTML dosyası bulunamadı: {target}")

    with open(target, 'r', encoding='utf-8', errors='replace') as f:
        return f.read()


def run_audit():
    html = get_html_content()

    print('================================================================')
    print('BETASOFT SATIS (MACOS SUITE) - SUPER-MANDATE ENGINE V3.0 DENETİMİ')
    print('================================================================\n')

    # 1. BYTE SIZE (ENG-01: 14KB TCP Window)
    byte_length = len(html.encode('utf-8'))
    budget_status = 'PASS (<= 14.336)' if byte_length <= 14336 else 'FAIL (90KB, 14KB bütçesini aşıyor)'
    print(f'[ENG-01 & ENG-02] HTML Toplam Boyutu: {byte_length} bytes')
    print(f'                  14KB Bütçe Durumu: {budget_status}')

    # 2. CORE SEO (ENG-04)
    title_match = re.search(r'<title>([^<]*)</title>', html, re.IGNORECASE)
    print(f"\n[ENG-04] Title: {title_match.group(1).strip() if title_match else 'YOK'}")

    desc_match = re.search(r'<meta[^>]*name=["\']description["\'][^>]*content=["\']([^"\']*)["\']', html, re.IGNORECASE)
    print(f"[ENG-04] Meta Description: {desc_match.group(1).strip() if desc_match else 'YOK'}")

    canon_match = re.search(r'<link[^>]*rel=["\']canonical["\'][^>]*href=["\']([^"\']*)["\']', html, re.IGNORECASE)
    print(f"[ENG-04] Canonical URL: {canon_match.group(1).strip() if canon_match else 'YOK'}")

    h1_matches = list(re.finditer(r'<h1[^>]*>([\s\S]*?)</h1>', html, re.IGNORECASE))
    print(f"[ENG-04] H1 Sayısı: {len(h1_matches)}")
    for i, m in enumerate(h1_matches, 1):
        cleaned_h1 = re.sub(r'<[^>]+>', '', m.group(1)).strip()
        print(f"         H1 #{i}: {cleaned_h1}")

    # 3. ColBERT MaxSim Headings (ENG-10)
    h2_count = len(re.findall(r'<h2[^>]*>', html, re.IGNORECASE))
    h3_count = len(re.findall(r'<h3[^>]*>', html, re.IGNORECASE))
    print(f"\n[ENG-10] ColBERT MaxSim Başlık Sayıları: H2 = {h2_count} , H3 = {h3_count}")

    # 4. LLMO & Machine Surfaces (ENG-07)
    described_by = re.findall(r'<link[^>]*rel=["\']describedby["\'][^>]*>', html, re.IGNORECASE)
    alternate_md = re.findall(r'<link[^>]*type=["\']text/markdown["\'][^>]*>', html, re.IGNORECASE)
    print(f"\n[ENG-07] LLMO Describedby Linki: {'MEVCUT' if len(described_by) > 0 else 'EKSİK'}")
    print(f"[ENG-07] LLMO Alternate Markdown Linki: {'MEVCUT' if len(alternate_md) > 0 else 'EKSİK'}")

    # 5. C2PA & Provenance (ENG-03)
    c2pa_meta = re.search(r'<meta[^>]*name=["\']c2pa-manifest["\'][^>]*>', html, re.IGNORECASE)
    dcterms_issued = re.search(r'<meta[^>]*name=["\']dcterms\.issued["\'][^>]*>', html, re.IGNORECASE)
    print(f"\n[ENG-03] C2PA Manifest Meta: {'MEVCUT' if c2pa_meta else 'EKSİK'}")
    print(f"[ENG-03] DCTERMS Issued Meta: {'MEVCUT' if dcterms_issued else 'EKSİK'}")

    # 6. RAG Chunks (ENG-09, ENG-10)
    chunk_ids = re.findall(r'data-chunk-id', html, re.IGNORECASE)
    print(f"\n[ENG-09] RAG data-chunk-id Sayısı: {len(chunk_ids)}")

    # 7. AEO Hero Answer in top 100px (ENG-06)
    hero_answer = re.search(r'class=["\'][^"\']*hero-answer', html, re.IGNORECASE)
    print(f"\n[ENG-06] AEO Hero Answer Sınıfı: {'MEVCUT' if hero_answer else 'EKSİK'}")

    # 8. JSON-LD Blocks (ENG-08 & ENG-15)
    json_ld_matches = list(re.finditer(r'<script[^>]*type=["\']application/ld\+json["\'][^>]*>([\s\S]*?)</script>', html, re.IGNORECASE))
    print(f"\n[ENG-08 & ENG-15] Toplam JSON-LD Blok Sayısı: {len(json_ld_matches)}")
    has_wikidata = False
    has_mid = False

    for i, m in enumerate(json_ld_matches, 1):
        content = m.group(1)
        if 'wikidata.org/wiki/Q' in content:
            has_wikidata = True
        if 'kgmid' in content or '/m/' in content:
            has_mid = True
        try:
            parsed = json.loads(content)
            node_type = parsed.get('@type')
            if not node_type and '@graph' in parsed and isinstance(parsed['@graph'], list):
                node_type = f"@graph ({len(parsed['@graph'])} nodes)"
            elif not node_type:
                node_type = 'Unknown'
            print(f"   Blok #{i} Type: {node_type}")
        except Exception as e:
            print(f"   Blok #{i} JSON Parse Hatası: {e}")

    print(f"[ENG-15] Knowledge Vault Wikidata QID: {'BAĞLI' if has_wikidata else 'EKSİK (sameAs içinde Wikidata yok)'}")
    print(f"[ENG-15] Knowledge Vault Google MID: {'BAĞLI' if has_mid else 'EKSİK'}")


if __name__ == '__main__':
    run_audit()
