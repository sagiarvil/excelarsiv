#!/usr/bin/env python3
"""
seo_enforcer.py --build-ai-pack dist sonrasında dist/llms.txt ve dist/llms-full.txt
dosyalarına validate-artifacts.mjs için zorunlu E-E-A-T, VKN, Sitemap ve Sektör dikeyleri
bölümlerini deterministik olarak ekler.
"""
from pathlib import Path

DIST_DIR = Path("dist")
BLOCK = """
## E-E-A-T ve Doğrulanmış Kurumsal Kimlik
- Kurumsal Ünvan: HtlandHtml / HtmlandHtml Yüksek Teknoloji Yazılım Ltd. Şti. (Barış Bağırlar — Excel Arşiv)
- Satıcı VKN / TCKN: 25403091318
- Sitemap: https://excelarsiv.com/sitemap.xml

## Sektör dikeyleri
- Finans, Nakit Akışı, Bütçe ve KPI Yönetimi: https://excelarsiv.com/sablonlar/
- İnşaat, Hakediş, E-Ticaret ve Kafe-Restoran Çözümleri: https://excelarsiv.com/sektor/
"""

for fname in ("llms.txt", "llms-full.txt"):
    p = DIST_DIR / fname
    if p.exists():
        txt = p.read_text(encoding="utf-8")
        if "25403091318" not in txt or "## Sektör dikeyleri" not in txt:
            p.write_text(txt.rstrip() + "\n" + BLOCK.lstrip(), encoding="utf-8")
print("[AI-PACK-EEAT] dist/llms.txt ve dist/llms-full.txt E-E-A-T/VKN/Sektör blokları güncellendi.")
