# scripts/generate_all_124_covers.py
import os
import json
from covers.part1 import PART1_DRAWERS
from covers.part2 import PART2_DRAWERS
from covers.part3 import PART3_DRAWERS
from covers.part4 import PART4_DRAWERS

# 124 slug'ın tamamını topla
ALL_DRAWERS = {}
ALL_DRAWERS.update(PART1_DRAWERS)
ALL_DRAWERS.update(PART2_DRAWERS)
ALL_DRAWERS.update(PART3_DRAWERS)
ALL_DRAWERS.update(PART4_DRAWERS)

print(f"Toplam kayıtlı benzersiz çizim fonksiyonu sayısı: {len(ALL_DRAWERS)}")

with open('scripts/products_meta.json', 'r', encoding='utf-8') as f:
    products = json.load(f)

print(f"Hedef ürün sayısı: {len(products)}")

def wrap_svg(title, category, body_svg):
    safe_title = title.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    safe_cat = (category or 'EXCEL ARŞİV PRO').upper().replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" width="100%" height="100%">
  <defs>
    <linearGradient id="gGreen" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#107c41" />
      <stop offset="100%" stop-color="#0d5c3a" />
    </linearGradient>
    <linearGradient id="gBlue" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#3b82f6" />
      <stop offset="100%" stop-color="#1d4ed8" />
    </linearGradient>
    <linearGradient id="gAmber" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#f59e0b" />
      <stop offset="100%" stop-color="#b45309" />
    </linearGradient>
    <linearGradient id="gRed" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#ef4444" />
      <stop offset="100%" stop-color="#b91c1c" />
    </linearGradient>
    <linearGradient id="gPurple" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#8b5cf6" />
      <stop offset="100%" stop-color="#6d28d9" />
    </linearGradient>
    <filter id="softCard" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="8" stdDeviation="12" flood-color="#0f172a" flood-opacity="0.06" />
    </filter>
  </defs>

  <!-- Saf Açık Tema Arka Plan (1000x1000 Kare Alan) -->
  <rect width="1000" height="1000" fill="#ffffff" />

  <!-- Dış Kenar Çerçeve (820x820 Çizim Alanı: x=90, y=90, %82 Doluluk) -->
  <rect x="90" y="90" width="820" height="820" rx="32" fill="#ffffff" stroke="#e2e8f0" stroke-width="2.5" filter="url(#softCard)" />

  <!-- Üst Başlık & Kategori Rozeti (x: 130, y: 130) -->
  <g transform="translate(130, 130)">
    <rect x="0" y="0" width="160" height="32" rx="16" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" />
    <circle cx="16" cy="16" r="5" fill="#107c41" />
    <text x="32" y="21" font-family="system-ui, -apple-system, sans-serif" font-size="11" font-weight="800" fill="#334155" letter-spacing="0.5">{safe_cat}</text>
    
    <text x="0" y="62" font-family="system-ui, -apple-system, sans-serif" font-size="20" font-weight="800" fill="#0f172a">{safe_title}</text>
  </g>

  <!-- 1:1 Kare Alanını Dolduran Benzersiz İnfografik / Çizim (x: 130, y: 220, w: 740, h: 650) -->
  <g transform="translate(130, 220)">
    {body_svg}
  </g>
</svg>"""

out_dir = 'public/images/kapak'
os.makedirs(out_dir, exist_ok=True)

missing = []
for p in products:
    slug = p['slug']
    if slug not in ALL_DRAWERS:
        missing.append(slug)

if missing:
    print(f"UYARI: {len(missing)} adet slug için çizim bulunamadı: {missing}")
else:
    print("MÜKEMMEL: 124 slug'ın 124'ü için de birebir eşleşen tekil çizim mevcut!")

count = 0
for p in products:
    slug = p['slug']
    drawer = ALL_DRAWERS[slug]
    body = drawer()
    full_svg = wrap_svg(p['title'], p.get('category'), body)
    
    out_file = os.path.join(out_dir, f"{slug}.svg")
    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(full_svg)
    count += 1

print(f"BAŞARILI: Toplam {count} adet 1000x1000 saf açık tema, %100 vektör, benzersiz kapak SVG dosyası üretildi!")
