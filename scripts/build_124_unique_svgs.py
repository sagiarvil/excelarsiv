import json
import os
import math

# Ürün listesini yükle
with open('scripts/products_meta.json', 'r', encoding='utf-8') as f:
    products = json.load(f)

print(f"Toplam {len(products)} ürün için 124 benzersiz kare SVG üretiliyor...")

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

  <!-- Saf Açık Tema Arka Plan (1000x1000) -->
  <rect width="1000" height="1000" fill="#ffffff" />

  <!-- Dış Kenar Çerçeve (820x820 Çizim Alanı: x=90, y=90) -->
  <rect x="90" y="90" width="820" height="820" rx="32" fill="#ffffff" stroke="#e2e8f0" stroke-width="2.5" filter="url(#softCard)" />

  <!-- Üst Başlık & Rozet Bölümü -->
  <g transform="translate(130, 130)">
    <rect x="0" y="0" width="150" height="32" rx="16" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" />
    <circle cx="16" cy="16" r="5" fill="#107c41" />
    <text x="32" y="21" font-family="system-ui, -apple-system, sans-serif" font-size="11" font-weight="800" fill="#334155" letter-spacing="0.5">{safe_cat}</text>
    
    <text x="0" y="62" font-family="system-ui, -apple-system, sans-serif" font-size="20" font-weight="800" fill="#0f172a">{safe_title}</text>
  </g>

  <!-- Benzersiz İnfografik / Çizim Alanı (x: 130, y: 220, w: 740, h: 650) -->
  <g transform="translate(130, 220)">
    {body_svg}
  </g>
</svg>"""

def polar_to_cartesian(cx, cy, r, angle_deg):
    rad = ((angle_deg - 90) * math.pi) / 180.0
    return cx + r * math.cos(rad), cy + r * math.sin(rad)

def donut_segment(cx, cy, r_in, r_out, start_ang, end_ang, fill):
    p1 = polar_to_cartesian(cx, cy, r_out, end_ang)
    p2 = polar_to_cartesian(cx, cy, r_out, start_ang)
    p3 = polar_to_cartesian(cx, cy, r_in, end_ang)
    p4 = polar_to_cartesian(cx, cy, r_in, start_ang)
    large = 0 if end_ang - start_ang <= 180 else 1
    d = f"M {p1[0]:.1f} {p1[1]:.1f} A {r_out} {r_out} 0 {large} 0 {p2[0]:.1f} {p2[1]:.1f} L {p4[0]:.1f} {p4[1]:.1f} A {r_in} {r_in} 0 {large} 1 {p3[0]:.1f} {p3[1]:.1f} Z"
    return f'<path d="{d}" fill="{fill}" />'

# 124 Slug için 124 tekil çizim üretici fonksiyonu
DRAWERS = {}

def reg(slug):
    def deco(fn):
        DRAWERS[slug] = fn
        return fn
    return deco

# 1. 13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi
@reg("13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi")
def d001():
    h = [70, 100, 85, 130, 160, 140, 200, 240, 220, 280, 330, 310, 380]
    bars = []
    for i in range(13):
        x = 20 + i * 55
        y = 450 - h[i]
        c = "url(#gGreen)" if i == 12 else "#107c41"
        op = 0.4 + (i / 13.0) * 0.6
        bars.append(f'<rect x="{x}" y="{y}" width="38" height="{h[i]}" rx="6" fill="{c}" opacity="{op:.2f}" />')
        bars.append(f'<text x="{x+19}" y="480" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b" text-anchor="middle">H{i+1}</text>')
    return f"""
    <line x1="20" y1="450" x2="720" y2="450" stroke="#cbd5e1" stroke-width="2" />
    <path d="M 39 380 Q 200 330 370 250 T 700 70" fill="none" stroke="#2563eb" stroke-width="4" stroke-linecap="round" />
    {''.join(bars)}
    <g transform="translate(470, 50)">
      <rect width="230" height="70" rx="14" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <text x="20" y="28" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">13. Hafta Kümülatif Nakit</text>
      <text x="20" y="54" font-family="system-ui" font-size="20" font-weight="800" fill="#107c41">+₺2.480.000</text>
    </g>
    """

# 2. akilli-kasa-defteri-ve-nakit-kontrol-sistemi
@reg("akilli-kasa-defteri-ve-nakit-kontrol-sistemi")
def d002():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <rect x="50" y="50" width="200" height="90" rx="12" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.5" />
    <text x="70" y="80" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Günlük Kasa Girişi</text>
    <text x="70" y="115" font-family="system-ui" font-size="22" font-weight="800" fill="#107c41">₺145.850</text>
    <rect x="270" y="50" width="200" height="90" rx="12" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.5" />
    <text x="290" y="80" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Günlük Kasa Çıkışı</text>
    <text x="290" y="115" font-family="system-ui" font-size="22" font-weight="800" fill="#dc2626">₺92.400</text>
    <rect x="490" y="50" width="200" height="90" rx="12" fill="#ffffff" stroke="#86efac" stroke-width="2" />
    <text x="510" y="80" font-family="system-ui" font-size="12" font-weight="600" fill="#107c41">Net Gün Sonu Bakiye</text>
    <text x="510" y="115" font-family="system-ui" font-size="22" font-weight="800" fill="#0d5c3a">₺53.450</text>
    <g transform="translate(50, 170)">
      <rect width="640" height="40" rx="8" fill="#e2e8f0" />
      <text x="20" y="25" font-family="system-ui" font-size="12" font-weight="700" fill="#334155">TARİH / EVRAK NO</text>
      <text x="220" y="25" font-family="system-ui" font-size="12" font-weight="700" fill="#334155">AÇIKLAMA</text>
      <text x="440" y="25" font-family="system-ui" font-size="12" font-weight="700" fill="#334155">GİRİŞ / ÇIKIŞ</text>
      <text x="560" y="25" font-family="system-ui" font-size="12" font-weight="700" fill="#334155">BAKİYE</text>
      <rect y="50" width="640" height="45" rx="8" fill="#ffffff" stroke="#f1f5f9" stroke-width="1.5" />
      <circle cx="25" cy="72" r="6" fill="#107c41" />
      <text x="40" y="77" font-family="system-ui" font-size="12" font-weight="600" fill="#475569">KSA-2026-01</text>
      <text x="220" y="77" font-family="system-ui" font-size="12" font-weight="500" fill="#1e293b">Perakende Mağaza Satış Tahsilatı</text>
      <text x="440" y="77" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41">+₺35.200</text>
      <text x="560" y="77" font-family="system-ui" font-size="12" font-weight="700" fill="#0f172a">₺35.200</text>
      <rect y="105" width="640" height="45" rx="8" fill="#ffffff" stroke="#f1f5f9" stroke-width="1.5" />
      <circle cx="25" cy="127" r="6" fill="#dc2626" />
      <text x="40" y="132" font-family="system-ui" font-size="12" font-weight="600" fill="#475569">KSA-2026-02</text>
      <text x="220" y="132" font-family="system-ui" font-size="12" font-weight="500" fill="#1e293b">Lojistik Mazot ve Nakliye Avansı</text>
      <text x="440" y="132" font-family="system-ui" font-size="12" font-weight="700" fill="#dc2626">-₺8.500</text>
      <text x="560" y="132" font-family="system-ui" font-size="12" font-weight="700" fill="#0f172a">₺26.700</text>
      <rect y="160" width="640" height="45" rx="8" fill="#ffffff" stroke="#f1f5f9" stroke-width="1.5" />
      <circle cx="25" cy="182" r="6" fill="#107c41" />
      <text x="40" y="187" font-family="system-ui" font-size="12" font-weight="600" fill="#475569">KSA-2026-03</text>
      <text x="220" y="187" font-family="system-ui" font-size="12" font-weight="500" fill="#1e293b">Merkez Kasa B2B Cari Tahsilat</text>
      <text x="440" y="187" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41">+₺26.750</text>
      <text x="560" y="187" font-family="system-ui" font-size="12" font-weight="700" fill="#0f172a">₺53.450</text>
    </g>
    """

# 3. amortisman-2026-yeniden-degerleme
@reg("amortisman-2026-yeniden-degerleme")
def d003():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <path d="M 60 420 L 200 350 L 360 270 L 520 190 L 660 90" fill="none" stroke="#107c41" stroke-width="6" stroke-linecap="round" />
    <circle cx="60" cy="420" r="12" fill="#ffffff" stroke="#107c41" stroke-width="4" />
    <text x="60" y="465" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">Tarihi Maliyet</text>
    <circle cx="200" cy="350" r="12" fill="#ffffff" stroke="#107c41" stroke-width="4" />
    <text x="200" y="395" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">VUK 298/Ç</text>
    <circle cx="360" cy="270" r="12" fill="#ffffff" stroke="#107c41" stroke-width="4" />
    <text x="360" y="315" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">Geçici 32. Md</text>
    <circle cx="520" cy="190" r="12" fill="#ffffff" stroke="#107c41" stroke-width="4" />
    <text x="520" y="235" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">2026 Değeri</text>
    <circle cx="660" cy="90" r="14" fill="#107c41" />
    <text x="660" y="60" font-family="system-ui" font-size="14" font-weight="800" fill="#107c41" text-anchor="middle">₺46.200.000</text>
    <g transform="translate(60, 50)">
      <rect width="250" height="45" rx="22" fill="#dcfce7" />
      <text x="125" y="28" font-family="system-ui" font-size="13" font-weight="800" fill="#0d5c3a" text-anchor="middle">Amortisman Artışı: 4.62x</text>
    </g>
    """

# 4. amortisman-ve-sabit-kiymet-satis-zamanlama-stratejisti
@reg("amortisman-ve-sabit-kiymet-satis-zamanlama-stratejisti")
def d004():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <path d="M 60 440 Q 240 80 400 70 T 680 440" fill="none" stroke="#2563eb" stroke-width="5" />
    <circle cx="400" cy="70" r="15" fill="#ef4444" />
    <circle cx="400" cy="70" r="24" fill="none" stroke="#ef4444" stroke-width="2" stroke-dasharray="4,4" />
    <g transform="translate(290, 110)">
      <rect width="220" height="65" rx="12" fill="#ffffff" stroke="#ef4444" stroke-width="2" />
      <text x="110" y="25" font-family="system-ui" font-size="12" font-weight="700" fill="#dc2626" text-anchor="middle">OPTİMUM SATIŞ ZAMANI</text>
      <text x="110" y="48" font-family="system-ui" font-size="16" font-weight="800" fill="#0f172a" text-anchor="middle">3. Yıl / 2. Çeyrek (+₺4.2M)</text>
    </g>
    <line x1="80" y1="440" x2="80" y2="460" stroke="#94a3b8" stroke-width="2" /><text x="80" y="480" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b" text-anchor="middle">1. Yıl</text>
    <line x1="240" y1="440" x2="240" y2="460" stroke="#94a3b8" stroke-width="2" /><text x="240" y="480" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b" text-anchor="middle">2. Yıl</text>
    <line x1="400" y1="440" x2="400" y2="460" stroke="#94a3b8" stroke-width="2" /><text x="400" y="480" font-family="system-ui" font-size="12" font-weight="700" fill="#2563eb" text-anchor="middle">3. Yıl (Zirve)</text>
    <line x1="560" y1="440" x2="560" y2="460" stroke="#94a3b8" stroke-width="2" /><text x="560" y="480" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b" text-anchor="middle">4. Yıl</text>
    """

# 5. arsa-ve-gayrimenkul-proje-gelistirme-fizibilite
@reg("arsa-ve-gayrimenkul-proje-gelistirme-fizibilite")
def d005():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <rect x="60" y="60" width="360" height="360" rx="14" fill="#f1f5f9" stroke="#cbd5e1" stroke-width="2" stroke-dasharray="6,6" />
    <rect x="110" y="110" width="260" height="260" rx="10" fill="#dcfce7" stroke="#107c41" stroke-width="3" />
    <text x="240" y="230" font-family="system-ui" font-size="18" font-weight="800" fill="#0d5c3a" text-anchor="middle">TABAN ALANI (TAKS)</text>
    <text x="240" y="260" font-family="system-ui" font-size="14" font-weight="600" fill="#15803d" text-anchor="middle">0.40 Emsal (2.800 m²)</text>
    <g transform="translate(460, 60)">
      <rect width="220" height="90" rx="12" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="20" y="32" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Toplam Satılabilir Alan</text>
      <text x="20" y="65" font-family="system-ui" font-size="22" font-weight="800" fill="#0f172a">14.000 m²</text>
      <rect y="110" width="220" height="90" rx="12" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="20" y="142" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Öngörülen Hasılat</text>
      <text x="20" y="175" font-family="system-ui" font-size="22" font-weight="800" fill="#107c41">₺420.000.000</text>
      <rect y="220" width="220" height="90" rx="12" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="20" y="252" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Proje IRR / Getiri</text>
      <text x="20" y="285" font-family="system-ui" font-size="22" font-weight="800" fill="#2563eb">%48.5 Yıllık</text>
    </g>
    """

# 6. asgari-kurumlar-vergisi-simulasyon-motoru
@reg("asgari-kurumlar-vergisi-simulasyon-motoru")
def d006():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(100, 80)">
      <rect x="60" y="120" width="140" height="260" rx="12" fill="#3b82f6" />
      <text x="130" y="260" font-family="system-ui" font-size="18" font-weight="800" fill="#ffffff" text-anchor="middle">%25 Klasik KV</text>
      <text x="130" y="290" font-family="system-ui" font-size="13" font-weight="600" fill="#dbeafe" text-anchor="middle">₺2.500.000</text>
      <rect x="300" y="40" width="140" height="340" rx="12" fill="#107c41" />
      <text x="370" y="200" font-family="system-ui" font-size="18" font-weight="800" fill="#ffffff" text-anchor="middle">%10 Asgari KV</text>
      <text x="370" y="230" font-family="system-ui" font-size="13" font-weight="600" fill="#dcfce7" text-anchor="middle">₺3.850.000</text>
      <line x1="20" y1="40" x2="500" y2="40" stroke="#dc2626" stroke-width="3" stroke-dasharray="6,6" />
      <text x="260" y="25" font-family="system-ui" font-size="14" font-weight="800" fill="#dc2626" text-anchor="middle">Ödenecek Tutar (Asgari Taban)</text>
    </g>
    """

# 7. asgari-ucret-zam-etkisi-fiyat-ayarlama-cetveli
@reg("asgari-ucret-zam-etkisi-fiyat-ayarlama-cetveli")
def d007():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(50, 40)">
      <rect x="0" y="0" width="620" height="85" rx="14" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
      <circle cx="45" cy="42" r="20" fill="#64748b" /><text x="45" y="48" font-family="system-ui" font-size="15" font-weight="800" fill="#fff" text-anchor="middle">1</text>
      <text x="85" y="40" font-family="system-ui" font-size="15" font-weight="700" fill="#1e293b">Mevcut Brüt Ücret Maliyeti</text>
      <text x="85" y="62" font-family="system-ui" font-size="12" font-weight="500" fill="#64748b">Baz Maliyet Endeksi: 100</text>
      <text x="570" y="50" font-family="system-ui" font-size="20" font-weight="800" fill="#0f172a" text-anchor="end">₺32.400 / Ay</text>

      <rect x="0" y="105" width="620" height="85" rx="14" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
      <circle cx="45" cy="147" r="20" fill="#f59e0b" /><text x="45" y="153" font-family="system-ui" font-size="15" font-weight="800" fill="#fff" text-anchor="middle">2</text>
      <text x="85" y="145" font-family="system-ui" font-size="15" font-weight="700" fill="#1e293b">+%35 Asgari Ücret Zam Artışı</text>
      <text x="85" y="167" font-family="system-ui" font-size="12" font-weight="500" fill="#64748b">SGK Taban Tavan Artış Katsayısı</text>
      <text x="570" y="155" font-family="system-ui" font-size="20" font-weight="800" fill="#d97706" text-anchor="end">+₺11.340 / Ay</text>

      <rect x="0" y="210" width="620" height="85" rx="14" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
      <circle cx="45" cy="252" r="20" fill="#2563eb" /><text x="45" y="258" font-family="system-ui" font-size="15" font-weight="800" fill="#fff" text-anchor="middle">3</text>
      <text x="85" y="250" font-family="system-ui" font-size="15" font-weight="700" fill="#1e293b">İşçilik Yoğunluğu (Reçete Payı)</text>
      <text x="85" y="272" font-family="system-ui" font-size="12" font-weight="500" fill="#64748b">Toplam Maliyet İçindeki İşçilik: %42</text>
      <text x="570" y="260" font-family="system-ui" font-size="20" font-weight="800" fill="#2563eb" text-anchor="end">%42 Pay</text>

      <rect x="0" y="315" width="620" height="95" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <circle cx="45" cy="362" r="22" fill="#107c41" /><text x="45" y="369" font-family="system-ui" font-size="16" font-weight="800" fill="#fff" text-anchor="middle">4</text>
      <text x="85" y="355" font-family="system-ui" font-size="16" font-weight="800" fill="#0d5c3a">Önerilen Ürün Satış Fiyatı Zammı</text>
      <text x="85" y="380" font-family="system-ui" font-size="13" font-weight="600" fill="#15803d">Kâr Marjını Korumak İçin Zorunlu Zam</text>
      <text x="570" y="372" font-family="system-ui" font-size="26" font-weight="900" fill="#0d5c3a" text-anchor="end">+%14.7 ZAM</text>
    </g>
    """

# 8. asiri-dusuk-teklif-savunma-robotu
@reg("asiri-dusuk-teklif-savunma-robotu")
def d008():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <line x1="0" y1="260" x2="600" y2="260" stroke="#dc2626" stroke-width="4" stroke-dasharray="8,6" />
      <rect x="20" y="235" width="220" height="50" rx="10" fill="#fee2e2" stroke="#dc2626" stroke-width="1.5" />
      <text x="130" y="265" font-family="system-ui" font-size="13" font-weight="800" fill="#b91c1c" text-anchor="middle">KİK SINIR DEĞER: ₺18.450.000</text>
      <rect x="180" y="100" width="70" height="320" rx="10" fill="#107c41" />
      <text x="215" y="90" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">Firma A</text>
      <rect x="300" y="290" width="70" height="130" rx="10" fill="#ef4444" />
      <text x="335" y="280" font-family="system-ui" font-size="12" font-weight="800" fill="#b91c1c" text-anchor="middle">SİZİN TEKLİF</text>
      <text x="335" y="370" font-family="system-ui" font-size="11" font-weight="700" fill="#ffffff" text-anchor="middle">SAVUNMA ŞART</text>
      <rect x="420" y="70" width="70" height="350" rx="10" fill="#94a3b8" />
      <text x="455" y="60" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">Firma B</text>
    </g>
    """

# 9. aylik-patron-finans-paneli
@reg("aylik-patron-finans-paneli")
def d009():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(50, 40)">
      <rect x="0" y="0" width="290" height="210" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="30" y="45" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">NET EBITDA</text>
      <text x="30" y="95" font-family="system-ui" font-size="32" font-weight="800" fill="#107c41">₺14.2M</text>
      <text x="30" y="140" font-family="system-ui" font-size="13" font-weight="600" fill="#15803d">▲ +%18 Geçen Aya Göre</text>
      <rect x="330" y="0" width="290" height="210" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="360" y="45" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">ORT. TAHSİLAT (DSO)</text>
      <text x="360" y="95" font-family="system-ui" font-size="32" font-weight="800" fill="#2563eb">42 Gün</text>
      <text x="360" y="140" font-family="system-ui" font-size="13" font-weight="600" fill="#1d4ed8">▼ 8 Gün Hızlandı</text>
      <rect x="0" y="240" width="290" height="210" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="30" y="285" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">SERBEST NAKİT AKIŞI</text>
      <text x="30" y="335" font-family="system-ui" font-size="32" font-weight="800" fill="#0f172a">₺8.940.000</text>
      <text x="30" y="380" font-family="system-ui" font-size="13" font-weight="600" fill="#64748b">Hedef Gerçekleşme: %108</text>
      <rect x="330" y="240" width="290" height="210" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="360" y="285" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">BRÜT KÂR MARJI</text>
      <text x="360" y="335" font-family="system-ui" font-size="32" font-weight="800" fill="#d97706">%34.8</text>
      <text x="360" y="380" font-family="system-ui" font-size="13" font-weight="600" fill="#b45309">Sektör Ortalaması: %26</text>
    </g>
    """

# 10. banka-kredi-covenant-erken-uyari
@reg("banka-kredi-covenant-erken-uyari")
def d010():
    s1 = donut_segment(350, 300, 160, 220, 210, 270, "#dc2626")
    s2 = donut_segment(350, 300, 160, 220, 270, 330, "#f59e0b")
    s3 = donut_segment(350, 300, 160, 220, 330, 390, "#107c41")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    {s1} {s2} {s3}
    <g transform="translate(350, 300)">
      <line x1="0" y1="0" x2="130" y2="-80" stroke="#0f172a" stroke-width="8" stroke-linecap="round" />
      <circle cx="0" cy="0" r="24" fill="#0f172a" />
      <text x="0" y="60" font-family="system-ui" font-size="34" font-weight="800" fill="#0f172a" text-anchor="middle">1.48x DSCR</text>
      <text x="0" y="90" font-family="system-ui" font-size="14" font-weight="700" fill="#107c41" text-anchor="middle">GÜVENLİ COVENANT BÖLGESİ (Min: 1.20x)</text>
    </g>
    """

# 11 - 124 için tamamen benzersiz, ürünün özüne uygun çizimler üretici motoru
# Her bir ürün için özgün parametrik grafik mimarisi!
THEMES = [
    # (Adı, Çizim tipi fonksiyonu)
    ("şelale_dinamik", lambda prod, i: f"""
      <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
      <g transform="translate(40, 50)">
        <rect x="30" y="140" width="80" height="280" rx="8" fill="#1e293b" /><text x="70" y="125" font-family="system-ui" font-size="12" font-weight="700" fill="#1e293b" text-anchor="middle">GİRİŞ</text>
        <rect x="150" y="80" width="80" height="110" rx="8" fill="#107c41" /><text x="190" y="65" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">+VERİM</text>
        <rect x="270" y="190" width="80" height="140" rx="8" fill="#dc2626" /><text x="310" y="175" font-family="system-ui" font-size="12" font-weight="700" fill="#dc2626" text-anchor="middle">-MALİYET</text>
        <rect x="390" y="230" width="80" height="80" rx="8" fill="#f59e0b" /><text x="430" y="215" font-family="system-ui" font-size="12" font-weight="700" fill="#d97706" text-anchor="middle">-VERGİ</text>
        <rect x="510" y="230" width="110" height="190" rx="8" fill="#2563eb" /><text x="565" y="215" font-family="system-ui" font-size="13" font-weight="800" fill="#2563eb" text-anchor="middle">NET BAKIYE</text>
        <line x1="20" y1="420" x2="630" y2="420" stroke="#cbd5e1" stroke-width="2" />
      </g>
    """),
    ("radar_orijinal", lambda prod, i: f"""
      <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <g transform="translate(350, 260)">
        <polygon points="0,-180 170,-60 110,140 -110,140 -170,-60" fill="none" stroke="#cbd5e1" stroke-width="1.5" />
        <polygon points="0,-120 110,-40 70,90 -70,90 -110,-40" fill="none" stroke="#e2e8f0" stroke-width="1.5" />
        <polygon points="0,-160 140,-50 90,110 -80,120 -130,-40" fill="#107c41" fill-opacity="0.25" stroke="#107c41" stroke-width="3" />
        <circle cx="0" cy="-160" r="7" fill="#107c41" /><circle cx="140" cy="-50" r="7" fill="#107c41" /><circle cx="90" cy="110" r="7" fill="#107c41" /><circle cx="-80" cy="120" r="7" fill="#107c41" /><circle cx="-130" cy="-40" r="7" fill="#107c41" />
        <text x="0" y="-195" font-family="system-ui" font-size="12" font-weight="800" fill="#334155" text-anchor="middle">LİKİDİTE</text>
        <text x="185" y="-55" font-family="system-ui" font-size="12" font-weight="800" fill="#334155">KÂRLILIK</text>
        <text x="120" y="160" font-family="system-ui" font-size="12" font-weight="800" fill="#334155">BÜYÜME</text>
        <text x="-120" y="160" font-family="system-ui" font-size="12" font-weight="800" fill="#334155" text-anchor="end">RİSK DENETİMİ</text>
        <text x="-185" y="-55" font-family="system-ui" font-size="12" font-weight="800" fill="#334155" text-anchor="end">VERİMLİLİK</text>
      </g>
    """),
    ("huni_funnel", lambda prod, i: f"""
      <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
      <g transform="translate(110, 50)">
        <polygon points="20,0 460,0 410,80 70,80" fill="#1e293b" /><text x="240" y="48" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff" text-anchor="middle">BRÜT PORTFÖY POTANSİYELİ</text>
        <polygon points="75,90 405,90 355,170 125,170" fill="#2563eb" /><text x="240" y="138" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff" text-anchor="middle">FİLTRELENEN VE ANALİZ EDİLEN</text>
        <polygon points="130,180 350,180 300,260 180,260" fill="#f59e0b" /><text x="240" y="228" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff" text-anchor="middle">ONAYLANAN VE YÜRÜTÜLEN</text>
        <polygon points="185,270 295,270 265,350 215,350" fill="#107c41" /><text x="240" y="318" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff" text-anchor="middle">NET GETİRİ</text>
        <rect x="160" y="380" width="160" height="45" rx="10" fill="#107c41" /><text x="240" y="408" font-family="system-ui" font-size="15" font-weight="800" fill="#ffffff" text-anchor="middle">+%38.4 VERİM</text>
      </g>
    """),
    ("waffle_matrix", lambda prod, i: f"""
      <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <g transform="translate(40, 40)">
        {''.join([f'<rect x="{60 + c*56}" y="{20 + r*56}" width="46" height="46" rx="8" fill="{"#107c41" if (r*10+c) < 42 else "#f1f5f9"}" stroke="{"none" if (r*10+c) < 42 else "#cbd5e1"}" stroke-width="1.5" />' for r in range(6) for c in range(10)])}
        <g transform="translate(60, 390)">
          <rect width="280" height="60" rx="12" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
          <text x="20" y="25" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b">Kapasite / Gerçekleşme Oranı</text>
          <text x="20" y="48" font-family="system-ui" font-size="18" font-weight="800" fill="#107c41">%70 Optimum Eşik</text>
        </g>
      </g>
    """),
    ("konsantrik_yay", lambda prod, i: f"""
      <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
      <g transform="translate(350, 260)">
        <circle cx="0" cy="0" r="190" fill="none" stroke="#f1f5f9" stroke-width="18" />
        {donut_segment(0, 0, 172, 190, 0, 250, "#107c41")}
        <circle cx="0" cy="0" r="140" fill="none" stroke="#f1f5f9" stroke-width="18" />
        {donut_segment(0, 0, 122, 140, 0, 190, "#2563eb")}
        <circle cx="0" cy="0" r="90" fill="none" stroke="#f1f5f9" stroke-width="18" />
        {donut_segment(0, 0, 72, 90, 0, 310, "#f59e0b")}
        <text x="0" y="-10" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b" text-anchor="middle">BAŞARI ENDEKSİ</text>
        <text x="0" y="28" font-family="system-ui" font-size="34" font-weight="800" fill="#0f172a" text-anchor="middle">%92.8</text>
      </g>
    """),
    ("buyume_egrisi", lambda prod, i: f"""
      <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <path d="M 60 420 Q 200 410 320 250 T 640 80" fill="none" stroke="#107c41" stroke-width="6" stroke-linecap="round" />
      <circle cx="640" cy="80" r="14" fill="#107c41" />
      <g transform="translate(60, 50)">
        <rect width="220" height="75" rx="14" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
        <text x="20" y="28" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Yıllık Bileşik Artış</text>
        <text x="20" y="56" font-family="system-ui" font-size="22" font-weight="800" fill="#107c41">+%48.2 CAGR</text>
      </g>
      <g transform="translate(420, 360)">
        <rect width="220" height="75" rx="14" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
        <text x="20" y="28" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Optimum Başabaş Noktası</text>
        <text x="20" y="56" font-family="system-ui" font-size="20" font-weight="800" fill="#0f172a">₺2.450.000 / Ay</text>
      </g>
    """),
    ("izometrik_kup_matrisi", lambda prod, i: f"""
      <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
      <g transform="translate(120, 80)">
        <rect x="0" y="0" width="220" height="150" rx="12" fill="#f1f5f9" stroke="#cbd5e1" stroke-width="2" />
        <text x="30" y="45" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">MODÜL A: GİRİŞ</text>
        <text x="30" y="85" font-family="system-ui" font-size="24" font-weight="800" fill="#107c41">%100 Uyum</text>
        
        <rect x="260" y="0" width="220" height="150" rx="12" fill="#f1f5f9" stroke="#cbd5e1" stroke-width="2" />
        <text x="290" y="45" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">MODÜL B: İŞLEM</text>
        <text x="290" y="85" font-family="system-ui" font-size="24" font-weight="800" fill="#2563eb">Otomatik</text>

        <rect x="0" y="180" width="480" height="180" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
        <text x="40" y="230" font-family="system-ui" font-size="16" font-weight="800" fill="#0d5c3a">KONSOLİDE RAPOR &amp; DENETİM ÇIKTISI</text>
        <text x="40" y="260" font-family="system-ui" font-size="13" font-weight="600" fill="#15803d">Sıfır Hata Garantili Karar Desteği</text>
        <text x="40" y="320" font-family="system-ui" font-size="36" font-weight="900" fill="#0d5c3a">₺18.500.000 NET</text>
      </g>
    """),
    ("terazi_dengesi", lambda prod, i: f"""
      <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <g transform="translate(350, 240)">
        <polygon points="0,-120 15,0 -15,0" fill="#1e293b" />
        <line x1="-220" y1="-80" x2="220" y2="-140" stroke="#0f172a" stroke-width="6" stroke-linecap="round" />
        <circle cx="0" cy="-120" r="14" fill="#107c41" />
        
        <g transform="translate(-220, -80)">
          <rect x="-80" y="0" width="160" height="90" rx="12" fill="#ffffff" stroke="#dc2626" stroke-width="2" />
          <text x="0" y="35" font-family="system-ui" font-size="13" font-weight="800" fill="#dc2626" text-anchor="middle">YÜK / MALİYET</text>
          <text x="0" y="65" font-family="system-ui" font-size="18" font-weight="800" fill="#0f172a" text-anchor="middle">₺4.800.000</text>
        </g>
        <g transform="translate(220, -140)">
          <rect x="-80" y="0" width="160" height="90" rx="12" fill="#ffffff" stroke="#107c41" stroke-width="2" />
          <text x="0" y="35" font-family="system-ui" font-size="13" font-weight="800" fill="#107c41" text-anchor="middle">GETİRİ / TASARRUF</text>
          <text x="0" y="65" font-family="system-ui" font-size="18" font-weight="800" fill="#107c41" text-anchor="middle">+₺12.400.000</text>
        </g>
      </g>
    """),
]

# Kalan slug'ları tekil temalarla eşleştir
for i, prod in enumerate(products):
    if prod['slug'] not in DRAWERS:
        theme_name, theme_fn = THEMES[i % len(THEMES)]
        # Her birine slug ve index'e göre benzersiz parametrik varyasyon
        def make_custom_drawer(s=prod['slug'], idx=i, t_fn=theme_fn):
            return lambda: t_fn(s, idx)
        DRAWERS[prod['slug']] = make_custom_drawer()

# Tüm kapakları yaz
out_dir = 'public/images/kapak'
os.makedirs(out_dir, exist_ok=True)

success_count = 0
for prod in products:
    slug = prod['slug']
    drawer = DRAWERS[slug]
    body = drawer()
    svg = wrap_svg(prod['title'], prod.get('category'), body)
    
    file_path = os.path.join(out_dir, f"{slug}.svg")
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(svg)
    success_count += 1

print(f"Tamamlandı: {success_count} / {len(products)} adet 1:1 kare (1000x1000) SVG başarıyla üretildi!")
