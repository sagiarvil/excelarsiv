# scripts/covers/part1.py
# 1'den 31'e kadar olan ürünlerin özel çizimleri

import math

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

PART1_DRAWERS = {}

def reg(slug):
    def deco(fn):
        PART1_DRAWERS[slug] = fn
        return fn
    return deco

# 1. 13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi
@reg("13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi")
def draw_001():
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
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
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
def draw_002():
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
def draw_003():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <path d="M 60 420 L 200 350 L 360 270 L 520 190 L 660 90" fill="none" stroke="#107c41" stroke-width="6" stroke-linecap="round" />
    <circle cx="60" cy="420" r="12" fill="#ffffff" stroke="#107c41" stroke-width="4" /><text x="60" y="465" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">Tarihi Maliyet</text>
    <circle cx="200" cy="350" r="12" fill="#ffffff" stroke="#107c41" stroke-width="4" /><text x="200" y="395" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">VUK 298/Ç</text>
    <circle cx="360" cy="270" r="12" fill="#ffffff" stroke="#107c41" stroke-width="4" /><text x="360" y="315" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">Geçici 32. Md</text>
    <circle cx="520" cy="190" r="12" fill="#ffffff" stroke="#107c41" stroke-width="4" /><text x="520" y="235" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">2026 Değeri</text>
    <circle cx="660" cy="90" r="14" fill="#107c41" /><text x="660" y="60" font-family="system-ui" font-size="14" font-weight="800" fill="#107c41" text-anchor="middle">₺46.200.000</text>
    <g transform="translate(60, 50)">
      <rect width="250" height="45" rx="22" fill="#dcfce7" />
      <text x="125" y="28" font-family="system-ui" font-size="13" font-weight="800" fill="#0d5c3a" text-anchor="middle">Amortisman Artışı: 4.62x</text>
    </g>
    """

# 4. amortisman-ve-sabit-kiymet-satis-zamanlama-stratejisti
@reg("amortisman-ve-sabit-kiymet-satis-zamanlama-stratejisti")
def draw_004():
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
def draw_005():
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
def draw_006():
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
def draw_007():
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
def draw_008():
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
def draw_009():
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
def draw_010():
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

# 11. banka-kredi-ve-taksit-takip-sistemi
@reg("banka-kredi-ve-taksit-takip-sistemi")
def draw_011():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(50, 40)">
      <!-- Amortisman Tüneli / Stacked Area -->
      <polygon points="50,400 600,400 600,280 50,120" fill="#dbeafe" />
      <polygon points="50,120 600,280 600,100 50,60" fill="#dcfce7" />
      <line x1="50" y1="120" x2="600" y2="280" stroke="#107c41" stroke-width="4" />
      <text x="250" y="340" font-family="system-ui" font-size="18" font-weight="800" fill="#1e3a8a">ANAPARA GERİ ÖDEME DİLİMİ</text>
      <text x="250" y="95" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">FAİZ YÜKÜ (AZALAN BAKİYE)</text>
      <g transform="translate(50, 420)">
        <rect width="250" height="50" rx="10" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" />
        <text x="20" y="30" font-family="system-ui" font-size="13" font-weight="700" fill="#334155">Kalan Anapara: ₺4.850.000</text>
      </g>
      <g transform="translate(350, 420)">
        <rect width="250" height="50" rx="10" fill="#f8fafc" stroke="#86efac" stroke-width="1.5" />
        <text x="20" y="30" font-family="system-ui" font-size="13" font-weight="700" fill="#107c41">Aylık Taksit: ₺185.400</text>
      </g>
    </g>
    """

# 12. borcla-sirket-satin-alma-ve-yatirim-getirisi
@reg("borcla-sirket-satin-alma-ve-yatirim-getirisi")
def draw_012():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(150, 50)">
      <!-- LBO Kaldıraç Piramidi -->
      <polygon points="200,40 280,140 120,140" fill="#107c41" />
      <text x="200" y="115" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff" text-anchor="middle">ÖZSERMAYE (%20)</text>

      <polygon points="115,150 285,150 335,260 65,260" fill="#2563eb" />
      <text x="200" y="215" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff" text-anchor="middle">MEZANİN FİNANSMAN (%30)</text>

      <polygon points="60,270 340,270 395,400 5,400" fill="#1e293b" />
      <text x="200" y="345" font-family="system-ui" font-size="15" font-weight="800" fill="#ffffff" text-anchor="middle">KIDEMLİ BANKA KREDİSİ (%50)</text>
    </g>
    <g transform="translate(520, 180)">
      <rect width="180" height="85" rx="12" fill="#ffffff" stroke="#107c41" stroke-width="2" />
      <text x="20" y="30" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b">Öngörülen IRR</text>
      <text x="20" y="65" font-family="system-ui" font-size="24" font-weight="900" fill="#107c41">%34.8 / Yıl</text>
    </g>
    """

# 13. cari-ba-bs-toplu-mutabakat
@reg("cari-ba-bs-toplu-mutabakat")
def draw_013():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(350, 240)">
      <!-- İkili Mutabakat Venn Kümesi -->
      <circle cx="-90" cy="0" r="160" fill="#3b82f6" fill-opacity="0.25" stroke="#2563eb" stroke-width="3" />
      <circle cx="90" cy="0" r="160" fill="#107c41" fill-opacity="0.25" stroke="#107c41" stroke-width="3" />
      <text x="-160" y="-10" font-family="system-ui" font-size="16" font-weight="800" fill="#1e40af">BİZİM KAYIT (BA)</text>
      <text x="-160" y="15" font-family="system-ui" font-size="13" font-weight="600" fill="#3b82f6">1.420 Fatura</text>
      <text x="160" y="-10" font-family="system-ui" font-size="16" font-weight="800" fill="#065f46">KARŞI CARİ (BS)</text>
      <text x="160" y="15" font-family="system-ui" font-size="13" font-weight="600" fill="#107c41">1.418 Fatura</text>
      <text x="0" y="-15" font-family="system-ui" font-size="16" font-weight="900" fill="#0f172a" text-anchor="middle">TAM EŞLEŞEN</text>
      <text x="0" y="15" font-family="system-ui" font-size="24" font-weight="900" fill="#107c41" text-anchor="middle">%99.8</text>
      <text x="0" y="40" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">Fark: 2 Evrak (₺4.200)</text>
    </g>
    """

# 14. cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi
@reg("cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi")
def draw_014():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <!-- 4x4 Risk Isı Matrisi -->
      <text x="300" y="20" font-family="system-ui" font-size="16" font-weight="800" fill="#0f172a" text-anchor="middle">MÜŞTERİ RİSK &amp; VADE MATRİSİ</text>
      <rect x="80" y="50" width="220" height="170" fill="#fee2e2" stroke="#fca5a5" stroke-width="2" rx="8" />
      <text x="190" y="130" font-family="system-ui" font-size="14" font-weight="800" fill="#b91c1c" text-anchor="middle">YÜKSEK RİSK / 90+ GÜN</text>
      <text x="190" y="155" font-family="system-ui" font-size="18" font-weight="900" fill="#b91c1c" text-anchor="middle">₺1.850.000</text>
      <rect x="320" y="50" width="220" height="170" fill="#fef3c7" stroke="#fcd34d" stroke-width="2" rx="8" />
      <text x="430" y="130" font-family="system-ui" font-size="14" font-weight="800" fill="#b45309" text-anchor="middle">ORTA RİSK / 60 GÜN</text>
      <text x="430" y="155" font-family="system-ui" font-size="18" font-weight="900" fill="#b45309" text-anchor="middle">₺3.400.000</text>
      <rect x="80" y="240" width="460" height="170" fill="#dcfce7" stroke="#86efac" stroke-width="2" rx="8" />
      <text x="310" y="320" font-family="system-ui" font-size="16" font-weight="800" fill="#15803d" text-anchor="middle">GÜVENLİ BÖLGE (DÜZENLİ TAHSİLAT 0-30 GÜN)</text>
      <text x="310" y="350" font-family="system-ui" font-size="24" font-weight="900" fill="#0d5c3a" text-anchor="middle">₺18.950.000</text>
    </g>
    """

# 15. cek-senet-ve-vade-risk-sistemi
@reg("cek-senet-ve-vade-risk-sistemi")
def draw_015():
    s1 = donut_segment(350, 260, 150, 200, 0, 180, "#107c41")
    s2 = donut_segment(350, 260, 150, 200, 180, 290, "#f59e0b")
    s3 = donut_segment(350, 260, 150, 200, 290, 360, "#dc2626")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    {s1} {s2} {s3}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">TOPLAM PORTFÖY</text>
      <text x="0" y="20" font-family="system-ui" font-size="28" font-weight="900" fill="#0f172a" text-anchor="middle">₺38.5M</text>
      <text x="0" y="45" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">Portföy Güvence Oranı: %88</text>
    </g>
    """

# 16. dava-masraf-harc-hesaplama
@reg("dava-masraf-harc-hesaplama")
def draw_016():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <!-- Nispi Harç Merdiveni -->
      <rect x="40" y="320" width="110" height="120" rx="8" fill="#cbd5e1" /><text x="95" y="380" font-family="system-ui" font-size="13" font-weight="800" fill="#1e293b" text-anchor="middle">BAŞVURU</text>
      <rect x="170" y="240" width="110" height="200" rx="8" fill="#93c5fd" /><text x="225" y="340" font-family="system-ui" font-size="13" font-weight="800" fill="#1e40af" text-anchor="middle">PEŞİN HARÇ</text>
      <rect x="300" y="160" width="110" height="280" rx="8" fill="#fde047" /><text x="355" y="300" font-family="system-ui" font-size="13" font-weight="800" fill="#854d0e" text-anchor="middle">GİDER AVANSI</text>
      <rect x="430" y="80" width="150" height="360" rx="8" fill="#86efac" /><text x="505" y="260" font-family="system-ui" font-size="15" font-weight="900" fill="#065f46" text-anchor="middle">VEKALET ÜCRETİ</text>
      <line x1="20" y1="440" x2="600" y2="440" stroke="#94a3b8" stroke-width="2" />
    </g>
    """

# 17. defter-beyan-e-arsiv-aktarim
@reg("defter-beyan-e-arsiv-aktarim")
def draw_017():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(50, 160)">
      <!-- XML Pipeline Akışı -->
      <rect x="20" y="40" width="160" height="110" rx="14" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="100" y="85" font-family="system-ui" font-size="15" font-weight="800" fill="#0f172a" text-anchor="middle">GİB e-Arşiv</text>
      <text x="100" y="110" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b" text-anchor="middle">XML / JSON Portal</text>
      <line x1="185" y1="95" x2="255" y2="95" stroke="#107c41" stroke-width="4" stroke-linecap="round" />
      <polygon points="255,95 240,88 240,102" fill="#107c41" />
      <rect x="260" y="20" width="180" height="150" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="350" y="70" font-family="system-ui" font-size="16" font-weight="900" fill="#0d5c3a" text-anchor="middle">OTOMATİK DOĞRULAYICI</text>
      <text x="350" y="100" font-family="system-ui" font-size="13" font-weight="700" fill="#15803d" text-anchor="middle">Sıfır Hata Eşleşmesi</text>
      <line x1="445" y1="95" x2="515" y2="95" stroke="#107c41" stroke-width="4" stroke-linecap="round" />
      <polygon points="515,95 500,88 500,102" fill="#107c41" />
      <rect x="520" y="40" width="160" height="110" rx="14" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="600" y="85" font-family="system-ui" font-size="15" font-weight="800" fill="#0f172a" text-anchor="middle">Defter Beyan</text>
      <text x="600" y="110" font-family="system-ui" font-size="12" font-weight="600" fill="#107c41" text-anchor="middle">✓ Hazır Aktarım</text>
    </g>
    """

# 18. depo-doluluk-lokasyon
@reg("depo-doluluk-lokasyon")
def draw_018():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <!-- Depo Raf Matrisi -->
      <text x="300" y="20" font-family="system-ui" font-size="16" font-weight="800" fill="#0f172a" text-anchor="middle">DEPO RAF GÖZÜ DOLULUK HARİTASI</text>
      <rect x="40" y="50" width="240" height="170" rx="10" fill="#ffffff" stroke="#107c41" stroke-width="2" />
      <text x="60" y="85" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">BLOK A (Hızlı Tüketim)</text>
      <text x="60" y="125" font-family="system-ui" font-size="28" font-weight="900" fill="#107c41">%94 Dolu</text>
      <rect x="320" y="50" width="240" height="170" rx="10" fill="#ffffff" stroke="#2563eb" stroke-width="2" />
      <text x="340" y="85" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">BLOK B (Yedek Parça)</text>
      <text x="340" y="125" font-family="system-ui" font-size="28" font-weight="900" fill="#2563eb">%68 Dolu</text>
      <rect x="40" y="240" width="520" height="160" rx="10" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="60" y="275" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">PALET DOLULUK VERİMİ</text>
      <rect x="60" y="300" width="480" height="35" rx="8" fill="#f1f5f9" />
      <rect x="60" y="300" width="400" height="35" rx="8" fill="#107c41" />
      <text x="470" y="323" font-family="system-ui" font-size="13" font-weight="800" fill="#ffffff">%83.3</text>
    </g>
    """

# 19. dernek-kooperatif-mali-yonetim
@reg("dernek-kooperatif-mali-yonetim")
def draw_019():
    s1 = donut_segment(350, 260, 140, 200, 0, 160, "#2563eb")
    s2 = donut_segment(350, 260, 140, 200, 160, 270, "#107c41")
    s3 = donut_segment(350, 260, 140, 200, 270, 360, "#f59e0b")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    {s1} {s2} {s3}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">FON DAĞILIMI</text>
      <text x="0" y="20" font-family="system-ui" font-size="28" font-weight="900" fill="#0f172a" text-anchor="middle">₺12.8M</text>
      <text x="0" y="45" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">Bütçe Disiplini: %100</text>
    </g>
    """

# 20. doviz-acik-pozisyonu-ve-kur-riski-stres-testi
@reg("doviz-acik-pozisyonu-ve-kur-riski-stres-testi")
def draw_020():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <text x="300" y="20" font-family="system-ui" font-size="16" font-weight="800" fill="#0f172a" text-anchor="middle">DÖVİZ ŞOKU DUYARLILIK YELPAZESİ</text>
      <rect x="40" y="60" width="160" height="240" rx="10" fill="#ffffff" stroke="#f59e0b" stroke-width="2" />
      <text x="120" y="100" font-family="system-ui" font-size="15" font-weight="800" fill="#d97706" text-anchor="middle">+%10 KUR</text>
      <text x="120" y="180" font-family="system-ui" font-size="20" font-weight="900" fill="#0f172a" text-anchor="middle">-₺1.2M</text>
      <rect x="220" y="60" width="160" height="280" rx="10" fill="#ffffff" stroke="#ea580c" stroke-width="2" />
      <text x="300" y="100" font-family="system-ui" font-size="15" font-weight="800" fill="#c2410c" text-anchor="middle">+%20 KUR</text>
      <text x="300" y="200" font-family="system-ui" font-size="20" font-weight="900" fill="#0f172a" text-anchor="middle">-₺2.4M</text>
      <rect x="400" y="60" width="160" height="320" rx="10" fill="#fee2e2" stroke="#dc2626" stroke-width="2" />
      <text x="480" y="100" font-family="system-ui" font-size="15" font-weight="800" fill="#b91c1c" text-anchor="middle">+%30 ŞOK</text>
      <text x="480" y="220" font-family="system-ui" font-size="22" font-weight="900" fill="#b91c1c" text-anchor="middle">-₺3.6M</text>
    </g>
    """

# 21. e-belge-zorunluluk-radari
@reg("e-belge-zorunluluk-radari")
def draw_021():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(350, 260)">
      <circle cx="0" cy="0" r="190" fill="none" stroke="#e2e8f0" stroke-width="2" stroke-dasharray="6,6" />
      <circle cx="0" cy="0" r="130" fill="none" stroke="#cbd5e1" stroke-width="2" />
      <circle cx="0" cy="0" r="70" fill="none" stroke="#107c41" stroke-width="3" />
      <line x1="0" y1="0" x2="160" y2="-90" stroke="#107c41" stroke-width="4" stroke-linecap="round" />
      <text x="0" y="-140" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">5M CİRO EŞİĞİ</text>
      <text x="0" y="-80" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">3M ZORUNLU EŞİK</text>
      <text x="0" y="15" font-family="system-ui" font-size="16" font-weight="900" fill="#0f172a" text-anchor="middle">E-FATURA ŞART</text>
    </g>
    """

# 22. e-fatura-satir-defteri-pdf-kaniti
@reg("e-fatura-satir-defteri-pdf-kaniti")
def draw_022():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(80, 50)">
      <!-- Belge Doğrulama Kartı -->
      <rect x="0" y="0" width="540" height="380" rx="16" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" />
      <rect x="30" y="30" width="140" height="30" rx="6" fill="#dcfce7" />
      <text x="100" y="50" font-family="system-ui" font-size="12" font-weight="800" fill="#15803d" text-anchor="middle">✓ HASH DOĞRULANDI</text>
      <line x1="30" y1="90" x2="510" y2="90" stroke="#e2e8f0" stroke-width="2" />
      <text x="30" y="130" font-family="system-ui" font-size="14" font-weight="700" fill="#334155">ETTN: 4a8f9b2c-7e1d-4893-b6a4-8f2c3d1e0a5b</text>
      <text x="30" y="170" font-family="system-ui" font-size="14" font-weight="700" fill="#334155">Satır Sayısı: 1.482 Satır Kalem Dökümü</text>
      <text x="30" y="210" font-family="system-ui" font-size="14" font-weight="700" fill="#334155">KDV Tevkifat Doğrulaması: %100 Uyumlu</text>
      <rect x="30" y="260" width="480" height="80" rx="12" fill="#f1f5f9" />
      <text x="50" y="305" font-family="system-ui" font-size="20" font-weight="800" fill="#0f172a">Yasal Delil &amp; İnceleme Kanıtı Hazır</text>
    </g>
    """

# 23. e-fatura-toplu-donusturucu
@reg("e-fatura-toplu-donusturucu")
def draw_023():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(100, 100)">
      <!-- Dişli Çark Dönüştürücü Şeması -->
      <circle cx="150" cy="150" r="100" fill="#dbeafe" stroke="#2563eb" stroke-width="4" />
      <text x="150" y="145" font-family="system-ui" font-size="18" font-weight="900" fill="#1e40af" text-anchor="middle">XML / UBL</text>
      <text x="150" y="170" font-family="system-ui" font-size="12" font-weight="700" fill="#3b82f6" text-anchor="middle">Giriş Dosyası</text>
      <circle cx="360" cy="150" r="100" fill="#dcfce7" stroke="#107c41" stroke-width="4" />
      <text x="360" y="145" font-family="system-ui" font-size="18" font-weight="900" fill="#065f46" text-anchor="middle">EXCEL / PDF</text>
      <text x="360" y="170" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">Net Çıktı</text>
      <path d="M 230 110 L 280 110" stroke="#0f172a" stroke-width="6" stroke-linecap="round" />
      <polygon points="280,110 265,100 265,120" fill="#0f172a" />
      <path d="M 280 190 L 230 190" stroke="#0f172a" stroke-width="6" stroke-linecap="round" />
      <polygon points="230,190 245,180 245,200" fill="#0f172a" />
    </g>
    """

# 24. e-ticaret-gercek-karlilik-fiyatlama
@reg("e-ticaret-gercek-karlilik-fiyatlama")
def draw_024():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(80, 50)">
      <rect x="40" y="20" width="100" height="380" rx="10" fill="#1e293b" /><text x="90" y="425" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">SATIŞ FİYATI</text>
      <rect x="180" y="160" width="80" height="240" rx="8" fill="#3b82f6" /><text x="220" y="425" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">ÜRÜN MALİYETİ</text>
      <rect x="280" y="240" width="80" height="160" rx="8" fill="#f59e0b" /><text x="320" y="425" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">KOMİSYON</text>
      <rect x="380" y="280" width="80" height="120" rx="8" fill="#dc2626" /><text x="420" y="425" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">KARGO/İADE</text>
      <rect x="480" y="140" width="80" height="260" rx="8" fill="#107c41" /><text x="520" y="425" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41" text-anchor="middle">NET KÂR</text>
    </g>
    """

# 25. elektrik-faturasi-dogrulama
@reg("elektrik-faturasi-dogrulama")
def draw_025():
    s1 = donut_segment(350, 260, 150, 200, 220, 290, "#107c41")
    s2 = donut_segment(350, 260, 150, 200, 290, 360, "#dc2626")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    {s1} {s2}
    <g transform="translate(350, 260)">
      <line x1="0" y1="0" x2="-100" y2="-90" stroke="#0f172a" stroke-width="6" stroke-linecap="round" />
      <circle cx="0" cy="0" r="18" fill="#0f172a" />
      <text x="0" y="40" font-family="system-ui" font-size="28" font-weight="900" fill="#107c41" text-anchor="middle">%14.2 İNDİRİFİK</text>
      <text x="0" y="65" font-family="system-ui" font-size="13" font-weight="700" fill="#15803d" text-anchor="middle">Reaktif Ceza Sınırının Altında (Güvenli)</text>
    </g>
    """

# 26. emlak-pipeline-komisyon
@reg("emlak-pipeline-komisyon")
def draw_026():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(100, 40)">
      <!-- Emlak Satış Hunisi -->
      <polygon points="20,0 480,0 420,80 80,80" fill="#1e293b" /><text x="250" y="48" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff" text-anchor="middle">İLAN PORTFÖYÜ (₺180M)</text>
      <polygon points="85,90 415,90 365,170 135,170" fill="#2563eb" /><text x="250" y="138" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff" text-anchor="middle">CİDDİ MÜŞTERİ GÖSTERİMİ</text>
      <polygon points="140,180 360,180 310,260 190,260" fill="#f59e0b" /><text x="250" y="228" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff" text-anchor="middle">TEKLİF &amp; SÖZLEŞME AŞAMASI</text>
      <polygon points="195,270 305,270 275,350 225,350" fill="#107c41" /><text x="250" y="318" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff" text-anchor="middle">KAPANIŞ</text>
      <rect x="160" y="375" width="180" height="50" rx="12" fill="#107c41" /><text x="250" y="405" font-family="system-ui" font-size="16" font-weight="800" fill="#ffffff" text-anchor="middle">NET KOMİSYON ₺3.6M</text>
    </g>
    """

# 27. emsal-sirket-carpanlariyla-degerleme
@reg("emsal-sirket-carpanlariyla-degerleme")
def draw_027():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <!-- Box-Plot Çarpan Bandı -->
      <line x1="80" y1="120" x2="560" y2="120" stroke="#cbd5e1" stroke-width="3" />
      <rect x="220" y="80" width="220" height="80" rx="8" fill="#dbeafe" stroke="#2563eb" stroke-width="2" />
      <line x1="320" y1="80" x2="320" y2="160" stroke="#1d4ed8" stroke-width="4" />
      <text x="320" y="60" font-family="system-ui" font-size="14" font-weight="800" fill="#1d4ed8" text-anchor="middle">MEDYAN: 8.4x EV/EBITDA</text>
      <g transform="translate(80, 240)">
        <rect width="480" height="120" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
        <text x="30" y="45" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">HEDEFLENEN ŞİRKET DEĞERİ (FV)</text>
        <text x="30" y="90" font-family="system-ui" font-size="36" font-weight="900" fill="#0f172a">₺142.500.000</text>
      </g>
    </g>
    """

# 28. fason-uretim-takip-sistemi
@reg("fason-uretim-takip-sistemi")
def draw_028():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(40, 50)">
      <!-- Kanban Fason Sipariş Kartları -->
      <rect x="20" y="40" width="180" height="340" rx="14" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" />
      <text x="40" y="75" font-family="system-ui" font-size="14" font-weight="800" fill="#475569">HAMMADDE ÇIKIŞI</text>
      <rect x="35" y="100" width="150" height="60" rx="8" fill="#f1f5f9" /><text x="50" y="135" font-family="system-ui" font-size="12" font-weight="700" fill="#334155">Parti #104 - 5.000 Adet</text>
      <rect x="230" y="40" width="180" height="340" rx="14" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" />
      <text x="250" y="75" font-family="system-ui" font-size="14" font-weight="800" fill="#2563eb">ATÖLYE DİKİMDE</text>
      <rect x="245" y="100" width="150" height="60" rx="8" fill="#dbeafe" /><text x="260" y="135" font-family="system-ui" font-size="12" font-weight="700" fill="#1e40af">Parti #102 - %85 Bitti</text>
      <rect x="440" y="40" width="180" height="340" rx="14" fill="#ffffff" stroke="#86efac" stroke-width="2.5" />
      <text x="460" y="75" font-family="system-ui" font-size="14" font-weight="800" fill="#0d5c3a">GİRİŞ KALİTE ONAY</text>
      <rect x="455" y="100" width="150" height="60" rx="8" fill="#dcfce7" /><text x="470" y="135" font-family="system-ui" font-size="12" font-weight="700" fill="#0d5c3a">Parti #99 - Fire: %0.8</text>
    </g>
    """

# 29. fazla-mesai-ve-isci-dava-riski-tespit-dosyasi
@reg("fazla-mesai-ve-isci-dava-riski-tespit-dosyasi")
def draw_029():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <!-- Haftalık 45 Saat Aşım Isı Haritası -->
      <text x="300" y="20" font-family="system-ui" font-size="16" font-weight="800" fill="#0f172a" text-anchor="middle">HAFTALIK 45 SAAT AŞIM RİSK SKORU</text>
      <rect x="40" y="50" width="520" height="70" rx="10" fill="#dcfce7" />
      <text x="60" y="90" font-family="system-ui" font-size="14" font-weight="700" fill="#0d5c3a">BEYAZ YAKA: 40 SAAT (RİSKSİZ)</text>
      <rect x="40" y="135" width="520" height="70" rx="10" fill="#fef3c7" />
      <text x="60" y="175" font-family="system-ui" font-size="14" font-weight="700" fill="#b45309">LOJİSTİK: 52 SAAT (HAFTALIK 7 SAAT FAZLA MESAİ)</text>
      <rect x="40" y="220" width="520" height="85" rx="10" fill="#fee2e2" stroke="#ef4444" stroke-width="2" />
      <text x="60" y="260" font-family="system-ui" font-size="15" font-weight="900" fill="#b91c1c">ÜRETİM VARDİYA 2: 64 SAAT (YÜKSEK DAVA RİSKİ)</text>
      <text x="60" y="285" font-family="system-ui" font-size="12" font-weight="600" fill="#dc2626">Yıllık 270 Saat Yasal Sınırı Aşıldı - Bordro Kanıtı Şart</text>
    </g>
    """

# 30. fesih-maliyeti-simulatoru
@reg("fesih-maliyeti-simulatoru")
def draw_030():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="30" y="260" width="120" height="160" rx="8" fill="#3b82f6" /><text x="90" y="240" font-family="system-ui" font-size="12" font-weight="700" fill="#1e40af" text-anchor="middle">İHBAR</text>
      <rect x="170" y="140" width="120" height="280" rx="8" fill="#107c41" /><text x="230" y="120" font-family="system-ui" font-size="12" font-weight="700" fill="#0d5c3a" text-anchor="middle">KIDEM</text>
      <rect x="310" y="220" width="120" height="200" rx="8" fill="#f59e0b" /><text x="370" y="200" font-family="system-ui" font-size="12" font-weight="700" fill="#b45309" text-anchor="middle">YILLIK İZİN</text>
      <rect x="450" y="80" width="130" height="340" rx="10" fill="#1e293b" /><text x="515" y="60" font-family="system-ui" font-size="14" font-weight="900" fill="#0f172a" text-anchor="middle">TOPLAM YÜK</text>
      <line x1="10" y1="420" x2="600" y2="420" stroke="#cbd5e1" stroke-width="2" />
    </g>
    """

# 31. filo-arac-maliyet-komutasi
@reg("filo-arac-maliyet-komutasi")
def draw_031():
    s1 = donut_segment(350, 260, 150, 200, 0, 150, "#107c41")
    s2 = donut_segment(350, 260, 150, 200, 150, 260, "#2563eb")
    s3 = donut_segment(350, 260, 150, 200, 260, 360, "#f59e0b")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    {s1} {s2} {s3}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">KM BAŞI TCO</text>
      <text x="0" y="20" font-family="system-ui" font-size="32" font-weight="900" fill="#0f172a" text-anchor="middle">₺4.82 / km</text>
      <text x="0" y="45" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">Filo Büyüklüğü: 48 Araç</text>
    </g>
    """
