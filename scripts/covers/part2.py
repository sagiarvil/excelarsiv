# scripts/covers/part2.py
# 32'den 62'ye kadar olan ürünlerin özel çizimleri

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

PART2_DRAWERS = {}

def reg(slug):
    def deco(fn):
        PART2_DRAWERS[slug] = fn
        return fn
    return deco

# 32. filo-yatirim-ve-arac-yenileme-fizibilite
@reg("filo-yatirim-ve-arac-yenileme-fizibilite")
def draw_032():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <path d="M 60 140 Q 300 180 640 420" fill="none" stroke="#dc2626" stroke-width="5" />
    <path d="M 60 420 Q 300 280 640 120" fill="none" stroke="#107c41" stroke-width="5" />
    <circle cx="340" cy="270" r="14" fill="#0f172a" />
    <text x="340" y="240" font-family="system-ui" font-size="14" font-weight="900" fill="#0f172a" text-anchor="middle">BAŞABAŞ: 4. YIL / 120.000 KM</text>
    <text x="80" y="120" font-family="system-ui" font-size="13" font-weight="700" fill="#dc2626">Eski Araç Bakım Yükü</text>
    <text x="600" y="100" font-family="system-ui" font-size="13" font-weight="700" fill="#107c41" text-anchor="end">Yeni Filo Verimliliği</text>
    """

# 33. franchise-yatirim-fizibilite
@reg("franchise-yatirim-fizibilite")
def draw_033():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="160" height="340" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="120" y="80" font-family="system-ui" font-size="14" font-weight="800" fill="#64748b" text-anchor="middle">ROYALTY</text>
      <text x="120" y="140" font-family="system-ui" font-size="28" font-weight="900" fill="#2563eb" text-anchor="middle">%5</text>
      <rect x="230" y="40" width="160" height="340" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="310" y="80" font-family="system-ui" font-size="14" font-weight="800" fill="#64748b" text-anchor="middle">REKLAM FONU</text>
      <text x="310" y="140" font-family="system-ui" font-size="28" font-weight="900" fill="#f59e0b" text-anchor="middle">%2</text>
      <rect x="420" y="40" width="160" height="340" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="500" y="80" font-family="system-ui" font-size="14" font-weight="800" fill="#0d5c3a" text-anchor="middle">ŞUBE NET KÂRI</text>
      <text x="500" y="140" font-family="system-ui" font-size="28" font-weight="900" fill="#0d5c3a" text-anchor="middle">%21.4</text>
      <text x="500" y="200" font-family="system-ui" font-size="13" font-weight="700" fill="#15803d" text-anchor="middle">Geri Dönüş:</text>
      <text x="500" y="230" font-family="system-ui" font-size="20" font-weight="900" fill="#0d5c3a" text-anchor="middle">18 Ay Payback</text>
    </g>
    """

# 34. gayrimenkul-tut-sat-karar
@reg("gayrimenkul-tut-sat-karar")
def draw_034():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="60" width="240" height="280" rx="16" fill="#ffffff" stroke="#2563eb" stroke-width="2.5" />
      <text x="160" y="110" font-family="system-ui" font-size="18" font-weight="900" fill="#1e40af" text-anchor="middle">SENARYO A: TUT</text>
      <text x="160" y="150" font-family="system-ui" font-size="13" font-weight="600" fill="#64748b" text-anchor="middle">10 Yıllık Kira NPV + Artış</text>
      <text x="160" y="220" font-family="system-ui" font-size="30" font-weight="900" fill="#0f172a" text-anchor="middle">₺24.8M</text>
      <rect x="320" y="60" width="240" height="280" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="3" />
      <text x="440" y="110" font-family="system-ui" font-size="18" font-weight="900" fill="#065f46" text-anchor="middle">SENARYO B: SAT</text>
      <text x="440" y="150" font-family="system-ui" font-size="13" font-weight="600" fill="#15803d" text-anchor="middle">Peşin Satış + Fon Getirisi</text>
      <text x="440" y="220" font-family="system-ui" font-size="30" font-weight="900" fill="#0d5c3a" text-anchor="middle">₺31.5M</text>
      <text x="440" y="280" font-family="system-ui" font-size="14" font-weight="800" fill="#107c41" text-anchor="middle">✓ %27 DAHA AVANTAJLI</text>
    </g>
    """

# 35. gayrimenkul-yatirim-fizibilite-ve-karlilik
@reg("gayrimenkul-yatirim-fizibilite-ve-karlilik")
def draw_035():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="520" height="100" rx="14" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <text x="70" y="80" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">NET PROJE DEĞERİ (NPV)</text>
      <text x="70" y="120" font-family="system-ui" font-size="28" font-weight="900" fill="#107c41">₺54.200.000</text>
      <rect x="40" y="160" width="240" height="180" rx="14" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <text x="70" y="200" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">İÇ VERİM ORANI (IRR)</text>
      <text x="70" y="250" font-family="system-ui" font-size="34" font-weight="900" fill="#2563eb">%38.4</text>
      <rect x="320" y="160" width="240" height="180" rx="14" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <text x="350" y="200" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">KİRA ÇARPANI (CAP RATE)</text>
      <text x="350" y="250" font-family="system-ui" font-size="34" font-weight="900" fill="#d97706">%8.2 Yıllık</text>
    </g>
    """

# 36. ges-uretim-performans
@reg("ges-uretim-performans")
def draw_036():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <path d="M 60 420 Q 200 400 350 100 T 640 420" fill="none" stroke="#f59e0b" stroke-width="6" />
    <circle cx="350" cy="100" r="14" fill="#f59e0b" />
    <text x="350" y="70" font-family="system-ui" font-size="16" font-weight="900" fill="#b45309" text-anchor="middle">ZİRVE SAAT: 13:00 (1.850 kW)</text>
    <g transform="translate(60, 320)">
      <rect width="240" height="70" rx="12" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="20" y="28" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Performans Oranı (PR)</text>
      <text x="20" y="55" font-family="system-ui" font-size="20" font-weight="900" fill="#107c41">%82.6 Optimum</text>
    </g>
    """

# 37. ges-yatirim-fizibilite-ve-proje-finansmani
@reg("ges-yatirim-fizibilite-ve-proje-finansmani")
def draw_037():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="520" height="120" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="85" font-family="system-ui" font-size="16" font-weight="800" fill="#065f46">25 YILLIK KÜMÜLATİF GES GELİRİ</text>
      <text x="70" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#0d5c3a">$4.850.000 USD</text>
      <rect x="40" y="180" width="240" height="160" rx="14" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="70" y="220" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">PROJE GERİ ÖDEME</text>
      <text x="70" y="270" font-family="system-ui" font-size="32" font-weight="900" fill="#2563eb">4.2 Yıl</text>
      <rect x="320" y="180" width="240" height="160" rx="14" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="350" y="220" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">LCOE BİRİM MALİYET</text>
      <text x="350" y="270" font-family="system-ui" font-size="32" font-weight="900" fill="#0f172a">$0.038/kWh</text>
    </g>
    """

# 38. ges-yatirim-fizibilite
@reg("ges-yatirim-fizibilite")
def draw_038():
    s1 = donut_segment(350, 260, 150, 200, 210, 390, "#f59e0b")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    {s1}
    <g transform="translate(350, 260)">
      <line x1="0" y1="0" x2="80" y2="-100" stroke="#0f172a" stroke-width="6" stroke-linecap="round" />
      <circle cx="0" cy="0" r="18" fill="#0f172a" />
      <text x="0" y="45" font-family="system-ui" font-size="32" font-weight="900" fill="#0f172a" text-anchor="middle">2.4 MWp</text>
      <text x="0" y="75" font-family="system-ui" font-size="13" font-weight="700" fill="#d97706" text-anchor="middle">Yıllık Üretim: 3.650.000 kWh</text>
    </g>
    """

# 39. gida-lot-maliyet-izlenebilirlik
@reg("gida-lot-maliyet-izlenebilirlik")
def draw_039():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="30" y="30" width="160" height="100" rx="10" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="110" y="70" font-family="system-ui" font-size="14" font-weight="800" fill="#475569" text-anchor="middle">LOT #G-2026</text>
      <text x="110" y="95" font-family="system-ui" font-size="11" font-weight="600" fill="#107c41" text-anchor="middle">Hammadde Giriş</text>
      <line x1="195" y1="80" x2="255" y2="80" stroke="#107c41" stroke-width="3" />
      <rect x="260" y="30" width="160" height="100" rx="10" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="340" y="70" font-family="system-ui" font-size="14" font-weight="800" fill="#475569" text-anchor="middle">REÇETE KAZANI</text>
      <text x="340" y="95" font-family="system-ui" font-size="11" font-weight="600" fill="#2563eb" text-anchor="middle">Pastörizasyon</text>
      <line x1="425" y1="80" x2="485" y2="80" stroke="#107c41" stroke-width="3" />
      <rect x="490" y="30" width="160" height="100" rx="10" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="570" y="70" font-family="system-ui" font-size="14" font-weight="800" fill="#0d5c3a" text-anchor="middle">SON ÜRÜN</text>
      <text x="570" y="95" font-family="system-ui" font-size="11" font-weight="700" fill="#15803d" text-anchor="middle">Birim: ₺18.42</text>
      <rect x="30" y="180" width="620" height="140" rx="14" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="60" y="225" font-family="system-ui" font-size="15" font-weight="700" fill="#334155">SKT / Raf Ömrü İklimleme Riski: Sıfır Kayıp</text>
      <text x="60" y="260" font-family="system-ui" font-size="15" font-weight="700" fill="#334155">Geri Çağırma (Recall) Süresi: 12 Saniye</text>
    </g>
    """

# 40. gunluk-gelir-gider-ve-gercek-karlilik-sistemi
@reg("gunluk-gelir-gider-ve-gercek-karlilik-sistemi")
def draw_040():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <path d="M 60 300 Q 180 80 320 300 T 640 180" fill="none" stroke="#107c41" stroke-width="5" />
    <path d="M 60 360 Q 180 200 320 360 T 640 320" fill="none" stroke="#dc2626" stroke-width="3" stroke-dasharray="6,6" />
    <g transform="translate(60, 50)">
      <rect width="200" height="65" rx="12" fill="#ffffff" stroke="#107c41" stroke-width="2" />
      <text x="20" y="26" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Günlük Net Kâr</text>
      <text x="20" y="52" font-family="system-ui" font-size="20" font-weight="900" fill="#107c41">+₺42.500</text>
    </g>
    """

# 41. hakedis-fiyat-farki-hak-kaybi-cetveli
@reg("hakedis-fiyat-farki-hak-kaybi-cetveli")
def draw_041():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="100" width="120" height="300" rx="10" fill="#94a3b8" />
      <text x="100" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#475569" text-anchor="middle">Sözleşme Fiyatı</text>
      <rect x="200" y="40" width="140" height="360" rx="10" fill="#107c41" />
      <text x="270" y="25" font-family="system-ui" font-size="14" font-weight="900" fill="#0d5c3a" text-anchor="middle">Pn Fiyat Farkı</text>
      <text x="270" y="200" font-family="system-ui" font-size="24" font-weight="900" fill="#ffffff" text-anchor="middle">+₺8.4M</text>
      <g transform="translate(370, 140)">
        <rect width="220" height="110" rx="14" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
        <text x="20" y="35" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">Hak Kaybı Engellendi</text>
        <text x="20" y="70" font-family="system-ui" font-size="22" font-weight="900" fill="#107c41">12 Gün İçinde</text>
      </g>
    </g>
    """

# 42. hizmet-ihracati-100-indirim-transfer-takip
@reg("hizmet-ihracati-100-indirim-transfer-takip")
def draw_042():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="40" width="520" height="120" rx="16" fill="#dbeafe" stroke="#2563eb" stroke-width="2" />
      <text x="70" y="85" font-family="system-ui" font-size="16" font-weight="800" fill="#1e40af">KVK 10/1-Ğ KAPSAMINDA İNDİRİM</text>
      <text x="70" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#1d4ed8">%80 MATRAH İNDİRİMİ</text>
      <g transform="translate(40, 190)">
        <rect width="520" height="120" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
        <text x="30" y="45" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">KURUM KAZANCI VERGİ AVANTAJI</text>
        <text x="30" y="90" font-family="system-ui" font-size="34" font-weight="900" fill="#0d5c3a">₺3.250.000 TASARRUF</text>
      </g>
    </g>
    """

# 43. hukuk-burosu-dosya-vekalet
@reg("hukuk-burosu-dosya-vekalet")
def draw_043():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="30" y="30" width="160" height="150" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="110" y="70" font-family="system-ui" font-size="14" font-weight="800" fill="#334155" text-anchor="middle">DERDEST</text>
      <text x="110" y="115" font-family="system-ui" font-size="32" font-weight="900" fill="#2563eb" text-anchor="middle">142</text>
      <rect x="230" y="30" width="160" height="150" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="310" y="70" font-family="system-ui" font-size="14" font-weight="800" fill="#334155" text-anchor="middle">İSTİNAF</text>
      <text x="310" y="115" font-family="system-ui" font-size="32" font-weight="900" fill="#f59e0b" text-anchor="middle">38</text>
      <rect x="430" y="30" width="160" height="150" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="510" y="70" font-family="system-ui" font-size="14" font-weight="800" fill="#0d5c3a" text-anchor="middle">TAHSİL EDİLEN</text>
      <text x="510" y="115" font-family="system-ui" font-size="32" font-weight="900" fill="#0d5c3a" text-anchor="middle">₺14.8M</text>
    </g>
    """

# 44. ihale-fiyat-farki-eskalasyon-pro
@reg("ihale-fiyat-farki-eskalasyon-pro")
def draw_044():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(80, 50)">
      <text x="270" y="30" font-family="system-ui" font-size="18" font-weight="900" fill="#0f172a" text-anchor="middle">ESKALASYON ENDEKS AĞIRLIKLARI (Pn)</text>
      <rect x="20" y="70" width="500" height="50" rx="8" fill="#107c41" />
      <text x="40" y="100" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff">İşçilik Endeksi (%35)</text>
      <rect x="20" y="135" width="400" height="50" rx="8" fill="#2563eb" />
      <text x="40" y="165" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff">Demir / Çelik Endeksi (%28)</text>
      <rect x="20" y="200" width="300" height="50" rx="8" fill="#f59e0b" />
      <text x="40" y="230" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff">Çimento Katkısı (%20)</text>
      <rect x="20" y="265" width="240" height="50" rx="8" fill="#8b5cf6" />
      <text x="40" y="295" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff">Akaryakıt (%17)</text>
    </g>
    """

# 45. ihaleye-kac-tl-teklif-vermeliyim
@reg("ihaleye-kac-tl-teklif-vermeliyim")
def draw_045():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <path d="M 60 420 Q 200 410 350 80 Q 500 410 640 420" fill="none" stroke="#2563eb" stroke-width="5" />
    <line x1="350" y1="80" x2="350" y2="420" stroke="#dc2626" stroke-width="3" stroke-dasharray="6,6" />
    <circle cx="350" cy="80" r="14" fill="#dc2626" />
    <g transform="translate(240, 110)">
      <rect width="220" height="60" rx="12" fill="#ffffff" stroke="#dc2626" stroke-width="2" />
      <text x="110" y="25" font-family="system-ui" font-size="12" font-weight="800" fill="#b91c1c" text-anchor="middle">OPTİMUM KAZANMA NOKTASI</text>
      <text x="110" y="47" font-family="system-ui" font-size="16" font-weight="900" fill="#0f172a" text-anchor="middle">₺24.850.000 (%18 Kâr)</text>
    </g>
    """

# 46. ihracat-siparis-karlilik-kur
@reg("ihracat-siparis-karlilik-kur")
def draw_046():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="60" width="160" height="280" rx="12" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" />
      <text x="120" y="100" font-family="system-ui" font-size="14" font-weight="800" fill="#475569" text-anchor="middle">FOB BEDELİ</text>
      <text x="120" y="150" font-family="system-ui" font-size="24" font-weight="900" fill="#0f172a" text-anchor="middle">$180.000</text>
      <rect x="230" y="60" width="160" height="280" rx="12" fill="#ffffff" stroke="#f59e0b" stroke-width="2" />
      <text x="310" y="100" font-family="system-ui" font-size="14" font-weight="800" fill="#d97706" text-anchor="middle">FORWARD KUR</text>
      <text x="310" y="150" font-family="system-ui" font-size="24" font-weight="900" fill="#d97706" text-anchor="middle">₺39.20</text>
      <rect x="420" y="60" width="160" height="280" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="500" y="100" font-family="system-ui" font-size="14" font-weight="800" fill="#0d5c3a" text-anchor="middle">NET KÂR</text>
      <text x="500" y="150" font-family="system-ui" font-size="24" font-weight="900" fill="#0d5c3a" text-anchor="middle">₺1.840.000</text>
    </g>
    """

# 47. insaat-hakedis-santiye-maliyet
@reg("insaat-hakedis-santiye-maliyet")
def draw_047():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <path d="M 60 420 Q 240 380 350 200 T 640 80" fill="none" stroke="#107c41" stroke-width="6" />
    <path d="M 60 420 Q 240 320 350 240 T 640 140" fill="none" stroke="#94a3b8" stroke-width="3" stroke-dasharray="6,6" />
    <g transform="translate(60, 50)">
      <rect width="240" height="70" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="1.5" />
      <text x="20" y="28" font-family="system-ui" font-size="12" font-weight="700" fill="#0d5c3a">S-Curve İlerleme</text>
      <text x="20" y="55" font-family="system-ui" font-size="20" font-weight="900" fill="#0d5c3a">%84 Gerçekleşen</text>
    </g>
    """

# 48. insaat-hakedis-yonetim-sistemi
@reg("insaat-hakedis-yonetim-sistemi")
def draw_048():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="30" y="50" width="120" height="320" rx="8" fill="#1e293b" /><text x="90" y="35" font-family="system-ui" font-size="12" font-weight="700" fill="#1e293b" text-anchor="middle">HAKEDİŞ 1</text>
      <rect x="170" y="10" width="120" height="360" rx="8" fill="#2563eb" /><text x="230" y="395" font-family="system-ui" font-size="12" font-weight="700" fill="#2563eb" text-anchor="middle">HAKEDİŞ 2</text>
      <rect x="310" y="80" width="120" height="290" rx="8" fill="#f59e0b" /><text x="370" y="395" font-family="system-ui" font-size="12" font-weight="700" fill="#d97706" text-anchor="middle">HAKEDİŞ 3</text>
      <rect x="450" y="120" width="130" height="250" rx="8" fill="#107c41" /><text x="515" y="395" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41" text-anchor="middle">KESİN HESAP</text>
    </g>
    """

# 49. isg-risk-degerlendirme-pro-6331
@reg("isg-risk-degerlendirme-pro-6331")
def draw_049():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <text x="300" y="20" font-family="system-ui" font-size="16" font-weight="900" fill="#0f172a" text-anchor="middle">5x5 L TİPİ MATRİS (6331 UYUMLU)</text>
      <rect x="60" y="50" width="220" height="160" rx="10" fill="#fee2e2" />
      <text x="170" y="120" font-family="system-ui" font-size="16" font-weight="900" fill="#b91c1c" text-anchor="middle">KABUL EDİLEMEZ</text>
      <rect x="300" y="50" width="220" height="160" rx="10" fill="#fef3c7" />
      <text x="410" y="120" font-family="system-ui" font-size="16" font-weight="900" fill="#b45309" text-anchor="middle">ÖNEMLİ RİSK</text>
      <rect x="60" y="230" width="460" height="160" rx="10" fill="#dcfce7" />
      <text x="290" y="315" font-family="system-ui" font-size="18" font-weight="900" fill="#0d5c3a" text-anchor="middle">ÖNLEM ALINMIŞ GÜVENLİ SAHA (%94)</text>
    </g>
    """

# 50. ithalat-depo-teslim-rafa-gelen-net-birim-maliyet
@reg("ithalat-depo-teslim-rafa-gelen-net-birim-maliyet")
def draw_050():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="30" y="200" width="100" height="220" rx="8" fill="#1e293b" /><text x="80" y="180" font-family="system-ui" font-size="12" font-weight="700" fill="#1e293b" text-anchor="middle">FOB BEDEL</text>
      <rect x="150" y="160" width="100" height="260" rx="8" fill="#2563eb" /><text x="200" y="140" font-family="system-ui" font-size="12" font-weight="700" fill="#2563eb" text-anchor="middle">NAVLUN</text>
      <rect x="270" y="110" width="100" height="310" rx="8" fill="#f59e0b" /><text x="320" y="90" font-family="system-ui" font-size="12" font-weight="700" fill="#d97706" text-anchor="middle">GÜMRÜK</text>
      <rect x="390" y="50" width="140" height="370" rx="12" fill="#107c41" /><text x="460" y="30" font-family="system-ui" font-size="14" font-weight="900" fill="#0d5c3a" text-anchor="middle">RAF BİRİM MALİYET</text>
      <text x="460" y="220" font-family="system-ui" font-size="26" font-weight="900" fill="#ffffff" text-anchor="middle">₺245.50</text>
    </g>
    """

# 51. ithalat-landed-cost-motoru
@reg("ithalat-landed-cost-motoru")
def draw_051():
    s1 = donut_segment(350, 260, 140, 200, 0, 140, "#1e293b")
    s2 = donut_segment(350, 260, 140, 200, 140, 220, "#2563eb")
    s3 = donut_segment(350, 260, 140, 200, 220, 310, "#f59e0b")
    s4 = donut_segment(350, 260, 140, 200, 310, 360, "#dc2626")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    {s1} {s2} {s3} {s4}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">LANDED COST</text>
      <text x="0" y="20" font-family="system-ui" font-size="28" font-weight="900" fill="#0f172a" text-anchor="middle">+%38.4 YÜK</text>
      <text x="0" y="45" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">Sıfır Sapma İthalat</text>
    </g>
    """

# 52. kacirilan-sgk-tesvikleri-ve-gercek-iscilik-maliyeti-analizi
@reg("kacirilan-sgk-tesvikleri-ve-gercek-iscilik-maliyeti-analizi")
def draw_052():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="30" y="40" width="540" height="100" rx="14" fill="#fee2e2" stroke="#ef4444" stroke-width="2" />
      <text x="60" y="75" font-family="system-ui" font-size="14" font-weight="800" fill="#b91c1c">KAÇIRILAN GEÇMİŞ SGK TEŞVİKLERİ</text>
      <text x="60" y="115" font-family="system-ui" font-size="28" font-weight="900" fill="#b91c1c">₺842.000 GERİ ALINABİLİR</text>
      <rect x="30" y="160" width="540" height="100" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="60" y="195" font-family="system-ui" font-size="14" font-weight="800" fill="#0d5c3a">AYLIK SÜREKLİ TEŞVİK TASARRUFU</text>
      <text x="60" y="235" font-family="system-ui" font-size="28" font-weight="900" fill="#0d5c3a">₺94.500 / AY KAZANÇ</text>
    </g>
    """

# 53. kargo-desi-maliyet-optimizasyonu
@reg("kargo-desi-maliyet-optimizasyonu")
def draw_053():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="60" width="220" height="260" rx="12" fill="#fee2e2" stroke="#f87171" stroke-width="2" />
      <text x="150" y="110" font-family="system-ui" font-size="16" font-weight="800" fill="#b91c1c" text-anchor="middle">HATALI DESİ (4 DESİ)</text>
      <text x="150" y="180" font-family="system-ui" font-size="32" font-weight="900" fill="#b91c1c" text-anchor="middle">₺68.50</text>
      <rect x="320" y="60" width="220" height="260" rx="12" fill="#dcfce7" stroke="#86efac" stroke-width="2.5" />
      <text x="430" y="110" font-family="system-ui" font-size="16" font-weight="800" fill="#0d5c3a" text-anchor="middle">OPTİMUM (2 DESİ)</text>
      <text x="430" y="180" font-family="system-ui" font-size="32" font-weight="900" fill="#0d5c3a" text-anchor="middle">₺42.00</text>
      <text x="430" y="250" font-family="system-ui" font-size="14" font-weight="800" fill="#15803d" text-anchor="middle">Tasarruf: %38.6</text>
    </g>
    """

# 54. kat-karsiligi-hasilat-paylasimi-simulator
@reg("kat-karsiligi-hasilat-paylasimi-simulator")
def draw_054():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="240" height="340" rx="16" fill="#ffffff" stroke="#2563eb" stroke-width="2.5" />
      <text x="160" y="90" font-family="system-ui" font-size="16" font-weight="800" fill="#1e40af" text-anchor="middle">ARSA SAHİBİ PAYI</text>
      <text x="160" y="160" font-family="system-ui" font-size="44" font-weight="900" fill="#2563eb" text-anchor="middle">%45</text>
      <text x="160" y="220" font-family="system-ui" font-size="18" font-weight="900" fill="#0f172a" text-anchor="middle">₺67.500.000</text>
      <rect x="320" y="40" width="240" height="340" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="440" y="90" font-family="system-ui" font-size="16" font-weight="800" fill="#0d5c3a" text-anchor="middle">MÜTEAHHİT PAYI</text>
      <text x="440" y="160" font-family="system-ui" font-size="44" font-weight="900" fill="#107c41" text-anchor="middle">%55</text>
      <text x="440" y="220" font-family="system-ui" font-size="18" font-weight="900" fill="#0d5c3a" text-anchor="middle">₺82.500.000</text>
    </g>
    """

# 55. kdv-iade-listesi-robotu-gib7
@reg("kdv-iade-listesi-robotu-gib7")
def draw_055():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="30" width="520" height="70" rx="10" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" />
      <text x="70" y="70" font-family="system-ui" font-size="15" font-weight="800" fill="#1e293b">1. İndirilecek KDV Listesi (GİB Format)</text>
      <rect x="40" y="115" width="520" height="70" rx="10" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" />
      <text x="70" y="155" font-family="system-ui" font-size="15" font-weight="800" fill="#1e293b">2. Yüklenilen KDV Dağıtım Tablosu</text>
      <rect x="40" y="200" width="520" height="70" rx="10" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" />
      <text x="70" y="240" font-family="system-ui" font-size="15" font-weight="800" fill="#1e293b">3. Satış Faturaları ve İhracat Beyannameleri</text>
      <rect x="40" y="285" width="520" height="90" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="70" y="340" font-family="system-ui" font-size="22" font-weight="900" fill="#0d5c3a">✓ 7 Liste Hatasız ve GİB Portal Uyumlu</text>
    </g>
    """

# 56. kdv-iadesi-azami-alacak-hesabi-dosya-hazirlayici
@reg("kdv-iadesi-azami-alacak-hesabi-dosya-hazirlayici")
def draw_056():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="40" width="520" height="140" rx="16" fill="#ffffff" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="90" font-family="system-ui" font-size="15" font-weight="800" fill="#64748b">AZAMİ ALINABİLİR İADE TAVANI</text>
      <text x="70" y="145" font-family="system-ui" font-size="38" font-weight="900" fill="#107c41">₺4.850.000</text>
      <rect x="40" y="200" width="520" height="120" rx="16" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="70" y="245" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">Dosya Onay Hazırlık Durumu</text>
      <text x="70" y="285" font-family="system-ui" font-size="20" font-weight="800" fill="#2563eb">YMM Raporu Öncesi Sıfır Risk</text>
    </g>
    """

# 57. kdv-iadesi-tutar-surec-simulasyonu
@reg("kdv-iadesi-tutar-surec-simulasyonu")
def draw_057():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <line x1="80" y1="200" x2="520" y2="200" stroke="#cbd5e1" stroke-width="4" />
      <circle cx="120" cy="200" r="16" fill="#107c41" /><text x="120" y="240" font-family="system-ui" font-size="12" font-weight="700" fill="#475569" text-anchor="middle">Talep</text>
      <circle cx="280" cy="200" r="16" fill="#2563eb" /><text x="280" y="240" font-family="system-ui" font-size="12" font-weight="700" fill="#475569" text-anchor="middle">YMM Raporu</text>
      <circle cx="480" cy="200" r="18" fill="#107c41" /><text x="480" y="240" font-family="system-ui" font-size="13" font-weight="900" fill="#0d5c3a" text-anchor="middle">Hesaba Nakit</text>
    </g>
    """

# 58. kdv-tevkifat-mahsup-iade-listesi
@reg("kdv-tevkifat-mahsup-iade-listesi")
def draw_058():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="40" width="160" height="300" rx="10" fill="#3b82f6" />
      <text x="120" y="100" font-family="system-ui" font-size="16" font-weight="800" fill="#ffffff" text-anchor="middle">5/10</text>
      <rect x="220" y="100" width="160" height="240" rx="10" fill="#107c41" />
      <text x="300" y="160" font-family="system-ui" font-size="16" font-weight="800" fill="#ffffff" text-anchor="middle">7/10</text>
      <rect x="400" y="160" width="160" height="180" rx="10" fill="#f59e0b" />
      <text x="480" y="220" font-family="system-ui" font-size="16" font-weight="800" fill="#ffffff" text-anchor="middle">9/10</text>
    </g>
    """

# 59. kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici
@reg("kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici")
def draw_059():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="520" height="110" rx="14" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <text x="70" y="80" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">TOPLAM FİRMA KIDEM YÜKÜ</text>
      <text x="70" y="125" font-family="system-ui" font-size="32" font-weight="900" fill="#0f172a">₺6.420.000</text>
      <rect x="40" y="170" width="240" height="180" rx="14" fill="#fee2e2" />
      <text x="70" y="210" font-family="system-ui" font-size="14" font-weight="800" fill="#b91c1c">İHBAR TAZMİNATI</text>
      <text x="70" y="260" font-family="system-ui" font-size="26" font-weight="900" fill="#b91c1c">₺1.180.000</text>
      <rect x="320" y="170" width="240" height="180" rx="14" fill="#dcfce7" />
      <text x="350" y="210" font-family="system-ui" font-size="14" font-weight="800" fill="#0d5c3a">AYLIK KARŞILIK</text>
      <text x="350" y="260" font-family="system-ui" font-size="26" font-weight="900" fill="#0d5c3a">₺142.000</text>
    </g>
    """

# 60. kik-asiri-dusuk-savunma-sinir-deger
@reg("kik-asiri-dusuk-savunma-sinir-deger")
def draw_060():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <line x1="20" y1="220" x2="560" y2="220" stroke="#dc2626" stroke-width="4" stroke-dasharray="6,6" />
      <text x="290" y="200" font-family="system-ui" font-size="14" font-weight="900" fill="#dc2626" text-anchor="middle">KİK SINIR DEĞER BARAJI: ₺12.4M</text>
      <rect x="60" y="120" width="80" height="260" rx="8" fill="#107c41" />
      <rect x="180" y="70" width="80" height="310" rx="8" fill="#107c41" />
      <rect x="300" y="240" width="80" height="140" rx="8" fill="#ef4444" />
      <rect x="420" y="90" width="80" height="290" rx="8" fill="#107c41" />
    </g>
    """

# 61. kira-avans-takip-dekont
@reg("kira-avans-takip-dekont")
def draw_061():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="30" y="40" width="540" height="320" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <text x="60" y="90" font-family="system-ui" font-size="16" font-weight="900" fill="#0f172a">MÜLK KİRA TAHSİLAT TAKVİMİ</text>
      <rect x="60" y="120" width="480" height="50" rx="8" fill="#dcfce7" />
      <text x="80" y="150" font-family="system-ui" font-size="14" font-weight="700" fill="#0d5c3a">Ofis A Blok: ₺85.000 (Zamanında Ödendi)</text>
      <rect x="60" y="185" width="480" height="50" rx="8" fill="#fee2e2" />
      <text x="80" y="215" font-family="system-ui" font-size="14" font-weight="700" fill="#b91c1c">Mağaza 4: ₺120.000 (8 Gün Gecikmede)</text>
      <rect x="60" y="250" width="480" height="50" rx="8" fill="#f1f5f9" />
      <text x="80" y="280" font-family="system-ui" font-size="14" font-weight="700" fill="#475569">Depo 2: ₺45.000 (Vadesi Gelmedi)</text>
    </g>
    """

# 62. kira-portfoyu-getiri-komutasi
@reg("kira-portfoyu-getiri-komutasi")
def draw_062():
    s1 = donut_segment(350, 260, 150, 200, 0, 240, "#107c41")
    s2 = donut_segment(350, 260, 150, 200, 240, 360, "#2563eb")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    {s1} {s2}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">PORTFÖY VERİMİ</text>
      <text x="0" y="20" font-family="system-ui" font-size="32" font-weight="900" fill="#0f172a" text-anchor="middle">%9.4 Brüt</text>
      <text x="0" y="45" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">Doluluk Oranı: %98</text>
    </g>
    """
