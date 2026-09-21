# scripts/covers/part4.py
# 94'ten 124'e kadar olan ürünlerin özel çizimleri

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

PART4_DRAWERS = {}

def reg(slug):
    def deco(fn):
        PART4_DRAWERS[slug] = fn
        return fn
    return deco

# 94. sekiz-d-duzeltici-faaliyet-dosya
@reg("sekiz-d-duzeltici-faaliyet-dosya")
def draw_094():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(100, 40)">
      <polygon points="250,40 380,90 430,220 380,350 250,400 120,350 70,220 120,90" fill="none" stroke="#2563eb" stroke-width="4" />
      <circle cx="250" cy="40" r="14" fill="#107c41" /><text x="250" y="25" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41" text-anchor="middle">D1: Ekip</text>
      <circle cx="380" cy="90" r="14" fill="#107c41" /><text x="400" y="85" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41">D2: Tanım</text>
      <circle cx="430" cy="220" r="14" fill="#107c41" /><text x="450" y="225" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41">D3: Önlem</text>
      <circle cx="380" cy="350" r="14" fill="#107c41" /><text x="400" y="365" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41">D4: Kök Neden</text>
      <circle cx="250" cy="400" r="14" fill="#107c41" /><text x="250" y="430" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41" text-anchor="middle">D5: Kalıcı Eylem</text>
      <text x="250" y="215" font-family="system-ui" font-size="28" font-weight="900" fill="#0f172a" text-anchor="middle">8D SİSTEMİ</text>
      <text x="250" y="245" font-family="system-ui" font-size="13" font-weight="700" fill="#107c41" text-anchor="middle">Otomotiv &amp; Havacılık</text>
    </g>
    """

# 95. sera-kurulum-fizibilite
@reg("sera-kurulum-fizibilite")
def draw_095():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="40" width="240" height="160" rx="12" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" />
      <text x="60" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">TOPRAKSIZ TARIM M² VERİMİ</text>
      <text x="60" y="130" font-family="system-ui" font-size="32" font-weight="900" fill="#107c41">55 kg / m²</text>
      <rect x="320" y="40" width="240" height="160" rx="12" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" />
      <text x="340" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">JEOTERMAL ISITMA AVANTAJI</text>
      <text x="340" y="130" font-family="system-ui" font-size="32" font-weight="900" fill="#f59e0b">-%42 Gider</text>
      <rect x="40" y="220" width="520" height="90" rx="12" fill="#dcfce7" />
      <text x="70" y="275" font-family="system-ui" font-size="20" font-weight="900" fill="#0d5c3a">Yatırım İtfa Süresi: 3.2 Yıl (TKDK Uyumlu)</text>
    </g>
    """

# 96. sevkiyat-fiyatlama-navlun
@reg("sevkiyat-fiyatlama-navlun")
def draw_096():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="60" width="220" height="260" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="150" y="110" font-family="system-ui" font-size="16" font-weight="800" fill="#475569" text-anchor="middle">PARSİYEL YÜK</text>
      <text x="150" y="180" font-family="system-ui" font-size="32" font-weight="900" fill="#0f172a" text-anchor="middle">₺18.500</text>
      <text x="150" y="230" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b" text-anchor="middle">Desi Başına Fiyat</text>
      <rect x="320" y="60" width="220" height="260" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="430" y="110" font-family="system-ui" font-size="16" font-weight="800" fill="#0d5c3a" text-anchor="middle">KOMPLE TIR</text>
      <text x="430" y="180" font-family="system-ui" font-size="32" font-weight="900" fill="#0d5c3a" text-anchor="middle">₺44.000</text>
      <text x="430" y="230" font-family="system-ui" font-size="13" font-weight="700" fill="#15803d" text-anchor="middle">%28 Birim Tasarruf</text>
    </g>
    """

# 97. sgk-prim-tesvikleri-optimizasyon-motoru
@reg("sgk-prim-tesvikleri-optimizasyon-motoru")
def draw_097():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="520" height="90" rx="12" fill="#ffffff" stroke="#107c41" stroke-width="2" />
      <text x="70" y="75" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">5510 SAYILI KANUN (%5 HAZİNE İNDİRİMİ)</text>
      <text x="70" y="110" font-family="system-ui" font-size="24" font-weight="900" fill="#107c41">₺142.500 / Ay Kazanç</text>
      <rect x="40" y="150" width="520" height="90" rx="12" fill="#ffffff" stroke="#2563eb" stroke-width="2" />
      <text x="70" y="185" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">6111 GENÇ İSTİHDAM TEŞVİKİ</text>
      <text x="70" y="220" font-family="system-ui" font-size="24" font-weight="900" fill="#2563eb">₺88.000 / Ay Kazanç</text>
      <rect x="40" y="260" width="520" height="70" rx="12" fill="#dcfce7" />
      <text x="70" y="303" font-family="system-ui" font-size="18" font-weight="900" fill="#0d5c3a">Yıllık Toplam İşveren Prim Avantajı: ₺2.766.000</text>
    </g>
    """

# 98. sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli
@reg("sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli")
def draw_098():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="160" height="280" rx="12" fill="#fee2e2" />
      <text x="120" y="140" font-family="system-ui" font-size="18" font-weight="900" fill="#b91c1c" text-anchor="middle">2/3 KAYIP</text>
      <text x="120" y="180" font-family="system-ui" font-size="13" font-weight="700" fill="#b91c1c" text-anchor="middle">Fesih Tehlikesi</text>
      <rect x="220" y="40" width="160" height="280" rx="12" fill="#fef3c7" />
      <text x="300" y="140" font-family="system-ui" font-size="18" font-weight="900" fill="#b45309" text-anchor="middle">1/2 KAYIP</text>
      <text x="300" y="180" font-family="system-ui" font-size="13" font-weight="700" fill="#b45309" text-anchor="middle">Genel Kurul Şart</text>
      <rect x="400" y="40" width="160" height="280" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="3" />
      <text x="480" y="140" font-family="system-ui" font-size="18" font-weight="900" fill="#0d5c3a" text-anchor="middle">SERMAYE OK</text>
      <text x="480" y="180" font-family="system-ui" font-size="13" font-weight="700" fill="#107c41" text-anchor="middle">Tam Koruma</text>
      <text x="300" y="370" font-family="system-ui" font-size="20" font-weight="900" fill="#0f172a" text-anchor="middle">TTK 376 SERMAYE TAMAMLAMA CETVELİ</text>
    </g>
    """

# 99. sirket-satin-alma-ve-ortaklik-devir-analizi
@reg("sirket-satin-alma-ve-ortaklik-devir-analizi")
def draw_099():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="40" width="520" height="120" rx="16" fill="#ffffff" stroke="#2563eb" stroke-width="2.5" />
      <text x="70" y="85" font-family="system-ui" font-size="15" font-weight="800" fill="#1e40af">İNDİRGENMİŞ NAKİT AKIŞI (DCF) DEĞERİ</text>
      <text x="70" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#0f172a">₺185.000.000</text>
      <rect x="40" y="180" width="520" height="120" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="225" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">ORTAKLIK DEVİR &amp; SATIN ALMA TEKLİFİ</text>
      <text x="70" y="270" font-family="system-ui" font-size="34" font-weight="900" fill="#0d5c3a">%40 HİSSE: ₺74.000.000</text>
    </g>
    """

# 100. site-apartman-yonetim-sistemi
@reg("site-apartman-yonetim-sistemi")
def draw_100():
    s1 = donut_segment(350, 260, 150, 200, 0, 310, "#107c41")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    {s1}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">AİDAT TAHSİLAT</text>
      <text x="0" y="20" font-family="system-ui" font-size="34" font-weight="900" fill="#0f172a" text-anchor="middle">%96.8</text>
      <text x="0" y="48" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">İşletme Projesi Denk</text>
    </g>
    """

# 101. spc-proses-yetenek-analizi
@reg("spc-proses-yetenek-analizi")
def draw_101():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <path d="M 60 400 Q 200 390 350 80 Q 500 390 640 400" fill="none" stroke="#2563eb" stroke-width="5" />
    <line x1="180" y1="60" x2="180" y2="420" stroke="#dc2626" stroke-width="3" stroke-dasharray="6,6" />
    <text x="180" y="45" font-family="system-ui" font-size="13" font-weight="800" fill="#dc2626" text-anchor="middle">LSL</text>
    <line x1="520" y1="60" x2="520" y2="420" stroke="#dc2626" stroke-width="3" stroke-dasharray="6,6" />
    <text x="520" y="45" font-family="system-ui" font-size="13" font-weight="800" fill="#dc2626" text-anchor="middle">USL</text>
    <circle cx="350" cy="80" r="12" fill="#107c41" />
    <text x="350" y="130" font-family="system-ui" font-size="24" font-weight="900" fill="#107c41" text-anchor="middle">Cpk = 1.67 (6 Sigma)</text>
    """

# 102. startup-yatirim-alma-ve-nakit-runway
@reg("startup-yatirim-alma-ve-nakit-runway")
def draw_102():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="40" width="240" height="160" rx="14" fill="#fee2e2" />
      <text x="60" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#b91c1c">NET AYLIK BURN RATE</text>
      <text x="60" y="130" font-family="system-ui" font-size="32" font-weight="900" fill="#b91c1c">-$42.000</text>
      <rect x="320" y="40" width="240" height="160" rx="14" fill="#dcfce7" />
      <text x="340" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#0d5c3a">KALAN RUNWAY</text>
      <text x="340" y="130" font-family="system-ui" font-size="32" font-weight="900" fill="#0d5c3a">18 Ay</text>
      <rect x="40" y="220" width="520" height="90" rx="14" fill="#f8fafc" stroke="#2563eb" stroke-width="2" />
      <text x="70" y="275" font-family="system-ui" font-size="20" font-weight="900" fill="#2563eb">Hedef Seri A Tur Değerleme: $6.5M</text>
    </g>
    """

# 103. stok-optimizasyon-abc-olu-stok-nakit
@reg("stok-optimizasyon-abc-olu-stok-nakit")
def draw_103():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <path d="M 60 420 Q 140 120 300 90 L 640 70" fill="none" stroke="#107c41" stroke-width="6" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="140" height="80" rx="10" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="110" y="75" font-family="system-ui" font-size="16" font-weight="900" fill="#0d5c3a" text-anchor="middle">A GRUBU</text>
      <text x="110" y="100" font-family="system-ui" font-size="12" font-weight="700" fill="#15803d" text-anchor="middle">%80 Değer / %20 SKU</text>
      <g transform="translate(360, 240)">
        <rect width="220" height="80" rx="10" fill="#fee2e2" stroke="#dc2626" stroke-width="2" />
        <text x="20" y="35" font-family="system-ui" font-size="13" font-weight="800" fill="#b91c1c">Kurtarılan Ölü Stok</text>
        <text x="20" y="65" font-family="system-ui" font-size="22" font-weight="900" fill="#b91c1c">₺3.400.000 Nakit</text>
      </g>
    </g>
    """

# 104. stok-satis-ve-nakit-baglanma-sistemi
@reg("stok-satis-ve-nakit-baglanma-sistemi")
def draw_104():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="40" width="520" height="120" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <text x="70" y="85" font-family="system-ui" font-size="15" font-weight="700" fill="#64748b">STOK DEVİR HIZI GÜN SAYISI</text>
      <text x="70" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#2563eb">44 Gün (Sektör: 68 Gün)</text>
      <rect x="40" y="180" width="520" height="120" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="225" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">SERBEST KALAN NAKİT</text>
      <text x="70" y="270" font-family="system-ui" font-size="34" font-weight="900" fill="#0d5c3a">+₺8.200.000 SERBEST</text>
    </g>
    """

# 105. sube-karlilik-ve-nakit-hesaplayici
@reg("sube-karlilik-ve-nakit-hesaplayici")
def draw_105():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="520" height="70" rx="10" fill="#dcfce7" />
      <text x="70" y="80" font-family="system-ui" font-size="16" font-weight="900" fill="#0d5c3a">Şube Kadıköy: +₺420.000 / Ay Net Kâr (KORU)</text>
      <rect x="40" y="125" width="520" height="70" rx="10" fill="#dcfce7" />
      <text x="70" y="165" font-family="system-ui" font-size="16" font-weight="900" fill="#0d5c3a">Şube Alsancak: +₺280.000 / Ay Net Kâr (KORU)</text>
      <rect x="40" y="210" width="520" height="90" rx="10" fill="#fee2e2" stroke="#dc2626" stroke-width="2" />
      <text x="70" y="250" font-family="system-ui" font-size="16" font-weight="900" fill="#b91c1c">Şube AVM: -₺140.000 / Ay Zarar (DERHAL KAPAT)</text>
      <text x="70" y="280" font-family="system-ui" font-size="12" font-weight="700" fill="#dc2626">Kapatma Kararı İle Yıllık ₺1.68M Zarar Engellenir</text>
    </g>
    """

# 106. sut-surusu-yonetim
@reg("sut-surusu-yonetim")
def draw_106():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <path d="M 60 380 Q 240 100 360 120 T 640 400" fill="none" stroke="#2563eb" stroke-width="5" />
    <g transform="translate(60, 50)">
      <rect width="220" height="70" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" />
      <text x="20" y="28" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Günlük Süt Verimi</text>
      <text x="20" y="55" font-family="system-ui" font-size="22" font-weight="900" fill="#2563eb">32.4 Litre / Baş</text>
    </g>
    <g transform="translate(420, 50)">
      <rect width="220" height="70" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="1.5" />
      <text x="20" y="28" font-family="system-ui" font-size="12" font-weight="700" fill="#0d5c3a">Yem Dönüşüm (FCR)</text>
      <text x="20" y="55" font-family="system-ui" font-size="22" font-weight="900" fill="#0d5c3a">1.38 İdeal Oran</text>
    </g>
    """

# 107. tahsilat-riski-vade-komuta-paneli
@reg("tahsilat-riski-vade-komuta-paneli")
def draw_107():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="40" width="240" height="150" rx="14" fill="#ffffff" stroke="#107c41" stroke-width="2" />
      <text x="60" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">ORTALAMA TAHSİLAT</text>
      <text x="60" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#107c41">36 Gün</text>
      <rect x="320" y="40" width="240" height="150" rx="14" fill="#ffffff" stroke="#dc2626" stroke-width="2" />
      <text x="340" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">RİSKLİ ALACAK</text>
      <text x="340" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#dc2626">₺480.000</text>
      <rect x="40" y="210" width="520" height="100" rx="14" fill="#dcfce7" />
      <text x="70" y="265" font-family="system-ui" font-size="22" font-weight="900" fill="#0d5c3a">Erken Uyarı Devrede: 0 İcra Takibi</text>
    </g>
    """

# 108. tarimsal-destek-uygunluk
@reg("tarimsal-destek-uygunluk")
def draw_108():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="520" height="85" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="70" y="85" font-family="system-ui" font-size="18" font-weight="900" fill="#0d5c3a">✓ MAZOT VE GÜBRE DESTEĞİ UYGUN (₺420.000)</text>
      <rect x="40" y="140" width="520" height="85" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="70" y="185" font-family="system-ui" font-size="18" font-weight="900" fill="#0d5c3a">✓ SERTİFİKALI TOHUM DESTEĞİ UYGUN (₺180.000)</text>
      <rect x="40" y="240" width="520" height="85" rx="12" fill="#dbeafe" stroke="#2563eb" stroke-width="2" />
      <text x="70" y="285" font-family="system-ui" font-size="18" font-weight="900" fill="#1e40af">✓ ÇKS PUAN KONTROLÜ: 94 / 100</text>
    </g>
    """

# 109. taseron-hakedis-kesinti-mutabakati
@reg("taseron-hakedis-kesinti-mutabakati")
def draw_109():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="30" y="40" width="120" height="320" rx="8" fill="#1e293b" /><text x="90" y="25" font-family="system-ui" font-size="12" font-weight="700" fill="#1e293b" text-anchor="middle">BRÜT İMALAT</text>
      <rect x="170" y="120" width="120" height="80" rx="8" fill="#dc2626" /><text x="230" y="105" font-family="system-ui" font-size="12" font-weight="700" fill="#dc2626" text-anchor="middle">-AVANS</text>
      <rect x="310" y="160" width="120" height="60" rx="8" fill="#f59e0b" /><text x="370" y="145" font-family="system-ui" font-size="12" font-weight="700" fill="#d97706" text-anchor="middle">-TEMİNAT</text>
      <rect x="450" y="140" width="120" height="220" rx="8" fill="#107c41" /><text x="510" y="125" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41" text-anchor="middle">NET TAŞERON</text>
      <text x="510" y="260" font-family="system-ui" font-size="22" font-weight="900" fill="#ffffff" text-anchor="middle">₺1.420.000</text>
    </g>
    """

# 110. tesvikli-bordro-avantajli-tesvik
@reg("tesvikli-bordro-avantajli-tesvik")
def draw_110():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="40" width="520" height="120" rx="16" fill="#dbeafe" stroke="#2563eb" stroke-width="2" />
      <text x="70" y="85" font-family="system-ui" font-size="15" font-weight="800" fill="#1e40af">5746 AR-GE &amp; TEKNOKENT İSTİSNASI</text>
      <text x="70" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#1d4ed8">₺184.000 / AY KAZANÇ</text>
      <rect x="40" y="180" width="520" height="120" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="225" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">EN AVANTAJLI KANUN SEÇİLDİ</text>
      <text x="70" y="270" font-family="system-ui" font-size="34" font-weight="900" fill="#0d5c3a">%38 İşveren Prim Tasarrufu</text>
    </g>
    """

# 111. tesvikli-bordro-optimizasyon
@reg("tesvikli-bordro-optimizasyon")
def draw_111():
    s1 = donut_segment(350, 260, 150, 200, 0, 280, "#107c41")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    {s1}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">OPTİMİZASYON ORANI</text>
      <text x="0" y="20" font-family="system-ui" font-size="32" font-weight="900" fill="#0f172a" text-anchor="middle">%100</text>
      <text x="0" y="45" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">Sıfır Çakışma, Maksimum Fayda</text>
    </g>
    """

# 112. trendyol-komisyon-sonrasi-net-kar
@reg("trendyol-komisyon-sonrasi-net-kar")
def draw_112():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="520" height="90" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" />
      <text x="70" y="75" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">ÜRÜN SATIŞ FİYATI (KDV DAHİL)</text>
      <text x="70" y="110" font-family="system-ui" font-size="26" font-weight="900" fill="#0f172a">₺599.00</text>
      <rect x="40" y="150" width="520" height="90" rx="12" fill="#fee2e2" />
      <text x="70" y="185" font-family="system-ui" font-size="14" font-weight="800" fill="#b91c1c">KOMİSYON + KARGO BAREMİ + STOPAJ</text>
      <text x="70" y="220" font-family="system-ui" font-size="26" font-weight="900" fill="#b91c1c">-₺184.50</text>
      <rect x="40" y="260" width="520" height="90" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="70" y="295" font-family="system-ui" font-size="14" font-weight="800" fill="#0d5c3a">BANKA HESABINA NET GEÇEN TUTAR</text>
      <text x="70" y="330" font-family="system-ui" font-size="28" font-weight="900" fill="#0d5c3a">₺414.50 (%69.2 Net)</text>
    </g>
    """

# 113. uretim-kapasite-yatirim-ve-karlilik
@reg("uretim-kapasite-yatirim-ve-karlilik")
def draw_113():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="40" width="520" height="100" rx="14" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
      <text x="70" y="80" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">MEVCUT DARBOĞAZ KAPASİTE</text>
      <text x="70" y="120" font-family="system-ui" font-size="30" font-weight="900" fill="#dc2626">12.000 Adet / Ay (%98)</text>
      <rect x="40" y="160" width="520" height="140" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="205" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">YENİ HAT YATIRIMI SONRASI</text>
      <text x="70" y="250" font-family="system-ui" font-size="34" font-weight="900" fill="#0d5c3a">35.000 Adet / Ay Kapasite</text>
      <text x="70" y="280" font-family="system-ui" font-size="13" font-weight="700" fill="#15803d">Yatırım Payback: 11 Ay</text>
    </g>
    """

# 114. uretim-kari-125-kv-optimizasyonu
@reg("uretim-kari-125-kv-optimizasyonu")
def draw_114():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="40" width="520" height="120" rx="16" fill="#dbeafe" stroke="#2563eb" stroke-width="2" />
      <text x="70" y="85" font-family="system-ui" font-size="15" font-weight="800" fill="#1e40af">İMALATÇI VE İHRACATÇI 1 PUAN İNDİRİM</text>
      <text x="70" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#1e40af">%24 İNDİRİMLİ KV ORANI</text>
      <rect x="40" y="180" width="520" height="120" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="225" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">NET KURUMLAR VERGİSİ TASARRUFU</text>
      <text x="70" y="270" font-family="system-ui" font-size="34" font-weight="900" fill="#0d5c3a">₺840.000 NAKİT AVANTAJ</text>
    </g>
    """

# 115. uretim-recetesi-ve-zam-yansitma-hesaplayici
@reg("uretim-recetesi-ve-zam-yansitma-hesaplayici")
def draw_115():
    s1 = donut_segment(350, 260, 140, 200, 0, 160, "#107c41")
    s2 = donut_segment(350, 260, 140, 200, 160, 270, "#2563eb")
    s3 = donut_segment(350, 260, 140, 200, 270, 360, "#f59e0b")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    {s1} {s2} {s3}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">ÖNERİLEN ZAM</text>
      <text x="0" y="20" font-family="system-ui" font-size="32" font-weight="900" fill="#0f172a" text-anchor="middle">+%18.4</text>
      <text x="0" y="45" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">Kâr Marjı %30 Korundu</text>
    </g>
    """

# 116. vergi-sgk-borcunu-tecil-etmeli-miyim-kredi-mi-tecil-mi
@reg("vergi-sgk-borcunu-tecil-etmeli-miyim-kredi-mi-tecil-mi")
def draw_116():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="50" width="240" height="240" rx="14" fill="#fee2e2" stroke="#dc2626" stroke-width="2" />
      <text x="160" y="100" font-family="system-ui" font-size="16" font-weight="800" fill="#b91c1c" text-anchor="middle">TİCARİ KREDİ</text>
      <text x="160" y="160" font-family="system-ui" font-size="34" font-weight="900" fill="#b91c1c" text-anchor="middle">%54 Faiz</text>
      <text x="160" y="200" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b" text-anchor="middle">Ağır Finansman Yükü</text>
      <rect x="320" y="50" width="240" height="240" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="440" y="100" font-family="system-ui" font-size="16" font-weight="800" fill="#0d5c3a" text-anchor="middle">6183 TECİL FAİZİ</text>
      <text x="440" y="160" font-family="system-ui" font-size="34" font-weight="900" fill="#107c41" text-anchor="middle">%48 Faiz</text>
      <text x="440" y="200" font-family="system-ui" font-size="14" font-weight="800" fill="#15803d" text-anchor="middle">✓ TECİL %12 DAHA UCUZ</text>
    </g>
    """

# 117. vergi-sgk-ve-maas-karsilik-ayirma-sistemi
@reg("vergi-sgk-ve-maas-karsilik-ayirma-sistemi")
def draw_117():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="520" height="90" rx="12" fill="#dbeafe" />
      <text x="70" y="80" font-family="system-ui" font-size="15" font-weight="800" fill="#1e40af">NET PERSONEL MAAŞ HAVUZU (AYIN 1'İ İÇİN)</text>
      <text x="70" y="110" font-family="system-ui" font-size="24" font-weight="900" fill="#1e40af">₺1.850.000 KİLİTLİ</text>
      <rect x="40" y="145" width="520" height="90" rx="12" fill="#dcfce7" />
      <text x="70" y="185" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">SGK &amp; MUHTASAR KARŞILIĞI (AYIN 26'SI İÇİN)</text>
      <text x="70" y="215" font-family="system-ui" font-size="24" font-weight="900" fill="#0d5c3a">₺940.000 KİLİTLİ</text>
      <rect x="40" y="250" width="520" height="70" rx="12" fill="#f1f5f9" />
      <text x="70" y="293" font-family="system-ui" font-size="16" font-weight="800" fill="#334155">Sıfır Gecikme Zammı Garantisi</text>
    </g>
    """

# 118. wacc-hesaplama-ve-sermaye-maliyeti
@reg("wacc-hesaplama-ve-sermaye-maliyeti")
def draw_118():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="40" width="240" height="160" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="60" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">ÖZKAYNAK MALİYETİ (Ke)</text>
      <text x="60" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#2563eb">%28.4</text>
      <rect x="320" y="40" width="240" height="160" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="340" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">BORÇLANMA MALİYETİ (Kd)</text>
      <text x="340" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#f59e0b">%18.2 Net</text>
      <rect x="40" y="220" width="520" height="110" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="265" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">AĞIRLIKLI ORTALAMA SERMAYE MALİYETİ (WACC)</text>
      <text x="70" y="305" font-family="system-ui" font-size="32" font-weight="900" fill="#0d5c3a">%22.8 İskonto Oranı</text>
    </g>
    """

# 119. yem-rasyonu-maliyet
@reg("yem-rasyonu-maliyet")
def draw_119():
    s1 = donut_segment(350, 260, 140, 200, 0, 150, "#107c41")
    s2 = donut_segment(350, 260, 140, 200, 150, 270, "#f59e0b")
    s3 = donut_segment(350, 260, 140, 200, 270, 360, "#2563eb")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    {s1} {s2} {s3}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">RASYON MALİYETİ</text>
      <text x="0" y="20" font-family="system-ui" font-size="30" font-weight="900" fill="#0f172a" text-anchor="middle">₺8.40 / kg</text>
      <text x="0" y="45" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">Linear Programming Optimizasyonu</text>
    </g>
    """

# 120. yeniden-degerleme-komuta-merkezi
@reg("yeniden-degerleme-komuta-merkezi")
def draw_120():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="40" width="520" height="120" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <text x="70" y="85" font-family="system-ui" font-size="15" font-weight="700" fill="#64748b">VUK 298/Ç VERGİSİZ ENDEKSLEME AVANTAJI</text>
      <text x="70" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#2563eb">₺28.400.000 MATRAH</text>
      <rect x="40" y="180" width="520" height="120" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="225" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">SAĞLANAN YILLIK AMORTİSMAN KALKANI</text>
      <text x="70" y="270" font-family="system-ui" font-size="34" font-weight="900" fill="#0d5c3a">₺7.100.000 NAKİT KÂR</text>
    </g>
    """

# 121. yeniden-degerleme-yapmali-miyim-vergi-tasarruf-analizi
@reg("yeniden-degerleme-yapmali-miyim-vergi-tasarruf-analizi")
def draw_121():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="50" width="240" height="240" rx="14" fill="#fee2e2" stroke="#dc2626" stroke-width="2" />
      <text x="160" y="100" font-family="system-ui" font-size="16" font-weight="800" fill="#b91c1c" text-anchor="middle">%2 VERGİ MALİYETİ</text>
      <text x="160" y="160" font-family="system-ui" font-size="34" font-weight="900" fill="#b91c1c" text-anchor="middle">₺600.000</text>
      <text x="160" y="200" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b" text-anchor="middle">Peşin Ödenen Vergi</text>
      <rect x="320" y="50" width="240" height="240" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="440" y="100" font-family="system-ui" font-size="16" font-weight="800" fill="#0d5c3a" text-anchor="middle">AMORTİSMAN KÂRI</text>
      <text x="440" y="160" font-family="system-ui" font-size="34" font-weight="900" fill="#107c41" text-anchor="middle">₺4.800.000</text>
      <text x="440" y="200" font-family="system-ui" font-size="14" font-weight="800" fill="#15803d" text-anchor="middle">✓ 8 KAT NET AVANTAJ</text>
    </g>
    """

# 122. yillara-sari-insaat-stopaj-nakit-akis-planlayici
@reg("yillara-sari-insaat-stopaj-nakit-akis-planlayici")
def draw_122():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="40" width="520" height="120" rx="16" fill="#fee2e2" stroke="#dc2626" stroke-width="2" />
      <text x="70" y="85" font-family="system-ui" font-size="15" font-weight="800" fill="#b91c1c">%5 HAKEDİŞ STOPAJ KESİNTİSİ</text>
      <text x="70" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#b91c1c">₺2.450.000 BLOKE</text>
      <rect x="40" y="180" width="520" height="120" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="225" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">GEÇİCİ VERGİ MAHSUP PLANI</text>
      <text x="70" y="270" font-family="system-ui" font-size="34" font-weight="900" fill="#0d5c3a">✓ Nakit Akışına İade</text>
    </g>
    """

# 123. ymm-tasdik-kontrol-robotu
@reg("ymm-tasdik-kontrol-robotu")
def draw_123():
    s1 = donut_segment(350, 260, 150, 200, 0, 360, "#107c41")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    {s1}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="15" font-weight="700" fill="#64748b" text-anchor="middle">YMM TASDİK SKORU</text>
      <text x="0" y="25" font-family="system-ui" font-size="36" font-weight="900" fill="#0d5c3a" text-anchor="middle">100 / 100</text>
      <text x="0" y="55" font-family="system-ui" font-size="13" font-weight="700" fill="#15803d" text-anchor="middle">Tüm Yasal Eşikler Doğrulandı</text>
    </g>
    """

# 124. yurt-disi-yapilanma-vergi-simulatoru
@reg("yurt-disi-yapilanma-vergi-simulatoru")
def draw_124():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <!-- ÇVÖA Kâr Transfer Şeması -->
      <rect x="40" y="40" width="220" height="150" rx="14" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="150" y="80" font-family="system-ui" font-size="14" font-weight="800" fill="#475569" text-anchor="middle">YURT DIŞI MERKEZ (LLC)</text>
      <text x="150" y="125" font-family="system-ui" font-size="28" font-weight="900" fill="#2563eb" text-anchor="middle">%9 Kurumlar V.</text>
      <line x1="265" y1="115" x2="335" y2="115" stroke="#107c41" stroke-width="4" stroke-linecap="round" />
      <polygon points="335,115 320,105 320,125" fill="#107c41" />
      <rect x="340" y="40" width="220" height="150" rx="14" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="450" y="80" font-family="system-ui" font-size="14" font-weight="800" fill="#475569" text-anchor="middle">TÜRKİYE İŞTİRAK</text>
      <text x="450" y="125" font-family="system-ui" font-size="28" font-weight="900" fill="#107c41" text-anchor="middle">İstisna Matrah</text>
      <rect x="40" y="220" width="520" height="100" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="275" font-family="system-ui" font-size="22" font-weight="900" fill="#0d5c3a">ÇVÖA Kapsamında Sıfır Çifte Vergilendirme</text>
    </g>
    """
