# scripts/covers/part3.py
# 63'ten 93'e kadar olan ürünlerin özel çizimleri

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

PART3_DRAWERS = {}

def reg(slug):
    def deco(fn):
        PART3_DRAWERS[slug] = fn
        return fn
    return deco

# 63. kkeg-ve-finansman-gider-kisitlamasi-vergi-savunma-seti
@reg("kkeg-ve-finansman-gider-kisitlamasi-vergi-savunma-seti")
def draw_063():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="520" height="120" rx="14" fill="#fee2e2" stroke="#dc2626" stroke-width="2" />
      <text x="70" y="85" font-family="system-ui" font-size="16" font-weight="800" fill="#b91c1c">%10 FİNANSMAN GİDER KISITLAMASI</text>
      <text x="70" y="130" font-family="system-ui" font-size="32" font-weight="900" fill="#b91c1c">₺1.420.000 KKEG</text>
      <rect x="40" y="180" width="520" height="120" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="70" y="225" font-family="system-ui" font-size="16" font-weight="800" fill="#0d5c3a">VERGİ SAVUNMA RAPORU</text>
      <text x="70" y="270" font-family="system-ui" font-size="32" font-weight="900" fill="#0d5c3a">✓ İnceleme Kanıtı Hazır</text>
    </g>
    """

# 64. klinik-saglik-merkezi-fizibilite-ve-karlilik
@reg("klinik-saglik-merkezi-fizibilite-ve-karlilik")
def draw_064():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="40" width="240" height="160" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="60" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">HASTA BAŞI ORT. GELİR</text>
      <text x="60" y="130" font-family="system-ui" font-size="32" font-weight="900" fill="#2563eb">₺4.850</text>
      <rect x="320" y="40" width="240" height="160" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="340" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">HEKİM HAKEDİŞ PAYI</text>
      <text x="340" y="130" font-family="system-ui" font-size="32" font-weight="900" fill="#f59e0b">%35</text>
      <rect x="40" y="220" width="520" height="120" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="70" y="265" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">NET POLİKLİNİK KÂRI</text>
      <text x="70" y="310" font-family="system-ui" font-size="32" font-weight="900" fill="#0d5c3a">₺1.850.000 / AY</text>
    </g>
    """

# 65. kobi-finans-yonetim-paketi
@reg("kobi-finans-yonetim-paketi")
def draw_065():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(350, 260)">
      <polygon points="0,-180 170,-60 110,140 -110,140 -170,-60" fill="none" stroke="#cbd5e1" stroke-width="2" />
      <polygon points="0,-140 130,-40 80,100 -80,100 -130,-40" fill="#107c41" fill-opacity="0.3" stroke="#107c41" stroke-width="3" />
      <text x="0" y="-195" font-family="system-ui" font-size="13" font-weight="800" fill="#0f172a" text-anchor="middle">LİKİDİTE (1.8x)</text>
      <text x="185" y="-55" font-family="system-ui" font-size="13" font-weight="800" fill="#0f172a">KÂRLILIK (%24)</text>
      <text x="120" y="160" font-family="system-ui" font-size="13" font-weight="800" fill="#0f172a">BÜYÜME (+%38)</text>
      <text x="-120" y="160" font-family="system-ui" font-size="13" font-weight="800" fill="#0f172a" text-anchor="end">BORÇLANMA (0.4)</text>
      <text x="-185" y="-55" font-family="system-ui" font-size="13" font-weight="800" fill="#0f172a" text-anchor="end">DSO (38 Gün)</text>
    </g>
    """

# 66. konkordato-nakit-akis-on-projesi
@reg("konkordato-nakit-akis-on-projesi")
def draw_066():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <text x="300" y="30" font-family="system-ui" font-size="18" font-weight="900" fill="#0f172a" text-anchor="middle">ALACAKLILAR TENZİLAT VE VADE PROJEKSİYONU</text>
      <rect x="40" y="60" width="520" height="70" rx="8" fill="#dbeafe" />
      <text x="60" y="100" font-family="system-ui" font-size="14" font-weight="700" fill="#1e40af">1. Yıl: 12 Ay Ödemesiz Dönem (Nakit Toparlanma)</text>
      <rect x="40" y="145" width="520" height="70" rx="8" fill="#dcfce7" />
      <text x="60" y="185" font-family="system-ui" font-size="14" font-weight="700" fill="#065f46">2. Yıl: %30 Anapara Taksitleri (₺14.000.000)</text>
      <rect x="40" y="230" width="520" height="70" rx="8" fill="#dcfce7" />
      <text x="60" y="270" font-family="system-ui" font-size="14" font-weight="700" fill="#065f46">3. Yıl: %40 Kalan Borç Tasfiyesi (Mahkeme Onaylı)</text>
    </g>
    """

# 67. konkordato-ttk376-kriz-paketi
@reg("konkordato-ttk376-kriz-paketi")
def draw_067():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="40" width="520" height="120" rx="16" fill="#fee2e2" stroke="#dc2626" stroke-width="2.5" />
      <text x="70" y="85" font-family="system-ui" font-size="16" font-weight="900" fill="#b91c1c">TTK 376 SERMAYE KAYBI: 2/3 EŞİĞİ AŞILDI</text>
      <text x="70" y="125" font-family="system-ui" font-size="28" font-weight="900" fill="#b91c1c">Acil Tedbir Şart</text>
      <rect x="40" y="180" width="520" height="120" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="225" font-family="system-ui" font-size="16" font-weight="900" fill="#0d5c3a">SERMAYE TAMAMLAMA FONU</text>
      <text x="70" y="265" font-family="system-ui" font-size="28" font-weight="900" fill="#0d5c3a">+₺8.500.000 İle Güvenli Bölge</text>
    </g>
    """

# 68. kvkk-veri-envanteri-verbis-uyum
@reg("kvkk-veri-envanteri-verbis-uyum")
def draw_068():
    s1 = donut_segment(350, 260, 150, 200, 0, 360, "#107c41")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    {s1}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="15" font-weight="700" fill="#64748b" text-anchor="middle">VERBİS UYUM</text>
      <text x="0" y="25" font-family="system-ui" font-size="34" font-weight="900" fill="#0d5c3a" text-anchor="middle">%100</text>
      <text x="0" y="55" font-family="system-ui" font-size="12" font-weight="700" fill="#15803d" text-anchor="middle">38 Veri Kategorisi İndeksli</text>
    </g>
    """

# 69. logo-sql-cari-yaslandirma-tahsilat-karar-motoru
@reg("logo-sql-cari-yaslandirma-tahsilat-karar-motoru")
def draw_069():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="30" y="40" width="110" height="320" rx="8" fill="#107c41" /><text x="85" y="390" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">0-30 GÜN</text>
      <rect x="160" y="100" width="110" height="260" rx="8" fill="#2563eb" /><text x="215" y="390" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">31-60 GÜN</text>
      <rect x="290" y="160" width="110" height="200" rx="8" fill="#f59e0b" /><text x="345" y="390" font-family="system-ui" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">61-90 GÜN</text>
      <rect x="420" y="240" width="110" height="120" rx="8" fill="#dc2626" /><text x="475" y="390" font-family="system-ui" font-size="12" font-weight="700" fill="#dc2626" text-anchor="middle">90+ GÜN</text>
    </g>
    """

# 70. makine-bakim-kalibrasyon-durus-maliyeti
@reg("makine-bakim-kalibrasyon-durus-maliyeti")
def draw_070():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="40" width="240" height="160" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="60" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">MTBF (ARIZASIZ SÜRE)</text>
      <text x="60" y="130" font-family="system-ui" font-size="32" font-weight="900" fill="#107c41">480 Saat</text>
      <rect x="320" y="40" width="240" height="160" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="340" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">MTTR (TAMİR SÜRESİ)</text>
      <text x="340" y="130" font-family="system-ui" font-size="32" font-weight="900" fill="#2563eb">42 Dakika</text>
      <rect x="40" y="220" width="520" height="90" rx="12" fill="#fee2e2" />
      <text x="70" y="275" font-family="system-ui" font-size="18" font-weight="800" fill="#b91c1c">Duruş Saati Maliyet Kaybı: ₺18.500 / Saat</text>
    </g>
    """

# 71. mini-mrp-bom-malzeme-kapasite
@reg("mini-mrp-bom-malzeme-kapasite")
def draw_071():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(100, 50)">
      <!-- Ürün Ağacı (BOM) Hiyerarşisi -->
      <rect x="180" y="20" width="160" height="60" rx="8" fill="#1e293b" /><text x="260" y="55" font-family="system-ui" font-size="14" font-weight="800" fill="#ffffff" text-anchor="middle">MAMUL (100 ADET)</text>
      <line x1="260" y1="80" x2="260" y2="130" stroke="#94a3b8" stroke-width="3" />
      <line x1="120" y1="130" x2="400" y2="130" stroke="#94a3b8" stroke-width="3" />
      <line x1="120" y1="130" x2="120" y2="160" stroke="#94a3b8" stroke-width="3" />
      <line x1="400" y1="130" x2="400" y2="160" stroke="#94a3b8" stroke-width="3" />
      <rect x="40" y="160" width="160" height="60" rx="8" fill="#2563eb" /><text x="120" y="195" font-family="system-ui" font-size="13" font-weight="700" fill="#ffffff" text-anchor="middle">Yarı Mamul A</text>
      <rect x="320" y="160" width="160" height="60" rx="8" fill="#2563eb" /><text x="400" y="195" font-family="system-ui" font-size="13" font-weight="700" fill="#ffffff" text-anchor="middle">Yarı Mamul B</text>
      <rect x="140" y="270" width="240" height="80" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="260" y="318" font-family="system-ui" font-size="16" font-weight="900" fill="#0d5c3a" text-anchor="middle">Kapasite İhtiyacı: %88</text>
    </g>
    """

# 72. mizan-anomali-denetim-oncesi-kontrol
@reg("mizan-anomali-denetim-oncesi-kontrol")
def draw_072():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="520" height="90" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="70" y="95" font-family="system-ui" font-size="18" font-weight="900" fill="#0d5c3a">BORÇ / ALACAK BAKİYESİ TAM DENK (₺0.00 FARK)</text>
      <rect x="40" y="150" width="520" height="80" rx="10" fill="#fee2e2" />
      <text x="70" y="195" font-family="system-ui" font-size="14" font-weight="800" fill="#b91c1c">! Ters Bakiye Veren 3 Hesap Yakalandı (320, 120, 102)</text>
      <rect x="40" y="250" width="520" height="80" rx="10" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" />
      <text x="70" y="295" font-family="system-ui" font-size="15" font-weight="700" fill="#334155">Vergi Denetimi Öncesi Sıfır Ceza Güvencesi</text>
    </g>
    """

# 73. mutfak-kayip-kacak-hesaplayici
@reg("mutfak-kayip-kacak-hesaplayici")
def draw_073():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(100, 50)">
      <!-- Fire Hunisi -->
      <polygon points="20,0 420,0 370,80 70,80" fill="#1e293b" /><text x="220" y="48" font-family="system-ui" font-size="13" font-weight="800" fill="#ffffff" text-anchor="middle">SATIN ALMA HAMMADDE (100 kg)</text>
      <polygon points="75,90 365,90 315,170 125,170" fill="#f59e0b" /><text x="220" y="138" font-family="system-ui" font-size="13" font-weight="800" fill="#ffffff" text-anchor="middle">ÖN HAZIRLIK &amp; AYIKLAMA (88 kg)</text>
      <polygon points="130,180 310,180 270,260 170,260" fill="#2563eb" /><text x="220" y="228" font-family="system-ui" font-size="13" font-weight="800" fill="#ffffff" text-anchor="middle">PİŞİRME FİRESİ (72 kg)</text>
      <rect x="135" y="290" width="170" height="50" rx="10" fill="#107c41" /><text x="220" y="322" font-family="system-ui" font-size="15" font-weight="900" fill="#ffffff" text-anchor="middle">NET PORKSİYON: %72</text>
    </g>
    """

# 74. nakliye-maliyeti-hesaplayici
@reg("nakliye-maliyeti-hesaplayici")
def draw_074():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="40" width="520" height="120" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
      <text x="70" y="85" font-family="system-ui" font-size="15" font-weight="700" fill="#64748b">SEFER BAŞINA MAZOT + OTOYOL DÖKÜMÜ</text>
      <text x="70" y="130" font-family="system-ui" font-size="32" font-weight="900" fill="#0f172a">₺14.850 / Sefer</text>
      <rect x="40" y="180" width="520" height="120" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="225" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">DÖNÜŞ YÜKÜ DAHİL NET KÂR</text>
      <text x="70" y="270" font-family="system-ui" font-size="32" font-weight="900" fill="#0d5c3a">+₺9.400 Kâr Marjı</text>
    </g>
    """

# 75. oee-durus-kok-neden-komutasi
@reg("oee-durus-kok-neden-komutasi")
def draw_075():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="40" width="160" height="280" rx="12" fill="#2563eb" />
      <text x="120" y="140" font-family="system-ui" font-size="28" font-weight="900" fill="#ffffff" text-anchor="middle">%88</text>
      <text x="120" y="180" font-family="system-ui" font-size="12" font-weight="700" fill="#dbeafe" text-anchor="middle">Kullanılabilirlik</text>
      <rect x="220" y="40" width="160" height="280" rx="12" fill="#f59e0b" />
      <text x="300" y="140" font-family="system-ui" font-size="28" font-weight="900" fill="#ffffff" text-anchor="middle">%92</text>
      <text x="300" y="180" font-family="system-ui" font-size="12" font-weight="700" fill="#fef3c7" text-anchor="middle">Performans</text>
      <rect x="400" y="40" width="160" height="280" rx="12" fill="#107c41" />
      <text x="480" y="140" font-family="system-ui" font-size="28" font-weight="900" fill="#ffffff" text-anchor="middle">%98</text>
      <text x="480" y="180" font-family="system-ui" font-size="12" font-weight="700" fill="#dcfce7" text-anchor="middle">Kalite</text>
      <text x="300" y="370" font-family="system-ui" font-size="22" font-weight="900" fill="#0f172a" text-anchor="middle">GENEL OEE: %79.3 DÜNYA STANDARDI</text>
    </g>
    """

# 76. on-uc-haftalik-nakit-odeme-onceligi
@reg("on-uc-haftalik-nakit-odeme-onceligi")
def draw_076():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="520" height="90" rx="12" fill="#fee2e2" stroke="#dc2626" stroke-width="2" />
      <text x="70" y="85" font-family="system-ui" font-size="15" font-weight="900" fill="#b91c1c">A GRUBU: KRİTİK TEDARİKÇİLER (HEMEN ÖDE: ₺4.2M)</text>
      <rect x="40" y="145" width="520" height="90" rx="12" fill="#fef3c7" stroke="#f59e0b" stroke-width="1.5" />
      <text x="70" y="190" font-family="system-ui" font-size="15" font-weight="800" fill="#b45309">B GRUBU: 30 GÜN VADELİ (KISMEN ÖDE: ₺2.8M)</text>
      <rect x="40" y="250" width="520" height="90" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="70" y="295" font-family="system-ui" font-size="15" font-weight="800" fill="#0d5c3a">C GRUBU: ERTELENEBİLİR DİLİM (GÜVENLİ: ₺1.9M)</text>
    </g>
    """

# 77. ortaklar-cari-ve-kasa-adat-faiz-faturasi-hesaplayici
@reg("ortaklar-cari-ve-kasa-adat-faiz-faturasi-hesaplayici")
def draw_077():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="40" width="520" height="130" rx="16" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
      <text x="70" y="85" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">TCMB REESKONT FAİZ ORANI ADAT HESABI</text>
      <text x="70" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#0f172a">₺284.500 ADAT FAİZİ</text>
      <rect x="40" y="190" width="520" height="110" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="70" y="235" font-family="system-ui" font-size="14" font-weight="800" fill="#0d5c3a">DÜZENLENECEK KDV DAHİL FATURA</text>
      <text x="70" y="275" font-family="system-ui" font-size="28" font-weight="900" fill="#0d5c3a">₺341.400 (Sıfır Vergi Cezası)</text>
    </g>
    """

# 78. ortakli-gayrimenkul-yatirimi-kar-dagitim
@reg("ortakli-gayrimenkul-yatirimi-kar-dagitim")
def draw_078():
    s1 = donut_segment(350, 260, 140, 200, 0, 180, "#107c41")
    s2 = donut_segment(350, 260, 140, 200, 180, 290, "#2563eb")
    s3 = donut_segment(350, 260, 140, 200, 290, 360, "#f59e0b")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    {s1} {s2} {s3}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">KÂR DAĞITIMI</text>
      <text x="0" y="20" font-family="system-ui" font-size="28" font-weight="900" fill="#0f172a" text-anchor="middle">₺18.5M</text>
      <text x="0" y="45" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">Hisse Oranında Waterfall</text>
    </g>
    """

# 79. ortulu-sermaye-emsal-faiz-tf-paketi
@reg("ortulu-sermaye-emsal-faiz-tf-paketi")
def draw_079():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <line x1="40" y1="200" x2="540" y2="200" stroke="#dc2626" stroke-width="4" stroke-dasharray="6,6" />
      <text x="290" y="180" font-family="system-ui" font-size="14" font-weight="900" fill="#dc2626" text-anchor="middle">ÖZSERMAYE 3 KATI SINIRI: ₺30M</text>
      <rect x="80" y="100" width="140" height="260" rx="10" fill="#107c41" />
      <text x="150" y="220" font-family="system-ui" font-size="18" font-weight="900" fill="#ffffff" text-anchor="middle">₺24M Borç</text>
      <text x="150" y="250" font-family="system-ui" font-size="12" font-weight="700" fill="#dcfce7" text-anchor="middle">(Güvenli Sınırda)</text>
      <g transform="translate(260, 240)">
        <rect width="280" height="110" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
        <text x="20" y="40" font-family="system-ui" font-size="14" font-weight="800" fill="#0d5c3a">Örtülü Sermaye Sayılmaz</text>
        <text x="20" y="75" font-family="system-ui" font-size="22" font-weight="900" fill="#0d5c3a">Faiz Gideri Yazılabilir</text>
      </g>
    </g>
    """

# 80. otel-fizibilite-ve-karlilik
@reg("otel-fizibilite-ve-karlilik")
def draw_080():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="40" width="240" height="160" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="60" y="80" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">RevPAR (ODA GELİRİ)</text>
      <text x="60" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#2563eb">€142</text>
      <rect x="320" y="40" width="240" height="160" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="340" y="80" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">YILLIK DOLULUK</text>
      <text x="340" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#107c41">%78.4</text>
      <rect x="40" y="220" width="520" height="90" rx="12" fill="#dcfce7" />
      <text x="70" y="275" font-family="system-ui" font-size="20" font-weight="900" fill="#0d5c3a">GOP (Brüt İşletme Kârı): %44.2</text>
    </g>
    """

# 81. pazaryeri-hakedis-mutabakat-motoru
@reg("pazaryeri-hakedis-mutabakat-motoru")
def draw_081():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="30" y="40" width="120" height="320" rx="8" fill="#1e293b" /><text x="90" y="25" font-family="system-ui" font-size="12" font-weight="700" fill="#1e293b" text-anchor="middle">SATIŞ (₺1.2M)</text>
      <rect x="170" y="120" width="120" height="120" rx="8" fill="#dc2626" /><text x="230" y="105" font-family="system-ui" font-size="12" font-weight="700" fill="#dc2626" text-anchor="middle">-KOMİSYON</text>
      <rect x="310" y="180" width="120" height="60" rx="8" fill="#f59e0b" /><text x="370" y="165" font-family="system-ui" font-size="12" font-weight="700" fill="#d97706" text-anchor="middle">-KARGO</text>
      <rect x="450" y="160" width="120" height="200" rx="8" fill="#107c41" /><text x="510" y="145" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41" text-anchor="middle">BANKA TRANSFER</text>
      <text x="510" y="260" font-family="system-ui" font-size="20" font-weight="900" fill="#ffffff" text-anchor="middle">₺892.400</text>
    </g>
    """

# 82. pazaryeri-net-kar-ve-eksik-hakedis-yakalayici
@reg("pazaryeri-net-kar-ve-eksik-hakedis-yakalayici")
def draw_082():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="40" width="520" height="120" rx="16" fill="#fee2e2" stroke="#dc2626" stroke-width="2.5" />
      <text x="70" y="85" font-family="system-ui" font-size="16" font-weight="900" fill="#b91c1c">YAKALANAN HAKSIZ CEZA KESİNTİSİ</text>
      <text x="70" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#b91c1c">₺48.750 İADE TALEP EDİLDİ</text>
      <rect x="40" y="180" width="520" height="120" rx="16" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="70" y="225" font-family="system-ui" font-size="16" font-weight="900" fill="#0d5c3a">SİPARİŞ BAŞINA GERÇEK NET KÂR</text>
      <text x="70" y="270" font-family="system-ui" font-size="34" font-weight="900" fill="#0d5c3a">%22.4 Net Marj</text>
    </g>
    """

# 83. perakende-magaza-acilis-fizibilite
@reg("perakende-magaza-acilis-fizibilite")
def draw_083():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="40" width="240" height="160" rx="12" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" />
      <text x="60" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">M² BAŞINA AYLIK CİRO</text>
      <text x="60" y="130" font-family="system-ui" font-size="32" font-weight="900" fill="#2563eb">₺18.500 / m²</text>
      <rect x="320" y="40" width="240" height="160" rx="12" fill="#ffffff" stroke="#cbd5e1" stroke-width="2" />
      <text x="340" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">KİRA / CİRO ORANI</text>
      <text x="340" y="130" font-family="system-ui" font-size="32" font-weight="900" fill="#107c41">%9.2 (Güvenli)</text>
      <rect x="40" y="220" width="520" height="90" rx="12" fill="#dcfce7" />
      <text x="70" y="275" font-family="system-ui" font-size="20" font-weight="900" fill="#0d5c3a">Yatırım Amortisman Süresi: 14 Ay</text>
    </g>
    """

# 84. pos-komisyon-ve-net-tahsilat-kontrol-sistemi
@reg("pos-komisyon-ve-net-tahsilat-kontrol-sistemi")
def draw_084():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <rect x="40" y="50" width="240" height="240" rx="14" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="160" y="100" font-family="system-ui" font-size="16" font-weight="800" fill="#1e293b" text-anchor="middle">ERTESİ GÜN GEÇİŞ</text>
      <text x="160" y="160" font-family="system-ui" font-size="36" font-weight="900" fill="#dc2626" text-anchor="middle">%3.49</text>
      <text x="160" y="200" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b" text-anchor="middle">Komisyon Maliyeti</text>
      <rect x="320" y="50" width="240" height="240" rx="14" fill="#dcfce7" stroke="#107c41" stroke-width="2.5" />
      <text x="440" y="100" font-family="system-ui" font-size="16" font-weight="800" fill="#0d5c3a" text-anchor="middle">28 GÜN BLOKELİ</text>
      <text x="440" y="160" font-family="system-ui" font-size="36" font-weight="900" fill="#107c41" text-anchor="middle">%0.49</text>
      <text x="440" y="200" font-family="system-ui" font-size="14" font-weight="800" fill="#15803d" text-anchor="middle">Yıllık ₺480.000 Tasarruf</text>
    </g>
    """

# 85. profesyonel-hizmet-sirketi-karlilik-ve-kapasite
@reg("profesyonel-hizmet-sirketi-karlilik-ve-kapasite")
def draw_085():
    s1 = donut_segment(350, 260, 150, 200, 0, 290, "#107c41")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    {s1}
    <g transform="translate(350, 260)">
      <text x="0" y="-15" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">FATURALANABİLİR SAAT</text>
      <text x="0" y="20" font-family="system-ui" font-size="32" font-weight="900" fill="#0f172a" text-anchor="middle">%81 Verim</text>
      <text x="0" y="45" font-family="system-ui" font-size="12" font-weight="700" fill="#107c41" text-anchor="middle">Danışman Başı: ₺220.000 / Ay</text>
    </g>
    """

# 86. proje-finansmani-dscr-llcr-paketi
@reg("proje-finansmani-dscr-llcr-paketi")
def draw_086():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 50)">
      <line x1="40" y1="200" x2="540" y2="200" stroke="#107c41" stroke-width="4" stroke-dasharray="6,6" />
      <text x="290" y="180" font-family="system-ui" font-size="14" font-weight="900" fill="#107c41" text-anchor="middle">GÜVENLİ DSCR KORİDORU: 1.35x</text>
      <rect x="60" y="60" width="180" height="120" rx="12" fill="#dbeafe" />
      <text x="150" y="110" font-family="system-ui" font-size="14" font-weight="800" fill="#1e40af" text-anchor="middle">LLCR (ÖMÜR BOYU)</text>
      <text x="150" y="150" font-family="system-ui" font-size="26" font-weight="900" fill="#1e40af" text-anchor="middle">1.62x</text>
      <rect x="340" y="60" width="180" height="120" rx="12" fill="#dcfce7" />
      <text x="430" y="110" font-family="system-ui" font-size="14" font-weight="800" fill="#065f46" text-anchor="middle">PLCR (PROJE)</text>
      <text x="430" y="150" font-family="system-ui" font-size="26" font-weight="900" fill="#065f46" text-anchor="middle">1.84x</text>
    </g>
    """

# 87. proje-ve-is-bazinda-gercek-karlilik-sistemi
@reg("proje-ve-is-bazinda-gercek-karlilik-sistemi")
def draw_087():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="30" width="520" height="75" rx="10" fill="#ffffff" stroke="#107c41" stroke-width="2" />
      <text x="70" y="65" font-family="system-ui" font-size="15" font-weight="800" fill="#1e293b">Proje #A-102: Net Kâr %32 (+₺420.000)</text>
      <rect x="40" y="120" width="520" height="75" rx="10" fill="#ffffff" stroke="#f59e0b" stroke-width="2" />
      <text x="70" y="155" font-family="system-ui" font-size="15" font-weight="800" fill="#1e293b">Proje #B-205: Net Kâr %14 (+₺110.000)</text>
      <rect x="40" y="210" width="520" height="75" rx="10" fill="#ffffff" stroke="#dc2626" stroke-width="2" />
      <text x="70" y="245" font-family="system-ui" font-size="15" font-weight="800" fill="#b91c1c">Proje #C-401: ZARAR -%6 (-₺85.000 Adam/Saat Aşımı)</text>
    </g>
    """

# 88. puantaj-vardiya-fazla-mesai-kanit-sistemi
@reg("puantaj-vardiya-fazla-mesai-kanit-sistemi")
def draw_088():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(100, 50)">
      <circle cx="250" cy="200" r="160" fill="none" stroke="#cbd5e1" stroke-width="2" />
      <line x1="250" y1="200" x2="250" y2="40" stroke="#107c41" stroke-width="4" />
      <line x1="250" y1="200" x2="390" y2="270" stroke="#2563eb" stroke-width="4" />
      <line x1="250" y1="200" x2="110" y2="270" stroke="#f59e0b" stroke-width="4" />
      <text x="250" y="30" font-family="system-ui" font-size="13" font-weight="800" fill="#107c41" text-anchor="middle">Vardiya 1 (08-16)</text>
      <text x="410" y="280" font-family="system-ui" font-size="13" font-weight="800" fill="#2563eb">Vardiya 2 (16-24)</text>
      <text x="90" y="280" font-family="system-ui" font-size="13" font-weight="800" fill="#f59e0b" text-anchor="end">Vardiya 3 (24-08)</text>
      <circle cx="250" cy="200" r="12" fill="#0f172a" />
    </g>
    """

# 89. recete-maliyeti-menu-muhendisligi
@reg("recete-maliyeti-menu-muhendisligi")
def draw_089():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <text x="300" y="20" font-family="system-ui" font-size="16" font-weight="900" fill="#0f172a" text-anchor="middle">BOSTON (BCG) MENÜ MATRİSİ</text>
      <rect x="50" y="40" width="240" height="160" rx="10" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
      <text x="170" y="110" font-family="system-ui" font-size="18" font-weight="900" fill="#0d5c3a" text-anchor="middle">★ YILDIZLAR (%42)</text>
      <rect x="310" y="40" width="240" height="160" rx="10" fill="#dbeafe" stroke="#2563eb" stroke-width="2" />
      <text x="430" y="110" font-family="system-ui" font-size="18" font-weight="900" fill="#1e40af" text-anchor="middle">İNEKLER (%35)</text>
      <rect x="50" y="220" width="240" height="160" rx="10" fill="#fef3c7" stroke="#f59e0b" stroke-width="2" />
      <text x="170" y="290" font-family="system-ui" font-size="18" font-weight="900" fill="#b45309" text-anchor="middle">? SORU İŞARETLERİ</text>
      <rect x="310" y="220" width="240" height="160" rx="10" fill="#fee2e2" stroke="#dc2626" stroke-width="2" />
      <text x="430" y="290" font-family="system-ui" font-size="18" font-weight="900" fill="#b91c1c" text-anchor="middle">KÖPEKLER (Çıkar)</text>
    </g>
    """

# 90. restoran-kafe-yatirim-fizibilite-ve-karlilik
@reg("restoran-kafe-yatirim-fizibilite-ve-karlilik")
def draw_090():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 60)">
      <rect x="40" y="40" width="240" height="160" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="60" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">MASA DEVİR HIZI</text>
      <text x="60" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#2563eb">3.8 Tur / Gün</text>
      <rect x="320" y="40" width="240" height="160" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2" />
      <text x="340" y="80" font-family="system-ui" font-size="13" font-weight="700" fill="#64748b">ORT. ADİSYON TUTARI</text>
      <text x="340" y="130" font-family="system-ui" font-size="34" font-weight="900" fill="#107c41">₺485</text>
      <rect x="40" y="220" width="520" height="90" rx="12" fill="#dcfce7" />
      <text x="70" y="275" font-family="system-ui" font-size="20" font-weight="900" fill="#0d5c3a">Yatırım Geri Dönüşü: 16 Ay Payback</text>
    </g>
    """

# 91. restoran-recete-maliyet-fire
@reg("restoran-recete-maliyet-fire")
def draw_091():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    <g transform="translate(60, 40)">
      <rect x="40" y="40" width="520" height="90" rx="12" fill="#ffffff" stroke="#107c41" stroke-width="2" />
      <text x="70" y="75" font-family="system-ui" font-size="14" font-weight="700" fill="#64748b">HEDEF GIDA MALİYETİ (FOOD COST)</text>
      <text x="70" y="110" font-family="system-ui" font-size="26" font-weight="900" fill="#107c41">%28.5 (İdeal Aralık)</text>
      <rect x="40" y="150" width="520" height="160" rx="12" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="70" y="195" font-family="system-ui" font-size="15" font-weight="800" fill="#0f172a">Porsiyon Gramaj Hassasiyeti: ±%1 Fire Limiti</text>
      <text x="70" y="240" font-family="system-ui" font-size="15" font-weight="800" fill="#2563eb">Aylık Engellenen Kaçak/Kayıp: ₺120.000</text>
    </g>
    """

# 92. saas-finansal-planlama-runway-ve-yatirim
@reg("saas-finansal-planlama-runway-ve-yatirim")
def draw_092():
    return """
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" />
    <path d="M 60 420 L 220 400 L 360 360 L 500 240 L 640 80" fill="none" stroke="#2563eb" stroke-width="6" stroke-linecap="round" />
    <circle cx="640" cy="80" r="14" fill="#2563eb" />
    <g transform="translate(60, 50)">
      <rect width="220" height="70" rx="12" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" />
      <text x="20" y="28" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Aylık Tekrarlayan Gelir</text>
      <text x="20" y="55" font-family="system-ui" font-size="22" font-weight="900" fill="#2563eb">$85.000 MRR</text>
    </g>
    <g transform="translate(420, 320)">
      <rect width="220" height="70" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="1.5" />
      <text x="20" y="28" font-family="system-ui" font-size="12" font-weight="700" fill="#0d5c3a">Nakit Runway</text>
      <text x="20" y="55" font-family="system-ui" font-size="22" font-weight="900" fill="#0d5c3a">22 Ay Güvenli</text>
    </g>
    """

# 93. sarj-istasyonu-yatirim
@reg("sarj-istasyonu-yatirim")
def draw_093():
    s1 = donut_segment(350, 260, 150, 200, 200, 380, "#107c41")
    return f"""
    <rect x="20" y="20" width="700" height="520" rx="20" fill="#f8fafc" stroke="#e2e8f0" stroke-width="2" />
    {s1}
    <g transform="translate(350, 260)">
      <line x1="0" y1="0" x2="90" y2="-90" stroke="#0f172a" stroke-width="6" stroke-linecap="round" />
      <circle cx="0" cy="0" r="18" fill="#0f172a" />
      <text x="0" y="45" font-family="system-ui" font-size="32" font-weight="900" fill="#0f172a" text-anchor="middle">180 kW DC</text>
      <text x="0" y="75" font-family="system-ui" font-size="13" font-weight="700" fill="#107c41" text-anchor="middle">Günlük Doluluk: 14 Araç / Soket</text>
    </g>
    """
