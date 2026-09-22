import os
import re
import html

def generate_svg_screen2(meta, cfg):
    name = meta['name']
    slug = meta['slug']
    inputs = meta['inputs']
    if not inputs:
        inputs = ['İşlem Tarihi ve Açıklama', 'Hesap / Cari Kodu', 'Miktar ve Birim Fiyat', 'Tutar ve KDV Oranı']

    # 10 satırlık gerçekçi veri tablosu
    row_data = [
        ('12.10.2026', 'TR-0428-A', inputs[0 % len(inputs)], '₺845.200', '₺152.136', '✓ DOĞRULANDI', '#dcfce7', '#15803d'),
        ('14.10.2026', 'TR-0429-B', inputs[1 % len(inputs)], '₺1.240.000', '₺223.200', '✓ DOĞRULANDI', '#dcfce7', '#15803d'),
        ('18.10.2026', 'TR-0430-C', inputs[2 % len(inputs)], '₺620.500', '₺111.690', '● İNCELENİYOR', '#fef3c7', '#b45309'),
        ('21.10.2026', 'TR-0431-D', inputs[3 % len(inputs)], '₺415.000', '₺74.700', '✓ DOĞRULANDI', '#dcfce7', '#15803d'),
        ('25.10.2026', 'TR-0432-E', inputs[0 % len(inputs)], '₺980.400', '₺176.472', '★ KRİTİK EŞİK', '#fee2e2', '#b91c1c'),
        ('28.10.2026', 'TR-0433-F', inputs[1 % len(inputs)], '₺1.850.000', '₺333.000', '✓ DOĞRULANDI', '#dcfce7', '#15803d'),
        ('02.11.2026', 'TR-0434-G', inputs[2 % len(inputs)], '₺740.000', '₺133.200', '✓ DOĞRULANDI', '#dcfce7', '#15803d'),
        ('05.11.2026', 'TR-0435-H', inputs[3 % len(inputs)], '₺520.800', '₺93.744', '● İNCELENİYOR', '#fef3c7', '#b45309'),
        ('09.11.2026', 'TR-0436-I', inputs[0 % len(inputs)], '₺1.120.000', '₺201.600', '✓ DOĞRULANDI', '#dcfce7', '#15803d'),
        ('12.11.2026', 'TR-0437-J', inputs[1 % len(inputs)], '₺890.000', '₺160.200', '✓ DOĞRULANDI', '#dcfce7', '#15803d'),
    ]

    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1200 750" width="1200" height="750" style="background:#0f172a; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;">
  <defs>
    <filter id="dropShadow2" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="3" stdDeviation="5" flood-color="#000000" flood-opacity="0.12" />
    </filter>
  </defs>

  <rect width="1200" height="750" fill="#f1f5f9"/>

  <!-- Üst Bar -->
  <rect width="1200" height="38" fill="#1e293b"/>
  <circle cx="20" cy="19" r="6" fill="#ef4444"/>
  <circle cx="38" cy="19" r="6" fill="#f59e0b"/>
  <circle cx="56" cy="19" r="6" fill="#10b981"/>

  <g transform="translate(85, 11)">
    <rect width="22" height="18" rx="3" fill="#107c41"/>
    <text x="11" y="13" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">X</text>
    <text x="32" y="13" font-size="12" font-weight="600" fill="#f8fafc">{html.escape(slug)}.xlsx — [VERİ GİRİŞİ &amp; OPERASYON MOTORU]</text>
  </g>
  <rect x="980" y="8" width="200" height="22" rx="11" fill="#334155"/>
  <text x="1080" y="23" font-size="10" font-weight="700" fill="#e2e8f0" text-anchor="middle">DİNAMİK VERİ MODELİ</text>

  <!-- Ribbon -->
  <rect y="38" width="1200" height="32" fill="#ffffff" stroke="#cbd5e1" stroke-width="1"/>
  <text x="25" y="58" font-size="11" font-weight="600" fill="#64748b">Dosya</text>
  <text x="75" y="58" font-size="11" font-weight="600" fill="#64748b">Pano</text>
  <rect x="125" y="44" width="70" height="26" rx="4" fill="#dcfce7"/>
  <text x="160" y="61" font-size="11" font-weight="800" fill="#15803d" text-anchor="middle">Girdiler</text>
  <text x="220" y="58" font-size="11" font-weight="600" fill="#64748b">Hesap Motoru</text>
  <text x="315" y="58" font-size="11" font-weight="600" fill="#64748b">Formüller</text>
  <text x="385" y="58" font-size="11" font-weight="600" fill="#64748b">Veri Denetimi</text>

  <!-- Formül Çubuğu -->
  <rect y="70" width="1200" height="26" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1"/>
  <rect x="15" y="73" width="55" height="20" rx="3" fill="#ffffff" stroke="#cbd5e1" stroke-width="1"/>
  <text x="42" y="87" font-size="11" font-weight="700" fill="#334155" text-anchor="middle">C12</text>
  <text x="85" y="87" font-size="12" font-style="italic" font-weight="700" fill="#0f766e">fx</text>
  <line x1="105" y1="74" x2="105" y2="92" stroke="#cbd5e1" stroke-width="1"/>
  <text x="115" y="87" font-size="11" font-family="'JetBrains Mono',monospace" fill="#0369a1">=DÜŞEYARA(A12; ParametreTablosu; 4; YANLIŞ) * EĞER(D12&gt;0; 1; 0)</text>

  <!-- Filtre ve Kontrol Şeridi -->
  <g transform="translate(20, 106)">
    <rect width="1160" height="42" rx="8" fill="#ffffff" stroke="#cbd5e1" stroke-width="1" filter="url(#dropShadow2)"/>
    <!-- Arama Kutusu -->
    <rect x="15" y="8" width="280" height="26" rx="4" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1"/>
    <text x="30" y="25" font-size="11" fill="#94a3b8">🔍 Kayıt veya hesap ara...</text>

    <!-- Hızlı İstatistik Rozetleri -->
    <rect x="320" y="8" width="180" height="26" rx="4" fill="#eff6ff"/>
    <text x="410" y="25" font-size="10.5" font-weight="700" fill="#1d4ed8" text-anchor="middle">TOPLAM: 1.450 KAYIT</text>

    <rect x="515" y="8" width="180" height="26" rx="4" fill="#ecfdf5"/>
    <text x="605" y="25" font-size="10.5" font-weight="700" fill="#047857" text-anchor="middle">HATALI HÜCRE: 0 ADET</text>

    <rect x="710" y="8" width="220" height="26" rx="4" fill="#fef3c7"/>
    <text x="820" y="25" font-size="10.5" font-weight="700" fill="#b45309" text-anchor="middle">FORMÜL DENETİMİ: TAM UYUM</text>

    <rect x="980" y="8" width="165" height="26" rx="4" fill="{cfg['accent']}"/>
    <text x="1062" y="25" font-size="10.5" font-weight="800" fill="#ffffff" text-anchor="middle">+ YENİ SATIR EKLE</text>
  </g>

  <!-- BÜYÜK EXCEL TABLOSU -->
  <g transform="translate(20, 160)">
    <rect width="1160" height="535" rx="8" fill="#ffffff" stroke="#cbd5e1" stroke-width="1" filter="url(#dropShadow2)"/>

    <!-- Tablo Header -->
    <rect width="1160" height="34" rx="8" fill="#1e293b"/>
    <text x="25" y="22" font-size="10.5" font-weight="800" fill="#ffffff">TARİH</text>
    <text x="140" y="22" font-size="10.5" font-weight="800" fill="#ffffff">REFERANS NO</text>
    <text x="310" y="22" font-size="10.5" font-weight="800" fill="#ffffff">İŞLEM AÇIKLAMASI VE DETAY</text>
    <text x="680" y="22" font-size="10.5" font-weight="800" fill="#ffffff" text-anchor="end">BRÜT TUTAR (₺)</text>
    <text x="830" y="22" font-size="10.5" font-weight="800" fill="#ffffff" text-anchor="end">KDV / MASRAF (₺)</text>
    <text x="1020" y="22" font-size="10.5" font-weight="800" fill="#ffffff" text-anchor="middle">DOĞRULAMA DURUMU</text>'''

    y_pos = 34
    for r in row_data:
        bg = '#f8fafc' if (y_pos // 36) % 2 == 0 else '#ffffff'
        svg += f'''
    <g transform="translate(0, {y_pos})">
      <rect width="1160" height="36" fill="{bg}"/>
      <line x1="0" y1="36" x2="1160" y2="36" stroke="#f1f5f9" stroke-width="1"/>
      <text x="25" y="23" font-size="11" font-weight="600" fill="#475569">{r[0]}</text>
      <text x="140" y="23" font-size="11" font-family="'JetBrains Mono',monospace" font-weight="700" fill="#0f172a">{r[1]}</text>
      <text x="310" y="23" font-size="11" font-weight="600" fill="#334155">{html.escape(r[2])}</text>
      <text x="680" y="23" font-size="11.5" font-weight="800" fill="#0f172a" text-anchor="end">{r[3]}</text>
      <text x="830" y="23" font-size="11" font-weight="600" fill="#64748b" text-anchor="end">{r[4]}</text>
      <rect x="955" y="8" width="130" height="20" rx="10" fill="{r[6]}"/>
      <text x="1020" y="22" font-size="9.5" font-weight="800" fill="{r[7]}" text-anchor="middle">{r[5]}</text>
    </g>'''
        y_pos += 36

    svg += f'''
    <!-- Tablo Alt Toplam -->
    <g transform="translate(0, 480)">
      <rect width="1160" height="42" fill="#0f172a"/>
      <text x="25" y="26" font-size="11.5" font-weight="800" fill="#ffffff">SAYFA TOPLAMI (10 / 1.450 KAYIT)</text>
      <text x="680" y="26" font-size="13" font-weight="900" fill="#38bdf8" text-anchor="end">₺9.247.100</text>
      <text x="830" y="26" font-size="12" font-weight="800" fill="#f87171" text-anchor="end">₺1.644.406</text>
      <text x="1020" y="26" font-size="10.5" font-weight="800" fill="#4ade80" text-anchor="middle">✓ MATEMATİKSEL KONTROL OK</text>
    </g>
  </g>

  <!-- Sheet Bar -->
  <rect y="705" width="1200" height="45" fill="#e2e8f0" stroke="#cbd5e1" stroke-width="1"/>
  <g transform="translate(15, 715)">
    <text x="5" y="18" font-size="13" font-weight="700" fill="#64748b">◀  ▶</text>
    <rect x="45" y="0" width="85" height="35" fill="#f8fafc" rx="4"/><text x="87" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">KAPAK</text>
    <rect x="135" y="0" width="85" height="35" fill="#f8fafc" rx="4"/><text x="177" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">PANO</text>
    <rect x="225" y="-3" width="110" height="38" fill="#ffffff" rx="4" stroke="#107c41" stroke-width="2"/>
    <text x="280" y="19" font-size="10.5" font-weight="900" fill="#107c41" text-anchor="middle">GİRDİLER (Aktif)</text>
    <rect x="345" y="0" width="105" height="35" fill="#f8fafc" rx="4"/><text x="397" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">HESAP_MOTORU</text>
    <rect x="460" y="0" width="85" height="35" fill="#f8fafc" rx="4"/><text x="502" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">RAPOR</text>
    <text x="1150" y="18" font-size="10" font-weight="600" fill="#64748b" text-anchor="end">SAYFA 2 / 24 · FORMÜL KORUMASI: AKTİF</text>
  </g>
</svg>
'''
    svg = re.sub(r'&(?!(amp|lt|gt|quot|apos);)', '&amp;', svg)
    return svg

def generate_svg_screen3(meta, cfg):
    name = meta['name']
    slug = meta['slug']

    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1200 750" width="1200" height="750" style="background:#0f172a; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;">
  <defs>
    <filter id="dropShadow3" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="3" stdDeviation="5" flood-color="#000000" flood-opacity="0.12" />
    </filter>
    <linearGradient id="stampGrad" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#107c41"/>
      <stop offset="100%" stop-color="#064e3b"/>
    </linearGradient>
  </defs>

  <rect width="1200" height="750" fill="#f1f5f9"/>

  <!-- Üst Pencere -->
  <rect width="1200" height="38" fill="#1e293b"/>
  <circle cx="20" cy="19" r="6" fill="#ef4444"/>
  <circle cx="38" cy="19" r="6" fill="#f59e0b"/>
  <circle cx="56" cy="19" r="6" fill="#10b981"/>

  <g transform="translate(85, 11)">
    <rect width="22" height="18" rx="3" fill="#107c41"/>
    <text x="11" y="13" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">X</text>
    <text x="32" y="13" font-size="12" font-weight="600" fill="#f8fafc">{html.escape(slug)}.xlsx — [YÖNETİM KURULU İCRA RAPORU &amp; SENARYO ANALİZİ]</text>
  </g>
  <rect x="980" y="8" width="200" height="22" rx="11" fill="#334155"/>
  <text x="1080" y="23" font-size="10" font-weight="700" fill="#e2e8f0" text-anchor="middle">KARAR &amp; SENARYO MOTORU</text>

  <!-- Rapor Başlık Alanı -->
  <g transform="translate(30, 60)">
    <rect width="1140" height="85" rx="8" fill="#ffffff" stroke="#cbd5e1" stroke-width="1" filter="url(#dropShadow3)"/>
    <rect x="0" y="0" width="8" height="85" rx="4" fill="{cfg['accent']}"/>
    <text x="25" y="32" font-size="18" font-weight="900" fill="#0f172a">{html.escape(name[:55].upper())} — NİHAİ KARAR VE DEĞERLENDİRME BELGESİ</text>
    <text x="25" y="52" font-size="11.5" font-weight="600" fill="#64748b">Yönetim Kurulu ve İcra Komitesi için Otomatik Türetilen Finansal Stres Testi Raporu</text>
    <text x="25" y="70" font-size="10.5" font-weight="700" fill="{cfg['accent']}">DÖNEM: 2026 Q3 · DENETİM PROTOKOLÜ: ISO/VUK/IFRS UYUMLU · DURUM: ONAYLANDI</text>

    <!-- Resmi Doğrulama Mührü -->
    <g transform="translate(970, 12)">
      <rect width="145" height="60" rx="8" fill="url(#stampGrad)"/>
      <text x="72" y="24" font-size="9" font-weight="800" fill="#a7f3d0" text-anchor="middle">DECISION OS v15.1</text>
      <text x="72" y="42" font-size="12" font-weight="900" fill="#ffffff" text-anchor="middle">✓ DOĞRULANDI</text>
      <text x="72" y="54" font-size="8" font-weight="700" fill="#a7f3d0" text-anchor="middle">0 MATEMATİKSEL HATA</text>
    </g>
  </g>

  <!-- 3 SENARYO ANALİZ KARTI -->
  <g transform="translate(30, 165)">
    <!-- Senaryo 1: Kötümser -->
    <g transform="translate(0, 0)">
      <rect width="360" height="240" rx="8" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow3)"/>
      <rect width="360" height="32" rx="8" fill="#fee2e2"/>
      <text x="15" y="21" font-size="11" font-weight="800" fill="#991b1b">SENARYO A: KÖTÜMSER (STRES TESTİ -%20)</text>
      
      <text x="20" y="65" font-size="11" font-weight="600" fill="#64748b">Ciro / Hacim Düşüşü:</text>
      <text x="340" y="65" font-size="11.5" font-weight="800" fill="#dc2626" text-anchor="end">-%20.0 (Kriz Eşiği)</text>

      <text x="20" y="95" font-size="11" font-weight="600" fill="#64748b">Beklenen Net Kâr / Getiri:</text>
      <text x="340" y="95" font-size="12" font-weight="800" fill="#0f172a" text-anchor="end">₺4.280.000</text>

      <text x="20" y="125" font-size="11" font-weight="600" fill="#64748b">Minimum Nakit Tamponu:</text>
      <text x="340" y="125" font-size="11.5" font-weight="800" fill="#15803d" text-anchor="end">₺1.850.000 (Yeterli)</text>

      <text x="20" y="155" font-size="11" font-weight="600" fill="#64748b">Likidite Karşılama Oranı:</text>
      <text x="340" y="155" font-size="11.5" font-weight="800" fill="#15803d" text-anchor="end">1.28x (Eşik: 1.10x)</text>

      <rect x="20" y="185" width="320" height="36" rx="6" fill="#fef2f2" stroke="#fecaca" stroke-width="1"/>
      <text x="180" y="207" font-size="10" font-weight="800" fill="#b91c1c" text-anchor="middle">STRES SENARYOSUNDA DAHİ İFLAS RİSKİ YOK</text>
    </g>

    <!-- Senaryo 2: Baz / Beklenen -->
    <g transform="translate(390, 0)">
      <rect width="360" height="240" rx="8" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow3)"/>
      <rect width="360" height="32" rx="8" fill="#dbeafe"/>
      <text x="15" y="21" font-size="11" font-weight="800" fill="#1e40af">SENARYO B: BAZ MODEL (HEDEFLENEN)</text>

      <text x="20" y="65" font-size="11" font-weight="600" fill="#64748b">Ciro / Hacim Gerçekleşmesi:</text>
      <text x="340" y="65" font-size="11.5" font-weight="800" fill="#2563eb" text-anchor="end">%100.0 (Tam Bütçe)</text>

      <text x="20" y="95" font-size="11" font-weight="600" fill="#64748b">Beklenen Net Kâr / Getiri:</text>
      <text x="340" y="95" font-size="12" font-weight="800" fill="#0f172a" text-anchor="end">₺12.450.000</text>

      <text x="20" y="125" font-size="11" font-weight="600" fill="#64748b">Serbest Nakit Akımı:</text>
      <text x="340" y="125" font-size="11.5" font-weight="800" fill="#15803d" text-anchor="end">₺8.240.000</text>

      <text x="20" y="155" font-size="11" font-weight="600" fill="#64748b">Yatırım Geri Dönüşü (ROI):</text>
      <text x="340" y="155" font-size="11.5" font-weight="800" fill="#15803d" text-anchor="end">%38.4 Yıllık</text>

      <rect x="20" y="185" width="320" height="36" rx="6" fill="#eff6ff" stroke="#bfdbfe" stroke-width="1"/>
      <text x="180" y="207" font-size="10" font-weight="800" fill="#1d4ed8" text-anchor="middle">OPTİMUM İCRA VE PLANLAMA REÇETESİ</text>
    </g>

    <!-- Senaryo 3: İyimser -->
    <g transform="translate(780, 0)">
      <rect width="360" height="240" rx="8" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow3)"/>
      <rect width="360" height="32" rx="8" fill="#dcfce7"/>
      <text x="15" y="21" font-size="11" font-weight="800" fill="#166534">SENARYO C: İYİMSER (BÜYÜME +%25)</text>

      <text x="20" y="65" font-size="11" font-weight="600" fill="#64748b">Ciro / Pazar Payı Artışı:</text>
      <text x="340" y="65" font-size="11.5" font-weight="800" fill="#15803d" text-anchor="end">+%25.0 (Agresif Büyüme)</text>

      <text x="20" y="95" font-size="11" font-weight="600" fill="#64748b">Beklenen Net Kâr / Getiri:</text>
      <text x="340" y="95" font-size="12" font-weight="800" fill="#15803d" text-anchor="end">₺18.920.000</text>

      <text x="20" y="125" font-size="11" font-weight="600" fill="#64748b">Nakit Temettü Kapasitesi:</text>
      <text x="340" y="125" font-size="11.5" font-weight="800" fill="#15803d" text-anchor="end">₺11.400.000</text>

      <text x="20" y="155" font-size="11" font-weight="600" fill="#64748b">Şirket Çarpan Değerlemesi:</text>
      <text x="340" y="155" font-size="11.5" font-weight="800" fill="#15803d" text-anchor="end">8.5x FAVÖK</text>

      <rect x="20" y="185" width="320" height="36" rx="6" fill="#ecfdf5" stroke="#a7f3d0" stroke-width="1"/>
      <text x="180" y="207" font-size="10" font-weight="800" fill="#047857" text-anchor="middle">MAKSİMUM DEĞER YARATMA POTANSİYELİ</text>
    </g>
  </g>

  <!-- ALT ANALİTİK ÇIKTILAR VE İMZA BLOĞU -->
  <g transform="translate(30, 425)">
    <rect width="1140" height="265" rx="8" fill="#ffffff" stroke="#cbd5e1" stroke-width="1" filter="url(#dropShadow3)"/>
    <text x="25" y="30" font-size="13" font-weight="800" fill="#0f172a">YÖNETİM KARAR MATRİSİ VE UYGULAMA REÇETESİ</text>

    <!-- 4 Madde -->
    <g transform="translate(25, 50)">
      <circle cx="8" cy="8" r="4" fill="{cfg['accent']}"/>
      <text x="22" y="12" font-size="11.5" font-weight="700" fill="#0f172a">Öncelik 1: Acil Likidite ve Borç Servis Planlaması</text>
      <text x="22" y="28" font-size="10.5" font-weight="500" fill="#64748b">Sistemde hesaplanan vade günlerine sadık kalınarak nakit blokajı 3 gün önceden hazır edilmelidir.</text>

      <circle cx="8" cy="56" r="4" fill="{cfg['accent']}"/>
      <text x="22" y="60" font-size="11.5" font-weight="700" fill="#0f172a">Öncelik 2: Formül Doğrulama ve Çapraz Sağlama</text>
      <text x="22" y="76" font-size="10.5" font-weight="500" fill="#64748b">Tüm hücreler kilitli formül korumasında olup girdi sayfalarında sıfır döngüsel başvuru (circular ref) garantilidir.</text>

      <circle cx="8" cy="104" r="4" fill="{cfg['accent']}"/>
      <text x="22" y="108" font-size="11.5" font-weight="700" fill="#0f172a">Öncelik 3: Resmi Mevzuat ve Vergi Kalkanı Entegrasyonu</text>
      <text x="22" y="124" font-size="10.5" font-weight="500" fill="#64748b">2026 güncel vergi dilimleri, amortisman ve SGK teşvik parametreleri tam otomatik yansıtılmaktadır.</text>

      <circle cx="8" cy="152" r="4" fill="{cfg['accent']}"/>
      <text x="22" y="156" font-size="11.5" font-weight="700" fill="#0f172a">Öncelik 4: Yönetici İmzası ve Rapor Paylaşımı</text>
      <text x="22" y="172" font-size="10.5" font-weight="500" fill="#64748b">Tek tuşla PDF ve A4 çıktı formatına tam uyumlu sayfa yapısı kurgulanmıştır.</text>
    </g>

    <!-- Sağ İmza / Onay Kartı -->
    <g transform="translate(830, 45)">
      <rect width="280" height="195" rx="6" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1"/>
      <text x="140" y="28" font-size="11" font-weight="800" fill="#0f172a" text-anchor="middle">SİSTEM ONAY VE İMZA KÜNYESİ</text>
      <line x1="20" y1="40" x2="260" y2="40" stroke="#e2e8f0" stroke-width="1"/>
      <text x="25" y="65" font-size="10" font-weight="600" fill="#64748b">Sistem Mimarı:</text>
      <text x="255" y="65" font-size="10" font-weight="700" fill="#0f172a" text-anchor="end">Excel Arşiv Lead Eng.</text>

      <text x="25" y="90" font-size="10" font-weight="600" fill="#64748b">Kalite Kapısı:</text>
      <text x="255" y="90" font-size="10" font-weight="800" fill="#15803d" text-anchor="end">4/4 PASS (0 Hata)</text>

      <text x="25" y="115" font-size="10" font-weight="600" fill="#64748b">Rapor Tarihi:</text>
      <text x="255" y="115" font-size="10" font-weight="700" fill="#0f172a" text-anchor="end">2026 Dönemi</text>

      <rect x="25" y="135" width="230" height="40" rx="4" fill="#ffffff" stroke="#cbd5e1" stroke-dasharray="4 2"/>
      <text x="140" y="160" font-size="10" font-style="italic" fill="#94a3b8" text-anchor="middle">Yönetici İmzası &amp; Mühür</text>
    </g>
  </g>

  <!-- Sheet Bar -->
  <rect y="705" width="1200" height="45" fill="#e2e8f0" stroke="#cbd5e1" stroke-width="1"/>
  <g transform="translate(15, 715)">
    <text x="5" y="18" font-size="13" font-weight="700" fill="#64748b">◀  ▶</text>
    <rect x="45" y="0" width="85" height="35" fill="#f8fafc" rx="4"/><text x="87" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">KAPAK</text>
    <rect x="135" y="0" width="85" height="35" fill="#f8fafc" rx="4"/><text x="177" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">PANO</text>
    <rect x="225" y="0" width="85" height="35" fill="#f8fafc" rx="4"/><text x="267" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">GİRDİLER</text>
    <rect x="315" y="-3" width="145" height="38" fill="#ffffff" rx="4" stroke="#107c41" stroke-width="2"/>
    <text x="387" y="19" font-size="10.5" font-weight="900" fill="#107c41" text-anchor="middle">YÖNETİM_RAPORU (Aktif)</text>
    <text x="1150" y="18" font-size="10" font-weight="600" fill="#64748b" text-anchor="end">SAYFA 24 / 24 · YÖNETİM SUNUMUNA HAZIR</text>
  </g>
</svg>
'''
    svg = re.sub(r'&(?!(amp|lt|gt|quot|apos);)', '&amp;', svg)
    return svg

def run_all_screens():
    import glob
    from generate_all_rich_dashboards import parse_mdx, MDX_DIR, SCREENSHOT_DIR, CATEGORY_CONFIGS
    files = glob.glob(os.path.join(MDX_DIR, '*.mdx'))
    print(f"Generating Screen 2 and Screen 3 for {len(files)} templates...")
    for f in files:
        meta = parse_mdx(f)
        cfg = CATEGORY_CONFIGS.get(meta['category'], CATEGORY_CONFIGS['finansal-analiz'])
        slug = meta['slug']

        # Screen 2
        s2_svg = generate_svg_screen2(meta, cfg)
        s2_path = f'/tmp/{slug}-2.svg'
        with open(s2_path, 'w', encoding='utf-8') as out:
            out.write(s2_svg)
        import subprocess
        subprocess.run(['qlmanage', '-t', '-s', '1200', '-o', '/tmp', s2_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        tmp_png = f'/tmp/{slug}-2.svg.png'
        dest_png = os.path.join(SCREENSHOT_DIR, f"{slug}-2.png")
        if os.path.exists(tmp_png):
            os.replace(tmp_png, dest_png)

        # Screen 3
        s3_svg = generate_svg_screen3(meta, cfg)
        s3_path = f'/tmp/{slug}-3.svg'
        with open(s3_path, 'w', encoding='utf-8') as out:
            out.write(s3_svg)
        subprocess.run(['qlmanage', '-t', '-s', '1200', '-o', '/tmp', s3_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        tmp_png3 = f'/tmp/{slug}-3.svg.png'
        dest_png3 = os.path.join(SCREENSHOT_DIR, f"{slug}-3.png")
        if os.path.exists(tmp_png3):
            os.replace(tmp_png3, dest_png3)

    print("All Screen 2 and Screen 3 files generated successfully!")

if __name__ == '__main__':
    run_all_screens()
