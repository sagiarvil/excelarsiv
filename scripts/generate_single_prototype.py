import os

svg_content = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1200 750" width="1200" height="750" style="background:#0f172a; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;">
  <defs>
    <!-- Gölgeler ve Gradyanlar -->
    <filter id="dropShadow" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="3" stdDeviation="5" flood-color="#000000" flood-opacity="0.12" />
    </filter>
    <filter id="glowGreen" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="0" stdDeviation="4" flood-color="#10b981" flood-opacity="0.4" />
    </filter>
    <linearGradient id="cardGrad1" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#ffffff"/>
      <stop offset="100%" stop-color="#f8fafc"/>
    </linearGradient>
    <linearGradient id="heroRibbon" x1="0" y1="0" x2="1" y2="0">
      <stop offset="0%" stop-color="#0f766e"/>
      <stop offset="100%" stop-color="#042f2e"/>
    </linearGradient>
  </defs>

  <!-- 1. DIŞ ÇERÇEVE & ARKA PLAN -->
  <rect width="1200" height="750" fill="#f1f5f9"/>

  <!-- 2. EXCEL ÜST PENCERE BAŞLIĞI -->
  <rect width="1200" height="38" fill="#1e293b"/>
  <!-- Pencere Düğmeleri -->
  <circle cx="20" cy="19" r="6" fill="#ef4444"/>
  <circle cx="38" cy="19" r="6" fill="#f59e0b"/>
  <circle cx="56" cy="19" r="6" fill="#10b981"/>

  <!-- Dosya Bilgisi & Rozet -->
  <g transform="translate(85, 11)">
    <rect width="22" height="18" rx="3" fill="#107c41"/>
    <text x="11" y="13" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">X</text>
    <text x="32" y="13" font-size="12" font-weight="600" fill="#f8fafc">Banka_Kredi_ve_Taksit_Takip_Sistemi_v15.1.xlsx — Excel Arşiv Decision OS</text>
  </g>
  <rect x="970" y="8" width="210" height="22" rx="11" fill="#334155"/>
  <circle cx="985" cy="19" r="4" fill="#10b981" filter="url(#glowGreen)"/>
  <text x="998" y="23" font-size="10" font-weight="700" fill="#e2e8f0">CANLI YÖNETİM KOKPİTİ</text>

  <!-- 3. EXCEL RIBBON & FORMULA BAR -->
  <rect y="38" width="1200" height="32" fill="#ffffff" stroke="#cbd5e1" stroke-width="1"/>
  <!-- Sekmeler -->
  <text x="25" y="58" font-size="11" font-weight="600" fill="#64748b">Dosya</text>
  <rect x="70" y="44" width="56" height="26" rx="4" fill="#dcfce7"/>
  <text x="98" y="61" font-size="11" font-weight="800" fill="#15803d" text-anchor="middle">Pano</text>
  <text x="145" y="58" font-size="11" font-weight="600" fill="#64748b">Krediler</text>
  <text x="205" y="58" font-size="11" font-weight="600" fill="#64748b">Taksit Planı</text>
  <text x="285" y="58" font-size="11" font-weight="600" fill="#64748b">Nakit Akışı</text>
  <text x="365" y="58" font-size="11" font-weight="600" fill="#64748b">Stres Testi</text>
  <text x="445" y="58" font-size="11" font-weight="600" fill="#64748b">Rapor</text>

  <!-- Formül Çubuğu -->
  <rect y="70" width="1200" height="26" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1"/>
  <rect x="15" y="73" width="55" height="20" rx="3" fill="#ffffff" stroke="#cbd5e1" stroke-width="1"/>
  <text x="42" y="87" font-size="11" font-weight="700" fill="#334155" text-anchor="middle">B4</text>
  <text x="85" y="87" font-size="12" font-style="italic" font-weight="700" fill="#0f766e">fx</text>
  <line x1="105" y1="74" x2="105" y2="92" stroke="#cbd5e1" stroke-width="1"/>
  <text x="115" y="87" font-size="11" font-family="'JetBrains Mono',monospace" fill="#0369a1">=BORÇ_SERVİS_KARŞILAMA_ORANI(ToplamNakitGirişi; ToplamKrediTaksitleri; %15_GüvenlikMarjı)</text>

  <!-- 4. DASHBOARD HEADER KONSOLU -->
  <g transform="translate(20, 105)">
    <!-- Sol Başlık Bloğu -->
    <text x="0" y="24" font-size="18" font-weight="900" fill="#0f172a" letter-spacing="-0.3">BANKA, KREDİ VE NAKİT TAKSİT YÖNETİM MERKEZİ</text>
    <text x="0" y="42" font-size="11" font-weight="600" fill="#64748b">Konsolide Banka Borç Yükü, Vade Dağılımı ve Likidite Karar Kokpiti · 2026 Mali Dönem</text>

    <!-- Sağ Konsol Filtreleri -->
    <rect x="850" y="6" width="140" height="34" rx="6" fill="#ffffff" stroke="#cbd5e1" stroke-width="1"/>
    <text x="862" y="27" font-size="11" font-weight="700" fill="#334155">Dönem: 2026 / Q3-Q4</text>
    <text x="975" y="27" font-size="10" fill="#64748b">▼</text>

    <rect x="1000" y="6" width="160" height="34" rx="6" fill="#107c41" filter="url(#dropShadow)"/>
    <text x="1080" y="27" font-size="11" font-weight="800" fill="#ffffff" text-anchor="middle">₺ TÜM KREDİLER (4)</text>
  </g>

  <!-- 5. DÖRT ADET KRİTİK YÖNETİCİ KPI KARTI -->
  <g transform="translate(20, 158)">
    <!-- KART 1: TOPLAM KREDİ PORTFÖYÜ -->
    <g transform="translate(0, 0)">
      <rect width="275" height="90" rx="8" fill="url(#cardGrad1)" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <rect x="0" y="0" width="5" height="90" rx="2" fill="#2563eb"/>
      <text x="16" y="23" font-size="10.5" font-weight="700" fill="#64748b">TOPLAM KREDİ PORTFÖYÜ</text>
      <rect x="195" y="10" width="68" height="18" rx="9" fill="#dbeafe"/>
      <text x="229" y="23" font-size="9" font-weight="800" fill="#1d4ed8" text-anchor="middle">4 BANKA</text>
      <text x="16" y="54" font-size="23" font-weight="900" fill="#0f172a">₺28.450.000</text>
      <text x="16" y="75" font-size="10" font-weight="700" fill="#16a34a">↑ %12.5</text>
      <text x="60" y="75" font-size="10" font-weight="500" fill="#64748b">aktif limit kullanım payı</text>
    </g>

    <!-- KART 2: BU AY ÖDENECEK TAKSİT -->
    <g transform="translate(295, 0)">
      <rect width="275" height="90" rx="8" fill="url(#cardGrad1)" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <rect x="0" y="0" width="5" height="90" rx="2" fill="#dc2626"/>
      <text x="16" y="23" font-size="10.5" font-weight="700" fill="#64748b">BU AY ÖDENECEK TAKSİT</text>
      <rect x="185" y="10" width="78" height="18" rx="9" fill="#fee2e2"/>
      <text x="224" y="23" font-size="9" font-weight="800" fill="#b91c1c" text-anchor="middle">VADE: 6 GÜN</text>
      <text x="16" y="54" font-size="23" font-weight="900" fill="#b91c1c">₺1.428.600</text>
      <text x="16" y="75" font-size="10" font-weight="700" fill="#dc2626">● 3 Taksit</text>
      <text x="72" y="75" font-size="10" font-weight="500" fill="#64748b">Nakit Blokesi: Hazır</text>
    </g>

    <!-- KART 3: KALAN TOPLAM ANAPARA BORCU -->
    <g transform="translate(590, 0)">
      <rect width="275" height="90" rx="8" fill="url(#cardGrad1)" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <rect x="0" y="0" width="5" height="90" rx="2" fill="#16a34a"/>
      <text x="16" y="23" font-size="10.5" font-weight="700" fill="#64748b">KALAN ANAPARA BAKİYESİ</text>
      <rect x="195" y="10" width="68" height="18" rx="9" fill="#dcfce7"/>
      <text x="229" y="23" font-size="9" font-weight="800" fill="#15803d" text-anchor="middle">%41 İTFA</text>
      <text x="16" y="54" font-size="23" font-weight="900" fill="#0f172a">₺16.820.400</text>
      <rect x="16" y="66" width="160" height="7" rx="3.5" fill="#e2e8f0"/>
      <rect x="16" y="66" width="66" height="7" rx="3.5" fill="#16a34a"/>
      <text x="186" y="73" font-size="9.5" font-weight="700" fill="#64748b">18/36 Ay</text>
    </g>

    <!-- KART 4: AĞIRLIKLI FAİZ VE DSCR -->
    <g transform="translate(885, 0)">
      <rect width="275" height="90" rx="8" fill="url(#cardGrad1)" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <rect x="0" y="0" width="5" height="90" rx="2" fill="#0d9488"/>
      <text x="16" y="23" font-size="10.5" font-weight="700" fill="#64748b">AĞIRLIKLI FAİZ / DSCR</text>
      <rect x="195" y="10" width="68" height="18" rx="9" fill="#ccfbf1"/>
      <text x="229" y="23" font-size="9" font-weight="800" fill="#0f766e" text-anchor="middle">DSCR 1.84x</text>
      <text x="16" y="54" font-size="23" font-weight="900" fill="#0f766e">%3.12 <tspan font-size="13" font-weight="600" fill="#64748b">/ Aylık</tspan></text>
      <text x="16" y="75" font-size="10" font-weight="700" fill="#0f766e">✓ GÜVENLİ</text>
      <text x="82" y="75" font-size="10" font-weight="500" fill="#64748b">Faiz Kalkanı: Aktif</text>
    </g>
  </g>

  <!-- 6. ORTA BÖLÜM: 12 AYLIK GRAFİK & TABLO -->
  <g transform="translate(20, 262)">
    <!-- SOL PANEL: 12 AYLIK GRAFİK -->
    <g transform="translate(0, 0)">
      <rect width="460" height="270" rx="10" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <text x="16" y="25" font-size="12" font-weight="800" fill="#0f172a">12 AYLIK TAKSİT YÜKÜ &amp; LİKİDİTE DENGESİ</text>
      <text x="16" y="40" font-size="10" font-weight="500" fill="#64748b">Aylık Kümülatif Kredi Çıkışı vs. Nakit Girişi Projeksiyonu</text>

      <rect x="285" y="15" width="10" height="10" rx="2" fill="#2563eb"/>
      <text x="300" y="24" font-size="9" font-weight="600" fill="#475569">Nakit Giriş</text>
      <rect x="375" y="15" width="10" height="10" rx="2" fill="#dc2626"/>
      <text x="390" y="24" font-size="9" font-weight="600" fill="#475569">Taksit Çıkış</text>

      <g transform="translate(35, 58)">
        <line x1="0" y1="0" x2="400" y2="0" stroke="#f1f5f9" stroke-width="1"/>
        <line x1="0" y1="40" x2="400" y2="40" stroke="#f1f5f9" stroke-width="1"/>
        <line x1="0" y1="80" x2="400" y2="80" stroke="#f1f5f9" stroke-width="1"/>
        <line x1="0" y1="120" x2="400" y2="120" stroke="#f1f5f9" stroke-width="1"/>
        <line x1="0" y1="160" x2="400" y2="160" stroke="#cbd5e1" stroke-width="1.5"/>

        <text x="-8" y="5" font-size="9" fill="#94a3b8" text-anchor="end">₺3M</text>
        <text x="-8" y="45" font-size="9" fill="#94a3b8" text-anchor="end">₺2M</text>
        <text x="-8" y="85" font-size="9" fill="#94a3b8" text-anchor="end">₺1M</text>
        <text x="-8" y="125" font-size="9" fill="#94a3b8" text-anchor="end">₺500K</text>
        <text x="-8" y="164" font-size="9" fill="#94a3b8" text-anchor="end">0</text>

        <!-- Sütunlar -->
        <rect x="15" y="50" width="16" height="110" rx="2" fill="#2563eb"/><rect x="33" y="85" width="16" height="75" rx="2" fill="#dc2626"/>
        <text x="32" y="176" font-size="9.5" font-weight="700" fill="#475569" text-anchor="middle">Oca</text>

        <rect x="75" y="40" width="16" height="120" rx="2" fill="#2563eb"/><rect x="93" y="90" width="16" height="70" rx="2" fill="#dc2626"/>
        <text x="92" y="176" font-size="9.5" font-weight="700" fill="#475569" text-anchor="middle">Şub</text>

        <rect x="135" y="30" width="16" height="130" rx="2" fill="#2563eb"/><rect x="153" y="65" width="16" height="95" rx="2" fill="#dc2626"/>
        <text x="152" y="176" font-size="9.5" font-weight="700" fill="#475569" text-anchor="middle">Mar</text>

        <rect x="195" y="45" width="16" height="115" rx="2" fill="#2563eb"/><rect x="213" y="95" width="16" height="65" rx="2" fill="#dc2626"/>
        <text x="212" y="176" font-size="9.5" font-weight="700" fill="#475569" text-anchor="middle">Nis</text>

        <rect x="255" y="25" width="16" height="135" rx="2" fill="#2563eb"/><rect x="273" y="105" width="16" height="55" rx="2" fill="#dc2626"/>
        <text x="272" y="176" font-size="9.5" font-weight="700" fill="#475569" text-anchor="middle">May</text>

        <rect x="315" y="20" width="16" height="140" rx="2" fill="#2563eb"/><rect x="333" y="115" width="16" height="45" rx="2" fill="#dc2626"/>
        <text x="332" y="176" font-size="9.5" font-weight="700" fill="#475569" text-anchor="middle">Haz</text>

        <path d="M 32 60 L 92 50 L 152 40 L 212 55 L 272 35 L 332 25" fill="none" stroke="#10b981" stroke-width="3"/>
        <circle cx="332" cy="25" r="4" fill="#10b981"/>
        <rect x="340" y="16" width="62" height="18" rx="4" fill="#10b981"/>
        <text x="371" y="29" font-size="8.5" font-weight="800" fill="#ffffff" text-anchor="middle">+₺1.85M Net</text>
      </g>
    </g>

    <!-- SAĞ PANEL: TEMİZ HİZALANMIŞ BANKA KREDİ TABLOSU -->
    <g transform="translate(475, 0)">
      <rect width="685" height="270" rx="10" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <text x="16" y="25" font-size="12" font-weight="800" fill="#0f172a">KONSOLİDE KREDİ PORTFÖYÜ VE VADE ÇİZELGESİ</text>
      <text x="16" y="40" font-size="10" font-weight="500" fill="#64748b">Anlık Banka Ekstreleri, Faiz Oranları ve Otomatik Taksit Takvimi</text>

      <!-- Tablo Başlık Satırı -->
      <g transform="translate(16, 52)">
        <rect width="653" height="26" rx="4" fill="#1e293b"/>
        <text x="12" y="17" font-size="9.5" font-weight="800" fill="#ffffff">BANKA</text>
        <text x="125" y="17" font-size="9.5" font-weight="800" fill="#ffffff">KREDİ TÜRÜ</text>
        <text x="290" y="17" font-size="9.5" font-weight="800" fill="#ffffff" text-anchor="end">ANAPARA</text>
        <text x="390" y="17" font-size="9.5" font-weight="800" fill="#ffffff" text-anchor="end">AYLIK TAKSİT</text>
        <text x="460" y="17" font-size="9.5" font-weight="800" fill="#ffffff" text-anchor="end">KALAN</text>
        <text x="535" y="17" font-size="9.5" font-weight="800" fill="#ffffff" text-anchor="middle">VADE</text>
        <text x="610" y="17" font-size="9.5" font-weight="800" fill="#ffffff" text-anchor="middle">DURUM</text>
      </g>

      <!-- Satır 1 -->
      <g transform="translate(16, 82)">
        <rect width="653" height="32" fill="#f8fafc"/>
        <text x="12" y="20" font-size="11" font-weight="700" fill="#0f172a">Garanti BBVA</text>
        <text x="125" y="20" font-size="10" font-weight="500" fill="#475569">Ticari Rotatif (BCH)</text>
        <text x="290" y="20" font-size="11" font-weight="800" fill="#0f172a" text-anchor="end">₺8.500.000</text>
        <text x="390" y="20" font-size="11" font-weight="800" fill="#dc2626" text-anchor="end">₺425.000</text>
        <text x="460" y="20" font-size="10" font-weight="600" fill="#64748b" text-anchor="end">14 Ay</text>
        <text x="535" y="20" font-size="10" font-weight="700" fill="#334155" text-anchor="middle">12.10.2026</text>
        <rect x="575" y="7" width="70" height="18" rx="9" fill="#dcfce7"/>
        <text x="610" y="20" font-size="9" font-weight="800" fill="#15803d" text-anchor="middle">✓ DÜZENLİ</text>
      </g>

      <!-- Satır 2 -->
      <g transform="translate(16, 118)">
        <rect width="653" height="32" fill="#ffffff"/>
        <text x="12" y="20" font-size="11" font-weight="700" fill="#0f172a">İş Bankası</text>
        <text x="125" y="20" font-size="10" font-weight="500" fill="#475569">Yatırım Teşvik Kredisi</text>
        <text x="290" y="20" font-size="11" font-weight="800" fill="#0f172a" text-anchor="end">₺12.000.000</text>
        <text x="390" y="20" font-size="11" font-weight="800" fill="#dc2626" text-anchor="end">₺580.400</text>
        <text x="460" y="20" font-size="10" font-weight="600" fill="#64748b" text-anchor="end">28 Ay</text>
        <text x="535" y="20" font-size="10" font-weight="700" fill="#334155" text-anchor="middle">18.10.2026</text>
        <rect x="575" y="7" width="70" height="18" rx="9" fill="#dcfce7"/>
        <text x="610" y="20" font-size="9" font-weight="800" fill="#15803d" text-anchor="middle">✓ DÜZENLİ</text>
      </g>

      <!-- Satır 3 -->
      <g transform="translate(16, 154)">
        <rect width="653" height="32" fill="#f8fafc"/>
        <text x="12" y="20" font-size="11" font-weight="700" fill="#0f172a">Yapı Kredi</text>
        <text x="125" y="20" font-size="10" font-weight="500" fill="#475569">Ekipman Finansmanı</text>
        <text x="290" y="20" font-size="11" font-weight="800" fill="#0f172a" text-anchor="end">₺4.250.000</text>
        <text x="390" y="20" font-size="11" font-weight="800" fill="#dc2626" text-anchor="end">₺235.200</text>
        <text x="460" y="20" font-size="10" font-weight="600" fill="#64748b" text-anchor="end">8 Ay</text>
        <text x="535" y="20" font-size="10" font-weight="700" fill="#334155" text-anchor="middle">24.10.2026</text>
        <rect x="575" y="7" width="70" height="18" rx="9" fill="#fef3c7"/>
        <text x="610" y="20" font-size="9" font-weight="800" fill="#b45309" text-anchor="middle">● YAKLAŞAN</text>
      </g>

      <!-- Satır 4 -->
      <g transform="translate(16, 190)">
        <rect width="653" height="32" fill="#ffffff"/>
        <text x="12" y="20" font-size="11" font-weight="700" fill="#0f172a">Akbank</text>
        <text x="125" y="20" font-size="10" font-weight="500" fill="#475569">İhracat Destek Kredisi</text>
        <text x="290" y="20" font-size="11" font-weight="800" fill="#0f172a" text-anchor="end">₺3.700.000</text>
        <text x="390" y="20" font-size="11" font-weight="800" fill="#dc2626" text-anchor="end">₺188.000</text>
        <text x="460" y="20" font-size="10" font-weight="600" fill="#64748b" text-anchor="end">16 Ay</text>
        <text x="535" y="20" font-size="10" font-weight="700" fill="#334155" text-anchor="middle">28.10.2026</text>
        <rect x="575" y="7" width="70" height="18" rx="9" fill="#dbeafe"/>
        <text x="610" y="20" font-size="9" font-weight="800" fill="#1d4ed8" text-anchor="middle">★ AVANTAJ</text>
      </g>

      <!-- Toplam Çubuğu -->
      <g transform="translate(16, 226)">
        <rect width="653" height="30" rx="4" fill="#0f172a"/>
        <text x="12" y="19" font-size="10.5" font-weight="800" fill="#ffffff">TOPLAM KONSOLİDE</text>
        <text x="290" y="19" font-size="11.5" font-weight="900" fill="#38bdf8" text-anchor="end">₺28.450.000</text>
        <text x="390" y="19" font-size="11.5" font-weight="900" fill="#f87171" text-anchor="end">₺1.428.600</text>
        <text x="570" y="19" font-size="9.5" font-weight="800" fill="#4ade80" text-anchor="middle">DSCR: 1.84x (GÜVENLİ)</text>
      </g>
    </g>
  </g>

  <!-- 7. ALT KISIM: 3 ADET KOKPİT KARTI -->
  <g transform="translate(20, 546)">
    <!-- SOL: DONUT GRAFİĞİ -->
    <g transform="translate(0, 0)">
      <rect width="360" height="145" rx="10" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <text x="16" y="22" font-size="11.5" font-weight="800" fill="#0f172a">BANKA BAZLI PORTFÖY PAYI</text>
      
      <g transform="translate(70, 82)">
        <circle cx="0" cy="0" r="44" fill="none" stroke="#e2e8f0" stroke-width="16"/>
        <circle cx="0" cy="0" r="44" fill="none" stroke="#2563eb" stroke-width="16" stroke-dasharray="116 276" stroke-dashoffset="0"/>
        <circle cx="0" cy="0" r="44" fill="none" stroke="#107c41" stroke-width="16" stroke-dasharray="83 276" stroke-dashoffset="-116"/>
        <circle cx="0" cy="0" r="44" fill="none" stroke="#f59e0b" stroke-width="16" stroke-dasharray="41 276" stroke-dashoffset="-199"/>
        <circle cx="0" cy="0" r="44" fill="none" stroke="#8b5cf6" stroke-width="16" stroke-dasharray="36 276" stroke-dashoffset="-240"/>
        <text x="0" y="4" font-size="10.5" font-weight="900" fill="#0f172a" text-anchor="middle">₺28.4M</text>
      </g>

      <g transform="translate(150, 42)">
        <rect x="0" y="0" width="9" height="9" rx="2" fill="#2563eb"/><text x="15" y="8" font-size="9.5" font-weight="600" fill="#334155">İş Bankası (%42)</text>
        <rect x="0" y="22" width="9" height="9" rx="2" fill="#107c41"/><text x="15" y="30" font-size="9.5" font-weight="600" fill="#334155">Garanti BBVA (%30)</text>
        <rect x="0" y="44" width="9" height="9" rx="2" fill="#f59e0b"/><text x="15" y="52" font-size="9.5" font-weight="600" fill="#334155">Yapı Kredi (%15)</text>
        <rect x="0" y="66" width="9" height="9" rx="2" fill="#8b5cf6"/><text x="15" y="74" font-size="9.5" font-weight="600" fill="#334155">Akbank (%13)</text>
      </g>
    </g>

    <!-- ORTA: ERKEN UYARI VE COVENANT KONTROLÜ -->
    <g transform="translate(375, 0)">
      <rect width="405" height="145" rx="10" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.2" filter="url(#dropShadow)"/>
      <text x="16" y="22" font-size="11.5" font-weight="800" fill="#0f172a">COVENANT VE RİSK ERKEN UYARI RADARI</text>

      <g transform="translate(16, 36)">
        <text x="0" y="12" font-size="10" font-weight="700" fill="#334155">Borç / Özkaynak Oranı</text>
        <text x="370" y="12" font-size="10.5" font-weight="800" fill="#15803d" text-anchor="end">1.42x (Eşik: 2.50x)</text>
        <rect x="0" y="18" width="370" height="7" rx="3.5" fill="#f1f5f9"/>
        <rect x="0" y="18" width="210" height="7" rx="3.5" fill="#16a34a"/>
      </g>

      <g transform="translate(16, 70)">
        <text x="0" y="12" font-size="10" font-weight="700" fill="#334155">Faiz Karşılama Oranı (ICR)</text>
        <text x="370" y="12" font-size="10.5" font-weight="800" fill="#15803d" text-anchor="end">4.80x (Eşik: 2.00x)</text>
        <rect x="0" y="18" width="370" height="7" rx="3.5" fill="#f1f5f9"/>
        <rect x="0" y="18" width="300" height="7" rx="3.5" fill="#16a34a"/>
      </g>

      <g transform="translate(16, 106)">
        <rect width="370" height="26" rx="6" fill="#ecfdf5" stroke="#a7f3d0" stroke-width="1"/>
        <text x="185" y="17" font-size="9.5" font-weight="800" fill="#047857" text-anchor="middle">✓ TÜM PROTOKOLLER VE KRİTİK GÖSTERGELER GÜVENLİ BÖLGEDE</text>
      </g>
    </g>

    <!-- SAĞ: OTOMATİK KARAR MOTORU TAVSİYESİ -->
    <g transform="translate(795, 0)">
      <rect width="365" height="145" rx="10" fill="url(#heroRibbon)" filter="url(#dropShadow)"/>
      <rect x="14" y="12" width="100" height="20" rx="10" fill="#ffffff" fill-opacity="0.18"/>
      <text x="64" y="25" font-size="9" font-weight="900" fill="#ffffff" text-anchor="middle">KARAR MOTORU</text>

      <text x="16" y="52" font-size="13" font-weight="900" fill="#ffffff">RE-FİNANSMAN FIRSATI TESPİT EDİLDİ</text>
      <text x="16" y="72" font-size="10.5" font-weight="500" fill="#dcfce7">İş Bankası kredisinde 150 baz puan faiz indirimi ile yıllık <tspan font-weight="800" fill="#ffffff">₺314.000 net nakit tasarrufu</tspan> elde edilebilir.</text>
      
      <rect x="16" y="96" width="333" height="32" rx="6" fill="#ffffff"/>
      <text x="182" y="116" font-size="10.5" font-weight="900" fill="#0f766e" text-anchor="middle">SENARYOYU ÇALIŞTIR &amp; RAPORU AL ›</text>
    </g>
  </g>

  <!-- 8. EXCEL ÇALIŞMA SAYFASI SEKMELERİ -->
  <rect y="705" width="1200" height="45" fill="#e2e8f0" stroke="#cbd5e1" stroke-width="1"/>
  <g transform="translate(15, 715)">
    <text x="5" y="18" font-size="13" font-weight="700" fill="#64748b">◀  ▶</text>

    <rect x="45" y="0" width="85" height="35" fill="#f8fafc" rx="4"/>
    <text x="87" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">KAPAK</text>

    <rect x="135" y="0" width="115" height="35" fill="#f8fafc" rx="4"/>
    <text x="192" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">BANKA_TANIM</text>

    <rect x="255" y="0" width="95" height="35" fill="#f8fafc" rx="4"/>
    <text x="302" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">KREDİLER</text>

    <rect x="355" y="0" width="110" height="35" fill="#f8fafc" rx="4"/>
    <text x="410" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">TAKSİT_PLANI</text>

    <!-- Aktif Sekme -->
    <rect x="470" y="-3" width="145" height="38" fill="#ffffff" rx="4" stroke="#107c41" stroke-width="2"/>
    <rect x="470" y="32" width="145" height="3" fill="#ffffff"/>
    <text x="542" y="19" font-size="10.5" font-weight="900" fill="#107c41" text-anchor="middle">PANO (DASHBOARD)</text>

    <rect x="620" y="0" width="115" height="35" fill="#f8fafc" rx="4"/>
    <text x="677" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">KARAR_MOTORU</text>

    <rect x="740" y="0" width="110" height="35" fill="#f8fafc" rx="4"/>
    <text x="795" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">NAKİT_AKIŞI</text>

    <rect x="855" y="0" width="80" height="35" fill="#f8fafc" rx="4"/>
    <text x="895" y="18" font-size="10" font-weight="600" fill="#475569" text-anchor="middle">RAPOR</text>

    <circle cx="950" cy="17" r="11" fill="#f1f5f9" stroke="#cbd5e1" stroke-width="1"/>
    <text x="950" y="21" font-size="14" font-weight="700" fill="#64748b" text-anchor="middle">+</text>

    <text x="1150" y="18" font-size="10" font-weight="600" fill="#64748b" text-anchor="end">HAZIR · 100% · HESAPLAMA: OTOMATİK</text>
  </g>
</svg>
'''

out_path = '/Users/macair1/projects/excel-arsiv/public/images/kapak/banka-kredi-ve-taksit-takip-sistemi.svg'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(svg_content)

print('Success')
