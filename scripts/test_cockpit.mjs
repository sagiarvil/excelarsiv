import sharp from 'sharp';
import fs from 'fs';

function escapeXml(unsafe) {
  if (!unsafe) return '';
  return String(unsafe)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function generateTestCover(slug, title, category, kpis, tableRows, chartType) {
  const safeTitle = escapeXml(title);
  const safeCat = escapeXml((category || 'FİNANSAL YÖNETİM').toUpperCase());

  // Sol Tablo Satırları
  let rowsSvg = '';
  tableRows.forEach((r, idx) => {
    const y = 370 + idx * 46;
    const bg = idx % 2 === 0 ? '#ffffff' : '#f8fafc';
    const statusColor = r.statusType === 'good' ? '#107c41' : r.statusType === 'bad' ? '#dc2626' : '#2563eb';
    const statusBg = r.statusType === 'good' ? '#dcfce7' : r.statusType === 'bad' ? '#fee2e2' : '#dbeafe';

    rowsSvg += `
      <rect x="70" y="${y}" width="460" height="40" rx="6" fill="${bg}" />
      <text x="85" y="${y + 25}" font-family="system-ui, -apple-system, sans-serif" font-size="12" font-weight="600" fill="#334155">${escapeXml(r.col1)}</text>
      <text x="260" y="${y + 25}" font-family="system-ui, -apple-system, sans-serif" font-size="12" font-weight="700" fill="#0f172a" text-anchor="end">${escapeXml(r.col2)}</text>
      <text x="360" y="${y + 25}" font-family="system-ui, -apple-system, sans-serif" font-size="12" font-weight="600" fill="#64748b" text-anchor="end">${escapeXml(r.col3)}</text>
      
      <rect x="390" y="${y + 9}" width="125" height="22" rx="11" fill="${statusBg}" />
      <text x="452" y="${y + 24}" font-family="system-ui, -apple-system, sans-serif" font-size="10" font-weight="800" fill="${statusColor}" text-anchor="middle">${escapeXml(r.status)}</text>
    `;
  });

  // Sağ Grafik Gövdesi
  let chartSvg = '';
  if (chartType === 'waterfall') {
    chartSvg = `
      <g transform="translate(560, 340)">
        <rect width="370" height="420" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
        <text x="25" y="35" font-family="system-ui" font-size="13" font-weight="800" fill="#0f172a">KÜMÜLATİF NAKİT AKIŞI</text>
        <line x1="25" y1="360" x2="345" y2="360" stroke="#cbd5e1" stroke-width="2" />
        
        <!-- Barlar -->
        <rect x="40" y="240" width="35" height="120" rx="4" fill="#1e293b" /><text x="57" y="380" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H1</text>
        <rect x="90" y="190" width="35" height="50" rx="4" fill="#107c41" /><text x="107" y="380" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H2</text>
        <rect x="140" y="240" width="35" height="60" rx="4" fill="#dc2626" /><text x="157" y="380" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H3</text>
        <rect x="190" y="150" width="35" height="90" rx="4" fill="#107c41" /><text x="207" y="380" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H4</text>
        <rect x="240" y="220" width="35" height="70" rx="4" fill="#dc2626" /><text x="257" y="380" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H5</text>
        <rect x="290" y="80" width="40" height="280" rx="6" fill="#107c41" /><text x="310" y="380" font-family="system-ui" font-size="10" font-weight="800" fill="#107c41" text-anchor="middle">TOPLAM</text>
        
        <!-- Eğri -->
        <path d="M 57 240 L 107 190 L 157 240 L 207 150 L 257 220 L 310 80" fill="none" stroke="#2563eb" stroke-width="3" stroke-linecap="round" />
        <circle cx="310" cy="80" r="6" fill="#2563eb" />
        
        <rect x="170" y="25" width="175" height="36" rx="8" fill="#dcfce7" />
        <text x="257" y="48" font-family="system-ui" font-size="14" font-weight="900" fill="#0d5c3a" text-anchor="middle">+₺3.850.000 NET</text>
      </g>
    `;
  } else if (chartType === 'gauge') {
    chartSvg = `
      <g transform="translate(560, 340)">
        <rect width="370" height="420" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
        <text x="25" y="35" font-family="system-ui" font-size="13" font-weight="800" fill="#0f172a">DENETİM VE UYUM İBRESİ</text>
        
        <circle cx="185" cy="220" r="120" fill="none" stroke="#e2e8f0" stroke-width="24" stroke-dasharray="377" stroke-dashoffset="94" transform="rotate(135 185 220)" />
        <circle cx="185" cy="220" r="120" fill="none" stroke="#107c41" stroke-width="24" stroke-dasharray="377" stroke-dashoffset="140" transform="rotate(135 185 220)" stroke-linecap="round" />
        
        <text x="185" y="210" font-family="system-ui" font-size="36" font-weight="900" fill="#0f172a" text-anchor="middle">100 / 100</text>
        <text x="185" y="240" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41" text-anchor="middle">TAM YASAL GÜVENCE</text>
        
        <rect x="40" y="320" width="290" height="65" rx="12" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.5" />
        <text x="60" y="348" font-family="system-ui" font-size="12" font-weight="700" fill="#475569">Mevzuat Uyumluluk:</text>
        <text x="60" y="370" font-family="system-ui" font-size="14" font-weight="900" fill="#0d5c3a">%100 VUK &amp; TTK Uyumlu</text>
      </g>
    `;
  }

  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" width="100%" height="100%">
  <defs>
    <filter id="softCard" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="12" stdDeviation="16" flood-color="#0f172a" flood-opacity="0.07" />
    </filter>
  </defs>

  <!-- Saf Beyaz Zemin -->
  <rect width="1000" height="1000" fill="#ffffff" />

  <!-- Genişletilmiş ve Dengeli Kart (%92 Doluluk: 920x920) -->
  <rect x="40" y="40" width="920" height="920" rx="24" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" filter="url(#softCard)" />

  <!-- A) Üst Sistem Çubuğu -->
  <g transform="translate(70, 68)">
    <circle cx="0" cy="0" r="6" fill="#ef4444" />
    <circle cx="18" cy="0" r="6" fill="#f59e0b" />
    <circle cx="36" cy="0" r="6" fill="#107c41" />

    <rect x="65" y="-14" width="140" height="28" rx="6" fill="#107c41" />
    <text x="135" y="4" font-family="system-ui" font-size="10" font-weight="800" fill="#ffffff" text-anchor="middle" letter-spacing="0.5">EXCEL PRO ENGINE</text>

    <rect x="215" y="-14" width="140" height="28" rx="6" fill="#f1f5f9" stroke="#cbd5e1" stroke-width="1" />
    <text x="285" y="4" font-family="system-ui" font-size="10" font-weight="700" fill="#475569" text-anchor="middle" letter-spacing="0.5">${safeCat}</text>

    <text x="860" y="4" font-family="system-ui" font-size="11" font-weight="700" fill="#94a3b8" text-anchor="end">DECISION OS v15.1 · LOKAL ÇALIŞIR</text>
  </g>

  <!-- B) Başlık ve Formül Çubuğu -->
  <g transform="translate(70, 130)">
    <text x="0" y="0" font-family="system-ui, -apple-system, sans-serif" font-size="25" font-weight="900" fill="#0f172a">${safeTitle}</text>
    <rect x="0" y="16" width="860" height="34" rx="6" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
    <text x="14" y="38" font-family="system-ui, monospace" font-size="12" font-weight="800" fill="#107c41">fx</text>
    <line x1="36" y1="22" x2="36" y2="44" stroke="#cbd5e1" stroke-width="1.5" />
    <text x="46" y="38" font-family="system-ui, monospace" font-size="11" font-weight="600" fill="#475569">=KARAR_OPTIMIZASYONU(VeriGirisleri!A2:E50; Senaryo=&quot;Dinamik&quot;; Eşik=Otomatik)</text>
  </g>

  <!-- C) 3'lü Üst KPI Vitrin Kartları -->
  <g transform="translate(70, 205)">
    <!-- KPI 1 -->
    <rect x="0" y="0" width="270" height="85" rx="12" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
    <text x="20" y="28" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b">${escapeXml(kpis[0].label)}</text>
    <text x="20" y="62" font-family="system-ui" font-size="24" font-weight="900" fill="#0f172a">${escapeXml(kpis[0].value)}</text>

    <!-- KPI 2 -->
    <rect x="295" y="0" width="270" height="85" rx="12" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
    <text x="315" y="28" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b">${escapeXml(kpis[1].label)}</text>
    <text x="315" y="62" font-family="system-ui" font-size="24" font-weight="900" fill="#2563eb">${escapeXml(kpis[1].value)}</text>

    <!-- KPI 3 -->
    <rect x="590" y="0" width="270" height="85" rx="12" fill="#dcfce7" stroke="#107c41" stroke-width="2" />
    <text x="610" y="28" font-family="system-ui" font-size="11" font-weight="800" fill="#0d5c3a">${escapeXml(kpis[2].label)}</text>
    <text x="610" y="62" font-family="system-ui" font-size="24" font-weight="900" fill="#0d5c3a">${escapeXml(kpis[2].value)}</text>
  </g>

  <!-- D) Ana Gövde: Sol Analitik Tablo + Sağ Dinamik Görsel -->
  <!-- Sol Tablo Başlığı -->
  <rect x="70" y="325" width="460" height="38" rx="8" fill="#1e293b" />
  <text x="85" y="349" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff">PARAMETRE / KALEM</text>
  <text x="260" y="349" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="end">TUTAR / DEĞER</text>
  <text x="360" y="349" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="end">PAY (%)</text>
  <text x="452" y="349" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="middle">DURUM</text>

  <!-- Sol Tablo Satırları -->
  ${rowsSvg}

  <!-- Sağ Grafik -->
  ${chartSvg}

  <!-- E) Alt Güvenlik ve Denetim Çubuğu -->
  <g transform="translate(70, 920)">
    <line x1="0" y1="0" x2="860" y2="0" stroke="#e2e8f0" stroke-width="1.5" />
    <circle cx="10" cy="18" r="4" fill="#107c41" />
    <text x="22" y="22" font-family="system-ui" font-size="11" font-weight="700" fill="#0d5c3a">Tam Otomatik Özet Pano &amp; Karar Raporu</text>
    <text x="860" y="22" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b" text-anchor="end">Formüller Açık · Dilediğiniz Gibi Özelleştirilebilir</text>
  </g>

</svg>`;
}

// Test 1: 13-haftalik-nakit-akisi
const svg1 = generateTestCover(
  '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi',
  '13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi',
  'Nakit Akışı Yönetimi',
  [
    { label: 'Başlangıç Nakit Stoku', value: '₺1.450.000' },
    { label: 'Haftalık Ortalama Tahsilat', value: '₺840.000' },
    { label: '13. Hafta Net Projeksiyon', value: '+₺3.850.000' }
  ],
  [
    { col1: 'Hafta 1-4: Cari Tahsilatlar', col2: '₺3.360.000', col3: '%32.4', status: '✓ GÜVENLİ', statusType: 'good' },
    { col1: 'Hafta 5-8: Maaş & SGK Çıkışı', col2: '-₺1.850.000', col3: '%17.8', status: 'REZERV AYRILDI', statusType: 'normal' },
    { col1: 'Hafta 9: Kritik İthalat Akreditif', col2: '-₺940.000', col3: '%9.1', status: 'ÖDEME ONAYLI', statusType: 'normal' },
    { col1: 'Hafta 10-12: Çek / Senet Giriş', col2: '+₺2.480.000', col3: '%23.9', status: '✓ GÜVENLİ', statusType: 'good' },
    { col1: 'Hafta 13: KDV & Kurumlar V.', col2: '-₺720.000', col3: '%6.9', status: 'KARŞILIK OK', statusType: 'good' },
    { col1: '13 Hafta Sonu Serbest Nakit', col2: '₺3.850.000', col3: '%100', status: '★ KORUMALI', statusType: 'good' }
  ],
  'waterfall'
);

// Test 2: YMM Tasdik Robotu
const svg2 = generateTestCover(
  'ymm-tasdik-kontrol-robotu',
  'YMM Tasdik Kontrol Robotu',
  'Denetim & Vergi Savunma',
  [
    { label: 'İncelenen Yasal Eşik Sayısı', value: '142 Madde' },
    { label: 'Riskli / Anomali Sayısı', value: '0 Hata' },
    { label: 'Nihai Tasdik Skoru', value: '100 / 100' }
  ],
  [
    { col1: 'VUK 298/Ç Yeniden Değerleme', col2: '₺28.400.000', col3: '%100', status: '✓ TAM UYUMLU', statusType: 'good' },
    { col1: 'KVK 10/1 İndirim ve İstisnalar', col2: '₺4.250.000', col3: '%100', status: '✓ DOĞRULANDI', statusType: 'good' },
    { col1: 'KDV İadesi Yüklenim Dağıtımı', col2: '₺2.850.000', col3: '%100', status: '✓ KANIT HAZIR', statusType: 'good' },
    { col1: 'Transfer Fiyatlandırması Raporu', col2: '₺12.400.000', col3: '%100', status: '✓ EŞİK ALTINDA', statusType: 'good' },
    { col1: 'Örtülü Sermaye 3 Kat Sınırı', col2: '0.8x Emsal', col3: '%100', status: '✓ GÜVENLİ', statusType: 'good' },
    { col1: 'Net Kurumlar Vergisi Matrahı', col2: '₺18.900.000', col3: '%100', status: '★ ONAYLANDI', statusType: 'good' }
  ],
  'gauge'
);

async function render() {
  await sharp(Buffer.from(svg1)).png().toFile('/tmp/test_cockpit1.png');
  await sharp(Buffer.from(svg2)).png().toFile('/tmp/test_cockpit2.png');
  console.log('SUCCESS: Renders completed without errors!');
}
render();
