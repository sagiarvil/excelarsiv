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

function generateTestCover(slug, title, category, kpis, tableRows, chartSvg) {
  const safeTitle = escapeXml(title);
  const safeCat = escapeXml((category || 'FİNANSAL YÖNETİM').toUpperCase());

  // Sol Tablo Satırları (x: 70..550 -> 480px genişlik)
  // Kolon 1 (Parametre): x=85, genişlik ~180px
  // Kolon 2 (Tutar): x=310 (text-anchor="end")
  // Kolon 3 (Pay): x=380 (text-anchor="end")
  // Kolon 4 (Durum): x=405..535 (130px rozet)
  let rowsSvg = '';
  tableRows.forEach((r, idx) => {
    const y = 375 + idx * 48;
    const bg = idx % 2 === 0 ? '#ffffff' : '#f8fafc';
    const statusColor = r.statusType === 'good' ? '#107c41' : r.statusType === 'bad' ? '#dc2626' : '#2563eb';
    const statusBg = r.statusType === 'good' ? '#dcfce7' : r.statusType === 'bad' ? '#fee2e2' : '#dbeafe';

    rowsSvg += `
      <rect x="70" y="${y}" width="480" height="42" rx="8" fill="${bg}" stroke="#f1f5f9" stroke-width="1" />
      <text x="85" y="${y + 26}" font-family="system-ui, -apple-system, sans-serif" font-size="11" font-weight="600" fill="#334155">${escapeXml(r.col1)}</text>
      <text x="310" y="${y + 26}" font-family="system-ui, -apple-system, sans-serif" font-size="12" font-weight="700" fill="#0f172a" text-anchor="end">${escapeXml(r.col2)}</text>
      <text x="380" y="${y + 26}" font-family="system-ui, -apple-system, sans-serif" font-size="11" font-weight="600" fill="#64748b" text-anchor="end">${escapeXml(r.col3)}</text>
      
      <rect x="405" y="${y + 10}" width="130" height="22" rx="11" fill="${statusBg}" />
      <text x="470" y="${y + 25}" font-family="system-ui, -apple-system, sans-serif" font-size="10" font-weight="800" fill="${statusColor}" text-anchor="middle">${escapeXml(r.status)}</text>
    `;
  });

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

    <rect x="215" y="-14" width="160" height="28" rx="6" fill="#f1f5f9" stroke="#cbd5e1" stroke-width="1" />
    <text x="295" y="4" font-family="system-ui" font-size="10" font-weight="700" fill="#475569" text-anchor="middle" letter-spacing="0.5">${safeCat}</text>

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
  <rect x="70" y="325" width="480" height="38" rx="8" fill="#1e293b" />
  <text x="85" y="349" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff">PARAMETRE / KALEM</text>
  <text x="310" y="349" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="end">TUTAR / DEĞER</text>
  <text x="380" y="349" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="end">PAY (%)</text>
  <text x="470" y="349" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="middle">DURUM</text>

  <!-- Sol Tablo Satırları -->
  ${rowsSvg}

  <!-- Sağ Grafik (x: 580, y: 325, w: 350, h: 480) -->
  <g transform="translate(580, 325)">
    ${chartSvg}
  </g>

  <!-- E) Alt Güvenlik ve Denetim Çubuğu -->
  <g transform="translate(70, 920)">
    <line x1="0" y1="0" x2="860" y2="0" stroke="#e2e8f0" stroke-width="1.5" />
    <circle cx="10" cy="18" r="4" fill="#107c41" />
    <text x="22" y="22" font-family="system-ui" font-size="11" font-weight="700" fill="#0d5c3a">Tam Otomatik Özet Pano &amp; Karar Raporu</text>
    <text x="860" y="22" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b" text-anchor="end">Formüller Açık · Dilediğiniz Gibi Özelleştirilebilir</text>
  </g>

</svg>`;
}

const chart1 = `
  <rect width="350" height="480" rx="16" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
  <text x="25" y="38" font-family="system-ui" font-size="14" font-weight="800" fill="#0f172a">KÜMÜLATİF NAKİT AKIŞI</text>
  <line x1="25" y1="410" x2="325" y2="410" stroke="#cbd5e1" stroke-width="2" />
  
  <rect x="35" y="290" width="35" height="120" rx="4" fill="#1e293b" /><text x="52" y="430" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H1</text>
  <rect x="85" y="230" width="35" height="60" rx="4" fill="#107c41" /><text x="102" y="430" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H2</text>
  <rect x="135" y="290" width="35" height="70" rx="4" fill="#dc2626" /><text x="152" y="430" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H3</text>
  <rect x="185" y="180" width="35" height="110" rx="4" fill="#107c41" /><text x="202" y="430" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H4</text>
  <rect x="235" y="240" width="35" height="60" rx="4" fill="#dc2626" /><text x="252" y="430" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b" text-anchor="middle">H5</text>
  <rect x="285" y="100" width="40" height="310" rx="6" fill="#107c41" /><text x="305" y="430" font-family="system-ui" font-size="10" font-weight="800" fill="#107c41" text-anchor="middle">TOPLAM</text>
  
  <path d="M 52 290 L 102 230 L 152 290 L 202 180 L 252 240 L 305 100" fill="none" stroke="#2563eb" stroke-width="3" stroke-linecap="round" />
  <circle cx="305" cy="100" r="6" fill="#2563eb" />
  
  <rect x="150" y="22" width="180" height="38" rx="8" fill="#dcfce7" />
  <text x="240" y="47" font-family="system-ui" font-size="14" font-weight="900" fill="#0d5c3a" text-anchor="middle">+₺3.850.000 NET</text>
`;

const svg = generateTestCover(
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
  chart1
);

async function run() {
  await sharp(Buffer.from(svg)).png().toFile('/tmp/test_cockpit_fixed.png');
  console.log('Fixed render completed!');
}
run();
