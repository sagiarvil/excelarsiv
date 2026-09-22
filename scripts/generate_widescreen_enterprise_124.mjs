import fs from 'fs';
import path from 'path';

const templatesDir = 'src/content/templates';
const outputDir = 'public/images/kapak';

if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

// 1. Parse all 124 template MDX files
const mdxFiles = fs.readdirSync(templatesDir).filter(f => f.endsWith('.mdx'));
console.log(`Bulunan MDX şablon sayısı: ${mdxFiles.length}`);

function parseMdx(filePath) {
  const content = fs.readFileSync(filePath, 'utf8');
  const lines = content.split('\n');
  let inFm = false;
  let name = '', summary = '', category = '', sheetCount = 18;
  let sheets = [];
  let inSheetMap = false;

  for (const line of lines) {
    if (line.trim() === '---') {
      if (!inFm) { inFm = true; continue; }
      else break;
    }
    if (!inFm) continue;
    if (line.startsWith('name:')) name = line.replace('name:', '').trim().replace(/^['"]|['"]$/g, '');
    if (line.startsWith('summary:')) summary = line.replace('summary:', '').trim().replace(/^['"]|['"]$/g, '');
    if (line.startsWith('category:')) category = line.replace('category:', '').trim().replace(/^['"]|['"]$/g, '');
    if (line.startsWith('sheetCount:')) sheetCount = parseInt(line.replace('sheetCount:', '').trim(), 10) || 18;
    if (line.startsWith('sheetMap:')) { inSheetMap = true; continue; }
    if (inSheetMap) {
      if (line.trim().startsWith('- name:')) {
        sheets.push(line.replace('- name:', '').trim().replace(/^['"]|['"]$/g, ''));
      } else if (!line.startsWith(' ') && !line.startsWith('\t')) {
        inSheetMap = false;
      }
    }
  }

  const slug = path.basename(filePath, '.mdx');
  return { slug, name, summary, category, sheetCount, sheets };
}

const templates = mdxFiles.map(f => parseMdx(path.join(templatesDir, f)));

function escapeXml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Deterministic Pseudo-Random Generator
function mulberry32(a) {
  return function() {
    let t = a += 0x6D2B79F5;
    t = Math.imul(t ^ t >>> 15, t | 1);
    t ^= t + Math.imul(t ^ t >>> 7, t | 61);
    return ((t ^ t >>> 14) >>> 0) / 4294967296;
  };
}

function getRng(slug) {
  let seed = 0;
  for (let i = 0; i < slug.length; i++) {
    seed = (seed * 31 + slug.charCodeAt(i)) >>> 0;
  }
  return mulberry32(seed);
}

// 10 Distinct Professional Color Palettes
const PALETTES = [
  { primary: '#0d9488', dark: '#115e59', light: '#f0fdfa', accent: '#0f766e', border: '#99f6e4', badge: '#ccfbf1', textDark: '#134e4a' }, // Teal
  { primary: '#1d4ed8', dark: '#1e3a8a', light: '#eff6ff', accent: '#2563eb', border: '#bfdbfe', badge: '#dbeafe', textDark: '#1e3a8a' }, // Royal Blue
  { primary: '#059669', dark: '#064e3b', light: '#ecfdf5', accent: '#10b981', border: '#a7f3d0', badge: '#d1fae5', textDark: '#064e3b' }, // Emerald
  { primary: '#334155', dark: '#0f172a', light: '#f8fafc', accent: '#475569', border: '#cbd5e1', badge: '#e2e8f0', textDark: '#0f172a' }, // Slate
  { primary: '#4338ca', dark: '#312e81', light: '#eef2ff', accent: '#4f46e5', border: '#c7d2fe', badge: '#e0e7ff', textDark: '#312e81' }, // Indigo
  { primary: '#b91c1c', dark: '#7f1d1d', light: '#fef2f2', accent: '#dc2626', border: '#fecaca', badge: '#fee2e2', textDark: '#7f1d1d' }, // Crimson
  { primary: '#b45309', dark: '#78350f', light: '#fffbeb', accent: '#d97706', border: '#fde68a', badge: '#fef3c7', textDark: '#78350f' }, // Amber
  { primary: '#0e7490', dark: '#164e63', light: '#ecfeff', accent: '#06b6d4', border: '#a5f3fc', badge: '#cffafe', textDark: '#164e63' }, // Cyan
  { primary: '#6d28d9', dark: '#4c1d95', light: '#f5f3ff', accent: '#7c3aed', border: '#ddd6fe', badge: '#ede9fe', textDark: '#4c1d95' }, // Violet
  { primary: '#be123c', dark: '#881337', light: '#fff1f2', accent: '#e11d48', border: '#fecdd3', badge: '#ffe4e6', textDark: '#881337' }  // Rose
];

console.log(`124 şablon için 16:10 Enterprise Excel Dashboard üretimi başlatılıyor...`);

let generatedCount = 0;

for (let i = 0; i < templates.length; i++) {
  const t = templates[i];
  const rng = getRng(t.slug);
  const pal = PALETTES[i % PALETTES.length];
  
  // Ürünün gerçek alanına göre özgün görsel mimarisi seçimi
  let archIndex = 0;
  const s = t.slug;
  if (s.includes('nakit') || s.includes('kasa') || s.includes('tahsilat') || s.includes('odeme') || s.includes('cek') || s.includes('senet') || s.includes('banka')) {
    archIndex = 1; // Waterfall Nakit Akışı & Likidite
  } else if (s.includes('vergi') || s.includes('vuk') || s.includes('sgk') || s.includes('kvk') || s.includes('degerleme') || s.includes('stopaj') || s.includes('kkeg') || s.includes('mahsup') || s.includes('ymm') || s.includes('beyan')) {
    archIndex = 2; // Çift Kadran Mevzuat & Vergi Kalkanı
  } else if (s.includes('recete') || s.includes('maliyet') || s.includes('bom') || s.includes('fire') || s.includes('uretim') || s.includes('lot') || s.includes('fason') || s.includes('stok') || s.includes('navlun')) {
    archIndex = 3; // 100% Yığılmış Reçete & Maliyet
  } else if (s.includes('yatirim') || s.includes('fizibilite') || s.includes('gayrimenkul') || s.includes('ges') || s.includes('otel') || s.includes('arsa') || s.includes('dcf') || s.includes('irr') || s.includes('wacc') || s.includes('patron')) {
    archIndex = 4; // S-Curve Sigmoid Payback & Fizibilite
  } else {
    archIndex = 0; // 5-Bölmeli Standart Yönetici Kokpiti (Kullanıcı örnek görseli gibi)
  }

  // Clean sheet tabs from product MDX
  const sheetNames = t.sheets && t.sheets.length > 0 
    ? t.sheets.slice(0, 6).map(s => s.replace(/_/g, ' ').toUpperCase())
    : ['PANO', 'VERİ GİRİŞİ', 'HESAPLAMA', 'KONTROLLER', 'SENARYOLAR', 'RAPOR'];

  // Ensure active tab
  const activeTab = sheetNames[0] || 'PANO';

  // 1. Dynamic Table 1
  const t1_rows = [
    { name: 'Temel Kapasite / Hacim', val: `${Math.floor(10 + rng() * 90)}.000 Birim` },
    { name: 'Operasyonel Bütçe', val: `₺${(1.2 + rng() * 8.5).toFixed(2)}M` },
    { name: 'Hedeflenen Kâr Marjı', val: `%${Math.floor(25 + rng() * 45)}.0` },
    { name: 'Mevzuat & Standart Uyum', val: 'Tam Uyumlu' }
  ];

  // 2. Dynamic Table 2
  const t2_rows = [
    { name: 'Duyarlılık Eşiği', val: `±%${Math.floor(5 + rng() * 15)} Tolerans` },
    { name: 'Kritik Risk Puanı', val: `${Math.floor(1 + rng() * 4)}/100 (Düşük)` },
    { name: 'Geri Dönüş (Payback)', val: `${(8 + rng() * 18).toFixed(1)} Ay` },
    { name: 'Denetim İmzası', val: 'Doğrulandı' }
  ];

  // 3. Dynamic Table 3
  const t3_rows = [
    { name: 'Yönetim Tavsiyesi', val: 'Onaylandı' },
    { name: 'Vergi Kalkanı Avantajı', val: `₺${Math.floor(150 + rng() * 850)}.000` },
    { name: 'Nakit Rezerv Oranı', val: `%${Math.floor(12 + rng() * 25)} Likit` },
    { name: 'Nihai Karar Durumu', val: '★ POZİTİF GÜVENLİ' }
  ];

  // Upper Table Rows Render
  function renderUpperTable(x, title, rows) {
    let rowsSvg = '';
    for (let r = 0; r < rows.length; r++) {
      const yLine = 24 + r * 18;
      const bg = r % 2 === 0 ? '#ffffff' : '#f8fafc';
      rowsSvg += `
      <rect x="0" y="${yLine}" width="280" height="18" fill="${bg}" />
      <line x1="0" y1="${yLine}" x2="280" y2="${yLine}" stroke="#e2e8f0" stroke-width="1" />
      <text x="12" y="${yLine + 13}" font-family="system-ui" font-size="10" font-weight="600" fill="#475569">${escapeXml(rows[r].name)}</text>
      <text x="268" y="${yLine + 13}" font-family="system-ui" font-size="10" font-weight="700" fill="#0f172a" text-anchor="end">${escapeXml(rows[r].val)}</text>
      `;
    }
    return `
    <g transform="translate(${x}, 0)">
      <rect width="280" height="98" rx="6" fill="#ffffff" stroke="#cbd5e1" stroke-width="1" />
      <rect width="280" height="24" rx="6" fill="${pal.light}" />
      <text x="12" y="16" font-family="system-ui" font-size="11" font-weight="800" fill="${pal.dark}">${escapeXml(title)}</text>
      ${rowsSvg}
    </g>
    `;
  }

  // 4. GENERATE DISTINCT MIDDLE COCKPIT BASED ON ARCHETYPE
  let middleSvg = '';

  if (archIndex === 0) {
    // 5-Bölmeli Standart Yönetici Kokpiti (Kullanıcının tam örnek görseli gibi)
    const kpiVal = `₺${(2.5 + rng() * 6).toFixed(2)}M`;
    const pctVal = Math.floor(75 + rng() * 22);
    const dashOffset = Math.round(408 * (1 - pctVal / 100));

    middleSvg = `
    <!-- Sütun 1: KPI & Donut -->
    <g transform="translate(0, 0)">
      <rect width="230" height="420" fill="#ffffff" stroke="#e2e8f0" stroke-width="1" />
      <rect width="230" height="55" fill="${pal.primary}" />
      <text x="115" y="25" font-family="system-ui" font-size="12" font-weight="800" fill="#ffffff" text-anchor="middle">TOPLAM HACİM</text>
      <text x="115" y="42" font-family="system-ui" font-size="11" font-weight="600" fill="${pal.badge}" text-anchor="middle">PORTFÖY LİKİDİTESİ</text>
      <text x="115" y="135" font-family="system-ui" font-size="34" font-weight="900" fill="#0f172a" text-anchor="middle">${kpiVal}</text>
      
      <rect y="180" width="230" height="40" fill="${pal.primary}" fill-opacity="0.9" />
      <text x="115" y="205" font-family="system-ui" font-size="12" font-weight="800" fill="#ffffff" text-anchor="middle">HEDEF REALİZASYONU</text>
      
      <g transform="translate(115, 310)">
        <circle cx="0" cy="0" r="65" fill="none" stroke="#e2e8f0" stroke-width="14" />
        <circle cx="0" cy="0" r="65" fill="none" stroke="${pal.primary}" stroke-width="14" stroke-dasharray="408" stroke-dashoffset="${dashOffset}" stroke-linecap="round" transform="rotate(-90)" />
        <text x="0" y="8" font-family="system-ui" font-size="22" font-weight="900" fill="#0f172a" text-anchor="middle">%${pctVal}.0</text>
        <text x="0" y="28" font-family="system-ui" font-size="10" font-weight="700" fill="${pal.primary}" text-anchor="middle">UYUMLULUK OK</text>
      </g>
    </g>

    <!-- Sütun 2: Yatay Risk Çubukları -->
    <g transform="translate(250, 20)">
      <text x="0" y="20" font-family="system-ui" font-size="12" font-weight="800" fill="#334155">RİSK &amp; ÖNCELİK DAĞILIMI</text>
      <text x="0" y="70" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Düşük Risk</text>
      <rect x="90" y="55" width="130" height="22" rx="3" fill="${pal.primary}" />
      <text x="230" y="71" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a">₺2.4M</text>

      <text x="0" y="140" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Orta Risk</text>
      <rect x="90" y="125" width="85" height="22" rx="3" fill="${pal.primary}" />
      <text x="185" y="141" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a">₺1.1M</text>

      <text x="0" y="210" font-family="system-ui" font-size="12" font-weight="600" fill="#64748b">Yüksek Risk</text>
      <rect x="90" y="195" width="45" height="22" rx="3" fill="#f59e0b" />
      <text x="145" y="211" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a">₺450K</text>

      <text x="0" y="280" font-family="system-ui" font-size="12" font-weight="700" fill="#ef4444">Kritik Eşik</text>
      <rect x="90" y="265" width="20" height="22" rx="3" fill="#ef4444" />
      <text x="120" y="281" font-family="system-ui" font-size="12" font-weight="800" fill="#ef4444">₺120K</text>
    </g>

    <!-- Sütun 3: Dikey Kolon Grafiği -->
    <g transform="translate(560, 20)">
      <line x1="-20" y1="-20" x2="-20" y2="400" stroke="#e2e8f0" stroke-width="1" />
      <text x="0" y="20" font-family="system-ui" font-size="12" font-weight="800" fill="#334155">PERİYODİK ANALİZ DAĞILIMI</text>
      
      <rect x="25" y="80" width="45" height="250" rx="3" fill="${pal.primary}" />
      <text x="47" y="65" font-family="system-ui" font-size="13" font-weight="900" fill="#0f172a" text-anchor="middle">14</text>
      <text x="47" y="350" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b" text-anchor="middle">1. ÇEYREK</text>

      <rect x="95" y="150" width="45" height="180" rx="3" fill="${pal.primary}" />
      <text x="117" y="135" font-family="system-ui" font-size="13" font-weight="900" fill="#0f172a" text-anchor="middle">9</text>
      <text x="117" y="350" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b" text-anchor="middle">2. ÇEYREK</text>

      <rect x="165" y="220" width="45" height="110" rx="3" fill="${pal.accent}" />
      <text x="187" y="205" font-family="system-ui" font-size="13" font-weight="900" fill="#0f172a" text-anchor="middle">5</text>
      <text x="187" y="350" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b" text-anchor="middle">3. ÇEYREK</text>
    </g>

    <!-- Sütun 4: Süreç Çubukları -->
    <g transform="translate(820, 20)">
      <line x1="-20" y1="-20" x2="-20" y2="400" stroke="#e2e8f0" stroke-width="1" />
      <text x="0" y="20" font-family="system-ui" font-size="12" font-weight="800" fill="#334155">SÜREÇ İLERLEME KONTROLÜ</text>

      <text x="0" y="65" font-family="system-ui" font-size="11" fill="#64748b">Beklemede</text>
      <rect x="90" y="52" width="25" height="18" rx="3" fill="${pal.primary}" fill-opacity="0.6" />
      <text x="125" y="66" font-family="system-ui" font-size="11" font-weight="800" fill="#334155">2</text>

      <text x="0" y="125" font-family="system-ui" font-size="11" fill="#64748b">İcrada</text>
      <rect x="90" y="112" width="130" height="18" rx="3" fill="${pal.primary}" />
      <text x="230" y="126" font-family="system-ui" font-size="11" font-weight="800" fill="#334155">9</text>

      <text x="0" y="185" font-family="system-ui" font-size="11" fill="#64748b">Onaylandı</text>
      <rect x="90" y="172" width="60" height="18" rx="3" fill="#f59e0b" />
      <text x="160" y="186" font-family="system-ui" font-size="11" font-weight="800" fill="#334155">4</text>

      <text x="0" y="245" font-family="system-ui" font-size="11" fill="#64748b">Askıda</text>
      <rect x="90" y="232" width="15" height="18" rx="3" fill="#94a3b8" />
      <text x="115" y="246" font-family="system-ui" font-size="11" font-weight="800" fill="#334155">1</text>

      <text x="0" y="305" font-family="system-ui" font-size="11" font-weight="700" fill="#10b981">Tamamlandı</text>
      <rect x="90" y="292" width="110" height="18" rx="3" fill="#10b981" />
      <text x="210" y="306" font-family="system-ui" font-size="11" font-weight="800" fill="#10b981">7</text>
    </g>

    <!-- Sütun 5: Detay Kalem Listesi -->
    <g transform="translate(1110, 20)">
      <line x1="-20" y1="-20" x2="-20" y2="400" stroke="#e2e8f0" stroke-width="1" />
      <text x="0" y="20" font-family="system-ui" font-size="12" font-weight="800" fill="#334155">KALEM DETAY DAĞILIMI</text>
      
      <g transform="translate(0, 45)">
        <text x="0" y="15" font-family="system-ui" font-size="10" fill="#475569">Ana Faaliyet Girişleri</text>
        <rect x="135" y="5" width="190" height="12" rx="2" fill="${pal.primary}" />
        <text x="335" y="15" font-family="system-ui" font-size="10" font-weight="800" fill="#0f172a">14</text>

        <text x="0" y="45" font-family="system-ui" font-size="10" fill="#475569">Sabit &amp; Kira Rezervleri</text>
        <rect x="135" y="35" width="120" height="12" rx="2" fill="${pal.primary}" />
        <text x="265" y="45" font-family="system-ui" font-size="10" font-weight="800" fill="#0f172a">8</text>

        <text x="0" y="75" font-family="system-ui" font-size="10" fill="#475569">Bordro &amp; SGK Karşılığı</text>
        <rect x="135" y="65" width="165" height="12" rx="2" fill="${pal.primary}" />
        <text x="310" y="75" font-family="system-ui" font-size="10" font-weight="800" fill="#0f172a">11</text>

        <text x="0" y="105" font-family="system-ui" font-size="10" fill="#475569">Vergi &amp; Tevkifat Mahsubu</text>
        <rect x="135" y="95" width="95" height="12" rx="2" fill="${pal.primary}" />
        <text x="240" y="105" font-family="system-ui" font-size="10" font-weight="800" fill="#0f172a">6</text>

        <text x="0" y="135" font-family="system-ui" font-size="10" fill="#475569">Tedarikçi &amp; Mal Alımı</text>
        <rect x="135" y="125" width="140" height="12" rx="2" fill="${pal.primary}" />
        <text x="285" y="135" font-family="system-ui" font-size="10" font-weight="800" fill="#0f172a">9</text>

        <text x="0" y="165" font-family="system-ui" font-size="10" fill="#475569">Finansman &amp; Kredi Taksiti</text>
        <rect x="135" y="155" width="80" height="12" rx="2" fill="${pal.primary}" />
        <text x="225" y="165" font-family="system-ui" font-size="10" font-weight="800" fill="#0f172a">5</text>

        <text x="0" y="195" font-family="system-ui" font-size="10" fill="#475569">Diğer İşletme Giderleri</text>
        <rect x="135" y="185" width="50" height="12" rx="2" fill="${pal.primary}" />
        <text x="195" y="195" font-family="system-ui" font-size="10" font-weight="800" fill="#0f172a">3</text>
      </g>
    </g>
    `;
  } else if (archIndex === 1) {
    // Waterfall Nakit Akışı ve Kümülatif Likidite Mimarisi
    middleSvg = `
    <!-- Sol Analitik Tablo -->
    <g transform="translate(30, 30)">
      <rect width="640" height="360" rx="8" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
      <rect width="640" height="36" rx="8" fill="#1e293b" />
      <text x="20" y="23" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff">KALEM / OPERASYONEL AKIŞ</text>
      <text x="360" y="23" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="end">TUTAR (TL)</text>
      <text x="470" y="23" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="end">PAY (%)</text>
      <text x="565" y="23" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="middle">DURUM</text>
      
      <!-- Satır 1 -->
      <text x="20" y="70" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">Banka &amp; Cari Tahsilat</text>
      <text x="360" y="70" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a" text-anchor="end">₺4.850.000</text>
      <text x="470" y="70" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b" text-anchor="end">%52.4</text>
      <rect x="510" y="52" width="110" height="24" rx="12" fill="#dcfce7" /><text x="565" y="68" font-family="system-ui" font-size="10" font-weight="800" fill="#107c41" text-anchor="middle">✓ TAHSİL OK</text>
      <line x1="20" y1="88" x2="620" y2="88" stroke="#f1f5f9" stroke-width="1" />

      <!-- Satır 2 -->
      <text x="20" y="120" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">Personel Maaş &amp; SGK Yükü</text>
      <text x="360" y="120" font-family="system-ui" font-size="12" font-weight="800" fill="#dc2626" text-anchor="end">-₺1.420.000</text>
      <text x="470" y="120" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b" text-anchor="end">%15.3</text>
      <rect x="510" y="102" width="110" height="24" rx="12" fill="#fee2e2" /><text x="565" y="118" font-family="system-ui" font-size="10" font-weight="800" fill="#dc2626" text-anchor="middle">BLOKE EDİLDİ</text>
      <line x1="20" y1="138" x2="620" y2="138" stroke="#f1f5f9" stroke-width="1" />

      <!-- Satır 3 -->
      <text x="20" y="170" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">Tedarikçi Çek &amp; Fatura Çıkışı</text>
      <text x="360" y="170" font-family="system-ui" font-size="12" font-weight="800" fill="#dc2626" text-anchor="end">-₺980.000</text>
      <text x="470" y="170" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b" text-anchor="end">%10.6</text>
      <rect x="510" y="152" width="110" height="24" rx="12" fill="#dbeafe" /><text x="565" y="168" font-family="system-ui" font-size="10" font-weight="800" fill="#2563eb" text-anchor="middle">ONAYLANDI</text>
      <line x1="20" y1="188" x2="620" y2="188" stroke="#f1f5f9" stroke-width="1" />

      <!-- Satır 4 -->
      <text x="20" y="220" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">Vergi &amp; KDV Mahsup Karşılığı</text>
      <text x="360" y="220" font-family="system-ui" font-size="12" font-weight="800" fill="#dc2626" text-anchor="end">-₺450.000</text>
      <text x="470" y="220" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b" text-anchor="end">%4.9</text>
      <rect x="510" y="202" width="110" height="24" rx="12" fill="#dcfce7" /><text x="565" y="218" font-family="system-ui" font-size="10" font-weight="800" fill="#107c41" text-anchor="middle">✓ KORUNDU</text>
      <line x1="20" y1="238" x2="620" y2="238" stroke="#f1f5f9" stroke-width="1" />

      <!-- Satır 5 -->
      <text x="20" y="270" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">Sabit Gider &amp; Tesis Rezervi</text>
      <text x="360" y="270" font-family="system-ui" font-size="12" font-weight="800" fill="#dc2626" text-anchor="end">-₺320.000</text>
      <text x="470" y="270" font-family="system-ui" font-size="11" font-weight="600" fill="#64748b" text-anchor="end">%3.5</text>
      <rect x="510" y="252" width="110" height="24" rx="12" fill="#fef3c7" /><text x="565" y="268" font-family="system-ui" font-size="10" font-weight="800" fill="#d97706" text-anchor="middle">AYRILDI</text>
      <line x1="20" y1="288" x2="620" y2="288" stroke="#f1f5f9" stroke-width="1" />

      <!-- Satır 6 -->
      <text x="20" y="325" font-family="system-ui" font-size="12" font-weight="900" fill="#0f172a">NET SERBEST LİKİDİTE</text>
      <text x="360" y="325" font-family="system-ui" font-size="14" font-weight="900" fill="#107c41" text-anchor="end">+₺1.680.000</text>
      <text x="470" y="325" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41" text-anchor="end">%100</text>
      <rect x="510" y="307" width="110" height="26" rx="13" fill="#107c41" /><text x="565" y="324" font-family="system-ui" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">★ GÜVENLİ</text>
    </g>

    <!-- Sağ Waterfall Grafiği -->
    <g transform="translate(710, 30)">
      <rect width="750" height="360" rx="12" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="30" y="35" font-family="system-ui" font-size="13" font-weight="800" fill="#0f172a">KÜMÜLATİF DÖNGÜ VE WATERFALL NAKİT BASAMAKLARI</text>
      <rect x="560" y="16" width="160" height="28" rx="6" fill="${pal.badge}" />
      <text x="640" y="35" font-family="system-ui" font-size="11" font-weight="800" fill="${pal.dark}" text-anchor="middle">+₺1.68M NET FARK</text>
      
      <line x1="40" y1="310" x2="710" y2="310" stroke="#cbd5e1" stroke-width="2" />

      <!-- Basamak 1: Giriş -->
      <rect x="60" y="110" width="70" height="200" rx="4" fill="#107c41" />
      <text x="95" y="95" font-family="system-ui" font-size="11" font-weight="800" fill="#107c41" text-anchor="middle">+4.85M</text>
      <text x="95" y="330" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b" text-anchor="middle">GİRİŞ</text>

      <!-- Basamak 2: Maaş -->
      <rect x="170" y="110" width="70" height="70" rx="4" fill="#dc2626" />
      <text x="205" y="95" font-family="system-ui" font-size="11" font-weight="800" fill="#dc2626" text-anchor="middle">-1.42M</text>
      <text x="205" y="330" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b" text-anchor="middle">MAAŞ</text>

      <!-- Basamak 3: Tedarik -->
      <rect x="280" y="180" width="70" height="50" rx="4" fill="#dc2626" />
      <text x="315" y="165" font-family="system-ui" font-size="11" font-weight="800" fill="#dc2626" text-anchor="middle">-980K</text>
      <text x="315" y="330" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b" text-anchor="middle">TEDARİK</text>

      <!-- Basamak 4: Vergi -->
      <rect x="390" y="230" width="70" height="28" rx="4" fill="#dc2626" />
      <text x="425" y="215" font-family="system-ui" font-size="11" font-weight="800" fill="#dc2626" text-anchor="middle">-450K</text>
      <text x="425" y="330" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b" text-anchor="middle">VERGİ</text>

      <!-- Basamak 5: Sabit -->
      <rect x="500" y="258" width="70" height="20" rx="4" fill="#f59e0b" />
      <text x="535" y="245" font-family="system-ui" font-size="11" font-weight="800" fill="#d97706" text-anchor="middle">-320K</text>
      <text x="535" y="330" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b" text-anchor="middle">SABİT</text>

      <!-- Basamak 6: Net Bakiye -->
      <rect x="610" y="225" width="80" height="85" rx="4" fill="#1e293b" />
      <text x="650" y="210" font-family="system-ui" font-size="12" font-weight="900" fill="#0f172a" text-anchor="middle">+1.68M</text>
      <text x="650" y="330" font-family="system-ui" font-size="11" font-weight="800" fill="#1e293b" text-anchor="middle">BAKİYE</text>

      <!-- Likidite Eğrisi -->
      <path d="M 95 110 L 205 180 L 315 230 L 425 258 L 535 278 L 650 225" fill="none" stroke="${pal.primary}" stroke-width="4" stroke-dasharray="6 4" />
      <circle cx="650" cy="225" r="7" fill="${pal.primary}" />
    </g>
    `;
  } else if (archIndex === 2) {
    // Çift Kadranlı Gösterge ve Mevzuat Uyumluluk Radarı (Ekran Resmi 20.42.55 referansı)
    middleSvg = `
    <!-- Sol Kadran Paneli -->
    <g transform="translate(40, 20)">
      <rect width="460" height="380" rx="12" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="30" y="35" font-family="system-ui" font-size="13" font-weight="800" fill="#0f172a">DENETİM &amp; MEVZUAT UYGUNLUK SKORU</text>
      
      <!-- Çift Kadran -->
      <g transform="translate(230, 190)">
        <circle cx="0" cy="0" r="110" fill="none" stroke="#e2e8f0" stroke-width="22" />
        <circle cx="0" cy="0" r="110" fill="none" stroke="${pal.primary}" stroke-width="22" stroke-dasharray="690" stroke-dashoffset="138" stroke-linecap="round" transform="rotate(-90)" />

        <circle cx="0" cy="0" r="75" fill="none" stroke="#f1f5f9" stroke-width="18" />
        <circle cx="0" cy="0" r="75" fill="none" stroke="#10b981" stroke-width="18" stroke-dasharray="471" stroke-dashoffset="47" stroke-linecap="round" transform="rotate(-90)" />

        <text x="0" y="10" font-family="system-ui" font-size="34" font-weight="900" fill="#0f172a" text-anchor="middle">%96</text>
        <text x="0" y="32" font-family="system-ui" font-size="11" font-weight="800" fill="#10b981" text-anchor="middle">TAM ONAYLI</text>
      </g>
      
      <g transform="translate(30, 335)">
        <circle cx="10" cy="10" r="6" fill="${pal.primary}" /><text x="25" y="14" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">VUK &amp; KVK Uyum (%80)</text>
        <circle cx="240" cy="10" r="6" fill="#10b981" /><text x="255" y="14" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">SGK Teşvik Payı (%90)</text>
      </g>
    </g>

    <!-- Sağ Denetim Matrisi Tablosu -->
    <g transform="translate(540, 20)">
      <rect width="910" height="380" rx="12" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
      <rect width="910" height="40" rx="12" fill="#1e293b" />
      <text x="25" y="25" font-family="system-ui" font-size="12" font-weight="800" fill="#ffffff">DENETİM KONTROL MADDELERİ</text>
      <text x="450" y="25" font-family="system-ui" font-size="12" font-weight="800" fill="#ffffff">YASAL DAYANAK</text>
      <text x="680" y="25" font-family="system-ui" font-size="12" font-weight="800" fill="#ffffff">TASARRUF TUTARI</text>
      <text x="835" y="25" font-family="system-ui" font-size="12" font-weight="800" fill="#ffffff" text-anchor="middle">DURUM</text>

      <!-- 5 Maddelik Denetim Listesi -->
      <g transform="translate(25, 65)">
        <text x="0" y="20" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">1. Geçici Vergi ve Matrah Mahsubu</text>
        <text x="425" y="20" font-family="system-ui" font-size="11" fill="#64748b">VUK Md. 120</text>
        <text x="655" y="20" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41">₺1.420.000</text>
        <rect x="760" y="2" width="100" height="24" rx="12" fill="#dcfce7" /><text x="810" y="18" font-family="system-ui" font-size="10" font-weight="800" fill="#107c41" text-anchor="middle">✓ TAM UYGUN</text>
        <line x1="0" y1="36" x2="860" y2="36" stroke="#f1f5f9" stroke-width="1" />

        <text x="0" y="75" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">2. Yeniden Değerleme Vergi Kalkanı</text>
        <text x="425" y="75" font-family="system-ui" font-size="11" fill="#64748b">VUK Geçici Md. 32</text>
        <text x="655" y="75" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41">₺840.000</text>
        <rect x="760" y="57" width="100" height="24" rx="12" fill="#dcfce7" /><text x="810" y="73" font-family="system-ui" font-size="10" font-weight="800" fill="#107c41" text-anchor="middle">✓ HESAPLANDI</text>
        <line x1="0" y1="91" x2="860" y2="91" stroke="#f1f5f9" stroke-width="1" />

        <text x="0" y="130" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">3. SGK Prim Teşvik Optimizasyonu</text>
        <text x="425" y="130" font-family="system-ui" font-size="11" fill="#64748b">5510 Sayılı Kanun</text>
        <text x="655" y="130" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41">₺380.000</text>
        <rect x="760" y="112" width="100" height="24" rx="12" fill="#dcfce7" /><text x="810" y="128" font-family="system-ui" font-size="10" font-weight="800" fill="#107c41" text-anchor="middle">✓ BORDRODA</text>
        <line x1="0" y1="146" x2="860" y2="146" stroke="#f1f5f9" stroke-width="1" />

        <text x="0" y="185" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">4. KKEG &amp; Finansman Gider Kısıtlaması</text>
        <text x="425" y="185" font-family="system-ui" font-size="11" fill="#64748b">KVK Md. 11/1-i</text>
        <text x="655" y="185" font-family="system-ui" font-size="12" font-weight="800" fill="#2563eb">₺210.000</text>
        <rect x="760" y="167" width="100" height="24" rx="12" fill="#dbeafe" /><text x="810" y="183" font-family="system-ui" font-size="10" font-weight="800" fill="#2563eb" text-anchor="middle">EŞİK KORUNDU</text>
        <line x1="0" y1="201" x2="860" y2="201" stroke="#f1f5f9" stroke-width="1" />

        <text x="0" y="240" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">5. KDV İade ve Mahsup Dosyası</text>
        <text x="425" y="240" font-family="system-ui" font-size="11" fill="#64748b">KDVK Md. 29/2</text>
        <text x="655" y="240" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41">₺950.000</text>
        <rect x="760" y="222" width="100" height="24" rx="12" fill="#107c41" /><text x="810" y="238" font-family="system-ui" font-size="10" font-weight="900" fill="#ffffff" text-anchor="middle">★ ONAYLANDI</text>
      </g>
    </g>
    `;
  } else if (archIndex === 3) {
    // 100% Yığılmış Çubuklar ve Kaynak Tahsis Matrisi (Ekran Resmi 20.44.02 referansı)
    middleSvg = `
    <!-- Sol 3 Yığın Çubuğu -->
    <g transform="translate(30, 20)">
      <rect width="680" height="380" rx="12" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="30" y="35" font-family="system-ui" font-size="13" font-weight="800" fill="#0f172a">BİRİM MALİYET &amp; KAYNAK DAĞILIMI (100% STACKED)</text>

      <!-- Yığın 1: 2024 -->
      <g transform="translate(80, 60)">
        <rect x="0" y="120" width="80" height="70" rx="3" fill="#cbd5e1" />
        <rect x="0" y="65" width="80" height="55" rx="3" fill="${pal.accent}" />
        <rect x="0" y="0" width="80" height="65" rx="3" fill="${pal.primary}" />
        <text x="40" y="220" font-family="system-ui" font-size="11" font-weight="800" fill="#64748b" text-anchor="middle">2024 FİİLİ</text>
        <text x="40" y="40" font-family="system-ui" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">%36</text>
        <text x="40" y="95" font-family="system-ui" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">%28</text>
        <text x="40" y="160" font-family="system-ui" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">%36</text>
      </g>

      <!-- Yığın 2: 2025 -->
      <g transform="translate(260, 60)">
        <rect x="0" y="100" width="80" height="90" rx="3" fill="#cbd5e1" />
        <rect x="0" y="45" width="80" height="55" rx="3" fill="${pal.accent}" />
        <rect x="0" y="0" width="80" height="45" rx="3" fill="${pal.primary}" />
        <text x="40" y="220" font-family="system-ui" font-size="11" font-weight="800" fill="#64748b" text-anchor="middle">2025 FİİLİ</text>
        <text x="40" y="30" font-family="system-ui" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">%24</text>
        <text x="40" y="75" font-family="system-ui" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">%28</text>
        <text x="40" y="150" font-family="system-ui" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">%48</text>
      </g>

      <!-- Yığın 3: 2026 PRO HEDEF -->
      <g transform="translate(440, 60)">
        <rect x="0" y="70" width="80" height="120" rx="3" fill="#cbd5e1" />
        <rect x="0" y="25" width="80" height="45" rx="3" fill="${pal.accent}" />
        <rect x="0" y="0" width="80" height="25" rx="3" fill="${pal.primary}" />
        <text x="40" y="220" font-family="system-ui" font-size="11" font-weight="900" fill="${pal.primary}" text-anchor="middle">2026 HEDEF</text>
        <text x="40" y="18" font-family="system-ui" font-size="10" font-weight="900" fill="#ffffff" text-anchor="middle">%12</text>
        <text x="40" y="52" font-family="system-ui" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">%24</text>
        <text x="40" y="135" font-family="system-ui" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">%64</text>
      </g>

      <!-- Açıklama -->
      <g transform="translate(50, 340)">
        <circle cx="10" cy="10" r="5" fill="${pal.primary}" /><text x="25" y="14" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">Hammadde (%12)</text>
        <circle cx="210" cy="10" r="5" fill="${pal.accent}" /><text x="225" y="14" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">İşçilik &amp; Enerji (%24)</text>
        <circle cx="430" cy="10" r="5" fill="#cbd5e1" /><text x="445" y="14" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">Net Katma Değer (%64)</text>
      </g>
    </g>

    <!-- Sağ Birim Reçete Tablosu -->
    <g transform="translate(740, 20)">
      <rect width="710" height="380" rx="12" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
      <rect width="710" height="40" rx="12" fill="#1e293b" />
      <text x="25" y="25" font-family="system-ui" font-size="12" font-weight="800" fill="#ffffff">REÇETE &amp; MALİYET PARAMETRELERİ</text>
      <text x="380" y="25" font-family="system-ui" font-size="12" font-weight="800" fill="#ffffff" text-anchor="end">BİRİM GİRDİ</text>
      <text x="520" y="25" font-family="system-ui" font-size="12" font-weight="800" fill="#ffffff" text-anchor="end">TOPLAM PAY</text>
      <text x="645" y="25" font-family="system-ui" font-size="12" font-weight="800" fill="#ffffff" text-anchor="middle">AKSİYON</text>

      <g transform="translate(25, 60)">
        <text x="0" y="25" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">Doğrudan Hammadde Girdisi</text>
        <text x="355" y="25" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a" text-anchor="end">₺94.50</text>
        <text x="495" y="25" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b" text-anchor="end">%48.2</text>
        <rect x="580" y="8" width="95" height="24" rx="12" fill="#dcfce7" /><text x="627" y="24" font-family="system-ui" font-size="10" font-weight="800" fill="#107c41" text-anchor="middle">✓ LİSTELENDİ</text>
        <line x1="0" y1="45" x2="660" y2="45" stroke="#f1f5f9" stroke-width="1" />

        <text x="0" y="80" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">İşçilik &amp; Adam/Saat Dağıtımı</text>
        <text x="355" y="80" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a" text-anchor="end">₺42.80</text>
        <text x="495" y="80" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b" text-anchor="end">%21.8</text>
        <rect x="580" y="63" width="95" height="24" rx="12" fill="#dbeafe" /><text x="627" y="79" font-family="system-ui" font-size="10" font-weight="800" fill="#2563eb" text-anchor="middle">OPTİMİZE</text>
        <line x1="0" y1="100" x2="660" y2="100" stroke="#f1f5f9" stroke-width="1" />

        <text x="0" y="135" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">Genel Üretim / Elektrik / Gaz</text>
        <text x="355" y="135" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a" text-anchor="end">₺24.20</text>
        <text x="495" y="135" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b" text-anchor="end">%12.3</text>
        <rect x="580" y="118" width="95" height="24" rx="12" fill="#fef3c7" /><text x="627" y="134" font-family="system-ui" font-size="10" font-weight="800" fill="#d97706" text-anchor="middle">ÖLÇÜLDÜ</text>
        <line x1="0" y1="155" x2="660" y2="155" stroke="#f1f5f9" stroke-width="1" />

        <text x="0" y="190" font-family="system-ui" font-size="11" font-weight="700" fill="#334155">Lojistik, Ambalaj &amp; Sevk</text>
        <text x="355" y="190" font-family="system-ui" font-size="12" font-weight="800" fill="#0f172a" text-anchor="end">₺18.50</text>
        <text x="495" y="190" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b" text-anchor="end">%9.4</text>
        <rect x="580" y="173" width="95" height="24" rx="12" fill="#dcfce7" /><text x="627" y="189" font-family="system-ui" font-size="10" font-weight="800" fill="#107c41" text-anchor="middle">✓ KORUNDU</text>
        <line x1="0" y1="210" x2="660" y2="210" stroke="#f1f5f9" stroke-width="1" />

        <text x="0" y="245" font-family="system-ui" font-size="12" font-weight="900" fill="#0f172a">ÖNERİLEN SATIŞ FİYATI</text>
        <text x="355" y="245" font-family="system-ui" font-size="14" font-weight="900" fill="#107c41" text-anchor="end">₺310.00</text>
        <text x="495" y="245" font-family="system-ui" font-size="12" font-weight="800" fill="#107c41" text-anchor="end">%100</text>
        <rect x="580" y="228" width="95" height="26" rx="13" fill="#107c41" /><text x="627" y="245" font-family="system-ui" font-size="10" font-weight="900" fill="#ffffff" text-anchor="middle">★ HEDEF KÂR</text>
      </g>
    </g>
    `;
  } else {
    // S-Curve Sigmoid Payback ve Yatırım Fizibilitesi
    const irrVal = `%${(28 + rng() * 25).toFixed(1)} IRR`;
    const paybackVal = `${Math.floor(10 + rng() * 12)} Ay`;

    middleSvg = `
    <!-- Sol S-Curve Yatırım Grafiği -->
    <g transform="translate(30, 20)">
      <rect width="800" height="380" rx="12" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
      <text x="30" y="35" font-family="system-ui" font-size="13" font-weight="800" fill="#0f172a">KÜMÜLATİF NAKİT DÖNGÜSÜ &amp; S-CURVE BAŞABAŞ GRAFİĞİ</text>
      
      <rect x="640" y="16" width="130" height="28" rx="6" fill="${pal.badge}" />
      <text x="705" y="35" font-family="system-ui" font-size="11" font-weight="800" fill="${pal.dark}" text-anchor="middle">${irrVal}</text>

      <line x1="50" y1="310" x2="750" y2="310" stroke="#cbd5e1" stroke-width="2" />
      <line x1="50" y1="180" x2="750" y2="180" stroke="#e2e8f0" stroke-width="1.5" stroke-dasharray="6 6" />
      <text x="740" y="172" font-family="system-ui" font-size="11" font-weight="800" fill="#059669" text-anchor="end">BAŞABAŞ EŞİĞİ</text>

      <!-- S-Curve Yolu -->
      <path d="M 60 295 Q 220 285 360 180 T 720 70" fill="none" stroke="${pal.primary}" stroke-width="5" stroke-linecap="round" />
      <circle cx="720" cy="70" r="8" fill="${pal.primary}" />
      <circle cx="360" cy="180" r="7" fill="#2563eb" />

      <!-- Başabaş Kartı -->
      <g transform="translate(380, 195)">
        <rect width="160" height="60" rx="8" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.5" />
        <text x="15" y="24" font-family="system-ui" font-size="10" font-weight="700" fill="#64748b">Geri Dönüş Noktası:</text>
        <text x="15" y="46" font-family="system-ui" font-size="14" font-weight="900" fill="#2563eb">${paybackVal}</text>
      </g>

      <g transform="translate(60, 340)">
        <text x="0" y="14" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b">Kurulum (Capex)</text>
        <text x="240" y="14" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b">Nötr Bakiye</text>
        <text x="500" y="14" font-family="system-ui" font-size="11" font-weight="800" fill="#107c41">Net Kâr Üretim Bölgesi ›</text>
      </g>
    </g>

    <!-- Sağ 3'lü Metrik Kartları -->
    <g transform="translate(860, 20)">
      <rect width="590" height="380" rx="12" fill="#ffffff" stroke="#e2e8f0" stroke-width="1.5" />
      
      <!-- Kart 1 -->
      <g transform="translate(25, 25)">
        <rect width="540" height="95" rx="8" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
        <text x="20" y="32" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b">TOPLAM PROJE YATIRIM BÜTÇESİ</text>
        <text x="20" y="70" font-family="system-ui" font-size="28" font-weight="900" fill="#0f172a">₺${(15 + rng() * 35).toFixed(1)}00.000</text>
        <rect x="420" y="32" width="100" height="30" rx="15" fill="#dcfce7" />
        <text x="470" y="52" font-family="system-ui" font-size="11" font-weight="800" fill="#107c41" text-anchor="middle">FİZİBİL</text>
      </g>

      <!-- Kart 2 -->
      <g transform="translate(25, 140)">
        <rect width="540" height="95" rx="8" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
        <text x="20" y="32" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b">YILLIK ORTALAMA FAVÖK / NAKİT ÜRETİMİ</text>
        <text x="20" y="70" font-family="system-ui" font-size="28" font-weight="900" fill="#2563eb">₺${(4.2 + rng() * 8).toFixed(2)}M</text>
        <rect x="420" y="32" width="100" height="30" rx="15" fill="#dbeafe" />
        <text x="470" y="52" font-family="system-ui" font-size="11" font-weight="800" fill="#2563eb" text-anchor="middle">POZİTİF</text>
      </g>

      <!-- Kart 3 -->
      <g transform="translate(25, 255)">
        <rect width="540" height="95" rx="8" fill="${pal.light}" stroke="${pal.primary}" stroke-width="2" />
        <text x="20" y="32" font-family="system-ui" font-size="11" font-weight="800" fill="${pal.dark}">İÇ VERİM ORANI (IRR) &amp; NET GETİRİ</text>
        <text x="20" y="70" font-family="system-ui" font-size="28" font-weight="900" fill="${pal.dark}">${irrVal}</text>
        <rect x="420" y="32" width="100" height="30" rx="15" fill="${pal.primary}" />
        <text x="470" y="52" font-family="system-ui" font-size="11" font-weight="900" fill="#ffffff" text-anchor="middle">★ YÜKSEK</text>
      </g>
    </g>
    `;
  }

  // 5. SHEET TABS RENDER (Directly from product MDX)
  let tabsSvg = '';
  let curX = 10;
  for (let s = 0; s < sheetNames.length; s++) {
    const sName = sheetNames[s];
    const tabWidth = Math.max(120, sName.length * 9 + 30);
    const isActive = s === 0;
    const bg = isActive ? '#ffffff' : '#e2e8f0';
    const fillTxt = isActive ? pal.primary : '#475569';
    const border = isActive ? `stroke="#cbd5e1" stroke-width="1"` : '';
    const lineActive = isActive ? `<line x1="${curX}" y1="38" x2="${curX + tabWidth}" y2="38" stroke="${pal.primary}" stroke-width="3" />` : '';

    tabsSvg += `
    <rect x="${curX}" y="6" width="${tabWidth}" height="33" rx="4" fill="${bg}" ${border} />
    ${lineActive}
    <text x="${curX + tabWidth / 2}" y="27" font-family="system-ui" font-size="10.5" font-weight="${isActive ? '800' : '700'}" fill="${fillTxt}" text-anchor="middle">${escapeXml(sName)}</text>
    `;
    curX += tabWidth + 10;
  }

  const safeTitle = escapeXml(t.name.toUpperCase());
  const safeCat = escapeXml(t.category.toUpperCase().replace(/-/g, ' '));

  // FULL WIDESCREEN SVG
  const fullSvg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1600 1000" width="100%" height="100%">
  <defs>
    <filter id="cardShadow" x="-5%" y="-5%" width="110%" height="110%">
      <feDropShadow dx="0" dy="6" stdDeviation="12" flood-color="#0f172a" flood-opacity="0.06" />
    </filter>
    <filter id="btnShadow" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="4" stdDeviation="6" flood-color="#ef4444" flood-opacity="0.3" />
    </filter>
  </defs>

  <!-- Arka Plan: Saf Beyaz -->
  <rect width="1600" height="1000" fill="#ffffff" />

  <!-- Ana Çalışma Kitabı Çerçevesi (16:10, 1540x940, 30px kenar boşluğu) -->
  <rect x="30" y="30" width="1540" height="940" rx="16" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.5" filter="url(#cardShadow)" />

  <!-- 1. ÜST BAŞLIK VE ŞERİT (Örnekteki 'SETTINGS AND SUMMARY PAGE' formatında) -->
  <g transform="translate(55, 60)">
    <text x="0" y="20" font-family="system-ui, -apple-system, sans-serif" font-size="18" font-weight="900" fill="#1e293b" letter-spacing="0.5">${safeTitle}</text>
    <text x="1490" y="20" font-family="system-ui" font-size="11" font-weight="700" fill="#94a3b8" text-anchor="end">MODÜL: ${safeCat} · v15.1</text>
    <line x1="0" y1="32" x2="1490" y2="32" stroke="${pal.primary}" stroke-width="3" stroke-linecap="round" />
  </g>

  <!-- 2. ÜST GİRDİ TABLOLARI VE AKSİYON BUTONU (Örnekteki 3'lü Kategori Tablosu + Kırmızı Buton) -->
  <g transform="translate(55, 115)">
    <text x="0" y="35" font-family="system-ui, sans-serif" font-size="12" font-weight="700" fill="#64748b">Parametre Girişi</text>
    <path d="M 95 32 L 115 32 M 110 27 L 116 32 L 110 37" stroke="#94a3b8" stroke-width="2" fill="none" stroke-linecap="round" />

    ${renderUpperTable(135, 'GİRDİ VE TAHMİNLER', t1_rows)}
    ${renderUpperTable(435, 'HESAPLAMA KRİTERLERİ', t2_rows)}
    ${renderUpperTable(735, 'DENETİM & ÇIKTILAR', t3_rows)}

    <!-- Kırmızı 3D Aksiyon Butonu (Örnekteki 'BRAIN DUMP INPUT PAGE' butonu) -->
    <g transform="translate(1270, 10)" filter="url(#btnShadow)">
      <rect width="220" height="75" rx="16" fill="#ef4444" />
      <rect x="2" y="2" width="216" height="36" rx="14" fill="#ffffff" fill-opacity="0.18" />
      <text x="110" y="35" font-family="system-ui, sans-serif" font-size="12" font-weight="900" fill="#ffffff" text-anchor="middle" letter-spacing="0.5">HESAPLAMA &amp; PANO</text>
      <text x="110" y="55" font-family="system-ui, sans-serif" font-size="10" font-weight="700" fill="#fee2e2" text-anchor="middle">FORMÜL MOTORUNA GEÇ ›</text>
    </g>
  </g>

  <!-- 3. ORTA GRAFİK VE KOKPİT BÖLÜMÜ -->
  <g transform="translate(55, 235)">
    <rect width="1490" height="420" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1.5" />
    ${middleSvg}
  </g>

  <!-- 4. ALT ÖZET VE İCRA KARAR ŞERİDİ (Örnekteki Turuncu Başlıklı Alt Bölüm) -->
  <g transform="translate(55, 675)">
    <text x="0" y="20" font-family="system-ui" font-size="13" font-weight="800" fill="#334155">ÖZET VE İCRA RAPORU : ${safeTitle}</text>
    <rect x="0" y="32" width="1490" height="170" fill="#ffffff" stroke="#cbd5e1" stroke-width="1" />
    
    <!-- Turuncu Sol İcra Kartı -->
    <rect x="0" y="32" width="240" height="170" fill="#f97316" />
    <text x="120" y="68" font-family="system-ui" font-size="11" font-weight="800" fill="#ffffff" text-anchor="middle">SİSTEM ANALİZ SKORU</text>
    <text x="120" y="86" font-family="system-ui" font-size="10" fill="#ffedd5" text-anchor="middle">GÜVEN VE DOĞRULUK</text>
    <text x="120" y="145" font-family="system-ui" font-size="36" font-weight="900" fill="#ffffff" text-anchor="middle">100 / 100</text>
    <text x="120" y="175" font-family="system-ui" font-size="11" font-weight="700" fill="#ffffff" text-anchor="middle">✓ TAM DOĞRULANDI</text>

    <!-- Orta Rapor Metinleri -->
    <g transform="translate(270, 55)">
      <text x="0" y="25" font-family="system-ui" font-size="13" font-weight="700" fill="#334155">Matematiksel Model Doğrulaması:</text>
      <text x="280" y="25" font-family="system-ui" font-size="13" font-weight="800" fill="#0d9488">Hatasız (Çapraz Sağlama OK)</text>

      <text x="0" y="65" font-family="system-ui" font-size="13" font-weight="700" fill="#334155">Makro Bağımlılığı:</text>
      <text x="280" y="65" font-family="system-ui" font-size="13" font-weight="800" fill="#0d9488">Sıfır VBA (Windows &amp; Mac Uyumlu)</text>

      <text x="0" y="105" font-family="system-ui" font-size="13" font-weight="700" fill="#334155">Karar Destek Çıktısı:</text>
      <text x="280" y="105" font-family="system-ui" font-size="13" font-weight="800" fill="#2563eb">Yönetim Sunumuna Hazır Dinamik Rapor</text>

      <!-- Sağ İkincil Özet Kutuları -->
      <rect x="620" y="10" width="270" height="105" rx="8" fill="#f8fafc" stroke="#e2e8f0" stroke-width="1" />
      <text x="640" y="35" font-family="system-ui" font-size="11" font-weight="700" fill="#64748b">İŞLETME SERMAYESİ KORUMASI</text>
      <text x="640" y="70" font-family="system-ui" font-size="22" font-weight="900" fill="#0f172a">₺${(1.5 + rng() * 4).toFixed(2)}M</text>
      <text x="640" y="95" font-family="system-ui" font-size="11" font-weight="800" fill="#107c41">✓ RİSK EŞİKLERİ DEVREDE</text>

      <!-- Sağ En Son Turuncu Kutu -->
      <rect x="910" y="10" width="280" height="105" rx="8" fill="#ffedd5" stroke="#fdba74" stroke-width="1.5" />
      <text x="930" y="35" font-family="system-ui" font-size="11" font-weight="800" fill="#c2410c">KURUMSAL KARAR ÖNERİSİ</text>
      <text x="930" y="65" font-family="system-ui" font-size="12" font-weight="800" fill="#9a3412">Haftalık veri girişleriyle simülasyonu</text>
      <text x="930" y="88" font-family="system-ui" font-size="12" font-weight="800" fill="#9a3412">canlı tutup riskleri önceden bertaraf edin.</text>
    </g>
  </g>

  <!-- 5. EN ALT EXCEL SAYFA SEKMELERİ (Gerçek Excel SheetMap Tabs) -->
  <g transform="translate(55, 905)">
    <rect width="1490" height="45" rx="6" fill="#f1f5f9" stroke="#cbd5e1" stroke-width="1" />
    ${tabsSvg}
    <text x="1470" y="28" font-family="system-ui" font-size="11" font-weight="800" fill="#64748b" text-anchor="end">DECISION OS v15.1 · %100 FORMÜLLER AÇIK</text>
  </g>
</svg>`;

  const outPath = path.join(outputDir, `${t.slug}.svg`);
  fs.writeFileSync(outPath, fullSvg, 'utf8');
  generatedCount++;
}

console.log(`BAŞARILI: 124 ürün için 16:10 Enterprise Widescreen Excel Dashboard kapak SVG'si üretildi! (Toplam: ${generatedCount})`);
