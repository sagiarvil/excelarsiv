import fs from 'node:fs';
import path from 'node:path';
import sharp from 'sharp';

const dataFile = path.resolve('excelarsiv-premium-tum-sayfalar/site/assets/data.js');
const siteAssetsDir = path.resolve('excelarsiv-premium-tum-sayfalar/site/assets');
const kaynakAssetsDir = path.resolve('excelarsiv-premium-tum-sayfalar/kaynak/assets');

const raw = fs.readFileSync(dataFile, 'utf8');
const json = raw.replace(/^window\.ExcelArsivData=/, '').replace(/;?\s*$/, '');
const DATA = JSON.parse(json);

// Distinct luxury color palettes for home page featured products
const PALETTES = {
  'kasa-takip-programi': { bg1: '#022c22', bg2: '#064e3b', accent: '#10b981', secondary: '#34d399', glow: 'rgba(16,185,129,0.35)', badge: 'GÜNLÜK KASA &amp; BAKİYE' },
  'stok-takip-ve-depo-yonetim-programi': { bg1: '#1e1b4b', bg2: '#312e81', accent: '#6366f1', secondary: '#a5b4fc', glow: 'rgba(99,102,241,0.35)', badge: 'KRİTİK STOK VE DEPO' },
  'personel-puantaj-ve-maas-hesaplama-tablosu-2026': { bg1: '#4a044e', bg2: '#701a75', accent: '#d946ef', secondary: '#f0abfc', glow: 'rgba(217,70,239,0.35)', badge: '2026 BORDRO &amp; PUANTAJ' },
  'fatura-ve-tahsilat-takip-tablosu': { bg1: '#082f49', bg2: '#0369a1', accent: '#0284c7', secondary: '#38bdf8', glow: 'rgba(2,132,199,0.35)', badge: 'VADE &amp; KDV TAHSİLAT' },
  '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi': { bg1: '#172554', bg2: '#1d4ed8', accent: '#3b82f6', secondary: '#60a5fa', glow: 'rgba(59,130,246,0.35)', badge: '13 HAFTALIK NAKİT' },
  'banka-kredi-ve-taksit-takip-sistemi': { bg1: '#451a03', bg2: '#78350f', accent: '#d97706', secondary: '#fbbf24', glow: 'rgba(217,119,6,0.35)', badge: 'KREDİ &amp; COVENANT' },
  'aylik-patron-finans-paneli': { bg1: '#0f172a', bg2: '#1e293b', accent: '#f59e0b', secondary: '#fbbf24', glow: 'rgba(245,158,11,0.35)', badge: 'C-LEVEL PATRON KOKPİT' },
  'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi': { bg1: '#14532d', bg2: '#166534', accent: '#22c55e', secondary: '#4ade80', glow: 'rgba(34,197,94,0.35)', badge: 'MÜŞTERİ RİSK MATRİSİ' },
  'dinamik-waffle-grafik-sablonu': { bg1: '#1e1b4b', bg2: '#3730a3', accent: '#818cf8', secondary: '#c7d2fe', glow: 'rgba(129,140,248,0.35)', badge: 'KPI WAFFLE HEDEF' },
  'kutup': { bg1: '#134e4a', bg2: '#0f766e', accent: '#14b8a6', secondary: '#5eead4', glow: 'rgba(20,184,166,0.35)', badge: 'POLAR ÇEMBER MODELİ' },
  'hedef-ve-ortalama-ustunu-dinamik-renklendiren-sutun-grafik': { bg1: '#31104b', bg2: '#581c87', accent: '#a855f7', secondary: '#d8b4fe', glow: 'rgba(168,85,247,0.35)', badge: 'DİNAMİK ORTALAMA' },
  'kosula-bagli-renk-degistiren-akilli-cubuk-grafik': { bg1: '#1c1917', bg2: '#292524', accent: '#ef4444', secondary: '#f87171', glow: 'rgba(239,68,68,0.35)', badge: 'AKILLI KOŞULLU GRAFİK' },
  'hedef-gostergeli-cubuk-grafik': { bg1: '#042f2e', bg2: '#115e59', accent: '#0d9488', secondary: '#2dd4bf', glow: 'rgba(13,148,136,0.35)', badge: 'BULLET CHART HEDEF' }
};

function generate3DImageSvg(p) {
  const pal = PALETTES[p.id] || { bg1: '#0f172a', bg2: '#1e293b', accent: '#38bdf8', secondary: '#818cf8', glow: 'rgba(56,189,248,0.35)', badge: 'EXCEL ÇÖZÜMÜ' };
  const title = p.shortTitle || p.title;

  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 600" width="960" height="600">
  <defs>
    <linearGradient id="bg_${p.id}" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="${pal.bg1}"/>
      <stop offset="100%" stop-color="${pal.bg2}"/>
    </linearGradient>

    <linearGradient id="bar_${p.id}" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" stop-color="${pal.accent}"/>
      <stop offset="100%" stop-color="${pal.secondary}"/>
    </linearGradient>

    <linearGradient id="glass_${p.id}" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="rgba(255,255,255,0.18)"/>
      <stop offset="100%" stop-color="rgba(255,255,255,0.03)"/>
    </linearGradient>

    <filter id="glow_${p.id}" x="-30%" y="-30%" width="160%" height="160%">
      <feGaussianBlur stdDeviation="24" result="blur"/>
      <feComposite in="SourceGraphic" in2="blur" operator="over"/>
    </filter>

    <filter id="shadow_${p.id}" x="-30%" y="-30%" width="160%" height="160%">
      <feDropShadow dx="0" dy="24" stdDeviation="22" flood-color="rgba(0,0,0,0.65)"/>
    </filter>
  </defs>

  <!-- Deep 3D Studio Stage -->
  <rect width="960" height="600" fill="url(#bg_${p.id})"/>

  <!-- High-Energy Radial Glow -->
  <circle cx="500" cy="240" r="380" fill="${pal.glow}" filter="url(#glow_${p.id})"/>

  <!-- 3D Perspective Studio Grid Floor -->
  <g opacity="0.22">
    <path d="M 60 480 L 480 240 L 900 480 L 480 620 Z" fill="none" stroke="rgba(255,255,255,0.5)" stroke-width="1.5"/>
    <path d="M 180 480 L 480 300 L 780 480 L 480 570 Z" fill="none" stroke="rgba(255,255,255,0.3)" stroke-width="1"/>
    <line x1="200" y1="380" x2="680" y2="580" stroke="rgba(255,255,255,0.2)" stroke-width="1"/>
    <line x1="320" y1="320" x2="800" y2="520" stroke="rgba(255,255,255,0.2)" stroke-width="1"/>
    <line x1="760" y1="380" x2="280" y2="580" stroke="rgba(255,255,255,0.2)" stroke-width="1"/>
    <line x1="640" y1="320" x2="160" y2="520" stroke="rgba(255,255,255,0.2)" stroke-width="1"/>
  </g>

  <!-- 3D Floating Glass Deck -->
  <g filter="url(#shadow_${p.id})">
    <polygon points="480,260 820,380 480,510 140,380" fill="url(#glass_${p.id})" stroke="rgba(255,255,255,0.4)" stroke-width="2.5"/>
    <polygon points="140,380 480,510 480,535 140,405" fill="rgba(0,0,0,0.6)" stroke="rgba(255,255,255,0.1)"/>
    <polygon points="480,510 820,380 820,405 480,535" fill="rgba(0,0,0,0.35)" stroke="rgba(255,255,255,0.1)"/>
  </g>

  <!-- High-Converting 3D Isometric Bar Pillars with Dynamic Heights -->
  <g transform="translate(260, 240)">
    <polygon points="0,170 38,186 38,16 0,0" fill="url(#bar_${p.id})"/>
    <polygon points="38,186 68,170 68,0 38,16" fill="${pal.secondary}" opacity="0.85"/>
    <polygon points="0,0 30,-16 68,0 38,16" fill="#ffffff" opacity="0.8"/>
  </g>

  <g transform="translate(350, 190)">
    <polygon points="0,220 42,238 42,18 0,0" fill="url(#bar_${p.id})"/>
    <polygon points="42,238 76,220 76,0 42,18" fill="${pal.secondary}" opacity="0.9"/>
    <polygon points="0,0 34,-18 76,0 42,18" fill="#ffffff" opacity="0.9"/>
  </g>

  <g transform="translate(440, 140)">
    <polygon points="0,270 46,290 46,20 0,0" fill="${pal.accent}"/>
    <polygon points="46,290 84,270 84,0 46,20" fill="${pal.secondary}"/>
    <polygon points="0,0 38,-20 84,0 46,20" fill="#ffffff"/>
  </g>

  <g transform="translate(540, 200)">
    <polygon points="0,210 42,228 42,18 0,0" fill="url(#bar_${p.id})"/>
    <polygon points="42,228 76,210 76,0 42,18" fill="${pal.secondary}" opacity="0.85"/>
    <polygon points="0,0 34,-18 76,0 42,18" fill="#ffffff" opacity="0.85"/>
  </g>

  <g transform="translate(630, 160)">
    <polygon points="0,250 40,268 40,18 0,0" fill="url(#bar_${p.id})"/>
    <polygon points="40,268 72,250 72,0 40,18" fill="${pal.secondary}" opacity="0.8"/>
    <polygon points="0,0 32,-18 72,0 40,18" fill="#ffffff" opacity="0.8"/>
  </g>

  <!-- 3D Glowing Trajectory Line with Data Points -->
  <path d="M 290 240 Q 380 180 470 140 T 660 160" fill="none" stroke="#ffffff" stroke-width="4" stroke-linecap="round" filter="url(#glow_${p.id})"/>
  <path d="M 290 240 Q 380 180 470 140 T 660 160" fill="none" stroke="${pal.accent}" stroke-width="3" stroke-linecap="round"/>

  <circle cx="290" cy="240" r="7" fill="#ffffff" filter="url(#glow_${p.id})"/>
  <circle cx="470" cy="140" r="9" fill="#ffffff" filter="url(#glow_${p.id})"/>
  <circle cx="660" cy="160" r="8" fill="#ffffff" filter="url(#glow_${p.id})"/>

  <!-- Floating Glass HUD Command Card -->
  <g transform="translate(160, 60)" filter="url(#shadow_${p.id})">
    <rect width="640" height="135" rx="16" fill="rgba(255,255,255,0.08)" stroke="rgba(255,255,255,0.25)" stroke-width="1.5"/>
    <rect x="2" y="2" width="636" height="4" rx="2" fill="${pal.accent}"/>

    <!-- Category Pill -->
    <rect x="24" y="20" width="180" height="26" rx="13" fill="${pal.glow}" stroke="${pal.accent}" stroke-width="1"/>
    <text x="114" y="37" fill="${pal.accent}" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="12" font-weight="900" text-anchor="middle" letter-spacing="1">${pal.badge}</text>

    <!-- Excel 3D Badge -->
    <g transform="translate(565, 20)">
      <rect width="48" height="32" rx="6" fill="#107c41" filter="url(#shadow_${p.id})"/>
      <text x="24" y="22" fill="#ffffff" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="16" font-weight="900" text-anchor="middle">XLSX</text>
    </g>

    <!-- Title -->
    <text x="24" y="78" fill="#ffffff" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="22" font-weight="900" letter-spacing="-0.4">
      ${title.length > 44 ? title.slice(0, 42) + '…' : title}
    </text>

    <!-- Features -->
    <text x="24" y="105" fill="rgba(255,255,255,0.75)" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="13" font-weight="600">
      3D İnteraktif Gösterge Paneli · Formüllü &amp; Otomatik Hesaplama · 2026 Sürüm
    </text>

    <!-- Price Tag 999 TL -->
    <g transform="translate(450, 92)">
      <rect width="165" height="28" rx="6" fill="rgba(16,185,129,0.22)" stroke="#10b981" stroke-width="1.2"/>
      <circle cx="16" cy="14" r="4" fill="#10b981"/>
      <text x="94" y="19" fill="#34d399" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="12" font-weight="800" text-anchor="middle">Tek Fiyat: 999 TL</text>
    </g>
  </g>

  <!-- Left 3D Radial Gauge -->
  <g transform="translate(130, 220)" filter="url(#shadow_${p.id})">
    <ellipse cx="45" cy="45" rx="40" ry="26" fill="none" stroke="rgba(255,255,255,0.12)" stroke-width="8"/>
    <ellipse cx="45" cy="45" rx="40" ry="26" fill="none" stroke="${pal.accent}" stroke-width="8" stroke-dasharray="160 70"/>
    <text x="45" y="49" fill="#ffffff" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="14" font-weight="900" text-anchor="middle">%100</text>
  </g>

  <!-- Right 3D Indicator Widget -->
  <g transform="translate(770, 220)" filter="url(#shadow_${p.id})">
    <polygon points="0,22 65,0 130,22 65,44" fill="url(#glass_${p.id})" stroke="${pal.secondary}" stroke-width="1.5"/>
    <polygon points="0,22 65,44 65,58 0,36" fill="rgba(0,0,0,0.5)"/>
    <polygon points="65,44 130,22 130,36 65,58" fill="rgba(0,0,0,0.3)"/>
    <text x="65" y="27" fill="#ffffff" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="13" font-weight="900" text-anchor="middle">PREMİUM</text>
  </g>
</svg>`;
}

async function renderAllWebps() {
  const webpProducts = DATA.products.filter(p => p.previewImage.endsWith('.webp'));
  console.log(`Generating 3D WebP visuals for ${webpProducts.length} home products...`);

  for (const p of webpProducts) {
    const filename = path.basename(p.previewImage);
    const svgCode = generate3DImageSvg(p);
    
    // Convert SVG buffer to high quality WebP using sharp
    const webpBuffer = await sharp(Buffer.from(svgCode))
      .webp({ quality: 95, effort: 6 })
      .toBuffer();

    fs.writeFileSync(path.join(siteAssetsDir, filename), webpBuffer);
    if (fs.existsSync(kaynakAssetsDir)) {
      fs.writeFileSync(path.join(kaynakAssetsDir, filename), webpBuffer);
    }
    console.log(`Generated 3D WebP: ${filename}`);
  }
}

renderAllWebps().then(() => console.log('Successfully completed 3D WebP visual generation!'));
