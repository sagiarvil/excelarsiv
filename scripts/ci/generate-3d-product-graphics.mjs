import fs from 'node:fs';
import path from 'node:path';

const dataFile = path.resolve('excelarsiv-premium-tum-sayfalar/site/assets/data.js');
const siteAssetsDir = path.resolve('excelarsiv-premium-tum-sayfalar/site/assets');
const kaynakAssetsDir = path.resolve('excelarsiv-premium-tum-sayfalar/kaynak/assets');

const raw = fs.readFileSync(dataFile, 'utf8');
const json = raw.replace(/^window\.ExcelArsivData=/, '').replace(/;?\s*$/, '');
const DATA = JSON.parse(json);

// Color themes based on category or hash
const PALETTES = [
  { bg1: '#0f172a', bg2: '#1e293b', accent: '#38bdf8', secondary: '#818cf8', glow: 'rgba(56,189,248,0.3)', light: '#e0f2fe' },
  { bg1: '#064e3b', bg2: '#065f46', accent: '#34d399', secondary: '#10b981', glow: 'rgba(52,211,153,0.3)', light: '#d1fae5' },
  { bg1: '#1e1b4b', bg2: '#312e81', accent: '#a78bfa', secondary: '#c084fc', glow: 'rgba(167,139,250,0.3)', light: '#ede9fe' },
  { bg1: '#701a75', bg2: '#86198f', accent: '#f472b6', secondary: '#fb7185', glow: 'rgba(244,114,182,0.3)', light: '#fce7f3' },
  { bg1: '#1c1917', bg2: '#292524', accent: '#fbbf24', secondary: '#f59e0b', glow: 'rgba(251,191,36,0.3)', light: '#fef3c7' },
  { bg1: '#1e3a5f', bg2: '#0f2744', accent: '#0ea5e9', secondary: '#38bdf8', glow: 'rgba(14,165,233,0.35)', light: '#e0f2fe' }
];

function getPalette(id, category) {
  let hash = 0;
  for (let i = 0; i < id.length; i++) hash = (hash * 31 + id.charCodeAt(i)) & 0xffffffff;
  return PALETTES[Math.abs(hash) % PALETTES.length];
}

function escapeXml(unsafe) {
  return String(unsafe).replace(/[<>&'"]/g, c => {
    switch (c) {
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '&': return '&amp;';
      case '\'': return '&apos;';
      case '"': return '&quot;';
    }
  });
}

function generate3DSvg(product) {
  const p = product;
  const pal = getPalette(p.id, p.category);
  const title = escapeXml(p.shortTitle || p.title || 'Excel Çözümü');
  const cat = escapeXml(p.categoryName || p.category || 'Finans & İş Çözümü');
  
  // Calculate dynamic 3D elements based on hash to make each uniquely distinct
  let h = 0;
  for (let i = 0; i < p.id.length; i++) h = (h * 17 + p.id.charCodeAt(i)) & 0xffff;
  
  const bar1 = 80 + (h % 90);
  const bar2 = 120 + ((h >> 2) % 110);
  const bar3 = 160 + ((h >> 4) % 120);
  const bar4 = 140 + ((h >> 6) % 130);
  const bar5 = 200 + ((h >> 8) % 100);

  const isFinance = /nakit|kasa|maliye|bilanco|kredi|faiz|finans|tahsilat|fatura/i.test(p.id + p.title);
  const isHR = /bordro|personel|maas|puantaj|izin|sgk|ik/i.test(p.id + p.title);
  const isStock = /stok|envanter|depo|siparis|tedarik|lojistik/i.test(p.id + p.title);

  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 600" width="100%" height="100%">
  <defs>
    <!-- Background Gradient -->
    <linearGradient id="bgGrad_${p.id}" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="${pal.bg1}"/>
      <stop offset="100%" stop-color="${pal.bg2}"/>
    </linearGradient>

    <!-- Glass Floor Gradient -->
    <linearGradient id="floorGrad_${p.id}" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" stop-color="rgba(255,255,255,0.12)"/>
      <stop offset="100%" stop-color="rgba(255,255,255,0.02)"/>
    </linearGradient>

    <!-- 3D Bar Gradients -->
    <linearGradient id="accentGrad_${p.id}" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="${pal.accent}"/>
      <stop offset="100%" stop-color="${pal.secondary}"/>
    </linearGradient>

    <linearGradient id="barFront_${p.id}" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" stop-color="${pal.accent}"/>
      <stop offset="100%" stop-color="${pal.secondary}"/>
    </linearGradient>

    <linearGradient id="barSide_${p.id}" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="rgba(0,0,0,0.35)"/>
      <stop offset="100%" stop-color="rgba(0,0,0,0.1)"/>
    </linearGradient>

    <!-- Glow Filter -->
    <filter id="glow_${p.id}" x="-30%" y="-30%" width="160%" height="160%">
      <feGaussianBlur stdDeviation="22" result="blur"/>
      <feComposite in="SourceGraphic" in2="blur" operator="over"/>
    </filter>

    <filter id="shadow_${p.id}" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="16" stdDeviation="18" flood-color="rgba(0,0,0,0.45)"/>
    </filter>
  </defs>

  <!-- Deep 3D Studio Canvas -->
  <rect width="960" height="600" fill="url(#bgGrad_${p.id})"/>

  <!-- Radial Atmosphere Light -->
  <circle cx="500" cy="220" r="380" fill="${pal.glow}" filter="url(#glow_${p.id})"/>

  <!-- 3D Isometric Floor Grid -->
  <g opacity="0.18">
    <path d="M 120 440 L 480 260 L 840 440 L 480 580 Z" fill="none" stroke="rgba(255,255,255,0.4)" stroke-width="1.5"/>
    <path d="M 210 440 L 480 305 L 750 440 L 480 535 Z" fill="none" stroke="rgba(255,255,255,0.3)" stroke-width="1"/>
    <path d="M 300 440 L 480 350 L 660 440 L 480 490 Z" fill="none" stroke="rgba(255,255,255,0.2)" stroke-width="1"/>
    
    <!-- Floor Matrix Grid Lines -->
    <line x1="240" y1="380" x2="600" y2="560" stroke="rgba(255,255,255,0.15)" stroke-width="1"/>
    <line x1="360" y1="320" x2="720" y2="500" stroke="rgba(255,255,255,0.15)" stroke-width="1"/>
    <line x1="720" y1="380" x2="360" y2="560" stroke="rgba(255,255,255,0.15)" stroke-width="1"/>
    <line x1="600" y1="320" x2="240" y2="500" stroke="rgba(255,255,255,0.15)" stroke-width="1"/>
  </g>

  <!-- Isometric 3D Platform Base -->
  <g filter="url(#shadow_${p.id})">
    <!-- Top Face -->
    <polygon points="480,270 780,380 480,490 180,380" fill="url(#floorGrad_${p.id})" stroke="rgba(255,255,255,0.3)" stroke-width="2"/>
    <!-- Left Front Face -->
    <polygon points="180,380 480,490 480,510 180,400" fill="rgba(0,0,0,0.5)" stroke="rgba(255,255,255,0.1)" stroke-width="1"/>
    <!-- Right Front Face -->
    <polygon points="480,490 780,380 780,400 480,510" fill="rgba(0,0,0,0.3)" stroke="rgba(255,255,255,0.1)" stroke-width="1"/>
  </g>

  <!-- Dynamic 3D Isometric Bar Graph Pillars Rising From Floor -->
  <!-- Pillar 1 -->
  <g transform="translate(290, ${410 - bar1})">
    <!-- Front Face -->
    <polygon points="0,${bar1} 34,${bar1 + 14} 34,14 0,0" fill="url(#barFront_${p.id})"/>
    <!-- Right Face (3D depth) -->
    <polygon points="34,${bar1 + 14} 60,${bar1} 60,0 34,14" fill="${pal.secondary}" opacity="0.8"/>
    <!-- Top Face -->
    <polygon points="0,0 26,-14 60,0 34,14" fill="#ffffff" opacity="0.6"/>
  </g>

  <!-- Pillar 2 -->
  <g transform="translate(370, ${430 - bar2})">
    <polygon points="0,${bar2} 36,${bar2 + 15} 36,15 0,0" fill="url(#barFront_${p.id})"/>
    <polygon points="36,${bar2 + 15} 64,${bar2} 64,0 36,15" fill="${pal.secondary}" opacity="0.8"/>
    <polygon points="0,0 28,-15 64,0 36,15" fill="#ffffff" opacity="0.75"/>
  </g>

  <!-- Pillar 3 (Peak Pillar) -->
  <g transform="translate(450, ${450 - bar3})">
    <polygon points="0,${bar3} 40,${bar3 + 17} 40,17 0,0" fill="${pal.accent}"/>
    <polygon points="40,${bar3 + 17} 72,${bar3} 72,0 40,17" fill="${pal.secondary}" opacity="0.85"/>
    <polygon points="0,0 32,-17 72,0 40,17" fill="#ffffff" opacity="0.9"/>
  </g>

  <!-- Pillar 4 -->
  <g transform="translate(540, ${425 - bar4})">
    <polygon points="0,${bar4} 36,${bar4 + 15} 36,15 0,0" fill="url(#barFront_${p.id})"/>
    <polygon points="36,${bar4 + 15} 64,${bar4} 64,0 36,15" fill="${pal.secondary}" opacity="0.75"/>
    <polygon points="0,0 28,-15 64,0 36,15" fill="#ffffff" opacity="0.7"/>
  </g>

  <!-- Pillar 5 -->
  <g transform="translate(620, ${400 - bar5})">
    <polygon points="0,${bar5} 34,${bar5 + 14} 34,14 0,0" fill="url(#barFront_${p.id})"/>
    <polygon points="34,${bar5 + 14} 60,${bar5} 60,0 34,14" fill="${pal.secondary}" opacity="0.8"/>
    <polygon points="0,0 26,-14 60,0 34,14" fill="#ffffff" opacity="0.6"/>
  </g>

  <!-- 3D Floating Curved Trend Line with glowing nodes -->
  <path d="M 320 ${410 - bar1} Q 400 ${360 - bar2} 480 ${430 - bar3} T 650 ${390 - bar5}" 
        fill="none" stroke="#ffffff" stroke-width="3" stroke-linecap="round" filter="url(#glow_${p.id})"/>
  <path d="M 320 ${410 - bar1} Q 400 ${360 - bar2} 480 ${430 - bar3} T 650 ${390 - bar5}" 
        fill="none" stroke="${pal.accent}" stroke-width="2" stroke-linecap="round"/>

  <!-- Glowing Indicator Spheres -->
  <circle cx="320" cy="${410 - bar1}" r="5" fill="#ffffff" filter="url(#glow_${p.id})"/>
  <circle cx="480" cy="${430 - bar3}" r="7" fill="#ffffff" filter="url(#glow_${p.id})"/>
  <circle cx="650" cy="${390 - bar5}" r="6" fill="#ffffff" filter="url(#glow_${p.id})"/>

  <!-- Floating Glassmorphic 3D Card HUD in Background -->
  <g transform="translate(180, 75)" filter="url(#shadow_${p.id})">
    <!-- Glass Background with subtle 3D tilt -->
    <rect width="600" height="150" rx="16" fill="rgba(255,255,255,0.06)" stroke="rgba(255,255,255,0.2)" stroke-width="1.5" backdrop-filter="blur(14px)"/>
    
    <!-- Top Accent Stripe -->
    <rect x="2" y="2" width="596" height="4" rx="2" fill="url(#accentGrad_${p.id})"/>

    <!-- Category Pill -->
    <rect x="28" y="24" width="130" height="26" rx="13" fill="${pal.glow}" stroke="${pal.accent}" stroke-width="1"/>
    <text x="93" y="41" fill="${pal.accent}" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="12" font-weight="800" text-anchor="middle" letter-spacing="0.5">${cat}</text>

    <!-- Excel 3D Badge -->
    <g transform="translate(530, 22)">
      <rect width="42" height="30" rx="6" fill="#107c41" filter="url(#shadow_${p.id})"/>
      <text x="21" y="21" fill="#ffffff" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="16" font-weight="900" text-anchor="middle">XLS</text>
    </g>

    <!-- Product Title in Card -->
    <text x="28" y="85" fill="#ffffff" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="22" font-weight="800" letter-spacing="-0.3">
      ${title.length > 42 ? title.slice(0, 40) + '…' : title}
    </text>

    <!-- Subtitle / Metrics -->
    <text x="28" y="115" fill="rgba(255,255,255,0.7)" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="13" font-weight="500">
      ${isFinance ? 'Otomatik Nakit Projeksiyonu · 3D Grafik Modülü · 2026 Uyumlu' : isHR ? 'Dinamik Bordro &amp; Puantaj Matrisi · 3D Analiz Paneli' : isStock ? 'Akıllı Stok &amp; Depo Optimizasyonu · 3D Karlılık Göstergeleri' : 'Kurumsal Karar Destek Modeli · Canlı Yönetim Kokpiti'}
    </text>

    <!-- High-Converting Value Badge (1.999 TL & Instant Delivery) -->
    <g transform="translate(420, 100)">
      <rect width="152" height="28" rx="6" fill="rgba(16,185,129,0.18)" stroke="#10b981" stroke-width="1"/>
      <circle cx="14" cy="14" r="4" fill="#10b981"/>
      <text x="82" y="19" fill="#34d399" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="12" font-weight="700" text-anchor="middle">Anında İndir · 1.999 TL</text>
    </g>
  </g>

  <!-- 3D Floating Mini KPI Ring / Donut (Left Side) -->
  <g transform="translate(140, 240)" filter="url(#shadow_${p.id})">
    <ellipse cx="40" cy="40" rx="36" ry="24" fill="none" stroke="rgba(255,255,255,0.1)" stroke-width="8"/>
    <ellipse cx="40" cy="40" rx="36" ry="24" fill="none" stroke="${pal.accent}" stroke-width="8" stroke-dasharray="140 60"/>
    <text x="40" y="44" fill="#ffffff" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="13" font-weight="800" text-anchor="middle">%98.4</text>
  </g>

  <!-- 3D Floating Mini KPI Stat (Right Side) -->
  <g transform="translate(760, 230)" filter="url(#shadow_${p.id})">
    <polygon points="0,20 60,0 120,20 60,40" fill="url(#floorGrad_${p.id})" stroke="${pal.secondary}" stroke-width="1.5"/>
    <polygon points="0,20 60,40 60,55 0,35" fill="rgba(0,0,0,0.4)"/>
    <polygon points="60,40 120,20 120,35 60,55" fill="rgba(0,0,0,0.2)"/>
    <text x="60" y="25" fill="#ffffff" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif" font-size="12" font-weight="800" text-anchor="middle">ROI +%42</text>
  </g>
</svg>`;
}

let generatedCount = 0;
for (const p of DATA.products) {
  if (p.previewImage && p.previewImage.endsWith('.svg')) {
    const filename = path.basename(p.previewImage);
    const svgCode = generate3DSvg(p);
    
    // Save to site/assets
    fs.writeFileSync(path.join(siteAssetsDir, filename), svgCode, 'utf8');
    // Save to kaynak/assets if exists
    if (fs.existsSync(kaynakAssetsDir)) {
      fs.writeFileSync(path.join(kaynakAssetsDir, filename), svgCode, 'utf8');
    }
    generatedCount++;
  }
}

console.log(`Successfully generated and replaced ${generatedCount} 3D SVG product graphics!`);
