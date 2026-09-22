import fs from 'fs';

// Let's test generating 4 distinct chart types from the user's examples in pure mathematical SVG:
// 1. Arrow Bar Chart (Ekran Resmi 20.43.01)
// 2. Dual Concentric Gauge (Ekran Resmi 20.42.55)
// 3. Growth Curve Arrow over Bars (Ekran Resmi 20.59.29)
// 4. Dot Waffle Matrix (Ekran Resmi 20.43.13)

const width = 1600;
const height = 1000;
const cardW = 1240;
const cardH = 780;
const cardX = (width - cardW) / 2;
const cardY = (height - cardH) / 2;

function getCardWrapper(title, tag, chartContent) {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">
  <defs>
    <filter id="cardShadow" x="-5%" y="-5%" width="110%" height="115%">
      <feDropShadow dx="0" dy="24" stdDeviation="32" flood-color="#0f172a" flood-opacity="0.06" />
      <feDropShadow dx="0" dy="4" stdDeviation="8" flood-color="#0f172a" flood-opacity="0.04" />
    </filter>
  </defs>
  
  <!-- Canvas Background -->
  <rect width="100%" height="100%" fill="#f8fafc" />
  
  <!-- Subtle Grid Pattern -->
  <g stroke="#e2e8f0" stroke-width="1" opacity="0.4">
    <line x1="0" y1="200" x2="1600" y2="200" />
    <line x1="0" y1="400" x2="1600" y2="400" />
    <line x1="0" y1="600" x2="1600" y2="600" />
    <line x1="0" y1="800" x2="1600" y2="800" />
  </g>
  
  <!-- Main Presentation Card -->
  <rect x="${cardX}" y="${cardY}" width="${cardW}" height="${cardH}" rx="32" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" filter="url(#cardShadow)" />
  
  <!-- Card Header -->
  <g transform="translate(${cardX + 60}, ${cardY + 60})">
    <rect x="0" y="0" width="160" height="32" rx="16" fill="#ecfdf5" stroke="#a7f3d0" stroke-width="1" />
    <text x="80" y="21" fill="#065f46" font-family="-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif" font-size="13" font-weight="700" text-anchor="middle">${tag}</text>
    <text x="0" y="70" fill="#0f172a" font-family="-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif" font-size="28" font-weight="800">${title}</text>
  </g>
  
  <!-- Chart Area (Centered) -->
  <g transform="translate(${cardX}, ${cardY + 120})">
    ${chartContent}
  </g>
</svg>`;
}

// 1. Redrawn Arrow Bar Chart
function drawArrowBars() {
  const bars = [
    { h: 360, col: '#064e3b', label: 'Q1' },
    { h: 240, col: '#047857', label: 'Q2' },
    { h: 180, col: '#10b981', label: 'Q3' },
    { h: 320, col: '#34d399', label: 'Q4' }
  ];
  let res = `<g transform="translate(320, 80)">`;
  // Horizontal gridlines
  for (let i = 0; i <= 4; i++) {
    const y = i * 90;
    res += `<line x1="0" y1="${y}" x2="600" y2="${y}" stroke="#e2e8f0" stroke-width="2" stroke-dasharray="6 6"/>`;
    res += `<text x="-20" y="${y + 5}" fill="#94a3b8" font-family="sans-serif" font-size="14" text-anchor="end">${(4 - i) * 25}%</text>`;
  }
  // Y-axis line
  res += `<line x1="0" y1="0" x2="0" y2="360" stroke="#cbd5e1" stroke-width="2"/>`;
  res += `<line x1="0" y1="360" x2="600" y2="360" stroke="#cbd5e1" stroke-width="2"/>`;
  
  // Arrow bars
  bars.forEach((b, idx) => {
    const bx = 60 + idx * 140;
    const by = 360 - b.h;
    const bw = 90;
    // Arrow top polygon
    res += `<polygon points="${bx},${by + 40} ${bx + bw/2},${by} ${bx + bw},${by + 40} ${bx + bw},360 ${bx},360" fill="${b.col}" />`;
    // Label
    res += `<text x="${bx + bw/2}" y="390" fill="#64748b" font-family="sans-serif" font-size="16" font-weight="600" text-anchor="middle">${b.label}</text>`;
  });
  res += `</g>`;
  return res;
}

// 2. Redrawn Dual Concentric Gauge
function drawConcentricGauge(pct1 = 34, pct2 = 72) {
  const cx = 620;
  const cy = 260;
  const r1 = 140;
  const r2 = 190;
  
  const circ1 = 2 * Math.PI * r1;
  const circ2 = 2 * Math.PI * r2;
  const offset1 = circ1 * (1 - pct1 / 100);
  const offset2 = circ2 * (1 - pct2 / 100);
  
  return `<g>
    <!-- Background track rings -->
    <circle cx="${cx}" cy="${cy}" r="${r1}" fill="none" stroke="#f1f5f9" stroke-width="24" />
    <circle cx="${cx}" cy="${cy}" r="${r2}" fill="none" stroke="#f1f5f9" stroke-width="28" />
    
    <!-- Active Progress Arcs -->
    <circle cx="${cx}" cy="${cy}" r="${r1}" fill="none" stroke="#10b981" stroke-width="24"
      stroke-dasharray="${circ1}" stroke-dashoffset="${offset1}" stroke-linecap="round"
      transform="rotate(-90 ${cx} ${cy})" />
    <circle cx="${cx}" cy="${cy}" r="${r2}" fill="none" stroke="#047857" stroke-width="28"
      stroke-dasharray="${circ2}" stroke-dashoffset="${offset2}" stroke-linecap="round"
      transform="rotate(-90 ${cx} ${cy})" />
      
    <!-- Center Percentage Text -->
    <text x="${cx}" y="${cy + 16}" fill="#0f172a" font-family="-apple-system, sans-serif" font-size="64" font-weight="900" text-anchor="middle">${pct1}%</text>
    <text x="${cx}" y="${cy + 52}" fill="#64748b" font-family="-apple-system, sans-serif" font-size="18" font-weight="600" text-anchor="middle">OPTİMUM ORAN</text>
  </g>`;
}

// 3. Redrawn Growth Curve Arrow over Bars
function drawGrowthCurve() {
  const heights = [70, 130, 200, 290, 360];
  let res = `<g transform="translate(340, 80)">`;
  // Grid
  for (let i = 0; i <= 4; i++) {
    const y = i * 90;
    res += `<line x1="0" y1="${y}" x2="560" y2="${y}" stroke="#e2e8f0" stroke-width="2"/>`;
  }
  res += `<line x1="0" y1="360" x2="560" y2="360" stroke="#94a3b8" stroke-width="2"/>`;
  
  // Bars
  const colors = ['#f87171', '#fb923c', '#fbbf24', '#38bdf8', '#34d399'];
  heights.forEach((h, i) => {
    const x = 40 + i * 105;
    const y = 360 - h;
    res += `<rect x="${x}" y="${y}" width="65" height="${h}" rx="6" fill="${colors[i]}" />`;
  });
  
  // Smooth curved arrow sweeping up
  res += `<path d="M 40 270 Q 220 260 480 50" fill="none" stroke="#ef4444" stroke-width="12" stroke-linecap="round" />`;
  // Arrow head
  res += `<polygon points="480,30 495,65 460,60" fill="#ef4444" />`;
  res += `</g>`;
  return res;
}

// 4. Redrawn Dot Waffle Matrix
function drawWaffleMatrix() {
  let res = `<g transform="translate(420, 60)">`;
  const cols = 10;
  const rows = 8;
  const dotR = 14;
  const gap = 38;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const x = c * gap;
      const y = r * gap;
      const isFilled = (r * cols + c) < 46;
      const col = isFilled ? (r < 3 ? '#047857' : (r < 6 ? '#10b981' : '#34d399')) : '#e2e8f0';
      res += `<circle cx="${x}" cy="${y}" r="${dotR}" fill="${col}" />`;
    }
  }
  res += `<text x="170" y="350" fill="#0f172a" font-family="sans-serif" font-size="28" font-weight="800" text-anchor="middle">%46 TAMAMLANMA ORANI</text>`;
  res += `</g>`;
  return res;
}

fs.writeFileSync('/tmp/chart1_arrows.svg', getCardWrapper('Haftalık Satış & Kapasite Analizi', 'VERİMLİLİK', drawArrowBars()));
fs.writeFileSync('/tmp/chart2_gauge.svg', getCardWrapper('Nakit Akışı Gerçekleşme Oranı', 'FİNANSAL GÖSTERGE', drawConcentricGauge()));
fs.writeFileSync('/tmp/chart3_growth.svg', getCardWrapper('Kârlılık & Büyüme Trendi', 'PERFORMANS', drawGrowthCurve()));
fs.writeFileSync('/tmp/chart4_waffle.svg', getCardWrapper('Kapasite Doluluk & Tahsis Matrisi', 'OPERASYON', drawWaffleMatrix()));
console.log('Saved 4 test charts to /tmp!');
