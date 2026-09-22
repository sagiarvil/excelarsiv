import fs from 'fs';
import sharp from 'sharp';

// Test 4 square archetypes (1000x1000, filling the square area, perfectly centered)

// 1. Arrow Bars (filling 820x820)
function drawArrowBarsSquare() {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" width="1000" height="1000">
    <rect width="100%" height="100%" fill="#ffffff" />
    <g transform="translate(100, 100)">
      <!-- Gridlines -->
      <line x1="0" y1="160" x2="800" y2="160" stroke="#f1f5f9" stroke-width="4" stroke-dasharray="12 12" />
      <line x1="0" y1="360" x2="800" y2="360" stroke="#f1f5f9" stroke-width="4" stroke-dasharray="12 12" />
      <line x1="0" y1="560" x2="800" y2="560" stroke="#f1f5f9" stroke-width="4" stroke-dasharray="12 12" />
      <line x1="0" y1="760" x2="800" y2="760" stroke="#cbd5e1" stroke-width="6" />
      
      <!-- 4 Arrow Bars -->
      <polygon points="60,320 120,240 180,320 180,760 60,760" fill="#064e3b" />
      <polygon points="260,460 320,380 380,460 380,760 260,760" fill="#047857" />
      <polygon points="460,540 520,460 580,540 580,760 460,760" fill="#10b981" />
      <polygon points="660,260 720,180 780,260 780,760 660,760" fill="#34d399" />
      
      <!-- Axis labels -->
      <text x="120" y="820" fill="#64748b" font-family="sans-serif" font-size="28" font-weight="800" text-anchor="middle">Q1</text>
      <text x="320" y="820" fill="#64748b" font-family="sans-serif" font-size="28" font-weight="800" text-anchor="middle">Q2</text>
      <text x="520" y="820" fill="#64748b" font-family="sans-serif" font-size="28" font-weight="800" text-anchor="middle">Q3</text>
      <text x="720" y="820" fill="#64748b" font-family="sans-serif" font-size="28" font-weight="800" text-anchor="middle">Q4</text>
    </g>
  </svg>`;
}

// 2. Concentric Gauge (filling 820x820)
function drawGaugeSquare() {
  const cx = 500, cy = 500;
  const r1 = 280, r2 = 380;
  const circ1 = 2 * Math.PI * r1;
  const circ2 = 2 * Math.PI * r2;
  const off1 = circ1 * (1 - 0.34);
  const off2 = circ2 * (1 - 0.72);
  
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" width="1000" height="1000">
    <rect width="100%" height="100%" fill="#ffffff" />
    <circle cx="${cx}" cy="${cy}" r="${r1}" fill="none" stroke="#f1f5f9" stroke-width="48" />
    <circle cx="${cx}" cy="${cy}" r="${r2}" fill="none" stroke="#f1f5f9" stroke-width="56" />
    
    <circle cx="${cx}" cy="${cy}" r="${r1}" fill="none" stroke="#10b981" stroke-width="48"
      stroke-dasharray="${circ1}" stroke-dashoffset="${off1}" stroke-linecap="round"
      transform="rotate(-90 ${cx} ${cy})" />
    <circle cx="${cx}" cy="${cy}" r="${r2}" fill="none" stroke="#047857" stroke-width="56"
      stroke-dasharray="${circ2}" stroke-dashoffset="${off2}" stroke-linecap="round"
      transform="rotate(-90 ${cx} ${cy})" />
      
    <text x="${cx}" y="${cy + 36}" fill="#0f172a" font-family="sans-serif" font-size="130" font-weight="900" text-anchor="middle">34%</text>
  </svg>`;
}

// 3. Document with Verified Badge (filling 820x820)
function drawDocumentSquare() {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" width="1000" height="1000">
    <rect width="100%" height="100%" fill="#ffffff" />
    <g transform="translate(180, 100)">
      <!-- Document Page -->
      <rect x="0" y="40" width="560" height="740" rx="24" fill="#ffffff" stroke="#1e293b" stroke-width="20" />
      <!-- Top Right Fold -->
      <path d="M 440 40 L 560 160 L 440 160 Z" fill="#f59e0b" stroke="#1e293b" stroke-width="16" />
      
      <!-- Text Lines -->
      <rect x="80" y="160" width="300" height="24" rx="12" fill="#64748b" />
      <rect x="80" y="240" width="400" height="24" rx="12" fill="#94a3b8" />
      <rect x="80" y="320" width="360" height="24" rx="12" fill="#94a3b8" />
      <rect x="80" y="400" width="280" height="24" rx="12" fill="#94a3b8" />
      <rect x="80" y="480" width="340" height="24" rx="12" fill="#94a3b8" />
      
      <!-- Verified Badge in Bottom Right -->
      <circle cx="500" cy="680" r="130" fill="#10b981" stroke="#1e293b" stroke-width="18" />
      <!-- Checkmark -->
      <path d="M 430 680 L 480 730 L 580 620" fill="none" stroke="#ffffff" stroke-width="28" stroke-linecap="round" stroke-linejoin="round" />
    </g>
  </svg>`;
}

// 4. Browser Spreadsheet Window (filling 820x820)
function drawSpreadsheetSquare() {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" width="1000" height="1000">
    <rect width="100%" height="100%" fill="#ffffff" />
    <g transform="translate(100, 100)">
      <!-- Window Frame -->
      <rect x="0" y="0" width="800" height="800" rx="32" fill="#ffffff" stroke="#1e293b" stroke-width="20" />
      <!-- Window Titlebar -->
      <rect x="0" y="0" width="800" height="100" rx="32" fill="#f8fafc" />
      <line x1="0" y1="100" x2="800" y2="100" stroke="#1e293b" stroke-width="16" />
      <!-- Window Dots -->
      <circle cx="60" cy="50" r="16" fill="#ef4444" />
      <circle cx="110" cy="50" r="16" fill="#f59e0b" />
      <circle cx="160" cy="50" r="16" fill="#10b981" />
      
      <!-- Inner Grid Area -->
      <rect x="40" y="140" width="720" height="620" rx="16" fill="#ecfdf5" stroke="#a7f3d0" stroke-width="8" />
      <!-- Spreadsheet Data Rows -->
      <rect x="80" y="200" width="480" height="32" rx="12" fill="#065f46" />
      <rect x="80" y="280" width="580" height="32" rx="12" fill="#047857" />
      <rect x="80" y="360" width="340" height="32" rx="12" fill="#059669" />
      <rect x="80" y="440" width="520" height="32" rx="12" fill="#10b981" />
      <rect x="80" y="520" width="620" height="32" rx="12" fill="#047857" />
      <rect x="80" y="600" width="420" height="32" rx="12" fill="#065f46" />
    </g>
  </svg>`;
}

fs.writeFileSync('/tmp/test_sq1.svg', drawArrowBarsSquare());
fs.writeFileSync('/tmp/test_sq2.svg', drawGaugeSquare());
fs.writeFileSync('/tmp/test_sq3.svg', drawDocumentSquare());
fs.writeFileSync('/tmp/test_sq4.svg', drawSpreadsheetSquare());

sharp('/tmp/test_sq1.svg').png().toFile('/tmp/test_sq1.png')
  .then(() => sharp('/tmp/test_sq2.svg').png().toFile('/tmp/test_sq2.png'))
  .then(() => sharp('/tmp/test_sq3.svg').png().toFile('/tmp/test_sq3.png'))
  .then(() => sharp('/tmp/test_sq4.svg').png().toFile('/tmp/test_sq4.png'))
  .then(() => console.log('Rendered 4 square test images!'))
  .catch(console.error);
