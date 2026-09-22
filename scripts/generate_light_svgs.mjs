import fs from 'fs';
import path from 'path';

const destDir = 'public/images/kapak';
if (!fs.existsSync(destDir)) {
  fs.mkdirSync(destDir, { recursive: true });
}

// 1. Get Slugs
const kapakFile = fs.readFileSync('src/lib/kapak.ts', 'utf8');
const match = kapakFile.match(/const PREMIUM_KAPAK_SLUGS = new Set<string>\(\[\s*([\s\S]*?)\s*\]\);/);
const slugs = match[1].split(',').map(s => s.trim().replace(/['"]/g, '')).filter(s => s.length > 0);

// RNG Generator
function mulberry32(a) {
  return function() {
    var t = a += 0x6D2B79F5;
    t = Math.imul(t ^ t >>> 15, t | 1);
    t ^= t + Math.imul(t ^ t >>> 7, t | 61);
    return ((t ^ t >>> 14) >>> 0) / 4294967296;
  }
}
function getRng(slug) {
    let seed = 0;
    for(let i=0; i<slug.length; i++) seed += slug.charCodeAt(i) * Math.pow(10, i%4);
    return mulberry32(seed);
}

function r(rng, min, max) { return min + rng() * (max - min); }

// Light Corporate Palettes (SaaS / Excel / Finance)
const palettes = [
  // 0: Excel Classic (Green & Slate)
  { bg: '#f8fafc', card: '#ffffff', p1: '#10b981', p2: '#059669', p3: '#34d399', text: '#334155', grid: '#e2e8f0' },
  // 1: Stripe Blue (Professional Finance)
  { bg: '#f1f5f9', card: '#ffffff', p1: '#3b82f6', p2: '#2563eb', p3: '#60a5fa', text: '#1e293b', grid: '#cbd5e1' },
  // 2: Modern Purple (SaaS Analytics)
  { bg: '#f8fafc', card: '#ffffff', p1: '#8b5cf6', p2: '#7c3aed', p3: '#a78bfa', text: '#334155', grid: '#e2e8f0' },
  // 3: Trust Teal (Corporate)
  { bg: '#f0fdfa', card: '#ffffff', p1: '#14b8a6', p2: '#0d9488', p3: '#5eead4', text: '#0f172a', grid: '#ccfbf1' },
  // 4: Warm Analytics (Orange/Amber)
  { bg: '#fffbeb', card: '#ffffff', p1: '#f59e0b', p2: '#d97706', p3: '#fbbf24', text: '#451a03', grid: '#fef3c7' }
];

function generateLightDashboardSVG(slug, rng) {
    const palIdx = Math.floor(rng() * palettes.length);
    const pal = palettes[palIdx];
    
    // Choose layout style
    const layout = Math.floor(rng() * 3); // 0: Sidebar, 1: Top Navbar, 2: Excel Grid Focus
    
    let content = '';
    
    if (layout === 0) { // Sidebar + Dashboard Cards
        // Sidebar
        content += `<rect x="0" y="0" width="400" height="1600" fill="${pal.card}" filter="url(#shadowRight)"/>`;
        for(let i=0; i<6; i++) {
            content += `<rect x="40" y="${100 + i*100}" width="${r(rng, 100, 280)}" height="30" rx="15" fill="${pal.grid}"/>`;
        }
        
        // Main Area Cards
        const numCards = 3 + Math.floor(rng() * 3);
        for(let i=0; i<numCards; i++) {
            const x = 500 + (i%2)*980;
            const y = 100 + Math.floor(i/2)*700;
            content += `<rect x="${x}" y="${y}" width="900" height="600" rx="40" fill="${pal.card}" filter="url(#shadowDown)"/>`;
            
            // Inside card: Bar Chart
            if (rng() > 0.5) {
                for(let b=0; b<6; b++) {
                    const bh = r(rng, 100, 400);
                    content += `<rect x="${x + 100 + b*120}" y="${y + 500 - bh}" width="60" height="${bh}" rx="10" fill="${pal.p1}"/>`;
                }
            } else { // Line Chart
                let d = `M ${x+100} ${y+400}`;
                for(let b=1; b<6; b++) {
                    d += ` L ${x+100 + b*140} ${y + 400 - r(rng,0,300)}`;
                }
                content += `<path d="${d}" fill="none" stroke="${pal.p2}" stroke-width="16" stroke-linejoin="round"/>`;
            }
        }
    } 
    else if (layout === 1) { // Top Navbar + Wide Analytics
        // Navbar
        content += `<rect x="0" y="0" width="2560" height="160" fill="${pal.card}" filter="url(#shadowDown)"/>`;
        content += `<circle cx="100" cy="80" r="40" fill="${pal.p1}"/>`;
        for(let i=0; i<4; i++) {
            content += `<rect x="${200 + i*200}" y="60" width="120" height="40" rx="20" fill="${pal.grid}"/>`;
        }
        
        // Huge Hero Chart
        content += `<rect x="100" y="260" width="2360" height="1100" rx="60" fill="${pal.card}" filter="url(#shadowDown)"/>`;
        // Grid lines inside
        for(let g=0; g<4; g++) {
            content += `<line x1="200" y1="${400 + g*250}" x2="2360" y2="${400 + g*250}" stroke="${pal.grid}" stroke-width="4"/>`;
        }
        
        // Beautiful Area Curve
        let d = `M 200 1100`;
        const pts = [];
        for(let p=0; p<10; p++) {
            pts.push({ x: 200 + p*240, y: 1100 - r(rng, 100, 700) });
        }
        for(let i=0; i<pts.length-1; i++) {
            const p1 = pts[i], p2 = pts[i+1];
            d += ` C ${p1.x + 120} ${p1.y}, ${p2.x - 120} ${p2.y}, ${p2.x} ${p2.y}`;
        }
        const fillD = d + ` L 2360 1360 L 200 1360 Z`;
        content += `<path d="${fillD}" fill="${pal.p3}" opacity="0.2"/>`;
        content += `<path d="${d}" fill="none" stroke="${pal.p1}" stroke-width="24" stroke-linecap="round"/>`;
    } 
    else { // Excel Grid Focus
        // A vector representation of a modern spreadsheet
        content += `<rect x="100" y="100" width="2360" height="1400" rx="40" fill="${pal.card}" filter="url(#shadowDown)"/>`;
        // Toolbar
        content += `<rect x="100" y="100" width="2360" height="120" fill="${pal.p1}" opacity="0.1"/>`;
        for(let i=0; i<8; i++) {
            content += `<rect x="${140 + i*100}" y="140" width="60" height="40" rx="10" fill="${pal.p2}"/>`;
        }
        // Grid Cells
        const cols = 12;
        const rows = 12;
        const cellW = 2360 / cols;
        const cellH = 1280 / rows;
        for(let c=0; c<cols; c++) {
            for(let r=0; r<rows; r++) {
                content += `<rect x="${100 + c*cellW}" y="${220 + r*cellH}" width="${cellW}" height="${cellH}" fill="none" stroke="${pal.grid}" stroke-width="4"/>`;
                // Add some fake data lines
                if (rng() > 0.4) {
                    content += `<rect x="${120 + c*cellW}" y="${220 + r*cellH + 40}" width="${cellW - 80}" height="16" rx="8" fill="${pal.text}" opacity="0.2"/>`;
                }
            }
        }
        // A floating chart over the grid
        content += `<rect x="600" y="400" width="1200" height="800" rx="40" fill="${pal.card}" filter="url(#shadowDown)"/>`;
        // Pie Chart inside
        content += `<circle cx="1200" cy="800" r="300" fill="${pal.p3}"/>`;
        content += `<path d="M 1200 800 L 1200 500 A 300 300 0 0 1 1460 650 Z" fill="${pal.p1}"/>`;
        content += `<path d="M 1200 800 L 1460 650 A 300 300 0 0 1 1100 1080 Z" fill="${pal.p2}"/>`;
    }

    // Modern Typography Overlay
    const hud = `
      <rect x="2560" y="1400" width="2560" height="200" transform="translate(-2560,0)"/>
      <text x="100" y="1520" fill="${pal.text}" font-family="sans-serif" font-size="48" font-weight="bold">SYSTEM REF: ${slug.toUpperCase()}</text>
      <text x="100" y="1570" fill="${pal.text}" opacity="0.6" font-family="sans-serif" font-size="32">ENTERPRISE EXCEL OS // LIGHT THEME v1.0 // 4K NATIVE</text>
    `;

    return `
<svg width="2560" height="1600" viewBox="0 0 2560 1600" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <filter id="shadowDown" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="20" stdDeviation="30" flood-opacity="0.05" />
    </filter>
    <filter id="shadowRight" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="20" dy="0" stdDeviation="30" flood-opacity="0.05" />
    </filter>
  </defs>
  
  <rect width="100%" height="100%" fill="${pal.bg}" />
  
  ${content}
  ${hud}
</svg>
    `.trim();
}

let count = 0;
for (const slug of slugs) {
    const rng = getRng(slug);
    const svgContent = generateLightDashboardSVG(slug, rng);
    fs.writeFileSync(path.join(destDir, `${slug}.svg`), svgContent);
    count++;
}

console.log(`✅ ${count} FLAWLESS LIGHT THEME SVG files generated!`);
