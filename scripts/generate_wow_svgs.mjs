import fs from 'fs';
import path from 'path';

const destDir = 'public/images/kapak';
if (!fs.existsSync(destDir)) {
  fs.mkdirSync(destDir, { recursive: true });
}

const kapakFile = fs.readFileSync('src/lib/kapak.ts', 'utf8');
const match = kapakFile.match(/const PREMIUM_KAPAK_SLUGS = new Set<string>\(\[\s*([\s\S]*?)\s*\]\);/);
const slugs = match[1].split(',').map(s => s.trim().replace(/['"]/g, '')).filter(s => s.length > 0);

// Deterministic RNG
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

// Enterprise Intelligence Palettes
const palettes = [
  // 0: Palantir Gotham (Dark Obsidian, Cyan, Neon Blue)
  { bg: '#020617', grid: '#0f172a', p1: '#06b6d4', p2: '#3b82f6', p3: '#0ea5e9', text: '#94a3b8' },
  // 1: Primer AI (Deep Violet, Magenta, Emerald)
  { bg: '#09090b', grid: '#18181b', p1: '#8b5cf6', p2: '#ec4899', p3: '#10b981', text: '#a1a1aa' },
  // 2: Recorded Future (Cyber Black, Warning Amber, Crimson)
  { bg: '#0a0a0a', grid: '#171717', p1: '#f59e0b', p2: '#ef4444', p3: '#f97316', text: '#d4d4d8' },
  // 3: Silicon Valley Stripe/Apple (Pristine White, Mesh Gradients)
  { bg: '#ffffff', grid: '#f1f5f9', p1: '#6366f1', p2: '#a855f7', p3: '#3b82f6', text: '#64748b' },
  // 4: Big Four FinTech (Navy Blue, Gold, Teal)
  { bg: '#001e36', grid: '#002b4d', p1: '#fbbf24', p2: '#2dd4bf', p3: '#38bdf8', text: '#94a3b8' },
];

function r(rng, min, max) { return min + rng() * (max - min); }

// Engine 1: Network Constellation (Gotham Style)
function genNetwork(rng, pal) {
    let svg = '';
    const numNodes = Math.floor(r(rng, 30, 80));
    const nodes = [];
    for(let i=0; i<numNodes; i++) {
        nodes.push({ x: r(rng, 400, 3440), y: r(rng, 400, 2000), vx: r(rng, -1, 1), vy: r(rng, -1, 1) });
    }
    // Edges
    for(let i=0; i<numNodes; i++) {
        for(let j=i+1; j<numNodes; j++) {
            const dist = Math.hypot(nodes[i].x - nodes[j].x, nodes[i].y - nodes[j].y);
            if(dist < 400) {
                const opacity = 1 - (dist/400);
                svg += `<line x1="${nodes[i].x}" y1="${nodes[i].y}" x2="${nodes[j].x}" y2="${nodes[j].y}" stroke="${pal.p1}" stroke-width="${r(rng,1,4)}" stroke-opacity="${opacity * 0.5}" filter="url(#glow)"/>`;
            }
        }
    }
    // Nodes
    for(let i=0; i<numNodes; i++) {
        const rad = r(rng, 4, 16);
        const col = rng() > 0.5 ? pal.p1 : pal.p2;
        svg += `<circle cx="${nodes[i].x}" cy="${nodes[i].y}" r="${rad}" fill="${col}" filter="url(#glow)"/>`;
        if(rng() > 0.8) {
            svg += `<circle cx="${nodes[i].x}" cy="${nodes[i].y}" r="${rad*3}" fill="none" stroke="${pal.p3}" stroke-width="2" stroke-dasharray="4 4" opacity="0.5"/>`;
            svg += `<text x="${nodes[i].x + 20}" y="${nodes[i].y - 20}" fill="${pal.text}" font-family="monospace" font-size="24">SYS.${Math.floor(r(rng,1000,9999))}</text>`;
        }
    }
    return svg;
}

// Engine 2: Orbital HUD (Primer AI Style)
function genOrbital(rng, pal) {
    let svg = '';
    const cx = 1920, cy = 1200;
    const numRings = Math.floor(r(rng, 5, 12));
    
    for(let i=0; i<numRings; i++) {
        const radius = r(rng, 200, 1000);
        const width = r(rng, 2, 24);
        const col = [pal.p1, pal.p2, pal.p3][Math.floor(rng()*3)];
        
        if (rng() > 0.3) {
            const dash1 = r(rng, 10, 200);
            const dash2 = r(rng, 10, 200);
            svg += `<circle cx="${cx}" cy="${cy}" r="${radius}" fill="none" stroke="${col}" stroke-width="${width}" stroke-dasharray="${dash1} ${dash2}" opacity="${r(rng,0.2,0.8)}" transform="rotate(${r(rng,0,360)} ${cx} ${cy})"/>`;
        } else {
            svg += `<circle cx="${cx}" cy="${cy}" r="${radius}" fill="none" stroke="${col}" stroke-width="${1}" opacity="0.3"/>`;
        }
        
        // Data points on orbit
        if (rng() > 0.5) {
            const angle = r(rng, 0, Math.PI*2);
            const px = cx + Math.cos(angle)*radius;
            const py = cy + Math.sin(angle)*radius;
            svg += `<circle cx="${px}" cy="${py}" r="${width*1.5}" fill="${pal.bg}" stroke="${col}" stroke-width="4" filter="url(#glow)"/>`;
            svg += `<line x1="${px}" y1="${py}" x2="${px+100}" y2="${py-100}" stroke="${col}" stroke-width="2"/>`;
            svg += `<text x="${px+110}" y="${py-100}" fill="${pal.text}" font-family="monospace" font-size="32">Q${Math.floor(r(rng,1,9))}-${Math.floor(r(rng,100,999))}</text>`;
        }
    }
    
    // Core radar sweep
    svg += `<path d="M ${cx} ${cy} L ${cx} ${cy-1000} A 1000 1000 0 0 1 ${cx+707} ${cy-707} Z" fill="url(#radarSweep)" opacity="0.1"/>`;
    
    return svg;
}

// Engine 3: Isometric Cyber Grid (Recorded Future)
function genIsometric(rng, pal) {
    let svg = `<g transform="translate(1920, 1200) scale(1, 0.5) rotate(45)">`;
    const gridSize = 10;
    const cellSize = 100;
    
    // Base Grid
    for(let x=-gridSize; x<=gridSize; x++) {
        for(let y=-gridSize; y<=gridSize; y++) {
            svg += `<rect x="${x*cellSize}" y="${y*cellSize}" width="${cellSize-4}" height="${cellSize-4}" fill="${pal.grid}" opacity="0.5"/>`;
        }
    }
    
    // Data Blocks
    const numBlocks = Math.floor(r(rng, 20, 60));
    for(let i=0; i<numBlocks; i++) {
        const bx = Math.floor(r(rng, -gridSize, gridSize));
        const by = Math.floor(r(rng, -gridSize, gridSize));
        const h = r(rng, 50, 800);
        const col = [pal.p1, pal.p2, pal.p3][Math.floor(rng()*3)];
        
        // Pseudo 3D Box in Isometric projection
        const px = bx * cellSize;
        const py = by * cellSize;
        const s = cellSize - 4;
        
        // Shadow/glow at base
        svg += `<rect x="${px}" y="${py}" width="${s}" height="${s}" fill="${col}" filter="url(#glow)" opacity="0.6"/>`;
        
        // We simulate the 3D height by drawing a line up (y-axis in standard, but in iso it's weird, so we just use overlapping SVGs with translation)
    }
    svg += `</g>`;
    
    // Overlay HUD
    for(let i=0; i<5; i++) {
        const y = r(rng, 200, 2200);
        svg += `<rect x="200" y="${y}" width="400" height="2" fill="${pal.p1}" opacity="0.5"/>`;
        svg += `<text x="200" y="${y-10}" fill="${pal.text}" font-family="monospace" font-size="24">SECTOR_${Math.floor(r(rng,10,99))}</text>`;
    }
    return svg;
}

// Engine 4: Glassmorphism / Spatial UI (Silicon Valley)
function genGlass(rng, pal) {
    let svg = '';
    // Giant abstract mesh blobs in background
    svg += `<circle cx="${r(rng, 500, 1500)}" cy="${r(rng, 500, 1500)}" r="${r(rng, 600, 1200)}" fill="${pal.p1}" opacity="0.15" filter="url(#hugeBlur)"/>`;
    svg += `<circle cx="${r(rng, 2000, 3000)}" cy="${r(rng, 1000, 2000)}" r="${r(rng, 600, 1200)}" fill="${pal.p2}" opacity="0.15" filter="url(#hugeBlur)"/>`;
    
    // Glass Cards
    const numCards = Math.floor(r(rng, 3, 7));
    for(let i=0; i<numCards; i++) {
        const w = r(rng, 600, 1400);
        const h = r(rng, 400, 800);
        const x = r(rng, 400, 3440 - w);
        const y = r(rng, 400, 2000 - h);
        
        svg += `<g transform="translate(${x}, ${y})">`;
        // Glass panel
        svg += `<rect width="${w}" height="${h}" rx="40" fill="${pal.bg}" fill-opacity="0.6" stroke="${pal.p3}" stroke-opacity="0.2" stroke-width="4" filter="url(#shadow)"/>`;
        // Card Header
        svg += `<circle cx="60" cy="60" r="15" fill="${pal.p1}"/>`;
        svg += `<circle cx="110" cy="60" r="15" fill="${pal.p2}"/>`;
        svg += `<text x="160" y="70" fill="${pal.text}" font-family="sans-serif" font-weight="bold" font-size="32">ANALYTICS // MODULE ${Math.floor(r(rng,1,9))}</text>`;
        
        // Inner Chart (Simple Sine Wave or Bars)
        if(rng() > 0.5) {
            let path = `M 100 ${h/2}`;
            for(let px=100; px<w-100; px+=50) {
                path += ` L ${px} ${h/2 + Math.sin(px/100 + i)*100}`;
            }
            svg += `<path d="${path}" fill="none" stroke="${pal.p3}" stroke-width="12" stroke-linecap="round"/>`;
        } else {
            for(let px=100; px<w-100; px+=80) {
                const bh = r(rng, 50, h-200);
                svg += `<rect x="${px}" y="${h - 100 - bh}" width="40" height="${bh}" rx="20" fill="${pal.p2}"/>`;
            }
        }
        svg += `</g>`;
    }
    return svg;
}

// Engine 5: Quantum Core / Concentric Data Rings (Deep Web)
function genQuantum(rng, pal) {
    let svg = `<g transform="translate(1920, 1200)">`;
    
    // Core Sunburst
    const numRays = 120;
    for(let i=0; i<numRays; i++) {
        if(rng() > 0.7) continue;
        const angle = (i/numRays) * Math.PI * 2;
        const r1 = r(rng, 100, 300);
        const r2 = r1 + r(rng, 100, 800);
        const col = [pal.p1, pal.p2, pal.p3][Math.floor(rng()*3)];
        
        const x1 = Math.cos(angle)*r1, y1 = Math.sin(angle)*r1;
        const x2 = Math.cos(angle)*r2, y2 = Math.sin(angle)*r2;
        
        svg += `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${col}" stroke-width="${r(rng, 2, 12)}" opacity="0.6"/>`;
    }
    
    // Tech Circles
    svg += `<circle cx="0" cy="0" r="100" fill="${pal.bg}" stroke="${pal.p1}" stroke-width="8" filter="url(#glow)"/>`;
    svg += `<circle cx="0" cy="0" r="40" fill="${pal.p2}" filter="url(#glow)"/>`;
    
    svg += `</g>`;
    
    // Targeting Brackets
    svg += `<path d="M 200 400 L 200 200 L 400 200" fill="none" stroke="${pal.text}" stroke-width="12"/>`;
    svg += `<path d="M 3640 400 L 3640 200 L 3440 200" fill="none" stroke="${pal.text}" stroke-width="12"/>`;
    svg += `<path d="M 200 2000 L 200 2200 L 400 2200" fill="none" stroke="${pal.text}" stroke-width="12"/>`;
    svg += `<path d="M 3640 2000 L 3640 2200 L 3440 2200" fill="none" stroke="${pal.text}" stroke-width="12"/>`;
    
    return svg;
}

// Master Generator
function generateWowSVG(slug, rng) {
    const palIdx = Math.floor(rng() * palettes.length);
    const pal = palettes[palIdx];
    const engineIdx = Math.floor(rng() * 5);
    
    let content = '';
    
    // Background Grid Pattern (Common to most)
    let gridDef = `
      <pattern id="gridPattern" width="100" height="100" patternUnits="userSpaceOnUse">
        <path d="M 100 0 L 0 0 0 100" fill="none" stroke="${pal.grid}" stroke-width="2"/>
      </pattern>
    `;
    let bgContent = `<rect width="100%" height="100%" fill="${pal.bg}" />`;
    if (palIdx !== 3) { // Skip strict grid for Silicon Valley light theme sometimes
        bgContent += `<rect width="100%" height="100%" fill="url(#gridPattern)" opacity="0.5"/>`;
    }

    if (engineIdx === 0) content = genNetwork(rng, pal);
    else if (engineIdx === 1) content = genOrbital(rng, pal);
    else if (engineIdx === 2) content = genIsometric(rng, pal);
    else if (engineIdx === 3) content = genGlass(rng, pal);
    else if (engineIdx === 4) content = genQuantum(rng, pal);
    
    // Universal HUD Elements
    const hud = `
      <text x="100" y="100" fill="${pal.text}" font-family="monospace" font-size="32" font-weight="bold">SYSTEM IDENT: ${slug.toUpperCase().substring(0, 20)}...</text>
      <text x="100" y="150" fill="${pal.text}" font-family="monospace" font-size="24">ENGINE_${engineIdx} | SEC_LEVEL_ALPHA | ${Math.random().toString(36).substring(2, 10).toUpperCase()}</text>
      
      <g transform="translate(3500, 100)">
        <rect width="200" height="40" fill="${pal.p1}" opacity="0.2"/>
        <rect width="${r(rng, 50, 180)}" height="40" fill="${pal.p1}"/>
        <text x="10" y="30" fill="${pal.bg}" font-family="monospace" font-size="20" font-weight="bold">PROCESSING</text>
      </g>
    `;

    return `
<svg width="3840" height="2400" viewBox="0 0 3840 2400" xmlns="http://www.w3.org/2000/svg">
  <defs>
    ${gridDef}
    <filter id="glow" x="-20%" y="-20%" width="140%" height="140%">
      <feGaussianBlur stdDeviation="15" result="blur" />
      <feComposite in="SourceGraphic" in2="blur" operator="over" />
    </filter>
    <filter id="shadow" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="40" stdDeviation="50" flood-opacity="0.15" />
    </filter>
    <filter id="hugeBlur" x="-50%" y="-50%" width="200%" height="200%">
      <feGaussianBlur stdDeviation="150" />
    </filter>
    <radialGradient id="radarSweep" cx="50%" cy="50%" r="50%">
      <stop offset="0%" stop-color="${pal.p1}" stop-opacity="1"/>
      <stop offset="100%" stop-color="${pal.bg}" stop-opacity="0"/>
    </radialGradient>
  </defs>
  
  ${bgContent}
  ${content}
  ${hud}
  
  <!-- Outer Frame -->
  <rect width="3840" height="2400" fill="none" stroke="${pal.text}" stroke-width="40" opacity="0.1"/>
</svg>
    `.trim();
}

let count = 0;
for (const slug of slugs) {
    const rng = getRng(slug);
    const svgContent = generateWowSVG(slug, rng);
    fs.writeFileSync(path.join(destDir, `${slug}.svg`), svgContent);
    count++;
}

console.log(`✅ ${count} DEEP WEB / WOOOOW 4K SVG files generated!`);
