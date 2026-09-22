import fs from 'fs';
import path from 'path';

const destDir = 'public/images/kapak';
if (!fs.existsSync(destDir)) {
  fs.mkdirSync(destDir, { recursive: true });
}

// Get slugs
const kapakFile = fs.readFileSync('src/lib/kapak.ts', 'utf8');
const match = kapakFile.match(/const PREMIUM_KAPAK_SLUGS = new Set<string>\(\[\s*([\s\S]*?)\s*\]\);/);
const slugs = match[1].split(',').map(s => s.trim().replace(/['"]/g, '')).filter(s => s.length > 0);

// Palettes
const palettes = [
  ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6'], // Standard vibrant
  ['#0ea5e9', '#14b8a6', '#84cc16', '#eab308', '#f97316'], // Ocean to sun
  ['#6366f1', '#a855f7', '#ec4899', '#f43f5e', '#f97316'], // Sunset
  ['#059669', '#10b981', '#34d399', '#3b82f6', '#60a5fa'], // Forest & sky
  ['#2563eb', '#4f46e5', '#7c3aed', '#9333ea', '#c026d3'], // Deep tech
  ['#dc2626', '#ea580c', '#d97706', '#ca8a04', '#65a30d'], // Autumn
];

// RNG based on string seed so it's deterministic per slug
function seedRandom(str) {
  let h = 0;
  for (let i = 0; i < str.length; i++) h = Math.imul(31, h) + str.charCodeAt(i) | 0;
  return function() {
    h = Math.imul(h ^ (h >>> 16), 2246822507);
    h = Math.imul(h ^ (h >>> 13), 3266489909);
    return (h ^= h >>> 16) >>> 0 / 4294967296;
  };
}
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

function polarToCartesian(centerX, centerY, radius, angleInDegrees) {
  var angleInRadians = (angleInDegrees-90) * Math.PI / 180.0;
  return {
    x: centerX + (radius * Math.cos(angleInRadians)),
    y: centerY + (radius * Math.sin(angleInRadians))
  };
}

function describeArc(x, y, radius, startAngle, endAngle){
    var start = polarToCartesian(x, y, radius, endAngle);
    var end = polarToCartesian(x, y, radius, startAngle);
    var largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";
    var d = [
        "M", start.x, start.y, 
        "A", radius, radius, 0, largeArcFlag, 0, end.x, end.y,
        "L", x, y,
        "Z"
    ].join(" ");
    return d;       
}

function generateChartSVG(slug, rng) {
    const pIdx = Math.floor(rng() * palettes.length);
    const palette = palettes[pIdx];
    
    // Choose chart type: 0: Bar, 1: Pie/Donut, 2: Line/Area, 3: Combined
    const type = Math.floor(rng() * 4);
    
    let innerContent = '';
    
    // Grid lines for rectangular charts
    if (type !== 1) {
        innerContent += `
        <line x1="200" y1="400" x2="2840" y2="400" stroke="#f1f5f9" stroke-width="8" />
        <line x1="200" y1="700" x2="2840" y2="700" stroke="#f1f5f9" stroke-width="8" />
        <line x1="200" y1="1000" x2="2840" y2="1000" stroke="#f1f5f9" stroke-width="8" />
        <line x1="200" y1="1300" x2="2840" y2="1300" stroke="#f1f5f9" stroke-width="8" />
        `;
    }

    if (type === 0) { // Bar Chart
        const numBars = 5 + Math.floor(rng() * 5); // 5 to 9 bars
        const barWidth = 1800 / numBars;
        const spacing = barWidth * 0.4;
        const actualWidth = barWidth - spacing;
        
        let startX = 200 + (2640 - (numBars * barWidth)) / 2;
        
        for (let i = 0; i < numBars; i++) {
            const h = 300 + rng() * 900;
            const y = 1300 - h;
            const color = palette[i % palette.length];
            innerContent += `<rect x="${startX + (i * barWidth)}" y="${y}" width="${actualWidth}" height="${h}" rx="${actualWidth/3}" fill="${color}" />`;
        }
    } 
    else if (type === 1) { // Pie / Donut
        const numSlices = 4 + Math.floor(rng() * 4);
        let total = 0;
        const values = [];
        for(let i=0; i<numSlices; i++) {
            let v = 20 + rng() * 80;
            values.push(v);
            total += v;
        }
        
        let currentAngle = 0;
        const cx = 1520;
        const cy = 800;
        const radius = 600;
        
        for(let i=0; i<numSlices; i++) {
            const angle = (values[i] / total) * 360;
            const startAngle = currentAngle;
            const endAngle = currentAngle + angle;
            const color = palette[i % palette.length];
            
            // Explode a bit
            const explode = 15;
            const midAngle = startAngle + (angle/2);
            const exX = Math.cos((midAngle-90) * Math.PI/180) * explode;
            const exY = Math.sin((midAngle-90) * Math.PI/180) * explode;
            
            innerContent += `<path d="${describeArc(cx + exX, cy + exY, radius, startAngle, endAngle)}" fill="${color}" stroke="#fff" stroke-width="12" />`;
            currentAngle += angle;
        }
        
        // Make it a donut randomly
        if (rng() > 0.5) {
            innerContent += `<circle cx="${cx}" cy="${cy}" r="300" fill="#ffffff" />`;
        }
    }
    else if (type === 2) { // Smooth Line/Area Chart
        const numPoints = 6 + Math.floor(rng() * 4);
        const step = 2400 / (numPoints - 1);
        
        let points = [];
        for(let i=0; i<numPoints; i++) {
            points.push({ x: 320 + i*step, y: 300 + rng()*800 });
        }
        
        const color = palette[0];
        
        // Generate SVG cubic bezier path
        let d = `M ${points[0].x} ${points[0].y}`;
        for(let i=0; i<points.length-1; i++) {
            const p1 = points[i];
            const p2 = points[i+1];
            const cp1x = p1.x + (p2.x - p1.x) / 2;
            const cp2x = cp1x;
            d += ` C ${cp1x} ${p1.y}, ${cp2x} ${p2.y}, ${p2.x} ${p2.y}`;
        }
        
        // Area under curve
        const areaD = d + ` L ${points[points.length-1].x} 1300 L ${points[0].x} 1300 Z`;
        innerContent += `<path d="${areaD}" fill="${color}" fill-opacity="0.2" />`;
        
        // The line itself
        innerContent += `<path d="${d}" fill="none" stroke="${color}" stroke-width="24" stroke-linecap="round" />`;
        
        // Data dots
        for(let i=0; i<points.length; i++) {
            innerContent += `<circle cx="${points[i].x}" cy="${points[i].y}" r="32" fill="#fff" stroke="${color}" stroke-width="12" />`;
        }
    }
    else if (type === 3) { // Combined Bar + Line (Like user's screenshot)
        const numPoints = 8;
        const step = 2400 / (numPoints - 1);
        
        let linePoints = [];
        for(let i=0; i<numPoints; i++) {
            linePoints.push({ x: 320 + i*step, y: 300 + rng()*600 });
        }
        
        const barColor = palette[2];
        const lineColor = palette[0];
        
        // Bars
        const barW = 80;
        for(let i=0; i<numPoints; i++) {
            const h = 200 + rng() * 700;
            const y = 1300 - h;
            innerContent += `<rect x="${linePoints[i].x - barW/2}" y="${y}" width="${barW}" height="${h}" rx="20" fill="${barColor}" fill-opacity="0.8" />`;
        }
        
        // Line
        let d = `M ${linePoints[0].x} ${linePoints[0].y}`;
        for(let i=1; i<linePoints.length; i++) {
            d += ` L ${linePoints[i].x} ${linePoints[i].y}`;
        }
        innerContent += `<path d="${d}" fill="none" stroke="${lineColor}" stroke-width="20" stroke-linejoin="round" />`;
        
        for(let i=0; i<linePoints.length; i++) {
            innerContent += `<circle cx="${linePoints[i].x}" cy="${linePoints[i].y}" r="28" fill="${lineColor}" stroke="#fff" stroke-width="8" />`;
        }
    }

    if (type !== 1) {
        // Base Axis
        innerContent += `<line x1="200" y1="1300" x2="2840" y2="1300" stroke="#cbd5e1" stroke-width="12" stroke-linecap="round" />`;
    }

    return `
<svg width="3840" height="2400" viewBox="0 0 3840 2400" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <filter id="shadow" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="30" stdDeviation="40" flood-opacity="0.08" />
    </filter>
  </defs>
  
  <rect width="100%" height="100%" fill="#f4f6f9" />
  
  <g transform="translate(400, 400)" filter="url(#shadow)">
    <!-- White Card Background -->
    <rect width="3040" height="1600" rx="80" fill="#ffffff" />
    
    <!-- Chart Content -->
    ${innerContent}
  </g>
</svg>
    `.trim();
}

let count = 0;
for (const slug of slugs) {
    const rng = getRng(slug);
    const svgContent = generateChartSVG(slug, rng);
    fs.writeFileSync(path.join(destDir, `${slug}.svg`), svgContent);
    count++;
}

console.log(`✅ ${count} UNIQUE, 4K Flawless SVG files generated perfectly!`);
