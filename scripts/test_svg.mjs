import fs from 'fs';

const svg = `
<svg width="3840" height="2400" viewBox="0 0 3840 2400" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="bg" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#f8fafc" />
      <stop offset="100%" stop-color="#e2e8f0" />
    </linearGradient>
    <linearGradient id="barGrad" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" stop-color="#3b82f6" />
      <stop offset="100%" stop-color="#1d4ed8" />
    </linearGradient>
    <filter id="shadow" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="20" stdDeviation="20" flood-opacity="0.1" />
    </filter>
  </defs>
  
  <rect width="100%" height="100%" fill="url(#bg)" />
  
  <g transform="translate(400, 400)" filter="url(#shadow)">
    <rect width="3040" height="1600" rx="60" fill="#ffffff" />
    
    <!-- Grid -->
    <line x1="200" y1="400" x2="2840" y2="400" stroke="#f1f5f9" stroke-width="8" />
    <line x1="200" y1="800" x2="2840" y2="800" stroke="#f1f5f9" stroke-width="8" />
    <line x1="200" y1="1200" x2="2840" y2="1200" stroke="#f1f5f9" stroke-width="8" />
    
    <!-- Bars -->
    <rect x="400" y="800" width="160" height="600" rx="40" fill="url(#barGrad)" />
    <rect x="800" y="500" width="160" height="900" rx="40" fill="url(#barGrad)" />
    <rect x="1200" y="900" width="160" height="500" rx="40" fill="url(#barGrad)" />
    <rect x="1600" y="400" width="160" height="1000" rx="40" fill="url(#barGrad)" />
    <rect x="2000" y="700" width="160" height="700" rx="40" fill="url(#barGrad)" />
    <rect x="2400" y="300" width="160" height="1100" rx="40" fill="url(#barGrad)" />
    
    <!-- Axis -->
    <line x1="200" y1="1400" x2="2840" y2="1400" stroke="#cbd5e1" stroke-width="12" stroke-linecap="round" />
  </g>
</svg>
`;

fs.writeFileSync('test.svg', svg);
console.log('Test SVG created');
