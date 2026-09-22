import fs from 'fs';
import path from 'path';

// Ortak renk paleti ve yardımcılar
const COLORS = {
  bg: '#ffffff',
  surface: '#f8fafc',
  surfaceBorder: '#e2e8f0',
  navy: '#0f172a',
  navyLight: '#1e293b',
  textMuted: '#64748b',
  textLight: '#94a3b8',
  excelGreen: '#107c41',
  excelDark: '#0d5c3a',
  excelLight: '#dcfce7',
  excelBorder: '#86efac',
  blue: '#2563eb',
  blueLight: '#dbeafe',
  amber: '#d97706',
  amberLight: '#fef3c7',
  red: '#dc2626',
  redLight: '#fee2e2',
  purple: '#7c3aed',
  purpleLight: '#f3e8ff',
  teal: '#0d9488',
  tealLight: '#ccfbf1',
  indigo: '#4f46e5',
  indigoLight: '#e0e7ff',
};

// Vektör yardımcı fonksiyonları
export function svgWrapper(content, title = '') {
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" width="100%" height="100%">
  <defs>
    <linearGradient id="excelGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#107c41" />
      <stop offset="100%" stop-color="#0d5c3a" />
    </linearGradient>
    <linearGradient id="blueGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#3b82f6" />
      <stop offset="100%" stop-color="#1d4ed8" />
    </linearGradient>
    <linearGradient id="amberGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#f59e0b" />
      <stop offset="100%" stop-color="#b45309" />
    </linearGradient>
    <linearGradient id="redGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#ef4444" />
      <stop offset="100%" stop-color="#b91c1c" />
    </linearGradient>
    <linearGradient id="purpleGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#8b5cf6" />
      <stop offset="100%" stop-color="#6d28d9" />
    </linearGradient>
    <filter id="softShadow" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="8" stdDeviation="12" flood-color="#0f172a" flood-opacity="0.06" />
    </filter>
    <filter id="glowGreen" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="4" stdDeviation="8" flood-color="#107c41" flood-opacity="0.25" />
    </filter>
  </defs>

  <!-- Saf Açık Tema Zemin (1000x1000) -->
  <rect width="1000" height="1000" fill="#ffffff" />

  <!-- Dış Çerçeve & Hafif Grid Kılavuzu (820x820 Çizim Alanı: x=90, y=90) -->
  <rect x="90" y="90" width="820" height="820" rx="28" fill="#ffffff" stroke="#e2e8f0" stroke-width="2" filter="url(#softShadow)" />
  <rect x="90" y="90" width="820" height="820" rx="28" fill="none" stroke="#f1f5f9" stroke-width="1" />

  <!-- Çizim İçeriği -->
  ${content}

</svg>`;
}

// Polar koordinat dönüştürücü
export function polarToCartesian(centerX, centerY, radius, angleInDegrees) {
  const angleInRadians = ((angleInDegrees - 90) * Math.PI) / 180.0;
  return {
    x: centerX + radius * Math.cos(angleInRadians),
    y: centerY + radius * Math.sin(angleInRadians),
  };
}

// Donut segmenti
export function describeArc(x, y, innerRadius, outerRadius, startAngle, endAngle) {
  const start = polarToCartesian(x, y, outerRadius, endAngle);
  const end = polarToCartesian(x, y, outerRadius, startAngle);
  const startInner = polarToCartesian(x, y, innerRadius, endAngle);
  const endInner = polarToCartesian(x, y, innerRadius, startAngle);

  const largeArcFlag = endAngle - startAngle <= 180 ? '0' : '1';

  return [
    'M', start.x, start.y,
    'A', outerRadius, outerRadius, 0, largeArcFlag, 0, end.x, end.y,
    'L', endInner.x, endInner.y,
    'A', innerRadius, innerRadius, 0, largeArcFlag, 1, startInner.x, startInner.y,
    'Z'
  ].join(' ');
}
