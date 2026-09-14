#!/usr/bin/env node
/**
 * HTML Bütçesi ve CSS Haricileştirme Optimizatörü
 * Amaç: Post-build sırasında HTML içine enjekte edilen devasa CSS bloklarını
 * harici önbelleklenebilir bir dosyaya çıkararak ilk HTML boyutunu 202KB'tan 110KB'a indirmek.
 */

import fs from 'node:fs';
import path from 'node:path';

const DIST_DIR = path.resolve('dist');
const HOME_HTML = path.join(DIST_DIR, 'index.html');
const ASTRO_DIR = path.join(DIST_DIR, '_astro');

if (!fs.existsSync(HOME_HTML)) {
  console.log('optimize-html-budget: dist/index.html bulunamadı, atlanıyor.');
  process.exit(0);
}

if (!fs.existsSync(ASTRO_DIR)) {
  fs.mkdirSync(ASTRO_DIR, { recursive: true });
}

const targetIds = [
  'sitewide-continuous-grid-v4',
  'mobile-premium-v18-css',
  'home-desktop-premium-v19-css',
  'home-conversion-v20-css',
  'home-native-info-hard-color-v33',
  'home-authority-label-contrast-v34',
  'dual-funnel-home-v17-css',
  'native-seo-infographics',
  'home-original-hero-v21-css',
  'data-contextual-imagery',
  'data-ea-typography-mandate'
];

let html = fs.readFileSync(HOME_HTML, 'utf8');
const initialSize = Buffer.byteLength(html, 'utf8');
const extractedCss = [];

for (const sid of targetIds) {
  const regex = new RegExp(`<style[^>]*${sid}[^>]*>([\\s\\S]*?)</style>`, 'i');
  const match = html.match(regex);
  if (match) {
    const css = match[1].trim();
    extractedCss.push(`/* ${sid} */\n${css}`);
    html = html.replace(match[0], '');
  }
}

if (extractedCss.length > 0) {
  const combinedCss = extractedCss.join('\n\n');
  const bundleFileName = 'site-premium-bundle.css';
  const bundlePath = path.join(ASTRO_DIR, bundleFileName);
  
  fs.writeFileSync(bundlePath, combinedCss, 'utf8');
  
  const linkTag = `<link rel="stylesheet" href="/_astro/${bundleFileName}">`;
  if (!html.includes(linkTag)) {
    html = html.replace('</head>', `  ${linkTag}\n</head>`);
  }
  
  fs.writeFileSync(HOME_HTML, html, 'utf8');
  const finalSize = Buffer.byteLength(html, 'utf8');
  const saved = initialSize - finalSize;
  const pct = ((saved / initialSize) * 100).toFixed(1);
  
  console.log(`✅ [HTML-BUDGET] dist/index.html optimize edildi: ${initialSize} B -> ${finalSize} B (Tasarruf: ${saved} B, %${pct})`);
  console.log(`✅ [HTML-BUDGET] Harici CSS oluşturuldu: /_astro/${bundleFileName} (${Buffer.byteLength(combinedCss, 'utf8')} B)`);
} else {
  console.log('ℹ️ [HTML-BUDGET] Ayrıştırılacak inline style bulunamadı.');
}
