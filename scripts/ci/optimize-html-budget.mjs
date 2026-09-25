#!/usr/bin/env node
/**
 * HTML Bütçesi ve CSS Haricileştirme Optimizatörü
 * Amaç: Post-build sırasında HTML içine enjekte edilen devasa CSS bloklarını
 * harici önbelleklenebilir bir dosyaya çıkararak ilk HTML boyutunu 202KB'tan 110KB'a indirmek.
 */

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';

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

// Extract any remaining inline <style> tags to further reduce HTML payload
const genericStyleRegex = /<style\b[^>]*>([\s\S]*?)<\/style>/gi;
let sm;
while ((sm = genericStyleRegex.exec(html)) !== null) {
  const css = sm[1].trim();
  if (css.length > 0) {
    extractedCss.push(`/* inlined-style */\n${css}`);
  }
}
html = html.replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '');

// Collapse redundant indentation whitespace between tags to stay strictly under the 250KB ceiling
html = html.replace(/>\s{2,}</g, '> <');

if (extractedCss.length > 0) {
  const combinedCss = extractedCss.join('\n\n');
  const bundleHash = crypto.createHash('md5').update(combinedCss).digest('hex').slice(0, 8);
  const hashedBundleFileName = `site-premium-bundle.${bundleHash}.css`;
  const hashedBundlePath = path.join(ASTRO_DIR, hashedBundleFileName);
  const fallbackBundlePath = path.join(ASTRO_DIR, 'site-premium-bundle.css');
  
  fs.writeFileSync(hashedBundlePath, combinedCss, 'utf8');
  fs.writeFileSync(fallbackBundlePath, combinedCss, 'utf8');
  
  const linkTag = `<link rel="stylesheet" href="/_astro/${hashedBundleFileName}">`;
  // Clean any previous site-premium-bundle link tag if present
  html = html.replace(/<link rel="stylesheet" href="\/_astro\/site-premium-bundle[^"]*">\n?/g, '');
  if (!html.includes(linkTag)) {
    html = html.replace('</head>', `  ${linkTag}\n</head>`);
  }
  
  fs.writeFileSync(HOME_HTML, html, 'utf8');
  const finalSize = Buffer.byteLength(html, 'utf8');
  const saved = initialSize - finalSize;
  const pct = ((saved / initialSize) * 100).toFixed(1);
  
  console.log(`✅ [HTML-BUDGET] dist/index.html optimize edildi: ${initialSize} B -> ${finalSize} B (Tasarruf: ${saved} B, %${pct})`);
  console.log(`✅ [HTML-BUDGET] Harici CSS oluşturuldu: /_astro/${hashedBundleFileName} (${Buffer.byteLength(combinedCss, 'utf8')} B)`);
} else {
  console.log('ℹ️ [HTML-BUDGET] Ayrıştırılacak inline style bulunamadı.');
}
