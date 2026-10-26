#!/usr/bin/env node
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync, cpSync, rmSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';

const root = process.cwd();
const premiumSiteDir = resolve(root, 'excelarsiv-premium-tum-sayfalar/site');
const distDir = resolve(root, 'dist');

if (!existsSync(premiumSiteDir)) {
  console.error('HATA: excelarsiv-premium-tum-sayfalar/site bulunamadı!');
  process.exit(1);
}

console.log('=== PREMİUM TASARIM ENTEGRASYONU BAŞLATILIYOR ===');

// 1. Copy assets to dist/assets
const premiumAssetsDir = join(premiumSiteDir, 'assets');
const distAssetsDir = join(distDir, 'assets');
if (existsSync(premiumAssetsDir)) {
  cpSync(premiumAssetsDir, distAssetsDir, { recursive: true });
  console.log('[PASS] assets klasörü dist/assets içine kopyalandı.');
}

// 2. Read sitemap to get canonical 334 URLs
const sitemap = readFileSync(join(distDir, 'sitemap-pages.xml'), 'utf8') + readFileSync(join(distDir, 'sitemap-products.xml'), 'utf8');
const locRegex = /<loc>https:\/\/excelarsiv\.com([^<]*)<\/loc>/g;
const validCanonicalRoutes = new Set();
let m;
while ((m = locRegex.exec(sitemap)) !== null) {
  validCanonicalRoutes.add(m[1] === '' ? '/' : m[1]);
}

const defaultRobotsMeta = '<meta name="robots" content="index, follow, max-image-preview:large, max-snippet:-1, max-video-preview:-1">';

function getJsonLd(html) {
  const match = html.match(/<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/i);
  return match ? match[0] : '';
}

function getFallbackJsonLd(url, title, description) {
  return `<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "WebPage",
  "name": ${JSON.stringify(title || "ExcelArşiv")},
  "description": ${JSON.stringify(description || "")},
  "url": ${JSON.stringify(url)},
  "isPartOf": {
    "@type": "WebSite",
    "name": "ExcelArşiv",
    "url": "https://excelarsiv.com/"
  }
}
</script>`;
}

let updatedPagesCount = 0;

for (const route of validCanonicalRoutes) {
  const relPath = route === '/' ? 'index.html' : join(route.replace(/^\//, ''), 'index.html');
  const premiumFile = join(premiumSiteDir, relPath);
  const distFile = join(distDir, relPath);

  if (existsSync(premiumFile)) {
    let premiumHtml = readFileSync(premiumFile, 'utf8');
    
    // Existing original JSON-LD from dist if available
    let jsonLd = '';
    if (existsSync(distFile)) {
      const originalDistHtml = readFileSync(distFile, 'utf8');
      jsonLd = getJsonLd(originalDistHtml);
    }
    if (!jsonLd) {
      const titleMatch = premiumHtml.match(/<title>([^<]*)<\/title>/i);
      const descMatch = premiumHtml.match(/<meta[^>]*name=["']description["'][^>]*content=["']([^"']*)["']/i);
      jsonLd = getFallbackJsonLd(`https://excelarsiv.com${route}`, titleMatch ? titleMatch[1] : '', descMatch ? descMatch[1] : '');
    }

    // Inject robots meta if not present
    if (!/<meta[^>]*name=["']robots["']/i.test(premiumHtml)) {
      premiumHtml = premiumHtml.replace('</head>', `  ${defaultRobotsMeta}\n</head>`);
    }

    // Inject JSON-LD if not present
    if (!/<script[^>]*type=["']application\/ld\+json["']/i.test(premiumHtml)) {
      premiumHtml = premiumHtml.replace('</head>', `  ${jsonLd}\n</head>`);
    }

    writeFileSync(distFile, premiumHtml, 'utf8');
    updatedPagesCount++;
  }
}

// 404.html
const premium404 = join(premiumSiteDir, '404.html');
if (existsSync(premium404)) {
  let h404 = readFileSync(premium404, 'utf8');
  if (!/<meta[^>]*name=["']robots["']/i.test(h404)) {
    h404 = h404.replace('</head>', `  <meta name="robots" content="noindex, follow">\n</head>`);
  }
  const jsonLd404 = getFallbackJsonLd('https://excelarsiv.com/404', 'Sayfa Bulunamadı | ExcelArşiv', 'Aradığınız sayfa bulunamadı.');
  if (!/<script[^>]*type=["']application\/ld\+json["']/i.test(h404)) {
    h404 = h404.replace('</head>', `  ${jsonLd404}\n</head>`);
  }
  writeFileSync(join(distDir, '404.html'), h404, 'utf8');
  console.log('[PASS] 404.html güncellendi.');
}

// Clean up any extra folders in dist that violate sitemap parity
if (existsSync(join(distDir, '404'))) {
  rmSync(join(distDir, '404'), { recursive: true, force: true });
}
if (existsSync(join(distDir, 'tek-dosya.html'))) {
  rmSync(join(distDir, 'tek-dosya.html'), { force: true });
}

// Remove any non-canonical subfolders from dist/pratik-programlar
const distPratik = join(distDir, 'pratik-programlar');
if (existsSync(distPratik)) {
  for (const entry of readdirSync(distPratik)) {
    const p = join(distPratik, entry);
    if (statSync(p).isDirectory()) {
      rmSync(p, { recursive: true, force: true });
    }
  }
}

// Remove non-canonical sablonlar/raporlama if present
if (existsSync(join(distDir, 'sablonlar', 'raporlama'))) {
  rmSync(join(distDir, 'sablonlar', 'raporlama'), { recursive: true, force: true });
}

console.log(`[PASS] ${updatedPagesCount} sayfa premium tasarımla başarıyla entegre edildi.`);
