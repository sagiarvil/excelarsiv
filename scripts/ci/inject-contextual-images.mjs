#!/usr/bin/env node
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'dist';
const routes = {
  home: join(DIST, 'index.html'),
  guides: join(DIST, 'rehber', 'index.html'),
  special: join(DIST, 'ozel-excel-sistemleri', 'index.html'),
  catalog: join(DIST, 'sablonlar', 'index.html'),
};

const assets = [
  join('public', 'images', 'site', 'excel-analytics-section.webp'),
  join('public', 'images', 'site', 'excel-guide-banner.webp'),
  join('public', 'images', 'site', 'excel-special-systems-hero.webp'),
];

function hash(path) {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
}

function mustExist(path) {
  if (!existsSync(path)) throw new Error(`CONTEXTUAL IMAGE GATE: missing ${path}`);
}

function replaceOnce(html, pattern, replacement, label) {
  const matches = html.match(pattern);
  if (!matches || matches.length !== 1) {
    throw new Error(`CONTEXTUAL IMAGE GATE: expected exactly one ${label} anchor; found ${matches?.length ?? 0}`);
  }
  return html.replace(pattern, replacement);
}

for (const path of [...Object.values(routes), ...assets]) mustExist(path);

const catalogBefore = hash(routes.catalog);

const commonStyles = `
<style data-contextual-imagery>
  .contextual-excel-visual{position:relative;overflow:hidden;border:1px solid rgba(22,101,52,.16);background:#eef6f1;box-shadow:0 18px 48px rgba(15,50,31,.10)}
  .contextual-excel-visual img{display:block;width:100%;height:100%;object-fit:cover}
  .contextual-excel-visual--home{margin-top:24px;aspect-ratio:16/10;border-radius:20px}
  .contextual-excel-visual--guide{margin-top:30px;aspect-ratio:16/6;border-radius:20px}
  .contextual-excel-visual--special{aspect-ratio:16/7;border-radius:18px 18px 0 0;border-width:0 0 1px 0;box-shadow:none}
  .contextual-excel-caption{position:absolute;left:14px;bottom:14px;padding:7px 10px;border-radius:999px;background:rgba(15,41,66,.88);color:#fff;font:700 10px/1.2 ui-sans-serif,system-ui,sans-serif;letter-spacing:.01em;backdrop-filter:blur(8px)}
  @media(max-width:900px){.contextual-excel-visual--home{aspect-ratio:16/9}.contextual-excel-visual--guide{aspect-ratio:16/8}}
  @media(max-width:620px){.contextual-excel-visual--home,.contextual-excel-visual--guide{margin-top:20px;aspect-ratio:16/9;border-radius:16px}.contextual-excel-visual--special{aspect-ratio:16/8}.contextual-excel-caption{left:10px;bottom:10px;font-size:9px}}
</style>`;

// Home: add a single premium visual only inside the existing "Size Özel" copy column.
let home = readFileSync(routes.home, 'utf8');
home = replaceOnce(home, /<\/head>/g, `${commonStyles}</head>`, 'home head');
home = replaceOnce(
  home,
  /(<p class="custom-lead">[\s\S]*?<\/p>)/g,
  `$1<figure class="contextual-excel-visual contextual-excel-visual--home" aria-label="Excel analiz ve karar sistemi görseli"><img src="/images/site/excel-analytics-section.webp" alt="Excel analiz, grafik ve karar sistemi görseli" width="1280" height="864" loading="lazy" decoding="async"><figcaption class="contextual-excel-caption">Analiz · kontrol · karar ekranı</figcaption></figure>`,
  'home custom-build lead',
);
writeFileSync(routes.home, home, 'utf8');

// Rehber: full-width symmetric banner under the existing intro stats; grid/cards remain untouched.
let guides = readFileSync(routes.guides, 'utf8');
guides = replaceOnce(guides, /<\/head>/g, `${commonStyles}</head>`, 'guide head');
guides = replaceOnce(
  guides,
  /(<div class="hub-stats mt-8">[\s\S]*?<strong>1:1<\/strong><span>rehber → ürün bağlantısı<\/span><\/div>\s*<\/div>)/g,
  `$1<figure class="contextual-excel-visual contextual-excel-visual--guide" aria-label="Excel uygulama rehberleri görseli"><img src="/images/site/excel-guide-banner.webp" alt="Microsoft Excel çalışma ekranı ve Excel simgesi" width="1280" height="720" loading="eager" fetchpriority="high" decoding="async"><figcaption class="contextual-excel-caption">Uygulama rehberleri · gerçek çalışma mantığı</figcaption></figure>`,
  'guide intro stats',
);
writeFileSync(routes.guides, guides, 'utf8');

// Special systems: visual becomes the top layer of the existing symptom card; no column/grid changes.
let special = readFileSync(routes.special, 'utf8');
special = replaceOnce(special, /<\/head>/g, `${commonStyles}</head>`, 'special head');
special = replaceOnce(
  special,
  /(<aside class="card overflow-hidden shadow-xl shadow-slate-200\/60">)/g,
  `$1<figure class="contextual-excel-visual contextual-excel-visual--special" aria-label="İhtiyaca özel Excel sistemi görseli"><img src="/images/site/excel-special-systems-hero.webp" alt="Dizüstü bilgisayarda Excel çalışma ekranı ve Excel simgesi" width="1280" height="720" fetchpriority="high" decoding="async"><figcaption class="contextual-excel-caption">İşleyişinize göre kurulan Excel sistemi</figcaption></figure>`,
  'special symptom card',
);
writeFileSync(routes.special, special, 'utf8');

const catalogAfter = hash(routes.catalog);
if (catalogBefore !== catalogAfter) {
  throw new Error(`CONTEXTUAL IMAGE GATE: /sablonlar changed unexpectedly (${catalogBefore.slice(0,12)} -> ${catalogAfter.slice(0,12)})`);
}

for (const [name, path] of Object.entries({ home: routes.home, guides: routes.guides, special: routes.special })) {
  const html = readFileSync(path, 'utf8');
  const count = (html.match(/data-contextual-imagery/g) || []).length;
  if (count !== 1) throw new Error(`CONTEXTUAL IMAGE GATE: ${name} style injection count=${count}`);
}

console.log('CONTEXTUAL IMAGE GATE PASS — home, rehber and special systems enriched; /sablonlar byte-identical');
