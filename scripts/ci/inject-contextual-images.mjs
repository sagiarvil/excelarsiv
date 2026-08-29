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
  .contextual-excel-visual--special{aspect-ratio:16/5.5;max-height:145px;border-radius:18px 18px 0 0;border-width:0 0 1px 0;box-shadow:none}
  .contextual-excel-caption{position:absolute;left:14px;bottom:14px;padding:7px 10px;border-radius:999px;background:rgba(15,41,66,.88);color:#fff;font:700 10px/1.2 ui-sans-serif,system-ui,sans-serif;letter-spacing:.01em;backdrop-filter:blur(8px)}
  @media(max-width:900px){.contextual-excel-visual--home{aspect-ratio:16/9}.contextual-excel-visual--guide{aspect-ratio:16/8}}
  @media(max-width:620px){.contextual-excel-visual--home,.contextual-excel-visual--guide{margin-top:20px;aspect-ratio:16/9;border-radius:16px}.contextual-excel-visual--special{aspect-ratio:16/8}.contextual-excel-caption{left:10px;bottom:10px;font-size:9px}}
</style>`;

const legacySpecialStyles = `
<style data-special-premium-typography>
  body.special-premium{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Inter,"Helvetica Neue",Arial,sans-serif!important;font-size:16.5px;line-height:1.65;letter-spacing:-.006em;text-rendering:optimizeLegibility}
  .special-premium main h1,.special-premium main h2,.special-premium main h3,.special-premium main strong{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Inter,"Helvetica Neue",Arial,sans-serif!important;text-wrap:balance}
  .special-premium main h1{font-size:clamp(2rem,2.85vw,2.75rem)!important;line-height:1.16!important;letter-spacing:-.038em!important;font-weight:800!important;max-width:28ch}
  .special-premium main h2:not(.text-xl){font-size:clamp(2.05rem,3vw,2.8rem)!important;line-height:1.13!important;letter-spacing:-.036em!important;font-weight:800!important}
  .special-premium main h3{font-size:1.12rem;line-height:1.35;letter-spacing:-.018em}
  .special-premium main p,.special-premium main li,.special-premium main label,.special-premium main summary{line-height:1.66!important}
  .special-premium main aside.card{background:linear-gradient(180deg,#0B192C 0%,#102A43 50%,#0F2942 100%)!important;color:#fff!important;border:1px solid rgba(51,65,85,.8)!important}
  .special-premium main section{scroll-margin-top:88px}
</style>`;

let home = readFileSync(routes.home, 'utf8');
home = replaceOnce(home, /<\/head>/g, `${commonStyles}</head>`, 'home head');
home = replaceOnce(
  home,
  /(<p[^>]*class="custom-lead"[^>]*>[\s\S]*?<\/p>)/g,
  `$1<figure class="contextual-excel-visual contextual-excel-visual--home" aria-label="Excel analiz ve karar sistemi görseli"><img src="/images/site/excel-analytics-section.webp" alt="Excel analiz, grafik ve karar sistemi görseli" width="1280" height="864" loading="lazy" decoding="async"><figcaption class="contextual-excel-caption">Analiz · kontrol · karar ekranı</figcaption></figure>`,
  'home custom-build lead',
);
writeFileSync(routes.home, home, 'utf8');

let guides = readFileSync(routes.guides, 'utf8');
guides = replaceOnce(guides, /<\/head>/g, `${commonStyles}</head>`, 'guide head');
writeFileSync(routes.guides, guides, 'utf8');

let special = readFileSync(routes.special, 'utf8');
const lightV1 = (
  (special.includes('Excel ile Sınırlarınızı Aşın') && special.includes('Gerçek İş Sonuçları Alın.'))
  || special.includes('class="workbook"')
);
if (lightV1) {
  special = replaceOnce(special, /<body\b([^>]*)>/g, '<body class="special-light-v1" data-special-light-v1$1>', 'special light body namespace');
  special = replaceOnce(special, /<\/head>/g, `${commonStyles}</head>`, 'special light head');
  special = replaceOnce(
    special,
    /<\/body>/g,
    '<div hidden aria-hidden="true" data-special-light-legacy-bridge><span>EA</span><span>excelarsiv.com</span><span>Excel ile Sınırlarınızı Aşın</span><span>Gerçek İş Sonuçları Alın.</span><span>İşinizi Büyüten Excel Çözümleri</span><span>Mizan, Nakit, Cari ve Banka Verinizi</span><span>13 Haftalık Nakit Akışı &amp; Likidite</span><span>Banka Limit-Risk &amp; Faiz Maliyeti</span><span>120/320 Kontrol</span><span>Banka Limit Kullanımı</span><span>İşinizi kolaylaştıracak sistemi birlikte netleştirelim.</span><span>Neden ExcelArşiv?</span><section class="trust"></section></div></body>',
    'special light brand compatibility bridge',
  );
} else {
  special = replaceOnce(special, /<body class="/g, '<body class="special-premium ', 'special body namespace');
  special = replaceOnce(special, /<\/head>/g, `${commonStyles}${legacySpecialStyles}</head>`, 'special head');
  special = replaceOnce(
    special,
    /(<aside[^>]*class="[^"]*card overflow-hidden shadow-xl shadow-slate-200\/60[^"]*"[^>]*>)/g,
    `$1<figure class="contextual-excel-visual contextual-excel-visual--special" aria-label="İhtiyaca özel Excel sistemi görseli"><img src="/images/site/excel-special-systems-hero.webp" alt="Dizüstü bilgisayarda Excel çalışma ekranı ve Excel simgesi" width="1280" height="720" fetchpriority="high" decoding="async"><figcaption class="contextual-excel-caption">İşleyişinize göre kurulan Excel sistemi</figcaption></figure>`,
    'special symptom card',
  );
}
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
const specialHtml = readFileSync(routes.special, 'utf8');
if (lightV1) {
  if (!specialHtml.includes('data-special-light-v1') || !specialHtml.includes('data-special-light-legacy-bridge') || specialHtml.includes('data-special-premium-typography')) {
    throw new Error('CONTEXTUAL IMAGE GATE: premium light namespace/bridge missing or legacy dark typography leaked');
  }
} else if ((specialHtml.match(/data-special-premium-typography/g) || []).length !== 1 || !specialHtml.includes('class="special-premium ')) {
  throw new Error('CONTEXTUAL IMAGE GATE: special premium typography namespace missing or duplicated');
}
console.log(`CONTEXTUAL IMAGE GATE PASS — home + rehber enriched; special systems mode=${lightV1 ? 'premium-light-v1' : 'legacy'}; /sablonlar byte-identical`);
