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

const specialPremiumStyles = `
<style data-special-premium-typography>
  body.special-premium{
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Inter,"Helvetica Neue",Arial,sans-serif!important;
    font-size:16.5px;
    line-height:1.65;
    letter-spacing:-.006em;
    font-kerning:normal;
    text-rendering:optimizeLegibility;
  }
  .special-premium main{font-feature-settings:"kern" 1,"liga" 1,"calt" 1}
  .special-premium main h1,
  .special-premium main h2,
  .special-premium main h3,
  .special-premium main strong{
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Inter,"Helvetica Neue",Arial,sans-serif!important;
    text-wrap:balance;
  }
  .special-premium main h1{
    font-size:clamp(2rem,2.85vw,2.75rem)!important;
    line-height:1.16!important;
    letter-spacing:-.038em!important;
    font-weight:800!important;
    max-width:28ch;
  }
  .special-premium main h2:not(.text-xl){
    font-size:clamp(2.05rem,3vw,2.8rem)!important;
    line-height:1.13!important;
    letter-spacing:-.036em!important;
    font-weight:800!important;
  }
  .special-premium main h2.text-xl{
    font-size:1.35rem!important;
    line-height:1.28!important;
    letter-spacing:-.024em!important;
  }
  .special-premium main h3{
    font-size:1.12rem;
    line-height:1.35;
    letter-spacing:-.018em;
  }
  .special-premium main p,
  .special-premium main li,
  .special-premium main label,
  .special-premium main summary{
    line-height:1.66!important;
  }
  .special-premium main p.text-slate-600,
  .special-premium main p.text-slate-500{
    letter-spacing:.001em;
  }
  .special-premium .text-xs{
    font-size:.82rem!important;
    line-height:1.56!important;
  }
  .special-premium .text-sm{
    font-size:.96rem!important;
    line-height:1.6!important;
  }
  .special-premium .text-base{
    font-size:1.04rem!important;
    line-height:1.7!important;
  }
  .special-premium .text-lg{
    font-size:1.14rem!important;
    line-height:1.72!important;
  }
  .special-premium main .font-mono{
    font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,"Liberation Mono",monospace!important;
    letter-spacing:.105em!important;
    line-height:1.45!important;
  }
  .special-premium main > section > div[class*="max-w-"]{
    padding-top:1.25rem!important;
    padding-bottom:1.25rem!important;
  }
  .special-premium main > section[class~="py-16"],
  .special-premium main > section[class~="py-20"],
  .special-premium main > section[class~="py-24"]{
    padding-top:1.75rem!important;
    padding-bottom:1.75rem!important;
  }
  .special-premium main > section:first-child > div[class*="max-w-"]{
    padding-top:1.5rem!important;
    padding-bottom:1.5rem!important;
  }
  .special-premium main .card{
    border-radius:20px;
  }
  .special-premium main aside.card{
    background: linear-gradient(180deg, #0B192C 0%, #102A43 50%, #0F2942 100%) !important;
    color: #ffffff !important;
    border: 1px solid rgba(51, 65, 85, 0.8) !important;
  }
  .special-premium main .card p + p{
    margin-top:.45rem;
  }
  .special-premium main details summary{
    padding-top:1.1rem;
    padding-bottom:1.1rem;
    letter-spacing:-.014em;
  }
  .special-premium main a,
  .special-premium main button{
    letter-spacing:-.012em;
  }
  .special-premium main section{
    scroll-margin-top:88px;
  }
  @media(min-width:1024px){
    .special-premium main > section:first-child .grid{gap:3.4rem!important}
    .special-premium main > section:first-child p.text-slate-600{max-width:66ch}
  }
  @media(max-width:900px){
    .special-premium main h1{font-size:clamp(2.55rem,8vw,3.55rem)!important;max-width:18ch}
    .special-premium main h2:not(.text-xl){font-size:clamp(1.9rem,5.8vw,2.45rem)!important}
    .special-premium main > section > div[class*="max-w-"],
    .special-premium main > section[class~="py-16"],
    .special-premium main > section[class~="py-20"],
    .special-premium main > section[class~="py-24"]{padding-top:1.25rem!important;padding-bottom:1.25rem!important}
  }
  @media(max-width:620px){
    body.special-premium{font-size:16px;line-height:1.62}
    .special-premium main h1{font-size:clamp(2.25rem,11vw,3rem)!important;line-height:1.08!important;letter-spacing:-.044em!important}
    .special-premium main h2:not(.text-xl){font-size:1.9rem!important;line-height:1.16!important}
    .special-premium main > section > div[class*="max-w-"],
    .special-premium main > section[class~="py-16"],
    .special-premium main > section[class~="py-20"],
    .special-premium main > section[class~="py-24"]{padding-top:1rem!important;padding-bottom:1rem!important}
    .special-premium .text-sm{font-size:.94rem!important}
    .special-premium .text-xs{font-size:.8rem!important}
  }
</style>`;

// Astro adds scoped data-* attributes in production; anchors therefore allow extra attributes.
// Home: add a single premium visual only inside the existing "Size Özel" copy column.
let home = readFileSync(routes.home, 'utf8');
home = replaceOnce(home, /<\/head>/g, `${commonStyles}</head>`, 'home head');
home = replaceOnce(
  home,
  /(<p[^>]*class="custom-lead"[^>]*>[\s\S]*?<\/p>)/g,
  `$1<figure class="contextual-excel-visual contextual-excel-visual--home" aria-label="Excel analiz ve karar sistemi görseli"><img src="/images/site/excel-analytics-section.webp" alt="Excel analiz, grafik ve karar sistemi görseli" width="1280" height="864" loading="lazy" decoding="async"><figcaption class="contextual-excel-caption">Analiz · kontrol · karar ekranı</figcaption></figure>`,
  'home custom-build lead',
);
writeFileSync(routes.home, home, 'utf8');

// Rehber: keep common styles without the banner visual.
let guides = readFileSync(routes.guides, 'utf8');
guides = replaceOnce(guides, /<\/head>/g, `${commonStyles}</head>`, 'guide head');
writeFileSync(routes.guides, guides, 'utf8');

// Special systems has two supported contracts:
// - legacy DOM: contextual visual + legacy typography namespace are injected here.
// - light premium v2: its product-window is already the contextual product visual and page typography is self-contained.
let special = readFileSync(routes.special, 'utf8');
const specialLightV2 = special.includes('class="product-window"');
if (specialLightV2) {
  special = replaceOnce(special, /<\/head>/g, `${commonStyles}</head>`, 'special v2 head');
} else {
  special = replaceOnce(special, /<body class="/g, '<body class="special-premium ', 'special body namespace');
  special = replaceOnce(special, /<\/head>/g, `${commonStyles}${specialPremiumStyles}</head>`, 'special head');
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
const specialIsV2 = specialHtml.includes('class="product-window"');
if (!specialIsV2 && ((specialHtml.match(/data-special-premium-typography/g) || []).length !== 1 || !specialHtml.includes('class="special-premium '))) {
  throw new Error('CONTEXTUAL IMAGE GATE: special premium typography namespace missing or duplicated');
}
if (specialIsV2 && specialHtml.includes('data-special-premium-typography')) {
  throw new Error('CONTEXTUAL IMAGE GATE: light premium v2 must not receive legacy special typography overrides');
}

console.log(`CONTEXTUAL IMAGE GATE PASS — home and rehber enriched; special systems contract=${specialIsV2 ? 'light-v2' : 'legacy'}; /sablonlar byte-identical`);
