#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const homeFile = path.resolve('dist/index.html');
const cssFile = path.resolve('src/styles/home-authority-label-contrast-v34.css');
const styleId = 'home-authority-label-contrast-v34';

const sitewidePublicCss = path.resolve('public/styles/sitewide-continuous-grid-v4.css');
const sitewideDistCss = path.resolve('dist/styles/sitewide-continuous-grid-v4.css');
const premiumGridPublicCss = path.resolve('public/styles/sitewide-premium-excel-grid-v5.css');
const premiumGridDistCss = path.resolve('dist/styles/sitewide-premium-excel-grid-v5.css');
const specialAssetPublic = path.resolve('public/images/ozel-excel/sahadan-uzmanlik-process-panel.webp');
const specialAssetDist = path.resolve('dist/images/ozel-excel/sahadan-uzmanlik-process-panel.webp');
const sitewideStyleId = 'sitewide-continuous-grid-v4';
const legacySitewideHref = '/styles/sitewide-continuous-grid-v4.css';
const legacyPremiumHref = '/styles/sitewide-premium-excel-grid-v5.css';

for (const file of [
  homeFile,
  cssFile,
  sitewidePublicCss,
  sitewideDistCss,
  premiumGridPublicCss,
  premiumGridDistCss,
  specialAssetPublic,
  specialAssetDist,
]) {
  if (!fs.existsSync(file)) throw new Error(`HOME/SITEWIDE VISUAL GATE: missing ${file}`);
}

function assertBalancedCss(file, label) {
  const value = fs.readFileSync(file, 'utf8');
  const withoutComments = value.replace(/\/\*[\s\S]*?\*\//g, '');
  const openBraces = (withoutComments.match(/\{/g) || []).length;
  const closeBraces = (withoutComments.match(/\}/g) || []).length;
  if (openBraces !== closeBraces) throw new Error(`${label}: CSS brace mismatch ${openBraces}/${closeBraces}`);
  return value;
}

const css = assertBalancedCss(cssFile, 'HOME AUTHORITY CONTRAST GATE');
for (const token of [
  '.authority-metrics article:nth-child(1) small',
  '.authority-metrics article:nth-child(2) small',
  '.authority-metrics article:nth-child(3) small',
  'color:#075c31!important',
  'color:#114fac!important',
  'color:#8a5200!important',
  'font-weight:950!important',
]) {
  if (!css.includes(token)) throw new Error(`HOME AUTHORITY CONTRAST GATE: required contract missing ${token}`);
}

let html = fs.readFileSync(homeFile, 'utf8');
html = html.replace(new RegExp(`<style\\b(?=[^>]*\\bid=["']${styleId}["'])[^>]*>[\\s\\S]*?<\\/style>`, 'gi'), '');
if (!html.includes('</head>')) throw new Error('HOME AUTHORITY CONTRAST GATE: homepage missing </head>');
html = html.replace('</head>', `<style id="${styleId}">\n${css}\n</style>\n</head>`);

if (!html.includes(`id="${styleId}"`)) throw new Error('HOME AUTHORITY CONTRAST GATE: inline style missing');
if (!html.includes('body.ea-home-color-v3 .home-v2 .authority-metrics article:nth-child(1) small')) {
  throw new Error('HOME AUTHORITY CONTRAST GATE: high-specificity selector missing in final homepage');
}
fs.writeFileSync(homeFile, html);

/* Sitewide continuous grid is the final visual layer.
   Inline base + v5 premium overlay into every emitted HTML route: one visual contract,
   no extra render-blocking stylesheet request, no section-by-section restart. */
const sitewideCss = assertBalancedCss(sitewidePublicCss, 'SITEWIDE GRID GATE');
const premiumGridCss = assertBalancedCss(premiumGridPublicCss, 'PREMIUM EXCEL GRID GATE');

if (fs.readFileSync(sitewidePublicCss, 'utf8') !== fs.readFileSync(sitewideDistCss, 'utf8')) {
  throw new Error('SITEWIDE GRID GATE: public/dist stylesheet parity failed');
}
if (fs.readFileSync(premiumGridPublicCss, 'utf8') !== fs.readFileSync(premiumGridDistCss, 'utf8')) {
  throw new Error('PREMIUM EXCEL GRID GATE: public/dist stylesheet parity failed');
}

for (const token of [
  '--ea-grid-paper:#fbfaf6',
  '--ea-grid-x:64px',
  'background-repeat:repeat!important',
  'body.special-light-v1[data-special-innovation="v4"] .hero::after',
  'display:none!important',
  '.process-card::after',
  'sahadan-uzmanlik-process-panel.webp',
  '.native-info--special::before',
  '--sp-blue:#176fe5',
  '--sp-green:#0a914a',
  '--sp-amber:#ee9d00',
  '--sp-coral:#f05a47',
  '.native-info__outcomes article:nth-child(4)',
]) {
  if (!sitewideCss.includes(token)) throw new Error(`SITEWIDE GRID GATE: required contract missing ${token}`);
}

for (const token of [
  '--ea-v5-grid-minor:rgba(24,73,126,.034)',
  '--ea-v5-grid-x:56px',
  '--ea-v5-grid-y:38px',
  'background-attachment:fixed',
  'body .home-v2>section::after',
  'body #icerik>section::after',
  'opacity:.72',
]) {
  if (!premiumGridCss.includes(token)) throw new Error(`PREMIUM EXCEL GRID GATE: required contract missing ${token}`);
}

for (const forbidden of ['clip-path:', 'skew(', 'polygon(']) {
  if (sitewideCss.includes(forbidden) || premiumGridCss.includes(forbidden)) {
    throw new Error(`SITEWIDE GRID GATE: diagonal/broken background primitive forbidden: ${forbidden}`);
  }
}

function htmlFiles(dir) {
  const out = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...htmlFiles(full));
    else if (entry.isFile() && entry.name.endsWith('.html')) out.push(full);
  }
  return out;
}

const pages = htmlFiles(path.resolve('dist'));
if (pages.length < 10) throw new Error(`SITEWIDE GRID GATE: unexpectedly low HTML route count ${pages.length}`);

const inlineMarkup = `<style id="${sitewideStyleId}">\n${sitewideCss}\n\n${premiumGridCss}\n</style>`;
for (const file of pages) {
  let page = fs.readFileSync(file, 'utf8');
  page = page.replace(new RegExp(`<style\\b(?=[^>]*\\bid=["']${sitewideStyleId}["'])[^>]*>[\\s\\S]*?<\\/style>`, 'gi'), '');
  page = page.replace(new RegExp(`<link\\b(?=[^>]*\\bid=["']${sitewideStyleId}["'])[^>]*>`, 'gi'), '');
  page = page.replace(new RegExp(`<link\\b(?=[^>]*\\bhref=["']${legacySitewideHref.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}["'])[^>]*>`, 'gi'), '');
  page = page.replace(new RegExp(`<link\\b(?=[^>]*\\bhref=["']${legacyPremiumHref.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}["'])[^>]*>`, 'gi'), '');
  if (!page.includes('</head>')) throw new Error(`SITEWIDE GRID GATE: ${file} missing </head>`);
  page = page.replace('</head>', `${inlineMarkup}\n</head>`);
  if (!page.includes(`id="${sitewideStyleId}"`) || !page.includes('--ea-grid-paper:#fbfaf6') || !page.includes('--ea-v5-grid-minor:rgba(24,73,126,.034)')) {
    throw new Error(`SITEWIDE GRID GATE: inline stylesheet injection failed for ${file}`);
  }
  if (page.includes(`href="${legacySitewideHref}"`) || page.includes(`href="${legacyPremiumHref}"`)) {
    throw new Error(`SITEWIDE GRID GATE: render-blocking sitewide stylesheet survived in ${file}`);
  }
  fs.writeFileSync(file, page);
}

const specialFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(specialFile)) throw new Error('SITEWIDE GRID GATE: special route missing');
const specialHtml = fs.readFileSync(specialFile, 'utf8');
for (const required of [
  `id="${sitewideStyleId}"`,
  'data-special-innovation="v4"',
  'Sahadan Gelen Uzmanlık',
  'data-native-info="special-decision-map"',
  'sahadan-uzmanlik-process-panel.webp',
  '--ea-v5-grid-x:56px',
]) {
  if (!specialHtml.includes(required)) throw new Error(`SITEWIDE GRID GATE: special final contract missing ${required}`);
}

console.log(`HOME AUTHORITY + PREMIUM SITEWIDE GRID PASS — ${pages.length} HTML routes receive one inline continuous Excel grid; colored sections keep a subtle aligned grid overlay; zero extra CSS request.`);
