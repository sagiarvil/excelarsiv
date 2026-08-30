#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const distFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(distFile)) throw new Error('SPECIAL RUNTIME MARKERS V10: dist route missing');

let html = fs.readFileSync(distFile, 'utf8');

const bodyMatch = html.match(/<body\b[^>]*>/u);
if (!bodyMatch) throw new Error('SPECIAL RUNTIME MARKERS V10: body tag missing');

const originalBodyTag = bodyMatch[0];
if (!/\bdata-special-source-v5\b/u.test(originalBodyTag)) {
  throw new Error('SPECIAL RUNTIME MARKERS V10: source ownership marker missing on body');
}
if (!/\bdata-finance-funnel-v7\b/u.test(originalBodyTag)) {
  throw new Error('SPECIAL RUNTIME MARKERS V10: finance funnel V7 marker missing on body');
}

const requiredRuntimeMarkers = ['data-hero-proof-layout-v8', 'data-cfo-positioning-v9'];
let nextBodyTag = originalBodyTag;
for (const marker of requiredRuntimeMarkers) {
  if (!new RegExp(`\\b${marker}\\b`, 'u').test(nextBodyTag)) {
    nextBodyTag = nextBodyTag.replace(/>$/u, ` ${marker}>`);
  }
}
html = html.replace(originalBodyTag, nextBodyTag);

const css = `
  <style id="special-runtime-markers-v10-css">
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .hero{padding:56px 0 42px}
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .hero-grid{column-gap:46px;row-gap:22px}
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .hero-grid>.proof-grid{grid-column:1/-1;width:100%;gap:16px;margin:4px 0 0}
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .hero-grid>.proof-grid .proof-chip{min-height:118px;padding:18px 19px 17px}
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .section{padding:56px 0}
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .section-head{margin-bottom:26px}
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .authority-comparison{padding:52px 0 56px}
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .authority-rail{margin:0 0 20px}
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .problem-grid,
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .area-grid,
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .process-grid,
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .why-grid{gap:14px}
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .node-flow{gap:20px}
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .delivery-grid{gap:16px}
    body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .cta{padding-top:56px;padding-bottom:56px}

    @media(max-width:1100px){
      body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .hero{padding:50px 0 38px}
      body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .hero-grid{grid-template-columns:1fr;gap:24px}
      body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .hero-grid>.proof-grid{grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
    }

    @media(max-width:720px){
      body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .hero{padding:42px 0 34px}
      body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .hero-grid{gap:20px}
      body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .hero-grid>.proof-grid{grid-template-columns:1fr;gap:10px}
      body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .hero-grid>.proof-grid .proof-chip{min-height:auto;padding:16px 17px}
      body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .section,
      body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .authority-comparison{padding:46px 0}
      body[data-hero-proof-layout-v8][data-cfo-positioning-v9] .section-head{margin-bottom:22px}
    }
  </style>`;

if (!html.includes('</head>')) throw new Error('SPECIAL RUNTIME MARKERS V10: head close missing');
html = html.replace('</head>', `${css}\n</head>`);

const finalBodyTag = html.match(/<body\b[^>]*>/u)?.[0] || '';
for (const marker of ['data-special-source-v5', 'data-finance-funnel-v7', ...requiredRuntimeMarkers]) {
  if (!new RegExp(`\\b${marker}\\b`, 'u').test(finalBodyTag)) {
    throw new Error(`SPECIAL RUNTIME MARKERS V10: runtime body marker missing: ${marker}`);
  }
}

const heroStart = html.indexOf('<section class="hero">');
const heroEnd = html.indexOf('</section>', heroStart);
const workbookStart = html.indexOf('<div class="workbook"', heroStart);
const proofStart = html.indexOf('<div class="proof-grid" aria-label="ExcelArşiv yaklaşımı">', heroStart);
if (!(heroStart >= 0 && workbookStart > heroStart && proofStart > workbookStart && proofStart < heroEnd)) {
  throw new Error('SPECIAL RUNTIME MARKERS V10: proof strip is not a full-width sibling after workbook');
}

const proofChipCount = (html.slice(proofStart, heroEnd).match(/class="proof-chip"/gu) || []).length;
if (proofChipCount !== 4) throw new Error(`SPECIAL RUNTIME MARKERS V10: proof chip count ${proofChipCount}/4`);

for (const selectorToken of [
  'body[data-hero-proof-layout-v8] .hero-grid>.proof-grid',
  'body[data-cfo-positioning-v9] .authority-rail',
  'body[data-cfo-positioning-v9] .area-card h3',
]) {
  if (!html.includes(selectorToken)) throw new Error(`SPECIAL RUNTIME MARKERS V10: expected active CSS missing: ${selectorToken}`);
}

fs.writeFileSync(distFile, html, 'utf8');
console.log('SPECIAL RUNTIME MARKERS V10 PASS — V8/V9 CSS now binds to body; proof strip spans the hero, authority rail renders correctly, and vertical rhythm is tightened.');
