#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const distFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(distFile)) throw new Error('SPECIAL HERO PROOF V8: dist route missing');

let html = fs.readFileSync(distFile, 'utf8');

if (!html.includes('data-finance-funnel-v7')) {
  throw new Error('SPECIAL HERO PROOF V8: finance funnel V7 contract missing');
}

if (html.includes('data-hero-proof-layout-v8')) {
  console.log('SPECIAL HERO PROOF V8: already applied');
  process.exit(0);
}

const proofMarker = '<div class="proof-grid" aria-label="ExcelArşiv yaklaşımı">';
const workbookMarker = '<div class="workbook"';

const proofStart = html.indexOf(proofMarker);
if (proofStart < 0) throw new Error('SPECIAL HERO PROOF V8: proof strip not found');

let workbookStart = html.indexOf(workbookMarker, proofStart);
if (workbookStart < 0) throw new Error('SPECIAL HERO PROOF V8: workbook not found after proof strip');

const heroCopyClose = html.lastIndexOf('</div>', workbookStart);
if (heroCopyClose <= proofStart) throw new Error('SPECIAL HERO PROOF V8: hero-copy boundary not found');

const proofBlock = html.slice(proofStart, heroCopyClose).trim();
const proofChipCount = (proofBlock.match(/class="proof-chip"/gu) || []).length;
if (proofChipCount !== 4) throw new Error(`SPECIAL HERO PROOF V8: expected 4 proof chips, got ${proofChipCount}`);

// Remove the proof strip from the left hero copy while preserving the hero-copy closing tag.
html = `${html.slice(0, proofStart)}${html.slice(heroCopyClose)}`;

workbookStart = html.indexOf(workbookMarker, proofStart);
if (workbookStart < 0) throw new Error('SPECIAL HERO PROOF V8: workbook disappeared during move');

const heroSectionEnd = html.indexOf('</section>', workbookStart);
if (heroSectionEnd < 0) throw new Error('SPECIAL HERO PROOF V8: hero section end not found');

// The final closing div immediately before </section> is the .hero-grid wrapper.
const heroGridClose = html.lastIndexOf('</div>', heroSectionEnd);
if (heroGridClose <= workbookStart) throw new Error('SPECIAL HERO PROOF V8: hero-grid closing boundary not found');

html = `${html.slice(0, heroGridClose)}\n        ${proofBlock}\n      ${html.slice(heroGridClose)}`;

const css = `
  <style id="special-hero-proof-layout-v8">
    body[data-hero-proof-layout-v8] .hero{padding:64px 0 50px}
    body[data-hero-proof-layout-v8] .hero-grid{grid-template-columns:minmax(0,.92fr) minmax(0,1.08fr);column-gap:48px;row-gap:24px;align-items:center}
    body[data-hero-proof-layout-v8] .hero-copy,body[data-hero-proof-layout-v8] .workbook{align-self:center}
    body[data-hero-proof-layout-v8] .hero-grid>.proof-grid{grid-column:1/-1;width:100%;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:2px 0 0;align-items:stretch}
    body[data-hero-proof-layout-v8] .hero-grid>.proof-grid .proof-chip{min-width:0;min-height:126px;height:100%;display:flex;flex-direction:column;justify-content:flex-start;padding:18px 20px 17px;border-radius:16px;background:#fff;box-shadow:0 8px 22px rgba(15,23,42,.055)}
    body[data-hero-proof-layout-v8] .hero-grid>.proof-grid .proof-chip strong{font-size:16px;line-height:1.28;font-weight:730;letter-spacing:-.025em;color:#0f172a}
    body[data-hero-proof-layout-v8] .hero-grid>.proof-grid .proof-chip span{margin-top:8px;font-size:15px;line-height:1.46;color:#64748b}
    body[data-hero-proof-layout-v8] .hero-grid>.proof-grid .proof-chip:nth-child(1){border-top-width:4px}
    body[data-hero-proof-layout-v8] .hero-grid>.proof-grid .proof-chip:nth-child(2){border-top-width:4px}
    body[data-hero-proof-layout-v8] .hero-grid>.proof-grid .proof-chip:nth-child(3){border-top-width:4px}
    body[data-hero-proof-layout-v8] .hero-grid>.proof-grid .proof-chip:nth-child(4){border-top-width:4px}

    body[data-hero-proof-layout-v8] .section{padding:68px 0}
    body[data-hero-proof-layout-v8] .section-head{margin-bottom:30px}
    body[data-hero-proof-layout-v8] .problem-grid,
    body[data-hero-proof-layout-v8] .area-grid,
    body[data-hero-proof-layout-v8] .process-grid,
    body[data-hero-proof-layout-v8] .why-grid{gap:16px}
    body[data-hero-proof-layout-v8] .node-flow{gap:24px}
    body[data-hero-proof-layout-v8] .node-card:not(:last-child)::after{right:-20px}
    body[data-hero-proof-layout-v8] .authority-summary{gap:10px;margin-bottom:16px}
    body[data-hero-proof-layout-v8] .delivery-grid{gap:18px}
    body[data-hero-proof-layout-v8] .cta{padding-top:64px;padding-bottom:64px}

    @media(max-width:1100px){
      body[data-hero-proof-layout-v8] .hero{padding:56px 0 44px}
      body[data-hero-proof-layout-v8] .hero-grid{grid-template-columns:1fr;gap:28px}
      body[data-hero-proof-layout-v8] .hero-grid>.proof-grid{grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin-top:0}
      body[data-hero-proof-layout-v8] .hero-grid>.proof-grid .proof-chip{min-height:112px}
    }

    @media(max-width:720px){
      body[data-hero-proof-layout-v8] .hero{padding:46px 0 38px}
      body[data-hero-proof-layout-v8] .hero-grid{gap:22px}
      body[data-hero-proof-layout-v8] .hero-grid>.proof-grid{grid-template-columns:1fr;gap:10px}
      body[data-hero-proof-layout-v8] .hero-grid>.proof-grid .proof-chip{min-height:auto;padding:16px 17px}
      body[data-hero-proof-layout-v8] .hero-grid>.proof-grid .proof-chip strong{font-size:16px}
      body[data-hero-proof-layout-v8] .hero-grid>.proof-grid .proof-chip span{font-size:14px;margin-top:5px}
      body[data-hero-proof-layout-v8] .section{padding:54px 0}
      body[data-hero-proof-layout-v8] .section-head{margin-bottom:24px}
      body[data-hero-proof-layout-v8] .problem-grid,
      body[data-hero-proof-layout-v8] .area-grid,
      body[data-hero-proof-layout-v8] .process-grid,
      body[data-hero-proof-layout-v8] .why-grid{gap:12px}
    }
  </style>`;

if (!html.includes('</head>')) throw new Error('SPECIAL HERO PROOF V8: </head> missing');
html = html.replace('</head>', `${css}\n</head>`);
html = html.replace('data-finance-funnel-v7', 'data-finance-funnel-v7 data-hero-proof-layout-v8');

const finalProofStart = html.indexOf(proofMarker);
const finalWorkbookStart = html.indexOf(workbookMarker);
const finalHeroSectionEnd = html.indexOf('</section>', finalWorkbookStart);
if (!(finalWorkbookStart >= 0 && finalProofStart > finalWorkbookStart && finalProofStart < finalHeroSectionEnd)) {
  throw new Error('SPECIAL HERO PROOF V8: proof strip is not a full-width sibling after workbook');
}
if (!html.includes('grid-column:1/-1')) throw new Error('SPECIAL HERO PROOF V8: full-row CSS contract missing');
if (!html.includes('grid-template-columns:repeat(4,minmax(0,1fr))')) throw new Error('SPECIAL HERO PROOF V8: desktop 4-column contract missing');

fs.writeFileSync(distFile, html, 'utf8');
console.log('SPECIAL HERO PROOF V8 PASS — proof strip moved below full hero, desktop 4-up / tablet 2-up / mobile 1-up, vertical rhythm tightened.');
