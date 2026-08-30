#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const distFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(distFile)) throw new Error('SPECIAL HOME HERO PARITY V15: dist route missing');

let html = fs.readFileSync(distFile, 'utf8');
const marker = 'data-home-hero-parity-v15';
const styleId = 'special-home-hero-parity-v15-css';

for (const token of ['<section class="hero">','class="wrap-wide hero-grid"','class="hero-copy"','class="workbook"','class="proof-grid" aria-label="ExcelArşiv yaklaşımı"']) {
  if (!html.includes(token)) throw new Error(`SPECIAL HOME HERO PARITY V15: required token missing: ${token}`);
}

const bodyOpen = html.match(/<body\b[^>]*>/u)?.[0];
if (!bodyOpen) throw new Error('SPECIAL HOME HERO PARITY V15: body tag missing');
if (!bodyOpen.includes(marker)) html = html.replace(bodyOpen, bodyOpen.replace(/>$/u, ` ${marker}>`));

// Move the four proof cards OUTSIDE the hero visual shell. This is the root fix:
// the cards were a third grid row, so the special hero could never keep the homepage 3:1 geometry.
const heroStart = html.indexOf('<section class="hero">');
const heroEnd = html.indexOf('</section>', heroStart);
if (heroStart < 0 || heroEnd < 0) throw new Error('SPECIAL HOME HERO PARITY V15: hero boundary missing');

const proofStart = html.indexOf('<div class="proof-grid" aria-label="ExcelArşiv yaklaşımı">', heroStart);
if (proofStart < 0 || proofStart > heroEnd) throw new Error('SPECIAL HOME HERO PARITY V15: proof grid is not inside hero');

function findMatchingDivEnd(source, start) {
  const re = /<div\b[^>]*>|<\/div>/gu;
  re.lastIndex = start;
  let depth = 0;
  let match;
  while ((match = re.exec(source))) {
    if (match[0].startsWith('</div')) depth -= 1;
    else depth += 1;
    if (depth === 0) return re.lastIndex;
  }
  return -1;
}

const proofEnd = findMatchingDivEnd(html, proofStart);
if (proofEnd < 0 || proofEnd > heroEnd) throw new Error('SPECIAL HOME HERO PARITY V15: proof grid closing div missing');
const proofBlock = html.slice(proofStart, proofEnd);
const proofChipCount = (proofBlock.match(/class="proof-chip"/gu) || []).length;
if (proofChipCount !== 4) throw new Error(`SPECIAL HOME HERO PARITY V15: expected 4 proof chips, got ${proofChipCount}`);

html = `${html.slice(0, proofStart)}${html.slice(proofEnd)}`;
const heroEndAfterRemoval = html.indexOf('</section>', heroStart);
if (heroEndAfterRemoval < 0) throw new Error('SPECIAL HOME HERO PARITY V15: hero boundary lost after proof move');
const heroCloseEnd = heroEndAfterRemoval + '</section>'.length;
const proofBand = `\n<div class="hero-proof-band" aria-label="ExcelArşiv yaklaşım özeti"><div class="wrap-wide">${proofBlock}</div></div>`;
html = `${html.slice(0, heroCloseEnd)}${proofBand}${html.slice(heroCloseEnd)}`;

const existingStyle = new RegExp(`<style\\b(?=[^>]*\\bid=["']${styleId}["'])[^>]*>[\\s\\S]*?<\\/style>`, 'gu');
html = html.replace(existingStyle, '');

const css = `<style id="${styleId}">
  /* Desktop geometry is intentionally the same visual frame as the homepage hero. */
  body[${marker}] .hero{
    position:relative!important;
    margin:28px 0 0!important;
    padding:0!important;
    background:transparent!important;
    overflow:visible!important;
  }
  body[${marker}] .hero::after{display:none!important}
  body[${marker}] .hero>.hero-grid{
    width:min(1680px,calc(100% - 64px))!important;
    aspect-ratio:3 / 1!important;
    min-height:0!important;
    max-height:none!important;
    margin-inline:auto!important;
    padding:30px 42px!important;
    display:grid!important;
    grid-template-columns:minmax(0,.92fr) minmax(0,1.08fr)!important;
    grid-template-rows:minmax(0,1fr)!important;
    column-gap:42px!important;
    row-gap:0!important;
    align-items:center!important;
    border:1px solid rgba(16,124,65,.12)!important;
    border-radius:24px!important;
    background:linear-gradient(90deg,#fff 0%,#fff 53%,#f4fbf7 100%)!important;
    box-shadow:0 24px 60px rgba(24,55,38,.13)!important;
    overflow:hidden!important;
  }
  body[${marker}] .hero-copy,
  body[${marker}] .workbook{min-width:0!important;max-height:100%;align-self:center!important}
  body[${marker}] .hero-copy{display:flex;flex-direction:column;justify-content:center}
  body[${marker}] .hero .eyebrow{
    width:max-content;max-width:100%;margin:0 0 14px!important;padding:7px 11px!important;
    font-size:12px!important;line-height:1.2!important;font-weight:760!important;letter-spacing:-.005em!important
  }
  body[${marker}] .hero h1{
    max-width:660px!important;margin:0!important;
    font-family:-apple-system,BlinkMacSystemFont,"SF Pro Display","SF Pro Text","Helvetica Neue",Arial,sans-serif!important;
    font-size:clamp(34px,2.75vw,46px)!important;
    line-height:1.02!important;font-weight:780!important;letter-spacing:-.045em!important;text-wrap:balance!important
  }
  body[${marker}] .hero-lead{
    max-width:650px!important;margin:15px 0 0!important;
    font-size:16px!important;line-height:1.48!important;letter-spacing:-.012em!important;color:#475569!important
  }
  body[${marker}] .hero-actions{display:flex!important;flex-wrap:wrap!important;gap:9px!important;margin-top:17px!important}
  body[${marker}] .hero-actions .btn{min-height:44px!important;padding:0 15px!important;border-radius:11px!important;font-size:13px!important;line-height:1.2!important}
  body[${marker}] .hero-microcopy{margin-top:10px!important;font-size:12px!important;line-height:1.42!important}
  body[${marker}] .hero-bullets{gap:7px 14px!important;margin-top:12px!important}
  body[${marker}] .hero-bullets li{font-size:11.5px!important;line-height:1.3!important}

  body[${marker}] .workbook{
    width:100%!important;max-width:100%!important;transform:none!important;
    border-radius:18px!important;box-shadow:0 16px 40px rgba(24,55,38,.10)!important
  }
  body[${marker}] .wb-top{padding:9px 11px!important}
  body[${marker}] .wb-title{font-size:12px!important}
  body[${marker}] .wb-title img{width:20px!important;height:20px!important}
  body[${marker}] .wb-note{font-size:10px!important}
  body[${marker}] .wb-body{padding:10px!important}
  body[${marker}] .metric-grid{gap:5px!important}
  body[${marker}] .metric{padding:7px!important;border-radius:9px!important}
  body[${marker}] .metric small{font-size:9.5px!important}
  body[${marker}] .metric strong{margin-top:3px!important;font-size:15px!important}
  body[${marker}] .metric span{margin-top:2px!important;font-size:9px!important}
  body[${marker}] .wb-chart-grid,
  body[${marker}] .wb-lists{gap:5px!important;margin-top:5px!important}
  body[${marker}] .panel{padding:8px!important;border-radius:9px!important}
  body[${marker}] .panel h3{margin-bottom:6px!important;font-size:10px!important}
  body[${marker}] .chart-legend{margin-bottom:5px!important;font-size:8.5px!important}
  body[${marker}] .data-row,
  body[${marker}] .alert-row{font-size:9px!important;line-height:1.25!important}

  /* The four proof cards are a separate band, not part of the 3:1 hero canvas. */
  body[${marker}] .hero-proof-band{margin:14px 0 48px!important}
  body[${marker}] .hero-proof-band>.wrap-wide{width:min(1320px,calc(100% - 48px))!important;margin-inline:auto!important}
  body[${marker}] .hero-proof-band .proof-grid{
    width:100%!important;display:grid!important;grid-template-columns:repeat(4,minmax(0,1fr))!important;
    gap:12px!important;margin:0!important
  }
  body[${marker}] .hero-proof-band .proof-chip{
    min-width:0!important;min-height:94px!important;height:100%!important;
    padding:13px 15px!important;border-radius:14px!important;background:#fff!important
  }
  body[${marker}] .hero-proof-band .proof-chip strong{font-size:14px!important;line-height:1.25!important}
  body[${marker}] .hero-proof-band .proof-chip span{margin-top:5px!important;font-size:12.5px!important;line-height:1.4!important}

  @media (min-width:721px) and (max-width:1180px){
    body[${marker}] .hero>.hero-grid{
      width:calc(100% - 40px)!important;aspect-ratio:auto!important;min-height:480px!important;
      padding:28px 30px!important;gap:28px!important
    }
    body[${marker}] .hero h1{font-size:clamp(32px,4vw,42px)!important}
    body[${marker}] .hero-proof-band .proof-grid{grid-template-columns:repeat(2,minmax(0,1fr))!important}
  }

  @media (max-width:900px){
    body[${marker}] .hero>.hero-grid{grid-template-columns:minmax(0,1fr)!important;aspect-ratio:auto!important;min-height:0!important}
    body[${marker}] .workbook{margin-top:8px!important}
  }

  @media (max-width:720px){
    body[${marker}] .hero{margin:0!important;padding:14px 0 0!important;background:linear-gradient(180deg,#f7faf8 0%,#fff 100%)!important}
    body[${marker}] .hero>.hero-grid{
      width:calc(100% - 28px)!important;min-height:0!important;padding:10px 0 0!important;gap:18px!important;
      border:0!important;border-radius:0!important;background:transparent!important;box-shadow:none!important;overflow:visible!important
    }
    body[${marker}] .hero .eyebrow{margin-bottom:10px!important;font-size:10.5px!important}
    body[${marker}] .hero h1{max-width:560px!important;font-size:clamp(31px,9.6vw,42px)!important;line-height:.98!important;letter-spacing:-.052em!important}
    body[${marker}] .hero-lead{max-width:560px!important;margin-top:13px!important;font-size:14px!important;line-height:1.55!important}
    body[${marker}] .hero-actions{display:grid!important;grid-template-columns:minmax(0,1fr)!important;gap:8px!important;margin-top:16px!important}
    body[${marker}] .hero-actions .btn{width:100%!important;min-height:48px!important;font-size:13px!important}
    body[${marker}] .hero-bullets{gap:7px 12px!important;margin-top:12px!important}
    body[${marker}] .hero-bullets li{font-size:12px!important}
    body[${marker}] .workbook{display:none!important}
    body[${marker}] .hero-proof-band{margin:14px 0 38px!important}
    body[${marker}] .hero-proof-band>.wrap-wide{width:calc(100% - 28px)!important}
    body[${marker}] .hero-proof-band .proof-grid{grid-template-columns:minmax(0,1fr)!important;gap:8px!important}
    body[${marker}] .hero-proof-band .proof-chip{min-height:0!important;padding:12px 13px!important}
  }
</style>`;

if (!html.includes('</head>')) throw new Error('SPECIAL HOME HERO PARITY V15: head close missing');
html = html.replace('</head>', `${css}\n</head>`);

for (const token of [
  marker,
  'aspect-ratio:3 / 1!important',
  'width:min(1680px,calc(100% - 64px))!important',
  'font-size:clamp(34px,2.75vw,46px)!important',
  'class="hero-proof-band"',
  'grid-template-columns:repeat(4,minmax(0,1fr))!important',
  'body[data-home-hero-parity-v15] .workbook{display:none!important}',
]) {
  if (!html.includes(token)) throw new Error(`SPECIAL HOME HERO PARITY V15: parity token missing: ${token}`);
}

const finalHeroStart = html.indexOf('<section class="hero">');
const finalHeroEnd = html.indexOf('</section>', finalHeroStart);
const finalProofStart = html.indexOf('<div class="hero-proof-band"', finalHeroEnd);
if (!(finalHeroStart >= 0 && finalHeroEnd > finalHeroStart && finalProofStart > finalHeroEnd)) {
  throw new Error('SPECIAL HOME HERO PARITY V15: proof cards still inflate the hero visual shell');
}
if (html.includes('overflow-x:hidden')) throw new Error('SPECIAL HOME HERO PARITY V15: forbidden global overflow hiding introduced');

fs.writeFileSync(distFile, html, 'utf8');
console.log('SPECIAL HOME HERO PARITY V15 PASS — special hero now uses homepage 3:1 desktop geometry and typography scale; proof cards are outside the hero canvas; content preserved.');
