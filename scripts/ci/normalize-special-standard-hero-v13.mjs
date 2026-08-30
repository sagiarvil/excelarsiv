#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const distFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(distFile)) throw new Error('SPECIAL STANDARD HERO V13: dist route missing');

let html = fs.readFileSync(distFile, 'utf8');
const styleId = 'special-standard-hero-v13-css';
const marker = 'data-standard-hero-v13';

for (const token of ['class="hero"', 'class="wrap-wide hero-grid"', 'class="hero-copy"', 'class="workbook"']) {
  if (!html.includes(token)) throw new Error(`SPECIAL STANDARD HERO V13: required hero token missing: ${token}`);
}

const bodyOpen = html.match(/<body\b[^>]*>/u)?.[0];
if (!bodyOpen) throw new Error('SPECIAL STANDARD HERO V13: body tag missing');
if (!bodyOpen.includes(marker)) {
  html = html.replace(bodyOpen, bodyOpen.replace(/>$/u, ` ${marker}>`));
}

const existingStyle = new RegExp(`<style\\b(?=[^>]*\\bid=["']${styleId}["'])[^>]*>[\\s\\S]*?<\\/style>`, 'gu');
html = html.replace(existingStyle, '');

const css = `<style id="${styleId}">
  body[${marker}] .hero{
    position:relative;
    overflow:visible;
    margin:28px 0 46px;
    padding:0;
    background:transparent;
  }
  body[${marker}] .hero::after{display:none!important}
  body[${marker}] .hero-grid{
    position:relative;
    z-index:1;
    width:min(1680px,calc(100% - 64px));
    min-height:533px;
    margin-inline:auto;
    padding:44px 48px;
    display:grid;
    grid-template-columns:minmax(0,.92fr) minmax(0,1.08fr);
    gap:48px;
    align-items:center;
    border:1px solid rgba(16,124,65,.12);
    border-radius:24px;
    background:
      linear-gradient(90deg,rgba(255,255,255,.98) 0%,rgba(255,255,255,.98) 52%,rgba(244,251,247,.98) 100%),
      linear-gradient(rgba(16,124,65,.035) 1px,transparent 1px),
      linear-gradient(90deg,rgba(16,124,65,.035) 1px,transparent 1px);
    background-size:auto,34px 34px,34px 34px;
    box-shadow:0 24px 60px rgba(24,55,38,.13);
    overflow:hidden;
  }
  body[${marker}] .hero-copy,
  body[${marker}] .workbook{min-width:0}
  body[${marker}] .hero h1{
    max-width:680px;
    font-size:clamp(38px,3.15vw,46px);
    line-height:1.02;
    letter-spacing:-.045em;
    text-wrap:balance;
  }
  body[${marker}] .hero-lead{
    max-width:680px;
    margin-top:20px;
    font-size:18px;
    line-height:1.55;
  }
  body[${marker}] .hero-actions{margin-top:24px}
  body[${marker}] .hero-bullets{margin-top:18px}
  body[${marker}] .proof-grid{
    grid-template-columns:repeat(2,minmax(0,1fr));
    gap:8px;
    margin-top:20px;
  }
  body[${marker}] .proof-chip{padding:10px 11px}
  body[${marker}] .proof-chip strong{font-size:12.5px}
  body[${marker}] .proof-chip span{font-size:11.5px}
  body[${marker}] .workbook{
    border-radius:20px;
    box-shadow:0 18px 46px rgba(24,55,38,.11);
  }
  body[${marker}] .wb-top{padding:12px 14px}
  body[${marker}] .wb-body{padding:14px}
  body[${marker}] .metric-grid{gap:7px}
  body[${marker}] .metric{padding:10px}
  body[${marker}] .metric strong{font-size:18px}
  body[${marker}] .wb-chart-grid,
  body[${marker}] .wb-lists{gap:7px;margin-top:7px}
  body[${marker}] .panel{padding:12px}

  @media (min-width:721px) and (max-width:1180px){
    body[${marker}] .hero-grid{width:calc(100% - 40px);padding:36px 34px;gap:34px}
    body[${marker}] .hero h1{font-size:clamp(36px,4vw,44px)}
  }

  @media (max-width:980px){
    body[${marker}] .hero-grid{
      grid-template-columns:minmax(0,1fr);
      min-height:0;
    }
    body[${marker}] .hero-copy{max-width:760px}
    body[${marker}] .workbook{width:100%}
  }

  @media (max-width:720px){
    body[${marker}] .hero{
      margin:0;
      padding:14px 0 22px;
      background:linear-gradient(180deg,#f7faf8 0%,#fff 100%);
      border-bottom:1px solid #e6ece8;
    }
    body[${marker}] .hero-grid{
      width:calc(100% - 28px);
      min-height:0;
      padding:10px 0 0;
      gap:18px;
      border:0;
      border-radius:0;
      background:transparent;
      box-shadow:none;
      overflow:visible;
    }
    body[${marker}] .eyebrow{margin-bottom:12px;font-size:12px}
    body[${marker}] .hero h1{
      max-width:560px;
      font-size:clamp(31px,9.6vw,42px);
      line-height:.98;
      letter-spacing:-.052em;
    }
    body[${marker}] .hero-lead{
      max-width:560px;
      margin-top:13px;
      font-size:14px;
      line-height:1.55;
    }
    body[${marker}] .hero-actions{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:9px;margin-top:18px}
    body[${marker}] .hero-actions .btn{width:100%;min-width:0;min-height:48px;padding:0 13px;font-size:13px;text-align:center}
    body[${marker}] .hero-bullets{gap:8px 14px;margin-top:14px}
    body[${marker}] .hero-bullets li{font-size:12px}
    body[${marker}] .proof-grid{grid-template-columns:minmax(0,1fr);gap:7px;margin-top:14px}
    body[${marker}] .proof-chip{display:grid;grid-template-columns:minmax(0,.9fr) minmax(0,1.1fr);gap:8px;align-items:center;padding:10px 11px}
    body[${marker}] .proof-chip span{margin-top:0}
    body[${marker}] .workbook{border-radius:18px;box-shadow:0 14px 34px rgba(24,55,38,.10)}
  }

  @media (max-width:430px){
    body[${marker}] .hero-actions{grid-template-columns:minmax(0,1fr)}
    body[${marker}] .proof-chip{grid-template-columns:minmax(0,1fr)}
    body[${marker}] .proof-chip span{margin-top:3px}
  }
</style>`;

if (!html.includes('</head>')) throw new Error('SPECIAL STANDARD HERO V13: head close missing');
html = html.replace('</head>', `${css}\n</head>`);

for (const token of [
  marker,
  'width:min(1680px,calc(100% - 64px))',
  'min-height:533px',
  'border-radius:24px',
  'margin:28px 0 46px',
  'font-size:clamp(31px,9.6vw,42px)',
]) {
  if (!html.includes(token)) throw new Error(`SPECIAL STANDARD HERO V13: standard geometry token missing: ${token}`);
}

if (html.includes('overflow-x:hidden')) throw new Error('SPECIAL STANDARD HERO V13: forbidden global overflow hiding introduced');

fs.writeFileSync(distFile, html, 'utf8');
console.log('SPECIAL STANDARD HERO V13 PASS — special page hero now follows the standard site hero width, rhythm, radius, shadow and responsive geometry while preserving content.');
