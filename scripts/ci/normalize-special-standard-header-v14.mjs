#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const distFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(distFile)) throw new Error('SPECIAL STANDARD HEADER V14: dist route missing');

let html = fs.readFileSync(distFile, 'utf8');
const marker = 'data-standard-header-v14';
const styleId = 'special-standard-header-v14-css';

for (const token of [
  'class="site-nav"',
  'class="wrap-wide nav-inner"',
  'class="brand"',
  'class="nav-links"',
  'class="nav-actions"',
  'class="mobile-menu"',
  'Ne zaman gerekir?',
  'Hangi alanlarda çalışıyoruz?',
  'Nasıl çalışıyoruz?',
  'Karşılaştırma',
  'Neden ExcelArşiv?',
  'SSS',
]) {
  if (!html.includes(token)) throw new Error(`SPECIAL STANDARD HEADER V14: required header token missing: ${token}`);
}

const bodyOpen = html.match(/<body\b[^>]*>/u)?.[0];
if (!bodyOpen) throw new Error('SPECIAL STANDARD HEADER V14: body tag missing');
if (!bodyOpen.includes(marker)) {
  html = html.replace(bodyOpen, bodyOpen.replace(/>$/u, ` ${marker}>`));
}

const existingStyle = new RegExp(`<style\\b(?=[^>]*\\bid=["']${styleId}["'])[^>]*>[\\s\\S]*?<\\/style>`, 'gu');
html = html.replace(existingStyle, '');

const css = `<style id="${styleId}">
  body[${marker}] .site-nav{
    position:sticky;
    top:0;
    z-index:100;
    background:rgba(255,255,255,.96);
    backdrop-filter:blur(12px);
    -webkit-backdrop-filter:blur(12px);
    border-bottom:1px solid #e5e5e3;
  }
  body[${marker}] .site-nav .nav-inner{
    width:min(1280px,calc(100% - 48px));
    min-height:68px;
    margin-inline:auto;
    display:flex;
    align-items:center;
    justify-content:space-between;
    gap:24px;
    min-width:0;
  }
  body[${marker}] .site-nav .brand{
    display:inline-flex;
    align-items:center;
    gap:9px;
    min-width:0;
    flex:0 0 auto;
  }
  body[${marker}] .site-nav .brand>img{
    width:40px;
    height:40px;
    object-fit:contain;
    flex:none;
  }
  body[${marker}] .site-nav .brand-copy{
    display:flex;
    flex-direction:column;
    align-items:flex-start;
    gap:0;
    min-width:0;
    line-height:1.15;
  }
  body[${marker}] .site-nav .brand-copy strong{
    font-size:18px;
    line-height:1.15;
    font-weight:800;
    letter-spacing:-.01em;
    color:#1b1b1f;
    white-space:nowrap;
  }
  body[${marker}] .site-nav .brand-copy small{
    font-size:9px;
    line-height:1.25;
    letter-spacing:.02em;
    color:#616166;
    white-space:nowrap;
  }
  body[${marker}] .site-nav .nav-links{
    display:flex;
    align-items:center;
    justify-content:center;
    gap:4px;
    margin-inline:auto;
    min-width:0;
    font-size:13px;
    line-height:1.2;
    color:#616166;
    letter-spacing:0;
  }
  body[${marker}] .site-nav .nav-links a{
    display:inline-flex;
    align-items:center;
    min-height:38px;
    padding:9px 9px;
    border-radius:9px;
    font-size:13px;
    line-height:1.2;
    font-weight:550;
    color:#616166;
    white-space:nowrap;
    transition:color 150ms ease-out,background 150ms ease-out;
  }
  body[${marker}] .site-nav .nav-links a:hover,
  body[${marker}] .site-nav .nav-links a:focus-visible{
    color:#0c5a30;
    background:#e9f5ee;
  }
  body[${marker}] .site-nav .nav-actions{
    display:flex;
    align-items:center;
    gap:8px;
    flex:0 0 auto;
  }
  body[${marker}] .site-nav .nav-cta{
    display:inline-flex;
    align-items:center;
    justify-content:center;
    min-height:42px!important;
    height:42px;
    padding-inline:16px!important;
    border-radius:11px;
    background:#107c41;
    color:#fff;
    font-size:13px!important;
    line-height:1;
    font-weight:700;
    box-shadow:none;
  }
  body[${marker}] .site-nav .nav-cta:hover{background:#0c5a30}
  body[${marker}] .site-nav .mobile-menu{display:none;position:relative}

  @media(max-width:1020px){
    body[${marker}] .site-nav .nav-links{display:none}
    body[${marker}] .site-nav .mobile-menu{display:block}
    body[${marker}] .site-nav .mobile-menu summary{
      display:grid;
      place-content:center;
      gap:4px;
      width:38px;
      height:38px;
      border:1px solid #d1d1cf;
      border-radius:10px;
      background:#fff;
      cursor:pointer;
      list-style:none;
    }
    body[${marker}] .site-nav .mobile-menu summary::-webkit-details-marker{display:none}
    body[${marker}] .site-nav .burger,
    body[${marker}] .site-nav .burger::before,
    body[${marker}] .site-nav .burger::after{
      width:16px;
      height:1.5px;
      background:#1b1b1f;
      border-radius:2px;
    }
    body[${marker}] .site-nav .burger::before{top:-5px}
    body[${marker}] .site-nav .burger::after{top:5px}
    body[${marker}] .site-nav .mobile-panel{
      position:absolute;
      top:46px;
      right:0;
      width:auto;
      min-width:240px;
      padding:8px;
      border:1px solid #d1d1cf;
      border-radius:14px;
      background:#fff;
      box-shadow:0 18px 48px rgba(24,55,38,.16);
    }
    body[${marker}] .site-nav .mobile-panel a{
      display:block;
      padding:11px 12px;
      border-radius:9px;
      color:#1b1b1f;
      font-size:13px;
      line-height:1.35;
      font-weight:550;
    }
    body[${marker}] .site-nav .mobile-panel a:hover{background:#e9f5ee}
  }

  @media(max-width:640px){
    body[${marker}] .site-nav .nav-inner{
      width:calc(100% - 28px);
      min-height:56px;
      gap:10px;
    }
    body[${marker}] .site-nav .brand{gap:7px}
    body[${marker}] .site-nav .brand>img{width:32px;height:32px}
    body[${marker}] .site-nav .brand-copy strong{font-size:16px}
    body[${marker}] .site-nav .brand-copy small{display:none}
    body[${marker}] .site-nav .nav-cta{
      min-height:36px!important;
      height:36px;
      padding-inline:11px!important;
      font-size:12px!important;
    }
    body[${marker}] .site-nav .mobile-menu summary{width:36px;height:36px}
  }

  @media(max-width:420px){
    body[${marker}] .site-nav .brand-copy strong{font-size:15px}
    body[${marker}] .site-nav .nav-cta{padding-inline:9px!important}
  }
</style>`;

if (!html.includes('</head>')) throw new Error('SPECIAL STANDARD HEADER V14: head close missing');
html = html.replace('</head>', `${css}\n</head>`);

for (const token of [
  marker,
  'width:min(1280px,calc(100% - 48px))',
  'min-height:68px',
  'width:40px',
  'font-size:18px',
  'gap:4px',
  'min-height:42px!important',
  '@media(max-width:1020px)',
  'width:calc(100% - 28px)',
  'min-height:56px',
]) {
  if (!html.includes(token)) throw new Error(`SPECIAL STANDARD HEADER V14: standard header token missing: ${token}`);
}

for (const label of [
  'Ne zaman gerekir?',
  'Hangi alanlarda çalışıyoruz?',
  'Nasıl çalışıyoruz?',
  'Karşılaştırma',
  'Neden ExcelArşiv?',
  'SSS',
]) {
  const occurrences = html.split(label).length - 1;
  if (occurrences < 2) throw new Error(`SPECIAL STANDARD HEADER V14: category label drifted or missing: ${label}`);
}

if (html.includes('overflow-x:hidden')) throw new Error('SPECIAL STANDARD HEADER V14: forbidden global overflow hiding introduced');

fs.writeFileSync(distFile, html, 'utf8');
console.log('SPECIAL STANDARD HEADER V14 PASS — special page header now matches the home header dimensions, spacing, typography, responsive breakpoint and mobile menu behavior while preserving all category labels.');
