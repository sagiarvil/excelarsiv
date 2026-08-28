#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const homeFile = path.resolve('dist/index.html');
const cssFile = path.resolve('src/styles/home-authority-label-contrast-v34.css');
const styleId = 'home-authority-label-contrast-v34';

for (const file of [homeFile, cssFile]) {
  if (!fs.existsSync(file)) throw new Error(`HOME AUTHORITY CONTRAST GATE: missing ${file}`);
}

const css = fs.readFileSync(cssFile, 'utf8');
const withoutComments = css.replace(/\/\*[\s\S]*?\*\//g, '');
const openBraces = (withoutComments.match(/\{/g) || []).length;
const closeBraces = (withoutComments.match(/\}/g) || []).length;
if (openBraces !== closeBraces) throw new Error(`HOME AUTHORITY CONTRAST GATE: CSS brace mismatch ${openBraces}/${closeBraces}`);

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
console.log('HOME AUTHORITY CONTRAST PASS — KATALOG / KAPSAM / SEÇENEK labels locked to dark readable finance tones.');
