#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const sourceFile = path.resolve('src/pages/ozel-excel-sistemleri.astro');
const distFile = path.resolve('dist/ozel-excel-sistemleri/index.html');

if (!fs.existsSync(sourceFile)) throw new Error('SPECIAL SOURCE V5: source file missing');
if (!fs.existsSync(distFile)) throw new Error('SPECIAL SOURCE V5: dist route missing');

const source = fs.readFileSync(sourceFile, 'utf8');
const html = source.replace(/^---[\s\S]*?---\s*/u, '');

const required = [
  '<!doctype html>',
  'data-special-source-v5',
  'class="workbook"',
  '--font:-apple-system,BlinkMacSystemFont',
  'font-size:17px',
  'Hazır dosya işinize uymuyorsa, Excel\'i işinize göre kuralım.',
  'Özel bir sisteme ne zaman ihtiyaç olur?',
  'Hazır şablon, standart geliştirme ve size özel Excel aynı şey değil.',
  '@media(max-width:800px)',
  'data-label="Size özel sistem"',
];
for (const token of required) {
  if (!html.includes(token)) throw new Error(`SPECIAL SOURCE V5: required token missing: ${token}`);
}

const forbidden = [
  '@font-face',
  'Manrope',
  'IBM Plex Mono',
  'data-special-light-legacy-bridge',
  'data-special-innovation=',
  'id="special-innovation-css"',
  'id="special-brand-sync-css"',
  'id="special-layout-stabilizer-css"',
  'id="native-seo-infographics"',
  'id="special-innovation-js"',
  'karar-laboratuvari',
  'min-width:780px',
  'min-width:850px',
  'diagnosis-head',
  'intent-card',
  'special-page-v4',
];
for (const token of forbidden) {
  if (html.includes(token)) throw new Error(`SPECIAL SOURCE V5: forbidden legacy token in source: ${token}`);
}

// Taşma riski: büyük sabit width/min-width. max-width ve media breakpoint'leri
// responsive sınırlardır ve bu gate tarafından hata sayılmaz.
const fixedWidthRisk = /(?:^|[;{])\s*(?:width|min-width):\s*(?:7\d{2}|8\d{2}|9\d{2}|1\d{3,})px/gi;
const riskyMatches = [...html.matchAll(fixedWidthRisk)].map((match) => match[0].trim());
if (riskyMatches.length) {
  throw new Error(`SPECIAL SOURCE V5: large rigid width risk: ${riskyMatches.join(', ')}`);
}

const unbalancedGridRisk = /grid-template-columns:[^;}]*\b(?:300|320|340|360|400)px\b/gi;
if (unbalancedGridRisk.test(html)) throw new Error('SPECIAL SOURCE V5: rigid side column detected');

fs.mkdirSync(path.dirname(distFile), { recursive: true });
fs.writeFileSync(distFile, html, 'utf8');

const finalHtml = fs.readFileSync(distFile, 'utf8');
for (const token of required) {
  if (!finalHtml.includes(token)) throw new Error(`SPECIAL SOURCE V5: final artifact lost required token: ${token}`);
}
for (const token of forbidden) {
  if (finalHtml.includes(token)) throw new Error(`SPECIAL SOURCE V5: final artifact contains legacy token: ${token}`);
}

console.log('SPECIAL SOURCE V5 PASS — source owns final route; Apple system typography, responsive comparison and legacy post-build visual mutations removed.');
