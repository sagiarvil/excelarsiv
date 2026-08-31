#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/index.html');
const cssFile = path.resolve('src/styles/home-desktop-premium-v19.css');
if (!fs.existsSync(file)) throw new Error('HOME DESKTOP PREMIUM V19: dist/index.html missing');
if (!fs.existsSync(cssFile)) throw new Error('HOME DESKTOP PREMIUM V19: desktop CSS missing');
let html = fs.readFileSync(file, 'utf8');
const desktopCss = fs.readFileSync(cssFile, 'utf8').trim();

for (const marker of ['data-dual-funnel-home-v17','data-mobile-premium-v18']) {
  if (!html.includes(marker)) throw new Error(`HOME DESKTOP PREMIUM V19: prerequisite missing: ${marker}`);
}
if (!/@media\(min-width:1021px\)/u.test(desktopCss)) throw new Error('HOME DESKTOP PREMIUM V19: desktop media query missing');
if (/@media\(max-width:/u.test(desktopCss)) throw new Error('HOME DESKTOP PREMIUM V19: desktop CSS must not contain mobile max-width overrides');
if (/overflow-x\s*:\s*hidden/u.test(desktopCss)) throw new Error('HOME DESKTOP PREMIUM V19: forbidden overflow-x:hidden fallback');

if (!html.includes('data-desktop-premium-v19')) {
  html = html.replace(/<body\b([^>]*)>/u, '<body$1 data-desktop-premium-v19>');
}
if (!html.includes('id="home-desktop-premium-v19-css"')) {
  html = html.replace('</head>', `<style id="home-desktop-premium-v19-css">${desktopCss}</style>\n</head>`);
}

for (const token of ['data-desktop-premium-v19','home-desktop-premium-v19-css','.hero-panel::after','.hero-search-rail','.finance-pillars','.high-ticket-bridge']) {
  if (!html.includes(token)) throw new Error(`HOME DESKTOP PREMIUM V19: required token missing: ${token}`);
}

fs.writeFileSync(file, html, 'utf8');
console.log('HOME DESKTOP PREMIUM V19 PASS — premium desktop presentation is isolated at >=1021px; V18 mobile remains untouched.');
