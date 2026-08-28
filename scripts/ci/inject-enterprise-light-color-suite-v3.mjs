#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const cssPublic = path.resolve('public/styles/enterprise-light-color-suite-v3.css');
const cssDist = path.resolve('dist/styles/enterprise-light-color-suite-v3.css');
const linkId = 'enterprise-light-color-suite-v3';
const href = '/styles/enterprise-light-color-suite-v3.css';

const routes = [
  { file: path.resolve('dist/index.html'), bodyClass: 'ea-home-color-v3', label: 'home' },
  { file: path.resolve('dist/hakkinda/index.html'), bodyClass: 'ea-about-color-v3', label: 'about' },
  { file: path.resolve('dist/rehber/index.html'), bodyClass: 'ea-guide-color-v3', label: 'guide' },
  { file: path.resolve('dist/nasil-calisir/index.html'), bodyClass: 'ea-how-color-v3', label: 'how' },
  { file: path.resolve('dist/sablonlar/index.html'), bodyClass: 'ea-catalog-color-v3', label: 'catalog' },
];

for (const file of [cssPublic, cssDist, ...routes.map((route) => route.file)]) {
  if (!fs.existsSync(file)) throw new Error(`ENTERPRISE LIGHT COLOR GATE: missing ${file}`);
}

const css = fs.readFileSync(cssPublic, 'utf8');
const withoutComments = css.replace(/\/\*[\s\S]*?\*\//g, '');
const openBraces = (withoutComments.match(/\{/g) || []).length;
const closeBraces = (withoutComments.match(/\}/g) || []).length;
if (openBraces !== closeBraces) throw new Error(`ENTERPRISE LIGHT COLOR GATE: brace mismatch ${openBraces}/${closeBraces}`);

for (const token of [
  'body.ea-home-color-v3 .difference',
  'body.ea-home-color-v3 .authority-panel',
  'body.ea-about-color-v3 .hakkinda__hero',
  'body.ea-guide-color-v3 #icerik>header',
  'body.ea-how-color-v3 #icerik>section>ol',
  'body.ea-catalog-color-v3 .template-grid>li>.card',
  '--ea3-rail:',
  '@media(prefers-reduced-motion:reduce)',
]) {
  if (!css.includes(token)) throw new Error(`ENTERPRISE LIGHT COLOR GATE: required CSS contract missing ${token}`);
}

// Explicit no-dark-surface contracts for the five requested pages.
for (const token of [
  'body.ea-home-color-v3 .difference{',
  'body.ea-home-color-v3 .authority-panel{',
  'body.ea-about-color-v3 .hakkinda__hero{',
  'body.ea-about-color-v3 .hakkinda__cta{',
  'body.ea-guide-color-v3 aside>div[class*="from-[#0B192C]"]{',
]) {
  if (!css.includes(token)) throw new Error(`ENTERPRISE LIGHT COLOR GATE: light-surface override missing ${token}`);
}

function addBodyClass(html, bodyClass) {
  return html.replace(/<body\b([^>]*)>/i, (tag) => {
    const classMatch = tag.match(/\bclass=(['"])(.*?)\1/i);
    if (classMatch) {
      const classes = classMatch[2].split(/\s+/).filter(Boolean);
      if (!classes.includes(bodyClass)) classes.push(bodyClass);
      return tag.replace(classMatch[0], `class=${classMatch[1]}${classes.join(' ')}${classMatch[1]}`);
    }
    return tag.replace('<body', `<body class="${bodyClass}"`);
  });
}

for (const route of routes) {
  let html = fs.readFileSync(route.file, 'utf8');
  html = addBodyClass(html, route.bodyClass);

  // Remove previous copies before appending a single deterministic final stylesheet link.
  html = html.replace(new RegExp(`<link\\b(?=[^>]*\\bid=["']${linkId}["'])[^>]*>`, 'gi'), '');
  html = html.replace(new RegExp(`<link\\b(?=[^>]*\\bhref=["']${href.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}["'])[^>]*>`, 'gi'), '');

  if (!html.includes('</head>')) throw new Error(`ENTERPRISE LIGHT COLOR GATE: ${route.label} missing </head>`);
  html = html.replace('</head>', `<link id="${linkId}" rel="stylesheet" href="${href}" />\n</head>`);

  for (const required of [route.bodyClass, `id="${linkId}"`, `href="${href}"`]) {
    if (!html.includes(required)) throw new Error(`ENTERPRISE LIGHT COLOR GATE: ${route.label} missing ${required}`);
  }

  fs.writeFileSync(route.file, html);
}

// The stylesheet must be physically present in the final hosting bundle.
const publicHash = fs.readFileSync(cssPublic, 'utf8');
const distHash = fs.readFileSync(cssDist, 'utf8');
if (publicHash !== distHash) throw new Error('ENTERPRISE LIGHT COLOR GATE: public/dist stylesheet parity failed');

console.log('ENTERPRISE LIGHT COLOR SUITE PASS — home/about/guide/how/catalog use final light-only Enterprise/Exclusive color layer.');
