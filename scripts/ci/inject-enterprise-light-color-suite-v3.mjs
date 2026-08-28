#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const cssPublic = path.resolve('public/styles/enterprise-light-color-suite-v3.css');
const cssDist = path.resolve('dist/styles/enterprise-light-color-suite-v3.css');
const qaCssPublic = path.resolve('public/styles/enterprise-light-color-suite-v3-qa.css');
const qaCssDist = path.resolve('dist/styles/enterprise-light-color-suite-v3-qa.css');
const linkId = 'enterprise-light-color-suite-v3';
const href = '/styles/enterprise-light-color-suite-v3.css';
const qaLinkId = 'enterprise-light-color-suite-v3-qa';
const qaHref = '/styles/enterprise-light-color-suite-v3-qa.css';

const routes = [
  { file: path.resolve('dist/index.html'), bodyClass: 'ea-home-color-v3', label: 'home' },
  { file: path.resolve('dist/hakkinda/index.html'), bodyClass: 'ea-about-color-v3', label: 'about' },
  { file: path.resolve('dist/rehber/index.html'), bodyClass: 'ea-guide-color-v3', label: 'guide' },
  { file: path.resolve('dist/nasil-calisir/index.html'), bodyClass: 'ea-how-color-v3', label: 'how' },
  { file: path.resolve('dist/sablonlar/index.html'), bodyClass: 'ea-catalog-color-v3', label: 'catalog' },
];

for (const file of [cssPublic, cssDist, qaCssPublic, qaCssDist, ...routes.map((route) => route.file)]) {
  if (!fs.existsSync(file)) throw new Error(`ENTERPRISE LIGHT COLOR GATE: missing ${file}`);
}

function assertBalancedCss(file, label) {
  const css = fs.readFileSync(file, 'utf8');
  const withoutComments = css.replace(/\/\*[\s\S]*?\*\//g, '');
  const openBraces = (withoutComments.match(/\{/g) || []).length;
  const closeBraces = (withoutComments.match(/\}/g) || []).length;
  if (openBraces !== closeBraces) throw new Error(`ENTERPRISE LIGHT COLOR GATE: ${label} brace mismatch ${openBraces}/${closeBraces}`);
  return css;
}

const css = assertBalancedCss(cssPublic, 'v3');
const qaCss = assertBalancedCss(qaCssPublic, 'v3.1-qa');

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

for (const token of [
  '--ea31-max:1180px',
  'body.ea-home-color-v3 .featured-grid article>div',
  'body.ea-about-color-v3 .hakkinda{width:min(1080px',
  'body.ea-guide-color-v3 #icerik aside{position:sticky',
  'body.ea-how-color-v3 #icerik>section>ol{display:grid',
  'body.ea-catalog-color-v3 .template-grid>li>.card{height:100%',
  '@media(max-width:520px)',
  '@media(prefers-reduced-motion:reduce)',
]) {
  if (!qaCss.includes(token)) throw new Error(`ENTERPRISE LIGHT COLOR QA GATE: required geometry contract missing ${token}`);
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

function removeCatalogFlow(html) {
  const before = html;
  html = html.replace(/<section\b[^>]*class=(['"])[^'\"]*\bcatalog-flow\b[^'\"]*\1[^>]*>[\s\S]*?<\/section>/i, '');
  if (html === before) throw new Error('CATALOG FLOW REMOVAL GATE: catalog-flow section was not found');
  for (const forbidden of ['data-catalog-infographic', 'HER ÜRÜN AYNI NET MANTIKLA ÇALIŞIR', 'Veriyi girin. Excel işlesin. Kararı görün.']) {
    if (html.includes(forbidden)) throw new Error(`CATALOG FLOW REMOVAL GATE: forbidden final HTML token remains: ${forbidden}`);
  }
  return html;
}

for (const route of routes) {
  let html = fs.readFileSync(route.file, 'utf8');
  html = addBodyClass(html, route.bodyClass);

  if (route.label === 'catalog') html = removeCatalogFlow(html);

  // Remove previous copies before appending deterministic final stylesheet links.
  for (const [id, linkHref] of [[linkId, href], [qaLinkId, qaHref]]) {
    html = html.replace(new RegExp(`<link\\b(?=[^>]*\\bid=["']${id}["'])[^>]*>`, 'gi'), '');
    html = html.replace(new RegExp(`<link\\b(?=[^>]*\\bhref=["']${linkHref.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}["'])[^>]*>`, 'gi'), '');
  }

  if (!html.includes('</head>')) throw new Error(`ENTERPRISE LIGHT COLOR GATE: ${route.label} missing </head>`);
  html = html.replace(
    '</head>',
    `<link id="${linkId}" rel="stylesheet" href="${href}" />\n<link id="${qaLinkId}" rel="stylesheet" href="${qaHref}" />\n</head>`,
  );

  for (const required of [route.bodyClass, `id="${linkId}"`, `href="${href}"`, `id="${qaLinkId}"`, `href="${qaHref}"`]) {
    if (!html.includes(required)) throw new Error(`ENTERPRISE LIGHT COLOR GATE: ${route.label} missing ${required}`);
  }

  const mainIndex = html.indexOf(`id="${linkId}"`);
  const qaIndex = html.indexOf(`id="${qaLinkId}"`);
  if (!(mainIndex >= 0 && qaIndex > mainIndex)) throw new Error(`ENTERPRISE LIGHT COLOR QA GATE: ${route.label} QA layer must load after v3`);

  fs.writeFileSync(route.file, html);
}

// Stylesheets must be physically present and byte-identical in final hosting bundle.
if (fs.readFileSync(cssPublic, 'utf8') !== fs.readFileSync(cssDist, 'utf8')) {
  throw new Error('ENTERPRISE LIGHT COLOR GATE: public/dist v3 stylesheet parity failed');
}
if (fs.readFileSync(qaCssPublic, 'utf8') !== fs.readFileSync(qaCssDist, 'utf8')) {
  throw new Error('ENTERPRISE LIGHT COLOR QA GATE: public/dist v3.1 stylesheet parity failed');
}

console.log('ENTERPRISE LIGHT COLOR SUITE PASS — v3 color layer + v3.1 geometry/mobile QA loaded; catalog flow removed from final templates HTML.');