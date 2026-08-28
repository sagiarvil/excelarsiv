#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const cssPublic = path.resolve('public/styles/enterprise-light-color-suite-v3.css');
const cssDist = path.resolve('dist/styles/enterprise-light-color-suite-v3.css');
const qaCssPublic = path.resolve('public/styles/enterprise-light-color-suite-v3-qa.css');
const qaCssDist = path.resolve('dist/styles/enterprise-light-color-suite-v3-qa.css');
const assuranceCssPublic = path.resolve('public/styles/catalog-help-premium-v32.css');
const assuranceCssDist = path.resolve('dist/styles/catalog-help-premium-v32.css');
const homeHardColorCssSource = path.resolve('src/styles/home-native-info-hard-color-v33.css');

const linkId = 'enterprise-light-color-suite-v3';
const href = '/styles/enterprise-light-color-suite-v3.css';
const qaLinkId = 'enterprise-light-color-suite-v3-qa';
const qaHref = '/styles/enterprise-light-color-suite-v3-qa.css';
const assuranceLinkId = 'catalog-help-premium-v32';
const assuranceHref = '/styles/catalog-help-premium-v32.css';
const homeHardColorStyleId = 'home-native-info-hard-color-v33';

const routes = [
  { file: path.resolve('dist/index.html'), bodyClass: 'ea-home-color-v3', label: 'home' },
  { file: path.resolve('dist/hakkinda/index.html'), bodyClass: 'ea-about-color-v3', label: 'about' },
  { file: path.resolve('dist/rehber/index.html'), bodyClass: 'ea-guide-color-v3', label: 'guide' },
  { file: path.resolve('dist/nasil-calisir/index.html'), bodyClass: 'ea-how-color-v3', label: 'how' },
  { file: path.resolve('dist/sablonlar/index.html'), bodyClass: 'ea-catalog-color-v3', label: 'catalog' },
];

for (const file of [cssPublic, cssDist, qaCssPublic, qaCssDist, assuranceCssPublic, assuranceCssDist, homeHardColorCssSource, ...routes.map((route) => route.file)]) {
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
const assuranceCss = assertBalancedCss(assuranceCssPublic, 'v3.2-catalog-assurance');
const homeHardColorCss = assertBalancedCss(homeHardColorCssSource, 'v3.3-home-hard-color');

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

for (const token of [
  'body.ea-catalog-color-v3 .catalog-help{',
  'body.ea-catalog-color-v3 .catalog-help__item:nth-child(2)',
  'body.ea-catalog-color-v3 .catalog-help__item:nth-child(3)',
  'linear-gradient(90deg,var(--ea3-green)',
  '@media(max-width:520px)',
]) {
  if (!assuranceCss.includes(token)) throw new Error(`CATALOG ASSURANCE GATE: premium board contract missing ${token}`);
}

for (const token of [
  'body.ea-home-color-v3 .native-info--home',
  '--hm-green:#0a914a',
  '--hm-blue:#176fe5',
  '--hm-amber:#ee9d00',
  '--hm-coral:#f05a47',
  '.native-info__core',
  '.native-info__outcomes article:nth-child(4)',
  '@media(max-width:620px)',
  '@media(prefers-reduced-motion:reduce)',
]) {
  if (!homeHardColorCss.includes(token)) throw new Error(`HOME HARD COLOR GATE: required homepage-only contract missing ${token}`);
}

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

  const links = [
    [linkId, href],
    [qaLinkId, qaHref],
    ...(route.label === 'catalog' ? [[assuranceLinkId, assuranceHref]] : []),
  ];

  for (const [id, linkHref] of links) {
    html = html.replace(new RegExp(`<link\\b(?=[^>]*\\bid=["']${id}["'])[^>]*>`, 'gi'), '');
    html = html.replace(new RegExp(`<link\\b(?=[^>]*\\bhref=["']${linkHref.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}["'])[^>]*>`, 'gi'), '');
  }
  html = html.replace(new RegExp(`<style\\b(?=[^>]*\\bid=["']${homeHardColorStyleId}["'])[^>]*>[\\s\\S]*?<\\/style>`, 'gi'), '');

  if (!html.includes('</head>')) throw new Error(`ENTERPRISE LIGHT COLOR GATE: ${route.label} missing </head>`);
  const linkMarkup = links.map(([id, linkHref]) => `<link id="${id}" rel="stylesheet" href="${linkHref}" />`).join('\n');
  const homeOnlyMarkup = route.label === 'home' ? `\n<style id="${homeHardColorStyleId}">\n${homeHardColorCss}\n</style>` : '';
  html = html.replace('</head>', `${linkMarkup}${homeOnlyMarkup}\n</head>`);

  for (const [id, linkHref] of links) {
    for (const required of [`id="${id}"`, `href="${linkHref}"`]) {
      if (!html.includes(required)) throw new Error(`ENTERPRISE LIGHT COLOR GATE: ${route.label} missing ${required}`);
    }
  }
  if (!html.includes(route.bodyClass)) throw new Error(`ENTERPRISE LIGHT COLOR GATE: ${route.label} missing ${route.bodyClass}`);

  const mainIndex = html.indexOf(`id="${linkId}"`);
  const qaIndex = html.indexOf(`id="${qaLinkId}"`);
  if (!(mainIndex >= 0 && qaIndex > mainIndex)) throw new Error(`ENTERPRISE LIGHT COLOR QA GATE: ${route.label} QA layer must load after v3`);
  if (route.label === 'catalog') {
    const assuranceIndex = html.indexOf(`id="${assuranceLinkId}"`);
    if (!(assuranceIndex > qaIndex)) throw new Error('CATALOG ASSURANCE GATE: premium assurance stylesheet must load after QA');
  }
  if (route.label === 'home') {
    const hardColorIndex = html.indexOf(`id="${homeHardColorStyleId}"`);
    if (!(hardColorIndex > qaIndex)) throw new Error('HOME HARD COLOR GATE: homepage hard-color style must load after v3.1 QA');
    if (!html.includes('body.ea-home-color-v3 .native-info--home')) throw new Error('HOME HARD COLOR GATE: inline homepage style missing decision-map selector');
  } else if (html.includes(`id="${homeHardColorStyleId}"`)) {
    throw new Error(`HOME HARD COLOR GATE: homepage-only style leaked into ${route.label}`);
  }

  fs.writeFileSync(route.file, html);
}

if (fs.readFileSync(cssPublic, 'utf8') !== fs.readFileSync(cssDist, 'utf8')) {
  throw new Error('ENTERPRISE LIGHT COLOR GATE: public/dist v3 stylesheet parity failed');
}
if (fs.readFileSync(qaCssPublic, 'utf8') !== fs.readFileSync(qaCssDist, 'utf8')) {
  throw new Error('ENTERPRISE LIGHT COLOR QA GATE: public/dist v3.1 stylesheet parity failed');
}
if (fs.readFileSync(assuranceCssPublic, 'utf8') !== fs.readFileSync(assuranceCssDist, 'utf8')) {
  throw new Error('CATALOG ASSURANCE GATE: public/dist v3.2 stylesheet parity failed');
}

console.log('ENTERPRISE LIGHT COLOR SUITE PASS — v3 + v3.1 + catalog v3.2 + homepage-only hard-color v3.3; catalog flow removed.');