import fs from 'node:fs';
import path from 'node:path';

const htmlFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
const innovationFile = path.resolve('public/styles/ozel-excel-innovation.css');
const brandFile = path.resolve('public/styles/ozel-excel-brand-sync.css');
const stabilizerFile = path.resolve('public/styles/ozel-excel-layout-stabilizer.css');

for (const file of [htmlFile, innovationFile, brandFile, stabilizerFile]) {
  if (!fs.existsSync(file)) throw new Error(`SPECIAL CSS REPAIR GATE: missing ${file}`);
}

const validateCss = (name, css) => {
  const withoutComments = css.replace(/\/\*[\s\S]*?\*\//g, '');
  const open = (withoutComments.match(/\{/g) || []).length;
  const close = (withoutComments.match(/\}/g) || []).length;
  if (open !== close) throw new Error(`SPECIAL CSS REPAIR GATE: ${name} brace mismatch ${open}/${close}`);
  if (!css.includes('special-light-v1') || !css.includes('data-special-innovation="v4"')) {
    throw new Error(`SPECIAL CSS REPAIR GATE: ${name} namespace missing`);
  }
  if (css.includes('</style>')) throw new Error(`SPECIAL CSS REPAIR GATE: ${name} contains unsafe </style>`);
};

const innovationCss = fs.readFileSync(innovationFile, 'utf8');
const brandCss = fs.readFileSync(brandFile, 'utf8');
const stabilizerCss = fs.readFileSync(stabilizerFile, 'utf8');
validateCss('innovation', innovationCss);
validateCss('brand', brandCss);
validateCss('stabilizer', stabilizerCss);

let html = fs.readFileSync(htmlFile, 'utf8');

html = html.replace(/<body\b([^>]*)>/i, (tag) => {
  let next = tag;
  const classMatch = next.match(/\bclass=(['"])(.*?)\1/i);
  if (classMatch) {
    const classes = classMatch[2].split(/\s+/).filter(Boolean);
    if (!classes.includes('special-light-v1')) {
      const merged = [...classes, 'special-light-v1'].join(' ');
      next = next.replace(classMatch[0], `class=${classMatch[1]}${merged}${classMatch[1]}`);
    }
  } else {
    next = next.replace('<body', '<body class="special-light-v1"');
  }
  if (!/\bdata-special-light-v1(?:\s|=|>)/i.test(next)) {
    next = next.replace('<body', '<body data-special-light-v1');
  }
  if (!/\bdata-special-innovation=(['"])v4\1/i.test(next)) {
    next = next.replace('<body', '<body data-special-innovation="v4"');
  }
  return next;
});

const inlineStylesheet = (id, sourcePath, css) => {
  const byId = new RegExp(`<link\\b(?=[^>]*\\bid=["']${id}["'])[^>]*>`, 'i');
  const byHref = new RegExp(`<link\\b(?=[^>]*\\brel=["']stylesheet["'])(?=[^>]*\\bhref=["']${sourcePath.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}["'])[^>]*>`, 'i');
  const existingStyle = new RegExp(`<style\\b(?=[^>]*\\bid=["']${id}["'])[^>]*>[\\s\\S]*?<\\/style>`, 'i');
  const tag = `<style id="${id}" data-inline-special-css data-source="${sourcePath}">\n${css}\n</style>`;
  if (existingStyle.test(html)) {
    html = html.replace(existingStyle, tag);
  } else if (byId.test(html)) {
    html = html.replace(byId, tag);
  } else if (byHref.test(html)) {
    html = html.replace(byHref, tag);
  } else if (!html.includes(`id="${id}"`)) {
    if (!html.includes('</head>')) throw new Error(`SPECIAL CSS REPAIR GATE: </head> missing for ${id}`);
    html = html.replace('</head>', `${tag}\n</head>`);
  }
};

inlineStylesheet('special-innovation-css', '/styles/ozel-excel-innovation.css', innovationCss);
inlineStylesheet('special-brand-sync-css', '/styles/ozel-excel-brand-sync.css', brandCss);
inlineStylesheet('special-layout-stabilizer-css', '/styles/ozel-excel-layout-stabilizer.css', stabilizerCss);

/* Ensure the stabilizer is physically the last special style in <head>, so later
   source/enterprise/editorial layers cannot override the deterministic layout. */
const stabilizerStylePattern = /<style\b(?=[^>]*\bid=["']special-layout-stabilizer-css["'])[^>]*>[\s\S]*?<\/style>/i;
const stabilizerMatch = html.match(stabilizerStylePattern);
if (!stabilizerMatch) throw new Error('SPECIAL CSS REPAIR GATE: stabilizer style missing after inline');
html = html.replace(stabilizerStylePattern, '');
if (!html.includes('</head>')) throw new Error('SPECIAL CSS REPAIR GATE: </head> missing while ordering stabilizer');
html = html.replace('</head>', `${stabilizerMatch[0]}\n</head>`);

for (const required of [
  'class="special-light-v1',
  'data-special-light-v1',
  'data-special-innovation="v4"',
  'id="special-innovation-css"',
  'id="special-brand-sync-css"',
  'id="special-layout-stabilizer-css"',
  'data-inline-special-css',
]) {
  if (!html.includes(required)) throw new Error(`SPECIAL CSS REPAIR GATE: final HTML missing ${required}`);
}

for (const href of [
  '/styles/ozel-excel-innovation.css',
  '/styles/ozel-excel-brand-sync.css',
  '/styles/ozel-excel-layout-stabilizer.css',
]) {
  if (html.includes(`href="${href}"`)) throw new Error(`SPECIAL CSS REPAIR GATE: external special CSS link survived: ${href}`);
}

for (const layoutContract of [
  '--stable-header:1440px',
  'grid-template-columns:auto minmax(0,1fr) auto!important',
  '.nav-links a:visited{color:#33463b!important}',
  'height:auto!important',
  'min-height:0!important',
  'margin-block:0!important;\n  margin-inline:auto!important',
]) {
  if (!stabilizerCss.includes(layoutContract)) throw new Error(`SPECIAL CSS REPAIR GATE: stabilizer layout contract missing ${layoutContract}`);
}

fs.writeFileSync(htmlFile, html);
console.log('SPECIAL CSS REPAIR PASS — innovation + brand + final layout stabilizer inlined; deterministic header/section geometry verified.');
