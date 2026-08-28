import fs from 'node:fs';
import path from 'node:path';

const htmlFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
const innovationFile = path.resolve('public/styles/ozel-excel-innovation.css');
const brandFile = path.resolve('public/styles/ozel-excel-brand-sync.css');

for (const file of [htmlFile, innovationFile, brandFile]) {
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
validateCss('innovation', innovationCss);
validateCss('brand', brandCss);

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
  const tag = `<style id="${id}" data-inline-special-css data-source="${sourcePath}">\n${css}\n</style>`;
  if (byId.test(html)) {
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

for (const required of [
  'class="special-light-v1',
  'data-special-light-v1',
  'data-special-innovation="v4"',
  'id="special-innovation-css"',
  'id="special-brand-sync-css"',
  'data-inline-special-css',
]) {
  if (!html.includes(required)) throw new Error(`SPECIAL CSS REPAIR GATE: final HTML missing ${required}`);
}

if (html.includes('href="/styles/ozel-excel-innovation.css"') || html.includes('href="/styles/ozel-excel-brand-sync.css"')) {
  throw new Error('SPECIAL CSS REPAIR GATE: external special CSS link survived; final render would remain network-dependent');
}

fs.writeFileSync(htmlFile, html);
console.log('SPECIAL CSS REPAIR PASS — innovation + brand CSS inlined, namespace verified, external CSS dependency removed.');
