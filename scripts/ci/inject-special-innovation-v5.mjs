import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/ozel-excel-sistemleri/index.html');
const fragment = path.resolve('scripts/ci/fragments/ozel-excel-decision-lab.html');
if (!fs.existsSync(file)) throw new Error('SPECIAL INNOVATION GATE: özel Excel build çıktısı bulunamadı');
if (!fs.existsSync(fragment)) throw new Error('SPECIAL INNOVATION GATE: decision-lab fragment bulunamadı');

let html = fs.readFileSync(file, 'utf8');
const lab = fs.readFileSync(fragment, 'utf8').trim();
const cssLink = '<link id="special-innovation-css" rel="stylesheet" href="/styles/ozel-excel-innovation.css" />';
const jsLink = '<script id="special-innovation-js" src="/scripts/ozel-excel-innovation.js" defer></script>';

if (!html.includes('data-special-innovation="v4"')) {
  const next = html.replace(/<body([^>]*)>/i, '<body$1 data-special-innovation="v4">');
  if (next === html) throw new Error('SPECIAL INNOVATION GATE: body anchor bulunamadı');
  html = next;
}
if (!html.includes('id="special-innovation-css"')) {
  if (!html.includes('</head>')) throw new Error('SPECIAL INNOVATION GATE: </head> bulunamadı');
  html = html.replace('</head>', `${cssLink}\n</head>`);
}
if (!html.includes('id="karar-laboratuvari"')) {
  const anchors = ['<section class="trust"', '<section class="solutions"', '<section class="benefits"'];
  const anchor = anchors.find((token) => html.includes(token));
  if (!anchor) throw new Error('SPECIAL INNOVATION GATE: insertion anchor bulunamadı');
  html = html.replace(anchor, `${lab}\n${anchor}`);
}
if (!html.includes('id="special-innovation-js"')) {
  if (!html.includes('</body>')) throw new Error('SPECIAL INNOVATION GATE: </body> bulunamadı');
  html = html.replace('</body>', `${jsLink}\n</body>`);
}

for (const token of ['+90 542 123 45 67', 'tel:+905421234567']) {
  if (html.includes(token)) throw new Error(`SPECIAL INNOVATION GATE: doğrulanmamış iletişim tokenı geri geldi: ${token}`);
}
for (const required of [
  'data-special-innovation="v4"',
  'id="special-innovation-css"',
  'id="karar-laboratuvari"',
  'id="iv-panel-nakit"',
  'id="iv-panel-cari"',
  'id="iv-panel-banka"',
  'id="iv-panel-yonetim"',
  'id="special-innovation-js"',
]) {
  if (!html.includes(required)) throw new Error(`SPECIAL INNOVATION GATE: zorunlu contract eksik: ${required}`);
}

fs.writeFileSync(file, html);
console.log('SPECIAL INNOVATION GATE PASS — decision lab + runtime sheet tabs + light enterprise visual system injected.');
