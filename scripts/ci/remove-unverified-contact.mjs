import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/ozel-excel-sistemleri/index.html');
const forbiddenPhone = '+90 542 123 45 67';
const forbiddenHref = 'tel:+905421234567';

if (!fs.existsSync(file)) {
  throw new Error('CONTACT GATE: özel Excel sayfası build çıktısı bulunamadı');
}

let html = fs.readFileSync(file, 'utf8');

html = html.replace(
  /<a\s+class=["']phone["']\s+href=["']tel:\+905421234567["'][\s\S]*?<\/a>/i,
  '',
);

if (html.includes(forbiddenPhone) || html.includes(forbiddenHref)) {
  throw new Error('CONTACT GATE: doğrulanmamış telefon numarası build çıktısında kaldı');
}

fs.writeFileSync(file, html);
console.log('CONTACT GATE PASS — doğrulanmamış telefon numarası özel Excel sayfasından kaldırıldı.');
