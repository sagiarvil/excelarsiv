import fs from 'node:fs';
import path from 'node:path';

const specialFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
const homeFile = path.resolve('dist/index.html');
const fragment = path.resolve('scripts/ci/fragments/ozel-excel-decision-lab.html');
if (!fs.existsSync(specialFile)) throw new Error('SPECIAL INNOVATION GATE: özel Excel build çıktısı bulunamadı');
if (!fs.existsSync(homeFile)) throw new Error('HOME COLOR GATE: ana sayfa build çıktısı bulunamadı');
if (!fs.existsSync(fragment)) throw new Error('SPECIAL INNOVATION GATE: decision-lab fragment bulunamadı');

let html = fs.readFileSync(specialFile, 'utf8');
let home = fs.readFileSync(homeFile, 'utf8');
const lab = fs.readFileSync(fragment, 'utf8').trim();
const cssLink = '<link id="special-innovation-css" rel="stylesheet" href="/styles/ozel-excel-innovation.css" />';
const brandCssLink = '<link id="special-brand-sync-css" rel="stylesheet" href="/styles/ozel-excel-brand-sync.css" />';
const jsLink = '<script id="special-innovation-js" src="/scripts/ozel-excel-innovation.js" defer></script>';
const homeColorLink = '<link id="home-finance-color-css" rel="stylesheet" href="/styles/home-finance-color-upgrade.css" />';

const replaceVisible = (from, to) => {
  if (!html.includes(to)) html = html.replaceAll(from, to);
};

if (!html.includes('data-special-innovation="v4"')) {
  const next = html.replace(/<body([^>]*)>/i, '<body$1 data-special-innovation="v4">');
  if (next === html) throw new Error('SPECIAL INNOVATION GATE: body anchor bulunamadı');
  html = next;
}
if (!html.includes('id="special-innovation-css"')) {
  if (!html.includes('</head>')) throw new Error('SPECIAL INNOVATION GATE: </head> bulunamadı');
  html = html.replace('</head>', `${cssLink}\n</head>`);
}
if (!html.includes('id="special-brand-sync-css"')) {
  if (!html.includes('</head>')) throw new Error('BRAND SYNC GATE: </head> bulunamadı');
  html = html.replace('</head>', `${brandCssLink}\n</head>`);
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

// Remove the unverified contact and the floating hero disclaimer permanently from the final build.
html = html.replace(/<a\s+class=["']phone["'][\s\S]*?<\/a>/i, '');
html = html.replace(/\s*<span class="product-badge">[\s\S]*?<\/span>/i, '');
html = html.replaceAll('Temsili ekran — proje kapsamına göre özelleştirilir', '');
html = html.replaceAll('Temsili ekran - proje kapsamına göre özelleştirilir', '');

// Match the special page header brand dimensions and identity to the main-site header.
html = html.replace(
  /<a class="brand" href="\/" aria-label="Excel Arşiv ana sayfa">[\s\S]*?<\/a>/i,
  '<a class="brand" href="/" aria-label="Excel Arşiv ana sayfa"><img src="/images/brand/excelarsiv-header-logo.png" alt="Excel Arşiv" width="40" height="40" /><span class="brand-copy"><strong>Excel Arşiv</strong><small>Hazır Excel işletme sistemleri</small></span></a>',
);

// Use the same real brand asset in the special-page footer without changing its existing footer structure.
const footerStart = html.indexOf('<footer class="footer">');
if (footerStart >= 0) {
  const headPart = html.slice(0, footerStart);
  let footerPart = html.slice(footerStart);
  footerPart = footerPart.replace('<img src="/images/excel-logo.png" alt="" />', '<img src="/images/brand/excelarsiv-header-logo.png" alt="Excel Arşiv logosu" width="40" height="40" />');
  footerPart = footerPart.replace('<strong>EXCELARŞİV</strong>', '<strong>Excel Arşiv</strong>');
  html = headPart + footerPart;
}

// CTA positioning: software knowledge + banking/KOBİ field experience in the same operating language.
replaceVisible('Ücretsiz Keşif Görüşmesi Al →', 'İşletmeme Özel Çözümü Konuşalım →');
replaceVisible('Ücretsiz Keşif Görüşmesi', 'İşletmenize Özel Çözümü Konuşalım');
replaceVisible('İşinizi kolaylaştıracak sistemi birlikte netleştirelim.', 'İşletmenizi anlayan biriyle, size özel Excel sistemi kurun.');
replaceVisible(
  '15 dakikalık ilk görüşmede mevcut iş akışınızı, en pahalı manuel kontrol noktasını ve hangi Excel mimarisinin en hızlı geri dönüşü sağlayacağını belirleyelim.',
  'Bu çalışma yalnız Excel veya yazılım işi değil. Bankacılık, kredi karar süreçleri, KOBİ danışmanlığı ve sahadaki işletme tecrübesiyle; mizanınızdan nakit akışınıza, cari yapınızdan banka limitlerinize kadar gerçek işleyişinizi önce anlıyor, sonra işletmenize özel Excel çözümünü kuruyoruz. Böylece ihtiyacınızı yazılımcıya tercüme etmek zorunda kalmazsınız; aynı dili konuşan biriyle doğrudan çözüm üretirsiniz.',
);
replaceVisible('İlk görüşme • kapsam teşhisi • çözüm haritası', 'Saha bilgisi • finans tecrübesi • yazılım disiplini');
replaceVisible('Neden ExcelArşiv?', 'Sahadan Gelen Uzmanlık');
replaceVisible('Finans mantığı ile ekran tasarımını aynı sistemde kuruyoruz.', 'Yazılımı bilen ama işletmenin finansal dilini de sahadan tanıyan bir bakışla tasarlıyoruz.');
replaceVisible('Mizan, muavin, nakit, vade ve faiz mantığı ekran tasarımından önce kurulur.', 'Bankacılık ve kredi karar mantığı; banka limiti, nakit ihtiyacı, faiz yükü ve finansman riskini yalnız hücre olarak değil, karar konusu olarak ele almamızı sağlar.');
replaceVisible('Girdi, hesaplanan alan ve yönetici raporu birbirinden ayrılır.', 'KOBİ danışmanlığı tecrübesi; patronun, muhasebenin ve finans ekibinin aynı veriye neden farklı sorular sorduğunu anlamamızı sağlar.');
replaceVisible('Hata ve istisna noktaları kullanıcıya görünür hale getirilir.', 'Mizan, cari, tahsilat, vade ve nakit akışı gibi işletme gerçekleri yazılım tasarımından önce modellenir; kullanıcıya yabancı bir sistem kurulmaz.');
replaceVisible('Proje bittiğinde kara kutu değil, işletilebilir bir çalışma düzeni kalır.', 'Sonuç hazır bir şablon değil; işletmenizin çalışma biçimine, kontrol noktalarına ve yönetim ihtiyacına göre kurulmuş özel bir Excel çözümüdür.');

// SEO intent ownership for bespoke business Excel systems.
html = html.replace(/<title>[\s\S]*?<\/title>/i, '<title>İşletmelere Özel Excel Çözümleri | Finans ve Yönetim Sistemleri | ExcelArşiv</title>');
html = html.replace(/<meta name="description" content="[^"]*"\s*\/?>/i, '<meta name="description" content="Bankacılık ve KOBİ danışmanlığı tecrübesiyle işletmenize özel Excel çözümleri: nakit akışı, cari, mizan, banka limitleri, raporlama ve yönetim sistemleri." />');

// Homepage: keep the hero composition, inject only the finance color system and premium accent layer.
if (!home.includes('id="home-finance-color-css"')) {
  if (!home.includes('</head>')) throw new Error('HOME COLOR GATE: </head> bulunamadı');
  home = home.replace('</head>', `${homeColorLink}\n</head>`);
}

for (const token of ['+90 542 123 45 67', 'tel:+905421234567', 'Temsili ekran — proje kapsamına göre özelleştirilir', 'Temsili ekran - proje kapsamına göre özelleştirilir']) {
  if (html.includes(token)) throw new Error(`SPECIAL BRAND GATE: kaldırılması gereken token build çıktısında kaldı: ${token}`);
}
for (const required of [
  'data-special-innovation="v4"',
  'id="special-innovation-css"',
  'id="special-brand-sync-css"',
  'id="karar-laboratuvari"',
  'id="iv-panel-nakit"',
  'id="iv-panel-cari"',
  'id="iv-panel-banka"',
  'id="iv-panel-yonetim"',
  'id="special-innovation-js"',
  '/images/brand/excelarsiv-header-logo.png',
  'Hazır Excel işletme sistemleri',
  'İşletmelere Özel Excel Çözümleri',
  'Sahadan Gelen Uzmanlık',
  'İşletmenizi anlayan biriyle, size özel Excel sistemi kurun.',
]) {
  if (!html.includes(required)) throw new Error(`SPECIAL BRAND GATE: zorunlu contract eksik: ${required}`);
}
if (!home.includes('id="home-finance-color-css"')) throw new Error('HOME COLOR GATE: ana sayfa renk katmanı eklenemedi');

fs.writeFileSync(specialFile, html);
fs.writeFileSync(homeFile, home);
console.log('SPECIAL/HOME DESIGN GATE PASS — brand sync + CTA expertise + hero cleanup + homepage finance color system verified.');
