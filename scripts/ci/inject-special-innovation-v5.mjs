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

function swap(oldText, newText, label) {
  if (html.includes(newText)) return;
  if (!html.includes(oldText)) throw new Error(`SPECIAL INNOVATION GATE: replacement anchor missing: ${label}`);
  html = html.replace(oldText, newText);
}

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

// Hero product mockup is self-explanatory; the floating disclaimer badge adds visual noise.
html = html.replace(/\s*<span class="product-badge">Temsili ekran\s*[—-]\s*proje kapsamına göre özelleştirilir<\/span>/i, '');

// SEO: own the "işletmelere özel Excel çözümleri" intent without keyword stuffing.
html = html.replace(
  /Özel Excel Finans (?:&amp;|&) Raporlama Sistemleri \| excelarsiv\.com/g,
  'İşletmelere Özel Excel Çözümleri | Finans ve Yönetim Sistemleri | ExcelArşiv',
);
swap(
  'İşleyişinize göre kurulan özel Excel finans, raporlama ve yönetim sistemleri. Nakit akışı, çek/senet, mizan, banka limitleri ve yönetim dashboard\'ları.',
  'Bankacılık ve KOBİ danışmanlığı tecrübesiyle işletmenize özel Excel çözümleri: nakit akışı, cari, mizan, banka limitleri, raporlama ve yönetim sistemleri.',
  'SEO description',
);

// Closing CTA first, then shorter CTA labels.
swap('Ücretsiz Keşif Görüşmesi Al →', 'İşletmeme Özel Çözümü Konuşalım →', 'closing CTA button');
swap('İşinizi kolaylaştıracak sistemi birlikte netleştirelim.', 'İşletmenizi anlayan biriyle, size özel Excel sistemi kurun.', 'closing CTA heading');
swap(
  '15 dakikalık ilk görüşmede mevcut iş akışınızı, en pahalı manuel kontrol noktasını ve hangi Excel mimarisinin en hızlı geri dönüşü sağlayacağını belirleyelim.',
  'Bu çalışma yalnız Excel veya yazılım işi değil. Bankacılık, kredi karar süreçleri, KOBİ danışmanlığı ve sahadaki işletme tecrübesiyle; mizanınızdan nakit akışınıza, cari yapınızdan banka limitlerinize kadar gerçek işleyişinizi önce anlıyor, sonra işletmenize özel Excel çözümünü kuruyoruz. Böylece ihtiyacınızı yazılımcıya tercüme etmek zorunda kalmazsınız; aynı dili konuşan biriyle doğrudan çözüm üretirsiniz.',
  'closing CTA paragraph',
);
swap('İlk görüşme • kapsam teşhisi • çözüm haritası', 'Saha bilgisi • finans tecrübesi • yazılım disiplini', 'closing CTA note');

// Navigation + hero CTA labels.
html = html.replaceAll('Ücretsiz Keşif Görüşmesi', 'İşletmenize Özel Çözümü Konuşalım');

// Field-expertise proof: banking + KOBİ advisory + software knowledge in one operating language.
swap('Neden ExcelArşiv?', 'Sahadan Gelen Uzmanlık', 'why eyebrow');
swap(
  'Finans mantığı ile ekran tasarımını aynı sistemde kuruyoruz.',
  'Yazılımı bilen ama işletmenin finansal dilini de sahadan tanıyan bir bakışla tasarlıyoruz.',
  'why heading',
);
swap(
  'Mizan, muavin, nakit, vade ve faiz mantığı ekran tasarımından önce kurulur.',
  'Bankacılık ve kredi karar mantığı; banka limiti, nakit ihtiyacı, faiz yükü ve finansman riskini yalnız hücre olarak değil, karar konusu olarak ele almamızı sağlar.',
  'why row banking',
);
swap(
  'Girdi, hesaplanan alan ve yönetici raporu birbirinden ayrılır.',
  'KOBİ danışmanlığı tecrübesi; patronun, muhasebenin ve finans ekibinin aynı veriye neden farklı sorular sorduğunu anlamamızı sağlar.',
  'why row advisory',
);
swap(
  'Hata ve istisna noktaları kullanıcıya görünür hale getirilir.',
  'Mizan, cari, tahsilat, vade ve nakit akışı gibi işletme gerçekleri yazılım tasarımından önce modellenir; kullanıcıya yabancı bir sistem kurulmaz.',
  'why row operations',
);
swap(
  'Proje bittiğinde kara kutu değil, işletilebilir bir çalışma düzeni kalır.',
  'Sonuç hazır bir şablon değil; işletmenizin çalışma biçimine, kontrol noktalarına ve yönetim ihtiyacına göre kurulmuş özel bir Excel çözümüdür.',
  'why row custom system',
);

for (const token of ['+90 542 123 45 67', 'tel:+905421234567']) {
  if (html.includes(token)) throw new Error(`SPECIAL INNOVATION GATE: doğrulanmamış iletişim tokenı geri geldi: ${token}`);
}
for (const token of ['Temsili ekran — proje kapsamına göre özelleştirilir', 'Temsili ekran - proje kapsamına göre özelleştirilir']) {
  if (html.includes(token)) throw new Error(`SPECIAL INNOVATION GATE: hero badge geri geldi: ${token}`);
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
  'İşletmelere Özel Excel Çözümleri',
  'İşletmenizi anlayan biriyle, size özel Excel sistemi kurun.',
  'Bankacılık, kredi karar süreçleri, KOBİ danışmanlığı',
  'Saha bilgisi • finans tecrübesi • yazılım disiplini',
  'Sahadan Gelen Uzmanlık',
]) {
  if (!html.includes(required)) throw new Error(`SPECIAL INNOVATION GATE: zorunlu contract eksik: ${required}`);
}

fs.writeFileSync(file, html);
console.log('SPECIAL INNOVATION GATE PASS — decision lab + field-expertise CTAs + SEO intent + hero badge removal verified.');
