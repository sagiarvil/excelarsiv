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

function replaceRequired(pattern, replacement, label) {
  if (!pattern.test(html)) throw new Error(`SPECIAL INNOVATION GATE: replacement anchor missing: ${label}`);
  html = html.replace(pattern, replacement);
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
html = html.replace(
  /\s*<span class="product-badge">Temsili ekran\s*[—-]\s*proje kapsamına göre özelleştirilir<\/span>/i,
  '',
);

// SEO: own the "işletmelere özel Excel çözümleri" intent without keyword stuffing.
replaceRequired(
  /<title>Özel Excel Finans &amp; Raporlama Sistemleri \| excelarsiv\.com<\/title>|<title>Özel Excel Finans & Raporlama Sistemleri \| excelarsiv\.com<\/title>/i,
  '<title>İşletmelere Özel Excel Çözümleri | Finans ve Yönetim Sistemleri | ExcelArşiv</title>',
  'SEO title',
);
replaceRequired(
  /<meta name="description" content="İşleyişinize göre kurulan özel Excel finans, raporlama ve yönetim sistemleri\. Nakit akışı, çek\/senet, mizan, banka limitleri ve yönetim dashboard(?:'|&#39;)ları\."\s*\/?>/i,
  '<meta name="description" content="Bankacılık ve KOBİ danışmanlığı tecrübesiyle işletmenize özel Excel çözümleri: nakit akışı, cari, mizan, banka limitleri, raporlama ve yönetim sistemleri." />',
  'SEO description',
);

// CTA language: software capability + banking/KOBİ field knowledge as the differentiator.
replaceRequired(
  /<a class="btn btn-primary nav-cta" href="\/iletisim">Ücretsiz Keşif Görüşmesi<\/a>/i,
  '<a class="btn btn-primary nav-cta" href="/iletisim">İşletmenize Özel Çözümü Konuşalım</a>',
  'navigation CTA',
);
replaceRequired(
  /<a class="btn btn-primary" href="\/iletisim">↗ Ücretsiz Keşif Görüşmesi<\/a>/i,
  '<a class="btn btn-primary" href="/iletisim">↗ İşletmenize Özel Çözümü Konuşalım</a>',
  'hero CTA',
);

replaceRequired(
  /<div class="why-card"><span class="eyebrow">Neden ExcelArşiv\?<\/span><h3 class="section-title" style="font-size:28px">Finans mantığı ile ekran tasarımını aynı sistemde kuruyoruz\.<\/h3><div class="why-list">[\s\S]*?<\/div><\/div>\s*<\/div>\s*<\/div>\s*<\/section>/i,
  `<div class="why-card"><span class="eyebrow">Sahadan Gelen Uzmanlık</span><h3 class="section-title" style="font-size:28px">Yazılımı bilen ama işletmenin finansal dilini de sahadan tanıyan bir bakışla tasarlıyoruz.</h3><div class="why-list"><div class="why-row"><i>✓</i><span>Bankacılık ve kredi karar mantığı; banka limiti, nakit ihtiyacı, faiz yükü ve finansman riskini yalnız hücre olarak değil, karar konusu olarak ele almamızı sağlar.</span></div><div class="why-row"><i>✓</i><span>KOBİ danışmanlığı tecrübesi; patronun, muhasebenin ve finans ekibinin aynı veriye neden farklı sorular sorduğunu anlamamızı sağlar.</span></div><div class="why-row"><i>✓</i><span>Mizan, cari, tahsilat, vade ve nakit akışı gibi işletme gerçekleri yazılım tasarımından önce modellenir; kullanıcıya yabancı bir sistem kurulmaz.</span></div><div class="why-row"><i>✓</i><span>Sonuç hazır bir şablon değil; işletmenizin çalışma biçimine, kontrol noktalarına ve yönetim ihtiyacına göre kurulmuş özel bir Excel çözümüdür.</span></div></div></div>
        </div>
      </div>
    </section>`,
  'field expertise card',
);

replaceRequired(
  /<section class="cta"><div class="wrap"><div class="cta-shell"><div class="cta-copy"><h2>İşinizi kolaylaştıracak sistemi birlikte netleştirelim\.<\/h2><p>15 dakikalık ilk görüşmede mevcut iş akışınızı, en pahalı manuel kontrol noktasını ve hangi Excel mimarisinin en hızlı geri dönüşü sağlayacağını belirleyelim\.<\/p><\/div><div class="cta-actions"><div><a class="btn btn-primary" href="\/iletisim">Ücretsiz Keşif Görüşmesi Al →<\/a><span class="cta-note">İlk görüşme • kapsam teşhisi • çözüm haritası<\/span><\/div><\/div><\/div><\/div><\/section>/i,
  `<section class="cta"><div class="wrap"><div class="cta-shell"><div class="cta-copy"><h2>İşletmenizi anlayan biriyle, size özel Excel sistemi kurun.</h2><p>Bu çalışma yalnız Excel veya yazılım işi değil. Bankacılık, kredi karar süreçleri, KOBİ danışmanlığı ve sahadaki işletme tecrübesiyle; mizanınızdan nakit akışınıza, cari yapınızdan banka limitlerinize kadar gerçek işleyişinizi önce anlıyor, sonra işletmenize özel Excel çözümünü kuruyoruz. Böylece ihtiyacınızı yazılımcıya tercüme etmek zorunda kalmazsınız; aynı dili konuşan biriyle doğrudan çözüm üretirsiniz.</p></div><div class="cta-actions"><div><a class="btn btn-primary" href="/iletisim">İşletmeme Özel Çözümü Konuşalım →</a><span class="cta-note">Saha bilgisi • finans tecrübesi • yazılım disiplini</span></div></div></div></div></section>`,
  'closing CTA',
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
]) {
  if (!html.includes(required)) throw new Error(`SPECIAL INNOVATION GATE: zorunlu contract eksik: ${required}`);
}

fs.writeFileSync(file, html);
console.log('SPECIAL INNOVATION GATE PASS — decision lab + field-expertise CTAs + SEO intent + hero badge removal verified.');
