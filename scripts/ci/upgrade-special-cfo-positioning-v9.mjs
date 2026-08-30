#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const distFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(distFile)) throw new Error('SPECIAL CFO POSITIONING V9: dist route missing');
let html = fs.readFileSync(distFile, 'utf8');

if (!html.includes('data-finance-funnel-v7') || !html.includes('data-hero-proof-layout-v8')) {
  throw new Error('SPECIAL CFO POSITIONING V9: V7/V8 prerequisite contract missing');
}
if (html.includes('data-cfo-positioning-v9')) {
  console.log('SPECIAL CFO POSITIONING V9: already applied');
  process.exit(0);
}

const WA_URL = 'https://wa.me/905419305372?text=Merhaba%2C%20ExcelAr%C5%9Fiv%20%C3%B6zel%20Excel%20sistemi%20i%C3%A7in%20i%C5%9F%20ak%C4%B1%C5%9F%C4%B1m%C4%B1%20anlatmak%20istiyorum.';
const sectionPattern = (id) => new RegExp(`<section\\b[^>]*\\bid="${id}"[^>]*>[\\s\\S]*?<\\/section>`, 'u');

const replaceOne = (pattern, replacement, label) => {
  const before = html;
  html = html.replace(pattern, replacement);
  if (html === before) throw new Error(`SPECIAL CFO POSITIONING V9: replacement missed: ${label}`);
};

replaceOne(/<title>[\s\S]*?<\/title>/u, '<title>Özel Excel Karar Sistemleri | Nakit Akışı, Banka Risk, CFO Modelleri | ExcelArşiv</title>', 'title');
replaceOne(/<meta name="description" content="[^"]*"\s*\/>/u, '<meta name="description" content="17 yıllık ticari bankacılık ve saha deneyimiyle 13 haftalık nakit akışı, banka limit-risk, dinamik fiyatlama, rolling forecast, DCF/DSCR ve ERP entegre CFO karar sistemleri." />', 'meta description');
replaceOne(/<meta property="og:title" content="[^"]*"\s*\/>/u, '<meta property="og:title" content="17 Yıllık Ticari Bankacılık Bakışıyla Özel Excel Karar Sistemleri | ExcelArşiv" />', 'og title');
replaceOne(/<meta property="og:description" content="[^"]*"\s*\/>/u, '<meta property="og:description" content="Yazılımcıya finans öğretmeden; nakit, banka riski, maliyet, bütçe, değerleme ve ERP verisini patron/CFO karar ekranına dönüştüren özel Excel sistemleri." />', 'og description');
replaceOne('data-hero-proof-layout-v8', 'data-hero-proof-layout-v8 data-cfo-positioning-v9', 'body V9 marker');

// Hero'yu kendi section sınırı içinde değiştir; önceki post-process sınıf/boşluk değişikliklerinden etkilenmez.
const heroStart = html.indexOf('<section class="hero">');
if (heroStart < 0) throw new Error('SPECIAL CFO POSITIONING V9: hero missing');
const heroClose = html.indexOf('</section>', heroStart);
if (heroClose < 0) throw new Error('SPECIAL CFO POSITIONING V9: hero end missing');
const heroEnd = heroClose + '</section>'.length;
let hero = html.slice(heroStart, heroEnd);
const heroBefore = hero;
hero = hero.replace(/<p class="eyebrow">[\s\S]*?<\/p>/u, '<p class="eyebrow">17 Yıllık Ticari Bankacılık &amp; Saha Deneyimi ile Geliştirildi</p>');
hero = hero.replace(/<h1\b[^>]*>[\s\S]*?<\/h1>/u, '<h1>Yazılımcıya finans öğretmekle uğraşmayın. <em>Şirketinizin dilini bilen özel Excel karar sistemleri.</em></h1>');
hero = hero.replace(/<p class="hero-lead">[\s\S]*?<\/p>/u, '<p class="hero-lead">Nakit akışını, banka limitlerini, kredi yükünü, maliyeti, bütçeyi ve ERP verisini ayrı ayrı izlemek yerine aynı karar mantığında birleştiriyoruz. Amaç daha fazla tablo değil; patronun ve CFO’nun “ne oluyor, neden oluyor, şimdi ne yapmalıyız?” sorusuna tek ekranda cevap veren finans sistemi kurmak.</p>');
hero = hero.replace(/<div class="hero-actions">[\s\S]*?<\/div>/u, `<div class="hero-actions"><a class="btn btn-primary btn-whatsapp" href="${WA_URL}" target="_blank" rel="noopener noreferrer">WhatsApp'tan Yazın — 15 Dakikada İhtiyacı Çıkaralım</a><a class="btn btn-secondary" href="#alanlar">Finans Sistemlerini İnceleyin <span aria-hidden="true">↓</span></a></div>`);
hero = hero.replace(/<p class="hero-microcopy">[\s\S]*?<\/p>/u, '<p class="hero-microcopy">Teknik şartname hazırlamayın. İşletmedeki darboğazı, mevcut dosyayı veya yönetim sorusunu anlatmanız yeterli.</p>');
if (hero === heroBefore || !hero.includes('Şirketinizin dilini bilen özel Excel karar sistemleri.')) {
  throw new Error('SPECIAL CFO POSITIONING V9: hero transformation failed');
}
html = `${html.slice(0, heroStart)}${hero}${html.slice(heroEnd)}`;

// V7'de zaten güçlü olan altı sistemi, kullanıcının talep ettiği CFO arama diline sabitle.
const cardRenames = [
  ['13 Haftalık Dinamik Nakit Akışı Modeli', 'Dinamik Nakit Akışı'],
  ['Çoklu Banka Limit-Risk, Kredi ve Teminat Matrisi', 'Banka Limit-Risk ve Teminat Havuzu'],
  ['Birim Maliyet, Dinamik Fiyatlama ve Başabaş Simülatörü', 'Birim Maliyet ve Dinamik Fiyatlama'],
  ['Senaryolu Bütçe ve Dinamik Projeksiyon', 'Senaryolu Bütçe ve Rolling Forecast'],
  ['Yatırım Fizibilitesi ve DCF Şirket Değerleme Modeli', 'Yatırım Fizibilitesi ve DCF Değerleme'],
  ['ERP Entegre CFO Yönetim Kokpiti', 'ERP Entegre CFO Yönetim Kokpiti'],
];
for (const [from, to] of cardRenames) {
  if (from !== to && html.includes(from)) html = html.replaceAll(from, to);
}

const comparison = `<section class="section section-soft authority-comparison" id="karsilastirma">
      <div class="wrap-wide">
        <div class="section-head"><span class="section-kicker">17 YILLIK FİNANS BAKIŞI</span><h2>Excel’i değil, şirketin finansal karar mantığını tasarlıyoruz.</h2><p>Ticari bankacılıkta kredi, limit, vade, risk ve nakit akışına bakılan çerçeveyi işletmenin gerçek saha verisiyle aynı modelde buluşturuyoruz. Fark yalnız teknik geliştirme değil; hangi verinin hangi yönetim kararına hizmet edeceğini bilmek.</p></div>
        <div class="authority-rail" aria-label="ExcelArşiv uzmanlık özeti">
          <div><strong>17 Yıl</strong><span>Ticari bankacılık, kredi, limit ve risk bakışı</span></div>
          <div><strong>Banka ↔ İşletme</strong><span>Finans, muhasebe ve saha dilini aynı modelde birleştirme</span></div>
          <div><strong>ERP → CFO Kararı</strong><span>Dağınık veriyi rapordan aksiyona taşıyan yönetim sistemi</span></div>
        </div>
        <div class="comparison-wrap"><div class="comparison" role="table" aria-label="Çözüm yaklaşımları karşılaştırması">
          <div class="comparison-head" role="row"><div role="columnheader">Kriter</div><div role="columnheader">Klasik yazılım / danışmanlık</div><div role="columnheader">Hazır Excel</div><div role="columnheader">ExcelArşiv özel karar sistemi</div></div>
          <div class="comparison-row" role="row"><div role="rowheader">Başlangıç noktası</div><div role="cell" data-label="Klasik yazılım / danışmanlık">Teknik kapsam, modül ve entegrasyon gereksinimleriyle başlayabilir.</div><div role="cell" data-label="Hazır Excel">Önceden tanımlanmış dosya düzenine uyum bekler.</div><div role="cell" data-label="ExcelArşiv özel karar sistemi"><strong>Önce finansal darboğaz teşhis edilir:</strong> nakit, limit, maliyet, bütçe, yatırım veya yönetim raporu.</div></div>
          <div class="comparison-row" role="row"><div role="rowheader">Finansal muhakeme</div><div role="cell" data-label="Klasik yazılım / danışmanlık">Proje ekibinin finans ve sektör tecrübesine göre değişir.</div><div role="cell" data-label="Hazır Excel">Standart formül ve önceden yazılmış kullanım senaryosu sunar.</div><div role="cell" data-label="ExcelArşiv özel karar sistemi"><strong>17 yıllık ticari bankacılık ve saha deneyimi</strong> doğrudan modelin karar kurallarına girer.</div></div>
          <div class="comparison-row" role="row"><div role="rowheader">Karar çıktısı</div><div role="cell" data-label="Klasik yazılım / danışmanlık">Geniş modül, kayıt ve raporlama ekosistemi sağlayabilir.</div><div role="cell" data-label="Hazır Excel">Önceden tanımlı raporları üretir.</div><div role="cell" data-label="ExcelArşiv özel karar sistemi">Patron/CFO için “ne oldu?” ile yetinmez; <strong>“neden oldu ve şimdi ne yapmalıyız?”</strong> sorusuna odaklanır.</div></div>
          <div class="comparison-row" role="row"><div role="rowheader">Kapsam ve maliyet</div><div role="cell" data-label="Klasik yazılım / danışmanlık">Lisans, entegrasyon ve proje kapsamı modele göre büyüyebilir.</div><div role="cell" data-label="Hazır Excel">Düşük giriş maliyeti; uyarlama sorumluluğu kullanıcıdadır.</div><div role="cell" data-label="ExcelArşiv özel karar sistemi">Yalnız kullanılan akış ve karar ekranları kurulur; gereksiz modül projeye eklenmez.</div></div>
          <div class="comparison-row" role="row"><div role="rowheader">Veri kaynağı</div><div role="cell" data-label="Klasik yazılım / danışmanlık">Entegrasyon mimarisine göre değişir.</div><div role="cell" data-label="Hazır Excel">Genellikle manuel giriş veya sabit şablon ister.</div><div role="cell" data-label="ExcelArşiv özel karar sistemi">Logo, SAP, Mikro, Zirve, banka ekstresi, Excel/CSV ve mevcut işletme kayıtları aynı karar hattında kullanılabilir.</div></div>
        </div></div>
      </div>
    </section>`;

if (!sectionPattern('karsilastirma').test(html)) throw new Error('SPECIAL CFO POSITIONING V9: comparison section missing');
html = html.replace(sectionPattern('karsilastirma'), comparison);
if (sectionPattern('neden').test(html)) html = html.replace(sectionPattern('neden'), '');

// Karşılaştırma + 17 yıl otoritesi hero'nun hemen arkasında görünür.
const comparisonMatch = html.match(sectionPattern('karsilastirma'));
if (!comparisonMatch) throw new Error('SPECIAL CFO POSITIONING V9: new comparison missing');
const comparisonBlock = comparisonMatch[0];
html = html.replace(sectionPattern('karsilastirma'), '');
const finalHeroStart = html.indexOf('<section class="hero">');
const finalHeroClose = html.indexOf('</section>', finalHeroStart);
const authorityInsert = finalHeroClose + '</section>'.length;
html = `${html.slice(0, authorityInsert)}\n${comparisonBlock}${html.slice(authorityInsert)}`;

html = html.replaceAll('href="#neden">Neden Barış Bağırlar?</a>', 'href="#karsilastirma">17 Yıllık Finans Bakışı</a>');
html = html.replaceAll('href="#neden">Neden ExcelArşiv?</a>', 'href="#karsilastirma">17 Yıllık Finans Bakışı</a>');

// V7 sanitizer'ının /iletisim'e çevirdiği satış CTA'larını kullanıcı talebiyle doğrudan WhatsApp'a geri bağla.
html = html.replaceAll('href="/iletisim"', `href="${WA_URL}" target="_blank" rel="noopener noreferrer"`);
html = html.replaceAll('>İhtiyacımı anlat</a>', ">WhatsApp'tan Yazın</a>");
html = html.replaceAll('>Sürecinizi Anlatın <span aria-hidden="true">→</span></a>', ">WhatsApp'tan Yazın — 15 Dakikada İhtiyacı Çıkaralım</a>");
html = html.replaceAll('>Dosyamı anlat</a>', ">WhatsApp'tan Yazın — İhtiyacı Netleştirelim</a>");
html = html.replaceAll('>İhtiyacımı Anlat <span aria-hidden="true">→</span></a>', ">WhatsApp'tan Yazın — 15 Dakikada İhtiyacı Çıkaralım</a>");

const css = `<style id="special-cfo-positioning-v9-css">
body[data-cfo-positioning-v9]{--ink:#0F172A;--ink-2:#1E293B;--green:#059669;--green-dark:#047857;--green-soft:#ECFDF5;color:#0F172A;font-size:17px}
body[data-cfo-positioning-v9] .hero{background:linear-gradient(105deg,#fff 0%,#fff 55%,#ecfdf5 100%)}
body[data-cfo-positioning-v9] .hero .eyebrow{padding:9px 14px;border:1px solid #a7f3d0;background:#ecfdf5;color:#065f46;font-size:15px;font-weight:780;letter-spacing:-.01em}
body[data-cfo-positioning-v9] .hero .eyebrow::before{background:#059669}
body[data-cfo-positioning-v9] .hero h1{max-width:760px;color:#0F172A;font-size:clamp(42px,4.8vw,62px);line-height:1.02;font-weight:780;letter-spacing:-.052em}
body[data-cfo-positioning-v9] .hero h1 em{color:#059669}
body[data-cfo-positioning-v9] .hero-lead{max-width:760px;color:#334155;font-size:19px;line-height:1.62}
body[data-cfo-positioning-v9] .btn{min-height:52px;padding-inline:22px;font-size:17px;font-weight:720}
body[data-cfo-positioning-v9] .btn-whatsapp,body[data-cfo-positioning-v9] .btn-primary{background:#059669;color:#fff;border-color:#059669;box-shadow:0 12px 28px rgba(5,150,105,.22)}
body[data-cfo-positioning-v9] .btn-whatsapp:hover,body[data-cfo-positioning-v9] .btn-primary:hover{background:#047857;border-color:#047857}
body[data-cfo-positioning-v9] .section-head h2{color:#0F172A;font-size:clamp(34px,3.8vw,46px);line-height:1.08;font-weight:760;letter-spacing:-.04em}
body[data-cfo-positioning-v9] .section-head p{max-width:860px;color:#475569;font-size:18px;line-height:1.62}
body[data-cfo-positioning-v9] .section-kicker{color:#047857;font-size:14px;font-weight:800;letter-spacing:.055em}
body[data-cfo-positioning-v9] .authority-comparison{padding-top:62px;background:#f8fafc}
body[data-cfo-positioning-v9] .authority-rail{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:1px;margin:0 0 22px;overflow:hidden;border:1px solid #0F172A;border-radius:18px;background:#334155;box-shadow:0 14px 34px rgba(15,23,42,.12)}
body[data-cfo-positioning-v9] .authority-rail>div{min-width:0;padding:24px 25px;background:#0F172A;color:#fff}
body[data-cfo-positioning-v9] .authority-rail strong{display:block;color:#34d399;font-size:24px;line-height:1.12;font-weight:800;letter-spacing:-.035em}
body[data-cfo-positioning-v9] .authority-rail span{display:block;margin-top:7px;color:#e2e8f0;font-size:16px;line-height:1.45}
body[data-cfo-positioning-v9] .comparison{border-color:#cbd5e1;box-shadow:0 12px 32px rgba(15,23,42,.07)}
body[data-cfo-positioning-v9] .comparison-head{background:#0F172A;color:#fff}
body[data-cfo-positioning-v9] .comparison-head>div,body[data-cfo-positioning-v9] .comparison-row>div{font-size:16px;line-height:1.55}
body[data-cfo-positioning-v9] .comparison-row>div:last-child{background:#ecfdf5;color:#0F172A}
body[data-cfo-positioning-v9] .area-card{border-top:4px solid #059669}
body[data-cfo-positioning-v9] .area-card h3{color:#0F172A;font-size:20px;line-height:1.28;font-weight:760}
body[data-cfo-positioning-v9] .area-card p{color:#475569;font-size:16px;line-height:1.58}
body[data-cfo-positioning-v9] .area-icon{background:#0F172A;color:#34d399;border-color:#0F172A;font-weight:800}
@media(max-width:900px){body[data-cfo-positioning-v9] .authority-rail{grid-template-columns:1fr}body[data-cfo-positioning-v9] .hero h1{font-size:44px}}
@media(max-width:660px){body[data-cfo-positioning-v9] .hero h1{font-size:38px}body[data-cfo-positioning-v9] .hero-lead,body[data-cfo-positioning-v9] .section-head p{font-size:17px}body[data-cfo-positioning-v9] .authority-rail>div{padding:20px}body[data-cfo-positioning-v9] .btn{width:100%;white-space:normal;text-align:center}}
</style>`;
replaceOne('</head>', `${css}\n</head>`, 'V9 CSS');

const required = [
  'data-cfo-positioning-v9','17 Yıllık Ticari Bankacılık &amp; Saha Deneyimi ile Geliştirildi','Şirketinizin dilini bilen özel Excel karar sistemleri.',
  "WhatsApp'tan Yazın — 15 Dakikada İhtiyacı Çıkaralım",'Dinamik Nakit Akışı','Banka Limit-Risk ve Teminat Havuzu','Birim Maliyet ve Dinamik Fiyatlama',
  'Senaryolu Bütçe ve Rolling Forecast','Yatırım Fizibilitesi ve DCF Değerleme','ERP Entegre CFO Yönetim Kokpiti','authority-rail',
  'Excel’i değil, şirketin finansal karar mantığını tasarlıyoruz.','--green:#059669','--ink:#0F172A','https://wa.me/905419305372'
];
for (const token of required) if (!html.includes(token)) throw new Error(`SPECIAL CFO POSITIONING V9: required token missing: ${token}`);

const comparisonIndex = html.indexOf('id="karsilastirma"');
const neZamanIndex = html.indexOf('id="ne-zaman"');
if (!(comparisonIndex > 0 && neZamanIndex > comparisonIndex)) throw new Error('SPECIAL CFO POSITIONING V9: authority not promoted above operations');
if (html.includes('id="neden"')) throw new Error('SPECIAL CFO POSITIONING V9: duplicate lower authority survived');
if (html.includes('overflow-x:hidden')) throw new Error('SPECIAL CFO POSITIONING V9: forbidden overflow hiding introduced');
const waCount = (html.match(/https:\/\/wa\.me\/905419305372/gu) || []).length;
if (waCount < 3) throw new Error(`SPECIAL CFO POSITIONING V9: expected >=3 direct WhatsApp CTAs, got ${waCount}`);

fs.writeFileSync(distFile, html, 'utf8');
console.log(`SPECIAL CFO POSITIONING V9 PASS — authority promoted, CFO language fixed, direct WhatsApp restored (${waCount}), navy/emerald hierarchy enforced.`);
