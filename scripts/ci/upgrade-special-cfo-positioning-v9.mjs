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

const replaceOne = (pattern, replacement, label) => {
  const before = html;
  html = html.replace(pattern, replacement);
  if (html === before) throw new Error(`SPECIAL CFO POSITIONING V9: replacement missed: ${label}`);
};
const sectionPattern = (id) => new RegExp(`<section\\b[^>]*\\bid="${id}"[^>]*>[\\s\\S]*?<\\/section>`, 'u');
const replaceSection = (id, replacement) => replaceOne(sectionPattern(id), replacement, `section#${id}`);
const extractSection = (id) => {
  const match = html.match(sectionPattern(id));
  if (!match) throw new Error(`SPECIAL CFO POSITIONING V9: section#${id} missing`);
  return match[0];
};

replaceOne(/<title>[\s\S]*?<\/title>/u, '<title>Özel Excel Karar Sistemleri | Nakit Akışı, Banka Risk, CFO Modelleri | ExcelArşiv</title>', 'title');
replaceOne(/<meta name="description" content="[^"]*"\s*\/>/u, '<meta name="description" content="17 yıllık ticari bankacılık ve saha deneyimiyle 13 haftalık nakit akışı, banka limit-risk, dinamik fiyatlama, rolling forecast, DCF/DSCR ve ERP entegre CFO karar sistemleri." />', 'meta description');
replaceOne(/<meta property="og:title" content="[^"]*"\s*\/>/u, '<meta property="og:title" content="17 Yıllık Ticari Bankacılık Bakışıyla Özel Excel Karar Sistemleri | ExcelArşiv" />', 'og title');
replaceOne(/<meta property="og:description" content="[^"]*"\s*\/>/u, '<meta property="og:description" content="Yazılımcıya finans öğretmeden; nakit, banka riski, maliyet, bütçe, değerleme ve ERP verisini patron/CFO karar ekranına dönüştüren özel Excel sistemleri." />', 'og description');

replaceOne('data-hero-proof-layout-v8', 'data-hero-proof-layout-v8 data-cfo-positioning-v9', 'body V9 marker');
replaceOne(/<p class="eyebrow">[\s\S]*?<\/p>/u, '<p class="eyebrow">17 Yıllık Ticari Bankacılık &amp; Saha Deneyimi ile Geliştirildi</p>', 'hero authority badge');
replaceOne(/(<div class="hero-copy">[\s\S]*?)<h1>[\s\S]*?<\/h1>/u, '$1<h1>Yazılımcıya finans öğretmekle uğraşmayın. <em>Şirketinizin dilini bilen özel Excel karar sistemleri.</em></h1>', 'hero headline');
replaceOne(/<p class="hero-lead">[\s\S]*?<\/p>/u, '<p class="hero-lead">Nakit akışını, banka limitlerini, kredi yükünü, maliyeti, bütçeyi ve ERP verisini ayrı ayrı izlemek yerine aynı karar mantığında birleştiriyoruz. Amaç daha fazla tablo değil; patronun ve CFO’nun “ne oluyor, neden oluyor, şimdi ne yapmalıyız?” sorusuna tek ekranda cevap veren finans sistemi kurmak.</p>', 'hero lead');
replaceOne(/<div class="hero-actions">[\s\S]*?<\/div>/u, `<div class="hero-actions"><a class="btn btn-primary btn-whatsapp" href="${WA_URL}" target="_blank" rel="noopener noreferrer">WhatsApp'tan Yazın — 15 Dakikada İhtiyacı Çıkaralım</a><a class="btn btn-secondary" href="#alanlar">Finans Sistemlerini İnceleyin <span aria-hidden="true">↓</span></a></div>`, 'hero CTAs');
replaceOne(/<p class="hero-microcopy">[\s\S]*?<\/p>/u, '<p class="hero-microcopy">Ön görüşmede teknik şartname istemiyoruz. İşletmedeki darboğazı, mevcut dosyayı veya yönetim sorusunu anlatmanız yeterli.</p>', 'hero microcopy');

replaceSection('alanlar', `<section class="section section-green" id="alanlar">
      <div class="wrap-wide">
        <div class="section-head"><span class="section-kicker">CFO / PATRON KARAR SİSTEMLERİ</span><h2>Günlük takip tablosu değil; finansal karar üreten altı sistem sınıfı.</h2><p>Her başlık gerçek bir yönetim sorusuna karşılık gelir. Nakit açığı, banka limiti, fiyatlama, bütçe sapması, yatırım getirisi ve işletme sermayesi aynı profesyonel finans diliyle modellenir.</p></div>
        <div class="area-grid">
          <article class="area-card"><div class="area-icon">13W</div><h3>Dinamik Nakit Akışı <span class="area-en">13-Week Cash Flow</span></h3><p>Likidite projeksiyonu, haftalık kasa açığı erken uyarısı, tahsilat senaryosu ve dinamik ödeme planı.</p><p class="area-benefit"><strong>Yönetim kararı:</strong> Maaş, çek, vergi ve tedarikçi ödemelerinin hangi haftada nakit baskısı yaratacağını önceden görün.</p></article>
          <article class="area-card"><div class="area-icon">₺</div><h3>Banka Limit-Risk ve Teminat Havuzu</h3><p>Nakdi/gayrinakdi limitler, rotatif faiz yükü, spot vadeler, çek risk havuzu, teminatlar ve kullanılabilir banka alanı.</p><p class="area-benefit"><strong>Yönetim kararı:</strong> Hangi bankada alan kaldığını, hangi limitin pahalılaştığını ve hangi vadenin sıkıştığını tek matriste görün.</p></article>
          <article class="area-card"><div class="area-icon">%</div><h3>Birim Maliyet ve Dinamik Fiyatlama</h3><p>Değişken maliyet endeksli teklif motoru, katkı payı, hedef marj ve başabaş <span lang="en">(break-even)</span> simülasyonu.</p><p class="area-benefit"><strong>Yönetim kararı:</strong> Teklif vermeden önce gerçek marjı görün; kur, hammadde, navlun veya finansman şokunun satış fiyatına etkisini test edin.</p></article>
          <article class="area-card"><div class="area-icon">Σ</div><h3>Senaryolu Bütçe ve Rolling Forecast</h3><p>Kur/faiz şoklarına duyarlı iyimser, baz ve kötü senaryolu dinamik bütçe; fiili sapma ve dönem sonu projeksiyonu.</p><p class="area-benefit"><strong>Yönetim kararı:</strong> Yıl başında donan bütçe yerine her ay güncellenen gelir, gider, nakit ve finansman görünümü kullanın.</p></article>
          <article class="area-card"><div class="area-icon">DCF</div><h3>Yatırım Fizibilitesi ve DCF Değerleme</h3><p>İskontolanmış nakit akımları, IRR, geri ödeme süresi, EBITDA normalizasyonu ve DSCR borç servis kapasitesi.</p><p class="area-benefit"><strong>Yönetim kararı:</strong> Makine, şube, yatırım veya şirket alımında getiriyi, borç taşıma gücünü ve değer aralığını aynı modelde sınayın.</p></article>
          <article class="area-card"><div class="area-icon">KPI</div><h3>ERP Entegre CFO Yönetim Kokpiti</h3><p>Logo, SAP, Mikro, Zirve veya mevcut Excel/CSV verisinden net işletme sermayesi, DSO, nakit, kredi ve KPI görünümü.</p><p class="area-benefit"><strong>Yönetim kararı:</strong> Manuel rapor birleştirmeyi azaltıp kritik finans göstergelerini tek patron/CFO ekranında toplayın.</p></article>
        </div>
      </div>
    </section>`);

replaceSection('karsilastirma', `<section class="section section-soft authority-comparison" id="karsilastirma">
      <div class="wrap-wide">
        <div class="section-head"><span class="section-kicker">17 YILLIK FİNANS BAKIŞI</span><h2>Excel’i değil, şirketin finansal karar mantığını tasarlıyoruz.</h2><p>Fark yalnız teknik geliştirme değildir. Ticari bankacılıkta kredi, limit, vade, risk ve nakit akışına bakılan çerçeveyi işletmenin gerçek saha verisiyle aynı modelde buluşturuyoruz.</p></div>
        <div class="authority-rail" aria-label="ExcelArşiv uzmanlık özeti">
          <div><strong>17 Yıl</strong><span>Ticari bankacılık, kredi, limit ve risk bakışı</span></div>
          <div><strong>Banka ↔ İşletme</strong><span>Finans, muhasebe ve saha dilini aynı modelde birleştirme</span></div>
          <div><strong>ERP → CFO Kararı</strong><span>Dağınık veriyi rapordan aksiyona taşıyan yönetim sistemi</span></div>
        </div>
        <div class="comparison-wrap"><div class="comparison" role="table" aria-label="Çözüm yaklaşımları karşılaştırması">
          <div class="comparison-head" role="row"><div role="columnheader">Kriter</div><div role="columnheader">Klasik yazılım / danışmanlık</div><div role="columnheader">Hazır Excel</div><div role="columnheader">ExcelArşiv özel karar sistemi</div></div>
          <div class="comparison-row" role="row"><div role="rowheader">Başlangıç noktası</div><div role="cell" data-label="Klasik yazılım / danışmanlık">Teknik kapsam, modül ve entegrasyon gereksinimleriyle başlayabilir.</div><div role="cell" data-label="Hazır Excel">Önceden tanımlanmış dosya düzenine uyum bekler.</div><div role="cell" data-label="ExcelArşiv özel karar sistemi"><strong>Önce finansal darboğaz teşhis edilir:</strong> nakit, limit, maliyet, bütçe, yatırım veya yönetim raporu.</div></div>
          <div class="comparison-row" role="row"><div role="rowheader">Finansal muhakeme</div><div role="cell" data-label="Klasik yazılım / danışmanlık">Proje ekibinin finans ve sektör tecrübesine göre değişir.</div><div role="cell" data-label="Hazır Excel">Standart formül ve önceden yazılmış kullanım senaryosu sunar.</div><div role="cell" data-label="ExcelArşiv özel karar sistemi"><strong>17 yıllık ticari bankacılık ve saha deneyimi</strong> doğrudan modelin karar kurallarına girer.</div></div>
          <div class="comparison-row" role="row"><div role="rowheader">Çıktı</div><div role="cell" data-label="Klasik yazılım / danışmanlık">Geniş modül, kayıt ve raporlama ekosistemi sağlayabilir.</div><div role="cell" data-label="Hazır Excel">Önceden tanımlı raporları üretir.</div><div role="cell" data-label="ExcelArşiv özel karar sistemi">Patron/CFO için “ne oldu?” ile yetinmez; <strong>“neden oldu ve şimdi ne yapmalıyız?”</strong> sorusuna odaklanır.</div></div>
          <div class="comparison-row" role="row"><div role="rowheader">Kapsam ve maliyet</div><div role="cell" data-label="Klasik yazılım / danışmanlık">Lisans, entegrasyon ve proje kapsamı modele göre büyüyebilir.</div><div role="cell" data-label="Hazır Excel">Düşük giriş maliyeti; uyarlama sorumluluğu kullanıcıdadır.</div><div role="cell" data-label="ExcelArşiv özel karar sistemi">Yalnız kullanılan akış ve karar ekranları kurulur; gereksiz modül projeye eklenmez.</div></div>
          <div class="comparison-row" role="row"><div role="rowheader">Veri kaynağı</div><div role="cell" data-label="Klasik yazılım / danışmanlık">Entegrasyon mimarisine göre değişir.</div><div role="cell" data-label="Hazır Excel">Genellikle manuel giriş veya sabit şablon ister.</div><div role="cell" data-label="ExcelArşiv özel karar sistemi">Logo, SAP, Mikro, Zirve, banka ekstresi, Excel/CSV ve mevcut işletme kayıtları aynı karar hattında kullanılabilir.</div></div>
        </div></div>
      </div>
    </section>`);

const comparisonBlock = extractSection('karsilastirma');
html = html.replace(sectionPattern('karsilastirma'), '');
const heroStart = html.indexOf('<section class="hero">');
if (heroStart < 0) throw new Error('SPECIAL CFO POSITIONING V9: hero missing');
const heroEnd = html.indexOf('</section>', heroStart);
if (heroEnd < 0) throw new Error('SPECIAL CFO POSITIONING V9: hero end missing');
const insertAt = heroEnd + '</section>'.length;
html = `${html.slice(0, insertAt)}\n${comparisonBlock}${html.slice(insertAt)}`;

if (sectionPattern('neden').test(html)) html = html.replace(sectionPattern('neden'), '');
html = html.replaceAll('href="#neden">Neden Barış Bağırlar?</a>', 'href="#karsilastirma">17 Yıllık Finans Bakışı</a>');
html = html.replaceAll('href="#neden">Neden ExcelArşiv?</a>', 'href="#karsilastirma">17 Yıllık Finans Bakışı</a>');

// Sanitizer V7 same-origin'e çevirdiği iletişim CTA'larını, kullanıcı talebi gereği yeniden doğrudan WhatsApp'a bağla.
html = html.replaceAll('href="/iletisim"', `href="${WA_URL}" target="_blank" rel="noopener noreferrer"`);
html = html.replaceAll('>İhtiyacımı anlat</a>', ">WhatsApp'tan Yazın</a>");
html = html.replaceAll('>Sürecinizi Anlatın <span aria-hidden="true">→</span></a>', ">WhatsApp'tan Yazın — 15 Dakikada İhtiyacı Çıkaralım</a>");
html = html.replaceAll('>Dosyamı anlat</a>', ">WhatsApp'tan Yazın — İhtiyacı Netleştirelim</a>");
html = html.replaceAll('>İhtiyacımı Anlat <span aria-hidden="true">→</span></a>', ">WhatsApp'tan Yazın — 15 Dakikada İhtiyacı Çıkaralım</a>");

const css = `
  <style id="special-cfo-positioning-v9-css">
    body[data-cfo-positioning-v9]{--ink:#0F172A;--ink-2:#1E293B;--green:#059669;--green-dark:#047857;--green-soft:#ECFDF5;color:#0F172A;font-size:17px}
    body[data-cfo-positioning-v9] .hero{background:linear-gradient(105deg,#fff 0%,#fff 55%,#ecfdf5 100%)}
    body[data-cfo-positioning-v9] .hero .eyebrow{padding:9px 14px;border:1px solid #a7f3d0;background:#ecfdf5;color:#065f46;font-size:15px;font-weight:780;letter-spacing:-.01em}
    body[data-cfo-positioning-v9] .hero .eyebrow::before{background:#059669}
    body[data-cfo-positioning-v9] .hero h1{max-width:760px;color:#0F172A;font-size:clamp(42px,4.8vw,62px);line-height:1.02;font-weight:780;letter-spacing:-.052em}
    body[data-cfo-positioning-v9] .hero h1 em{color:#059669}
    body[data-cfo-positioning-v9] .hero-lead{max-width:760px;color:#334155;font-size:19px;line-height:1.62}
    body[data-cfo-positioning-v9] .hero-microcopy{font-size:15px;color:#475569}
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
    body[data-cfo-positioning-v9] .comparison-head>div{font-size:16px;font-weight:760}
    body[data-cfo-positioning-v9] .comparison-row>div{font-size:16px;line-height:1.55;color:#334155}
    body[data-cfo-positioning-v9] .comparison-row>div:last-child{background:#ecfdf5;color:#0F172A}
    body[data-cfo-positioning-v9] .area-grid{gap:18px}
    body[data-cfo-positioning-v9] .area-card{border:1px solid #cbd5e1;border-top:4px solid #059669;background:#fff;box-shadow:0 10px 26px rgba(15,23,42,.06)}
    body[data-cfo-positioning-v9] .area-card h3{color:#0F172A;font-size:20px;line-height:1.28;font-weight:760}
    body[data-cfo-positioning-v9] .area-card p{color:#475569;font-size:16px;line-height:1.58}
    body[data-cfo-positioning-v9] .area-icon{background:#0F172A;color:#34d399;border-color:#0F172A;font-weight:800}
    body[data-cfo-positioning-v9] .area-en{color:#64748b;font-size:15px}
    body[data-cfo-positioning-v9] .area-benefit{font-size:16px!important}
    body[data-cfo-positioning-v9] .proof-chip strong{color:#0F172A;font-size:16px}
    body[data-cfo-positioning-v9] .proof-chip span{font-size:15px}
    @media(max-width:900px){body[data-cfo-positioning-v9] .authority-rail{grid-template-columns:1fr}body[data-cfo-positioning-v9] .hero h1{font-size:44px}}
    @media(max-width:660px){body[data-cfo-positioning-v9] .hero h1{font-size:38px}body[data-cfo-positioning-v9] .hero-lead{font-size:17px}body[data-cfo-positioning-v9] .section-head p{font-size:17px}body[data-cfo-positioning-v9] .authority-rail>div{padding:20px}body[data-cfo-positioning-v9] .authority-rail strong{font-size:22px}body[data-cfo-positioning-v9] .btn{width:100%;white-space:normal;text-align:center}}
  </style>`;
replaceOne('</head>', `${css}\n</head>`, 'V9 CSS');

const required = [
  'data-cfo-positioning-v9',
  '17 Yıllık Ticari Bankacılık &amp; Saha Deneyimi ile Geliştirildi',
  'Şirketinizin dilini bilen özel Excel karar sistemleri.',
  "WhatsApp'tan Yazın — 15 Dakikada İhtiyacı Çıkaralım",
  'Dinamik Nakit Akışı',
  'Banka Limit-Risk ve Teminat Havuzu',
  'Birim Maliyet ve Dinamik Fiyatlama',
  'Senaryolu Bütçe ve Rolling Forecast',
  'Yatırım Fizibilitesi ve DCF Değerleme',
  'ERP Entegre CFO Yönetim Kokpiti',
  'authority-rail',
  'Excel’i değil, şirketin finansal karar mantığını tasarlıyoruz.',
  'special-cfo-positioning-v9-css',
  '--green:#059669',
  '--ink:#0F172A',
  'https://wa.me/905419305372',
];
for (const token of required) {
  if (!html.includes(token)) throw new Error(`SPECIAL CFO POSITIONING V9: required token missing: ${token}`);
}

const comparisonIndex = html.indexOf('id="karsilastirma"');
const neZamanIndex = html.indexOf('id="ne-zaman"');
if (!(comparisonIndex > 0 && neZamanIndex > comparisonIndex)) {
  throw new Error('SPECIAL CFO POSITIONING V9: authority comparison is not promoted above operational content');
}
if (html.includes('id="neden"')) throw new Error('SPECIAL CFO POSITIONING V9: duplicate lower authority section still present');
if (html.includes('overflow-x:hidden')) throw new Error('SPECIAL CFO POSITIONING V9: forbidden overflow hiding introduced');
const directWhatsAppCount = (html.match(/https:\/\/wa\.me\/905419305372/gu) || []).length;
if (directWhatsAppCount < 3) throw new Error(`SPECIAL CFO POSITIONING V9: expected >=3 direct WhatsApp paths, got ${directWhatsAppCount}`);

fs.writeFileSync(distFile, html, 'utf8');
console.log(`SPECIAL CFO POSITIONING V9 PASS — banking authority promoted above operations, six CFO finance systems normalized, direct WhatsApp restored (${directWhatsAppCount}), and navy/emerald hierarchy enforced.`);
