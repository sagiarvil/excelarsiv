#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const ROOT = process.cwd();
const homeFile = path.resolve('dist/index.html');
const specialFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
const OLD_PHONE = '905419305372';
const WA_PHONE = '905393333303'; // 0539 333 33 03 -> international wa.me format

if (!fs.existsSync(homeFile)) throw new Error('DUAL FUNNEL V16: home dist route missing');
if (!fs.existsSync(specialFile)) throw new Error('DUAL FUNNEL V16: special dist route missing');

const encodeWa = (message) => `https://wa.me/${WA_PHONE}?text=${encodeURIComponent(message)}`;
const wa = {
  hero: encodeWa('Merhaba Barış Bey, excelarsiv.com üzerinden ulaşıyorum. Şirketimiz için özel Excel finans karar sistemi hakkında görüşmek istiyorum.'),
  mid: encodeWa('Merhaba Barış Bey, şirketimizin mevcut tablosu ve finansal darboğazı hakkında 15 dakikalık bir ön teşhis görüşmesi yapmak istiyorum.'),
  bottom: encodeWa('Merhaba Barış Bey, sürecimizi anlatmak ve özel Excel karar motoru mimarimizi netleştirmek istiyorum.'),
  floating: encodeWa('Merhaba Barış Bey, mobilden ulaşıyorum. Özel sistem ihtiyacımızı 15 dakikada netleştirelim.'),
  home: encodeWa('Merhaba Barış Bey, excelarsiv.com ana sayfasından ulaşıyorum. Şirketimizin finansal tıkanıklığını ve uygun Excel karar sistemini netleştirmek istiyorum.'),
};

const replaceOne = (html, pattern, replacement, label) => {
  const next = html.replace(pattern, replacement);
  if (next === html) throw new Error(`DUAL FUNNEL V16: replacement missed: ${label}`);
  return next;
};

const upsertMeta = (html, attr, key, content) => {
  const escaped = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const pattern = new RegExp(`<meta\\s+${attr}="${escaped}"\\s+content="[^"]*"\\s*\\/?\\s*>`, 'iu');
  const tag = `<meta ${attr}="${key}" content="${content}" />`;
  if (pattern.test(html)) return html.replace(pattern, tag);
  return html.replace('</head>', `  ${tag}\n</head>`);
};

const addOrReplaceJsonLdGraph = (html, mutator, label) => {
  const pattern = /<script type="application\/ld\+json">([\s\S]*?)<\/script>/u;
  const match = html.match(pattern);
  if (!match) throw new Error(`DUAL FUNNEL V16: JSON-LD missing on ${label}`);
  let data;
  try { data = JSON.parse(match[1]); } catch (error) { throw new Error(`DUAL FUNNEL V16: JSON-LD parse failed on ${label}: ${error.message}`); }
  if (!Array.isArray(data?.['@graph'])) throw new Error(`DUAL FUNNEL V16: JSON-LD @graph missing on ${label}`);
  mutator(data['@graph']);
  return html.replace(pattern, `<script type="application/ld+json">${JSON.stringify(data)}</script>`);
};

const ensureGraphNode = (graph, node, predicate) => {
  const index = graph.findIndex(predicate);
  if (index >= 0) graph[index] = node;
  else graph.push(node);
};

// -------------------- HOME --------------------
let home = fs.readFileSync(homeFile, 'utf8');

home = replaceOne(home, /<title>[\s\S]*?<\/title>/u, '<title>Finansal Karar ve Excel Sistemleri | Excel Arşiv</title>', 'home title');
home = upsertMeta(home, 'name', 'description', 'İşletmeler için 13 haftalık nakit akışı, banka limit-risk, birim maliyet, bütçe, değerleme ve özel Excel karar sistemleri. Hazır modeli seçin veya işletmenize özel sistem kurun.');
home = upsertMeta(home, 'property', 'og:type', 'website');
home = upsertMeta(home, 'property', 'og:url', 'https://excelarsiv.com/');
home = upsertMeta(home, 'property', 'og:title', 'İşletmeler İçin Finansal Karar ve Excel Sistemleri | Excel Arşiv');
home = upsertMeta(home, 'property', 'og:description', '17 yıllık ticari bankacılık ve saha finans bakışıyla nakit, kredi, maliyet, kârlılık ve bütçe kararları için Excel sistemleri.');
home = upsertMeta(home, 'property', 'og:image', 'https://excelarsiv.com/images/hero.jpg');
home = upsertMeta(home, 'name', 'twitter:card', 'summary_large_image');
home = upsertMeta(home, 'name', 'twitter:title', 'İşletmeler İçin Finansal Karar ve Excel Sistemleri | Excel Arşiv');
home = upsertMeta(home, 'name', 'twitter:description', 'Hazır finans modellerini inceleyin veya işletmenize özel Excel karar sistemi kurun.');
home = upsertMeta(home, 'name', 'twitter:image', 'https://excelarsiv.com/images/hero.jpg');

const financePillars = `
    <section class="finance-pillars" id="finans-sistemleri" data-experience-stage>
      <div class="home-shell">
        <div class="finance-pillars__head">
          <div><p class="eyebrow">FİNANSAL KARAR SİSTEMLERİ</p><h2>Dört ana finansal karar alanından başlayın.</h2></div>
          <p>Hazır sistem arayan işletmeler için ihtiyaçları finansal karar başlıklarına ayırdık. Daha özel bir veri akışınız varsa doğrudan özel sistem hattına geçebilirsiniz.</p>
        </div>
        <div class="finance-pillars__grid">
          <article class="finance-pillar finance-pillar--green">
            <span class="finance-pillar__no">01</span><h3>Nakit &amp; Likidite Sistemleri</h3>
            <ul><li>13 Haftalık Dinamik Nakit Akışı Modeli</li><li>Günlük Kasa &amp; Banka Likidite Takip Paneli</li><li>Çek &amp; Senet Vade Yaşlandırma Tablosu</li></ul>
            <a href="/sablonlar?q=nakit" data-cta="home_finance_cash" data-location="finance_pillars">Nakit sistemlerini gör →</a>
          </article>
          <article class="finance-pillar finance-pillar--blue">
            <span class="finance-pillar__no">02</span><h3>Banka, Kredi &amp; Teminat Sistemleri</h3>
            <ul><li>Çoklu Banka Limit-Risk ve Kredi Portföyü</li><li>Rotatif Kredi Faiz ve Finansman Maliyet Simülatörü</li><li>Teminat Mektubu &amp; İpotek Karşılama Matrisi</li></ul>
            <a href="/sablonlar?q=kredi" data-cta="home_finance_credit" data-location="finance_pillars">Banka sistemlerini gör →</a>
          </article>
          <article class="finance-pillar finance-pillar--violet">
            <span class="finance-pillar__no">03</span><h3>Birim Maliyet &amp; Dinamik Fiyatlama</h3>
            <ul><li>Değişken Maliyet ve Hammadde Endeksli Fiyat Teklif Motoru</li><li>Ürün &amp; Müşteri Bazlı Katkı Payı ve Kârlılık Matrisi</li><li>Şirket &amp; Proje Başabaş (Break-Even) Hesaplayıcı</li></ul>
            <a href="/sablonlar?q=maliyet" data-cta="home_finance_cost" data-location="finance_pillars">Maliyet sistemlerini gör →</a>
          </article>
          <article class="finance-pillar finance-pillar--amber">
            <span class="finance-pillar__no">04</span><h3>Bütçe, Projeksiyon &amp; Değerleme</h3>
            <ul><li>3 Senaryolu Dinamik Bütçe &amp; Rolling Forecast Modeli</li><li>DCF İskontolanmış Nakit Akımları &amp; Yatırım Fizibilitesi</li><li>Net İşletme Sermayesi (NÖS) &amp; DSO Gösterge Kokpiti</li></ul>
            <a href="/sablonlar?q=butce" data-cta="home_finance_budget" data-location="finance_pillars">Bütçe sistemlerini gör →</a>
          </article>
        </div>
      </div>
    </section>`;

if (!home.includes('class="finance-pillars"')) {
  home = replaceOne(home, '<section class="catalog-proof"', `${financePillars}\n\n    <section class="catalog-proof"`, 'home finance pillars insertion');
}

const highTicketBridge = `
    <section class="high-ticket-bridge" data-experience-stage aria-labelledby="high-ticket-title">
      <div class="home-shell high-ticket-bridge__inner">
        <div class="high-ticket-bridge__copy">
          <p class="high-ticket-kicker">KURUMSAL ÇÖZÜMLER</p>
          <h2 id="high-ticket-title">Hazır Şablonlar Sürecinize Dar mı Geliyor?</h2>
          <p>Logo, SAP, Mikro veya Zirve verileriniz dağınıksa; şirketinize has tahsilat, banka, üretim veya kârlılık akışınız varsa, hazır kalıplarla vakit kaybetmeyin. 17 yıllık ticari bankacılık ve saha finans tecrübesiyle şirketinize özel karar motoru kuralım.</p>
        </div>
        <div class="high-ticket-bridge__proof">
          <ul><li>Yazılımcıya finans anlatmakla uğraşmazsınız.</li><li>İş kuralları gerçek veriniz ve uç senaryolarla test edilir.</li><li>Kullanmayacağınız modül ve aylık lisans yükü eklenmez.</li></ul>
          <a href="/ozel-excel-sistemleri" data-cta="home_high_ticket_bridge" data-location="mid">İhtiyaca Özel Excel Sistemleri Sayfasını İnceleyin →</a>
        </div>
      </div>
    </section>`;

if (!home.includes('class="high-ticket-bridge"')) {
  home = replaceOne(home, '<section class="authority"', `${highTicketBridge}\n\n    <section class="authority"`, 'home high-ticket bridge insertion');
}

const homeBottomCta = `
    <section class="home-finance-close" data-experience-stage aria-labelledby="home-finance-close-title">
      <div class="home-shell home-finance-close__inner">
        <div><p class="eyebrow">15 DAKİKALIK ÖN TEŞHİS</p><h2 id="home-finance-close-title">Şirketinizin Finansal Tıkanıklığını 15 Dakikada Teşhis Edelim.</h2><p>İster hazır model seçin, ister sürecinizi WhatsApp'tan yazın. Doğru çözüm yolunu ve gerekli kapsamı netleştirelim.</p></div>
        <a class="home-finance-close__cta" href="${wa.home}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="home_bottom_whatsapp" data-location="bottom">WhatsApp'tan Doğrudan Danışın</a>
      </div>
    </section>`;

if (!home.includes('class="home-finance-close"')) {
  home = replaceOne(home, '<section class="faq-close"', `${homeBottomCta}\n\n    <section class="faq-close"`, 'home bottom CTA insertion');
}

home = addOrReplaceJsonLdGraph(home, (graph) => {
  const orgId = 'https://excelarsiv.com/#organization';
  const personId = 'https://excelarsiv.com/#baris-bagirlar';
  const serviceId = 'https://excelarsiv.com/#financial-decision-systems';
  ensureGraphNode(graph, {
    '@type':'Organization','@id':orgId,name:'Excel Arşiv',url:'https://excelarsiv.com/',
    description:'İşletmeler için hazır ve ihtiyaca özel Excel finansal karar sistemleri.'
  }, (node) => node?.['@id'] === orgId);
  ensureGraphNode(graph, {
    '@type':'Person','@id':personId,name:'Barış Bağırlar',jobTitle:'Ticari Bankacılık Uzmanı ve Finansal Sistem Mimarı',
    worksFor:{'@id':orgId},url:'https://excelarsiv.com/ozel-excel-sistemleri',
    description:'17 yıllık ticari bankacılık ve saha finans deneyimiyle nakit akışı, banka limit-risk, maliyet, bütçe, değerleme ve yönetim raporlama süreçlerini Excel karar sistemlerine dönüştürür.',
    knowsAbout:['13 haftalık nakit akışı','banka limit-risk ve teminat yönetimi','birim maliyet ve dinamik fiyatlama','rolling forecast ve senaryolu bütçe','DCF, IRR ve DSCR finansal modelleme','ERP verisinden CFO yönetim raporlaması']
  }, (node) => node?.['@id'] === personId);
  ensureGraphNode(graph, {
    '@type':'ProfessionalService','@id':serviceId,name:'Excel Arşiv Finansal Karar ve Excel Sistemleri',url:'https://excelarsiv.com/',
    provider:{'@id':personId},brand:{'@id':orgId},areaServed:{'@type':'Country',name:'Türkiye'},
    description:'Hazır finans modelleri ile işletmeye özel Excel karar sistemlerini aynı çözüm mimarisinde sunar.',
    knowsAbout:['Nakit ve likidite sistemleri','Banka, kredi ve teminat sistemleri','Birim maliyet ve dinamik fiyatlama','Bütçe, projeksiyon ve değerleme','İşletmeye özel Excel karar sistemleri']
  }, (node) => node?.['@id'] === serviceId || node?.['@type'] === 'ProfessionalService');
}, 'home');

const homeCss = `
<style id="dual-funnel-home-v16-css">
  .finance-pillars{padding:64px 0;border-bottom:1px solid #dfe5e0;background:#fff}
  .finance-pillars__head{display:grid;grid-template-columns:minmax(0,.9fr) minmax(0,1.1fr);gap:42px;align-items:end;margin-bottom:24px}
  .finance-pillars__head h2{margin:10px 0 0;color:#0f172a;font-size:clamp(30px,3.6vw,48px);line-height:1.04;letter-spacing:-.04em}
  .finance-pillars__head>p{margin:0;color:#647068;font-size:15px;line-height:1.7}
  .finance-pillars__grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
  .finance-pillar{position:relative;min-width:0;display:flex;flex-direction:column;padding:22px;border:1px solid #dfe5e0;border-radius:18px;background:#fff;box-shadow:0 10px 28px rgba(18,42,26,.05)}
  .finance-pillar::before{content:"";position:absolute;left:0;right:0;top:0;height:4px;border-radius:18px 18px 0 0;background:#059669}
  .finance-pillar--blue::before{background:#2563eb}.finance-pillar--violet::before{background:#7c3aed}.finance-pillar--amber::before{background:#d97706}
  .finance-pillar__no{color:#64748b;font:800 11px/1 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.08em}
  .finance-pillar h3{margin:13px 0 0;color:#0f172a;font-size:19px;line-height:1.18;letter-spacing:-.025em}
  .finance-pillar ul{display:grid;gap:9px;margin:17px 0 22px;padding:0;list-style:none;color:#526176;font-size:13px;line-height:1.5}
  .finance-pillar li{position:relative;padding-left:15px}.finance-pillar li::before{content:"";position:absolute;left:0;top:.62em;width:6px;height:6px;border-radius:50%;background:#c9d4cc}
  .finance-pillar>a{margin-top:auto;color:#075f39;font-size:13px;font-weight:800;text-decoration:none}
  .high-ticket-bridge{padding:68px 0;background:#0f172a;color:#fff}
  .high-ticket-bridge__inner{display:grid;grid-template-columns:minmax(0,1.15fr) minmax(0,.85fr);gap:68px;align-items:center}
  .high-ticket-kicker{margin:0;color:#34d399;font:850 11px/1.2 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.14em}
  .high-ticket-bridge h2{max-width:760px;margin:12px 0 0;color:#fff;font-size:clamp(32px,4vw,50px);line-height:1.02;letter-spacing:-.045em}
  .high-ticket-bridge__copy>p:last-child{max-width:760px;margin:20px 0 0;color:#cbd5e1;font-size:16px;line-height:1.7}
  .high-ticket-bridge__proof{padding:26px;border:1px solid #334155;border-radius:18px;background:#111c2f}
  .high-ticket-bridge__proof ul{display:grid;gap:12px;margin:0 0 22px;padding:0;list-style:none;color:#e2e8f0;font-size:14px;line-height:1.45}
  .high-ticket-bridge__proof li{position:relative;padding-left:22px}.high-ticket-bridge__proof li::before{content:"✓";position:absolute;left:0;color:#34d399;font-weight:900}
  .high-ticket-bridge__proof>a{min-height:48px;display:flex;align-items:center;justify-content:center;padding:0 18px;border-radius:12px;background:#10b981;color:#052e1c;font-size:13px;font-weight:850;text-decoration:none}
  .home-finance-close{padding:52px 0;border-bottom:1px solid #dfe5e0;background:linear-gradient(90deg,#f0fdf4,#eff6ff)}
  .home-finance-close__inner{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:36px;align-items:center}
  .home-finance-close h2{margin:9px 0 0;color:#0f172a;font-size:clamp(28px,3.2vw,42px);line-height:1.05;letter-spacing:-.04em}
  .home-finance-close p:not(.eyebrow){max-width:760px;margin:12px 0 0;color:#526176;font-size:15px;line-height:1.65}
  .home-finance-close__cta{min-height:50px;display:inline-flex;align-items:center;justify-content:center;padding:0 22px;border-radius:12px;background:#059669;color:#fff;font-size:14px;font-weight:850;text-decoration:none;box-shadow:0 12px 28px rgba(5,150,105,.2)}
  @media(max-width:980px){.finance-pillars__grid{grid-template-columns:repeat(2,minmax(0,1fr))}.high-ticket-bridge__inner{grid-template-columns:1fr;gap:28px}}
  @media(max-width:720px){.finance-pillars{padding:38px 0}.finance-pillars__head{grid-template-columns:1fr;gap:12px}.finance-pillars__head h2{font-size:28px}.finance-pillars__grid{grid-template-columns:1fr;gap:10px}.finance-pillar{padding:18px}.high-ticket-bridge{padding:42px 0}.high-ticket-bridge h2{font-size:32px}.high-ticket-bridge__proof{padding:20px}.home-finance-close{padding:38px 0}.home-finance-close__inner{grid-template-columns:1fr;gap:20px}.home-finance-close__cta{width:100%}}
</style>`;

if (!home.includes('dual-funnel-home-v16-css')) home = home.replace('</head>', `${homeCss}\n</head>`);

// -------------------- SPECIAL SYSTEMS --------------------
let special = fs.readFileSync(specialFile, 'utf8');

special = replaceOne(special, /<title>[\s\S]*?<\/title>/u, '<title>İhtiyaca Özel Excel Finans ve Karar Sistemleri | Excel Arşiv</title>', 'special title');
special = upsertMeta(special, 'name', 'description', '17 yıllık ticari bankacılık ve saha finans deneyimiyle 13 haftalık nakit akışı, banka limit-risk, maliyet, bütçe, DCF/DSCR ve ERP entegre CFO yönetim sistemleri.');
special = upsertMeta(special, 'property', 'og:type', 'website');
special = upsertMeta(special, 'property', 'og:url', 'https://excelarsiv.com/ozel-excel-sistemleri');
special = upsertMeta(special, 'property', 'og:title', 'İhtiyaca Özel Excel Karar Sistemleri — Barış Bağırlar');
special = upsertMeta(special, 'property', 'og:description', '17 yıllık ticari bankacılık ve saha deneyimiyle işletmenize özel dinamik finans, nakit, banka risk ve kârlılık karar sistemleri.');
special = upsertMeta(special, 'property', 'og:image', 'https://excelarsiv.com/images/hero.jpg');
special = upsertMeta(special, 'name', 'twitter:card', 'summary_large_image');
special = upsertMeta(special, 'name', 'twitter:title', 'İhtiyaca Özel Excel Karar Sistemleri — Barış Bağırlar');
special = upsertMeta(special, 'name', 'twitter:description', '17 yıllık ticari bankacılık ve saha finans deneyimiyle işletmeye özel Excel karar sistemleri.');
special = upsertMeta(special, 'name', 'twitter:image', 'https://excelarsiv.com/images/hero.jpg');

special = replaceOne(special, '<p class="eyebrow">İşletmeye özel Excel sistemleri</p>', '<p class="eyebrow">17 Yıllık Ticari Bankacılık &amp; Saha Finans Güvencesi</p>', 'special hero eyebrow');
special = replaceOne(special, /<h1>Hazır dosya işinize uymuyorsa,[\s\S]*?<\/h1>/u, '<h1>Yazılımcıya finans öğretmekle uğraşmayın. <em>Şirketinizin dilini bilen özel Excel karar sistemleri.</em></h1>', 'special hero h1');
special = replaceOne(special, /<p class="hero-lead">[\s\S]*?<\/p>/u, '<p class="hero-lead">Nakit akışını, banka limit-riskini, maliyet ve kârlılığı yalnız raporlamak değil; hangi verinin hangi yönetim kararını etkilediğini kurala bağlamak için çalışıyoruz. Dağınık veriyi patron ve CFO için okunabilir tek karar ekranına dönüştürüyoruz.</p>', 'special hero lead');

const heroPrimaryPattern = /<a class="btn btn-primary" href="\/iletisim">İhtiyacımı anlatmak istiyorum <span aria-hidden="true">→<\/span><\/a>/u;
special = replaceOne(special, heroPrimaryPattern, `<a class="btn btn-primary btn-whatsapp" href="${wa.hero}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="hero_whatsapp" data-location="top">WhatsApp'tan Yazın — 15 Dakikada İhtiyacı Çıkaralım <span aria-hidden="true">→</span></a>`, 'special hero WhatsApp CTA');

// Header and mobile menu use the verified number too.
special = special.replace(/<a class="btn btn-primary nav-cta" href="\/iletisim">[^<]*<\/a>/u, `<a class="btn btn-primary nav-cta btn-whatsapp" href="${wa.hero}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="nav_whatsapp" data-location="header">WhatsApp'tan Yazın</a>`);
special = special.replace(/<a href="\/iletisim">İhtiyacımı anlat<\/a>/u, `<a href="${wa.hero}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="mobile_menu_whatsapp" data-location="header">WhatsApp'tan Yazın</a>`);

// Move the four proof cards out of the left hero column into a full-width strip below the hero grid.
const heroStart = special.indexOf('<section class="hero">');
const heroEnd = heroStart >= 0 ? special.indexOf('</section>', heroStart) : -1;
if (heroStart < 0 || heroEnd < 0) throw new Error('DUAL FUNNEL V16: special hero section boundary missing');
let heroSection = special.slice(heroStart, heroEnd + '</section>'.length);
const proofMatch = heroSection.match(/(<div class="proof-grid"[\s\S]*?<\/div>)\s*<\/div>\s*(<div class="workbook"\b)/u);
if (!proofMatch) throw new Error('DUAL FUNNEL V16: special proof grid could not be detached from hero copy');
const proofGrid = proofMatch[1].replace('class="proof-grid"', 'class="proof-grid proof-grid--full"');
heroSection = heroSection.replace(proofMatch[0], `</div>\n\n        ${proofMatch[2]}`);
heroSection = heroSection.replace('</section>', `  <div class="wrap-wide hero-proof-strip">${proofGrid}</div>\n    </section>`);
special = `${special.slice(0, heroStart)}${heroSection}${special.slice(heroEnd + '</section>'.length)}`;

const authorityRail = `<div class="authority-rail" aria-label="ExcelArşiv uzmanlık özeti">
          <div class="authority-item"><strong>17 Yıl:</strong><span>Ticari bankacılık, kredi, limit ve risk bakışı</span></div>
          <div class="authority-item"><strong>Banka ↔ İşletme:</strong><span>Finans, muhasebe ve saha dilini aynı modelde birleştirme</span></div>
          <div class="authority-item"><strong>ERP → CFO Kararı:</strong><span>Dağınık veriyi rapordan aksiyona taşıyan yönetim sistemi</span></div>
        </div>`;
if (!special.includes('aria-label="ExcelArşiv uzmanlık özeti"')) {
  special = replaceOne(special, '<div class="comparison-wrap">', `${authorityRail}\n        <div class="comparison-wrap">`, 'special authority rail');
}

const midCta = `<section class="special-mid-cta" aria-label="Ön teşhis görüşmesi"><div class="wrap-wide special-mid-cta__inner"><div><strong>Hazır sistem değil, işletmenize göre karar akışı mı gerekiyor?</strong><span>Mevcut tablonuzu veya darboğazı anlatın; gerekli finansal sistemi birlikte netleştirelim.</span></div><a href="${wa.mid}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="mid_whatsapp" data-location="problem_solver">WhatsApp'tan Yazın — İhtiyacı Netleştirelim</a></div></section>`;
if (!special.includes('data-cta="mid_whatsapp"')) {
  const areaStart = special.indexOf('<section class="section section-green" id="alanlar">');
  const areaEnd = areaStart >= 0 ? special.indexOf('</section>', areaStart) : -1;
  if (areaStart < 0 || areaEnd < 0) throw new Error('DUAL FUNNEL V16: special finance areas section missing');
  const insertPos = areaEnd + '</section>'.length;
  special = `${special.slice(0, insertPos)}\n${midCta}${special.slice(insertPos)}`;
}

special = special.replace(/<a class="btn btn-primary" href="\/iletisim">Mevcut dosyamı anlatmak istiyorum<\/a>/u, `<a class="btn btn-primary btn-whatsapp" href="${wa.mid}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="fit_whatsapp" data-location="problem_solver">WhatsApp'tan Yazın — İhtiyacı Netleştirelim</a>`);
special = special.replace(/<a class="btn btn-primary" href="\/iletisim">İhtiyacımı anlat <span aria-hidden="true">→<\/span><\/a>/u, `<a class="btn btn-primary btn-whatsapp" href="${wa.bottom}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="footer_whatsapp" data-location="bottom">WhatsApp'tan Yazın — 15 Dakikada İhtiyacı Çıkaralım <span aria-hidden="true">→</span></a>`);

const floating = `<a class="floating-whatsapp" href="${wa.floating}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="floating_whatsapp" data-location="mobile_sticky" aria-label="WhatsApp ile 15 dakikalık ön teşhis görüşmesi başlat"><span aria-hidden="true">◉</span><strong>15 Dk'da Teşhis</strong></a>`;
if (!special.includes('class="floating-whatsapp"')) special = special.replace('</body>', `  ${floating}\n</body>`);

special = addOrReplaceJsonLdGraph(special, (graph) => {
  const orgId = 'https://excelarsiv.com/#organization';
  const personId = 'https://excelarsiv.com/#baris-bagirlar';
  const serviceId = 'https://excelarsiv.com/ozel-excel-sistemleri#professional-service';
  ensureGraphNode(graph, {'@type':'Organization','@id':orgId,name:'Excel Arşiv',url:'https://excelarsiv.com/'}, (node) => node?.['@id'] === orgId);
  ensureGraphNode(graph, {
    '@type':'Person','@id':personId,name:'Barış Bağırlar',jobTitle:'Ticari Bankacılık Uzmanı ve Finansal Sistem Mimarı',url:'https://excelarsiv.com/ozel-excel-sistemleri',worksFor:{'@id':orgId},
    description:'17 yıllık ticari bankacılık ve saha finans deneyimiyle şirketlerin nakit, kredi, limit-risk, maliyet, bütçe, yatırım ve yönetim raporlama ihtiyaçlarını özel Excel karar sistemlerine dönüştürür.',
    knowsAbout:['13 haftalık nakit akışı','banka limit-risk ve teminat yönetimi','rolling forecast ve senaryolu bütçe','DCF, IRR ve DSCR finansal modelleme','ERP verisinden CFO yönetim raporlaması']
  }, (node) => node?.['@id'] === personId || node?.['@type'] === 'Person');
  ensureGraphNode(graph, {
    '@type':'ProfessionalService','@id':serviceId,name:'Excel Arşiv - Özel Excel Karar Sistemleri',url:'https://excelarsiv.com/ozel-excel-sistemleri',provider:{'@id':personId},brand:{'@id':orgId},areaServed:{'@type':'Country',name:'Türkiye'},
    description:'17 yıllık ticari bankacılık ve saha finans deneyimiyle şirketlere özel dinamik nakit akışı, banka limit-risk, maliyet, bütçe, değerleme ve ERP entegre CFO karar sistemleri tasarımı.',
    knowsAbout:['Dinamik Nakit Akışı (13-Week Cash Flow)','Banka Limit-Risk ve Teminat Havuzu','Birim Maliyet ve Dinamik Fiyatlama','Senaryolu Bütçe ve Rolling Forecast','Yatırım Fizibilitesi ve DCF Değerleme','ERP Entegre CFO Yönetim Kokpiti']
  }, (node) => node?.['@id'] === serviceId || node?.['@type'] === 'ProfessionalService');
}, 'special');

const specialCss = `
<style id="special-360-v16-css">
  body{--v16-emerald:#059669;--v16-navy:#0f172a}
  .hero{padding:28px 0 20px!important;background:linear-gradient(90deg,#fff 0%,#fff 54%,#f4fbf7 100%)!important}
  .hero>.wrap-wide{width:min(1680px,calc(100% - 64px))!important}
  .hero-grid{grid-template-columns:minmax(0,.92fr) minmax(440px,1.08fr)!important;gap:36px!important;align-items:center!important}
  .hero h1{max-width:720px!important;font-size:clamp(34px,3.2vw,46px)!important;line-height:1.04!important;letter-spacing:-.046em!important;color:#0f172a!important}
  .hero-lead{max-width:700px!important;margin-top:18px!important;font-size:16px!important;line-height:1.58!important;color:#475569!important}
  .hero-actions{margin-top:22px!important;gap:10px!important}.hero-actions .btn{min-height:46px!important;font-size:14px!important;padding:0 17px!important}
  .workbook{max-width:100%!important;min-width:0!important;overflow:hidden!important;border-radius:20px!important}
  .hero-proof-strip{width:min(1680px,calc(100% - 64px));margin:28px auto 48px}.hero-proof-strip .proof-grid{grid-template-columns:repeat(4,minmax(0,1fr))!important;gap:14px!important;margin:0!important}
  .hero-proof-strip .proof-chip{min-height:118px;padding:18px!important;border-radius:16px!important}.hero-proof-strip .proof-chip strong{font-size:16px!important}.hero-proof-strip .proof-chip span{font-size:14px!important;line-height:1.45!important}
  .authority-rail{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;max-width:980px;margin:0 auto 28px;padding:18px;border:1px solid #dbe3eb;border-radius:18px;background:#fff;box-shadow:0 10px 24px rgba(15,23,42,.05)}
  .authority-item{display:flex;align-items:flex-start;gap:9px;min-width:0;color:#475569;font-size:15px;line-height:1.45}.authority-item strong{flex:0 0 auto;white-space:nowrap;color:#0f172a;font-weight:800}.authority-item span{min-width:0;overflow-wrap:anywhere}
  .faq-grid{align-items:start!important}.faq-copy,.faq-list{align-self:start!important;min-width:0}
  .special-mid-cta{padding:0 0 56px;background:#edf9f2}.special-mid-cta__inner{display:flex;align-items:center;justify-content:space-between;gap:28px;padding:24px 28px;border:1px solid #cfe7d8;border-radius:18px;background:#fff;box-shadow:0 12px 30px rgba(15,23,42,.06)}.special-mid-cta__inner>div{display:grid;gap:5px}.special-mid-cta__inner strong{color:#0f172a;font-size:18px}.special-mid-cta__inner span{color:#64748b;font-size:14px;line-height:1.5}.special-mid-cta__inner>a{flex:0 0 auto;min-height:46px;display:inline-flex;align-items:center;justify-content:center;padding:0 18px;border-radius:12px;background:#059669;color:#fff;font-size:14px;font-weight:800;text-decoration:none}
  .btn-whatsapp{background:#059669!important;color:#fff!important;border-color:#059669!important}.btn-whatsapp:hover{background:#047857!important}
  .floating-whatsapp{display:none;position:fixed;right:16px;bottom:calc(16px + env(safe-area-inset-bottom));z-index:140;align-items:center;gap:8px;min-height:48px;padding:0 15px;border:1px solid rgba(255,255,255,.5);border-radius:999px;background:#059669;color:#fff;font-size:13px;text-decoration:none;box-shadow:0 16px 36px rgba(5,150,105,.3)}.floating-whatsapp span{font-size:18px}.floating-whatsapp strong{font-weight:850}
  @media(max-width:980px){.hero-grid{grid-template-columns:1fr!important}.hero>.wrap-wide,.hero-proof-strip{width:calc(100% - 40px)!important}.authority-rail{grid-template-columns:1fr}.special-mid-cta__inner{align-items:flex-start;flex-direction:column}.special-mid-cta__inner>a{width:100%}}
  @media(max-width:660px){.hero{padding:18px 0 10px!important}.hero>.wrap-wide,.hero-proof-strip{width:calc(100% - 28px)!important}.hero h1{font-size:clamp(31px,9vw,40px)!important}.hero-lead{font-size:15px!important}.workbook{display:none!important}.hero-proof-strip{margin:20px auto 38px}.hero-proof-strip .proof-grid{grid-template-columns:1fr!important;gap:10px!important}.hero-proof-strip .proof-chip{min-height:0;padding:16px!important}.authority-rail{margin-bottom:22px;padding:15px}.authority-item{display:grid;grid-template-columns:auto minmax(0,1fr);gap:7px}.special-mid-cta{padding-bottom:38px}.special-mid-cta__inner{padding:20px}.floating-whatsapp{display:inline-flex}.footer{padding-bottom:84px!important}}
</style>`;
if (!special.includes('special-360-v16-css')) special = special.replace('</head>', `${specialCss}\n</head>`);

// -------------------- RELEASE CONTRACT --------------------
const requiredHomeTokens = [
  'İşletmeler İçin', 'Finansal Karar', 'İşletmenize Özel Sistem Kuralım', 'finance-pillars', 'high-ticket-bridge',
  'home_finance_cash', 'home_finance_credit', 'home_finance_cost', 'home_finance_budget', 'ProfessionalService',
];
const requiredSpecialTokens = [
  '17 Yıllık Ticari Bankacılık &amp; Saha Finans Güvencesi', 'Yazılımcıya finans öğretmekle uğraşmayın.', 'ERP → CFO Kararı:',
  `https://wa.me/${WA_PHONE}?text=`, 'data-cta="hero_whatsapp"', 'data-cta="mid_whatsapp"', 'data-cta="footer_whatsapp"', 'data-cta="floating_whatsapp"',
  'special-360-v16-css', 'ProfessionalService',
];
for (const token of requiredHomeTokens) if (!home.includes(token)) throw new Error(`DUAL FUNNEL V16: home required token missing: ${token}`);
for (const token of requiredSpecialTokens) if (!special.includes(token)) throw new Error(`DUAL FUNNEL V16: special required token missing: ${token}`);
if (home.includes(OLD_PHONE) || special.includes(OLD_PHONE)) throw new Error('DUAL FUNNEL V16: obsolete WhatsApp number leaked into final HTML');
if (home.includes('overflow-x:hidden') || special.includes('overflow-x:hidden')) throw new Error('DUAL FUNNEL V16: forbidden global overflow-x:hidden introduced');

const specialWaCount = (special.match(new RegExp(`https://wa\\.me/${WA_PHONE}\\?text=`, 'gu')) || []).length;
if (specialWaCount < 4) throw new Error(`DUAL FUNNEL V16: expected >=4 special WhatsApp routes, got ${specialWaCount}`);

// Prevent the retired number from re-entering current source/build scripts.
const textExt = new Set(['.mjs','.js','.ts','.astro','.html','.json','.md','.cjs','.css']);
const ignored = new Set(['node_modules','.git','dist','.astro']);
const offenders = [];
const walk = (dir) => {
  for (const entry of fs.readdirSync(dir, {withFileTypes:true})) {
    if (ignored.has(entry.name)) continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) { walk(full); continue; }
    if (!textExt.has(path.extname(entry.name))) continue;
    const text = fs.readFileSync(full, 'utf8');
    if (text.includes(OLD_PHONE)) offenders.push(path.relative(ROOT, full));
  }
};
walk(ROOT);
if (offenders.length) throw new Error(`DUAL FUNNEL V16: obsolete WhatsApp number remains in source: ${offenders.join(', ')}`);

fs.writeFileSync(homeFile, home, 'utf8');
fs.writeFileSync(specialFile, special, 'utf8');
console.log(`DUAL FUNNEL V16 PASS — home dual funnel + 4 finance pillars + high-ticket bridge + entity SEO; special P0 layout + ${specialWaCount} verified WhatsApp routes on ${WA_PHONE} + mobile sticky CTA.`);
