#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/index.html');
if (!fs.existsSync(file)) throw new Error('HOME DUAL FUNNEL V17: dist/index.html missing');
let html = fs.readFileSync(file, 'utf8');

const WA_PHONE = '905393333303';
const wa = (message) => `https://wa.me/${WA_PHONE}?text=${encodeURIComponent(message)}`;
const bottomWa = wa('Merhaba Barış Bey, excelarsiv.com ana sayfasından ulaşıyorum. Şirketimizin finansal tıkanıklığını ve uygun Excel karar sistemini netleştirmek istiyorum.');

const requireReplace = (pattern, replacement, label) => {
  const next = html.replace(pattern, replacement);
  if (next === html) throw new Error(`HOME DUAL FUNNEL V17: replacement missed: ${label}`);
  html = next;
};

if (html.includes('data-dual-funnel-home-v17')) {
  console.log('HOME DUAL FUNNEL V17: already applied');
  process.exit(0);
}

requireReplace(/<body\b([^>]*)>/u, '<body$1 data-dual-funnel-home-v17>', 'body marker');
requireReplace(/<title>[\s\S]*?<\/title>/u, '<title>Finansal Karar ve Excel Sistemleri | Excel Arşiv</title>', 'title');

const financePillars = `
<section class="finance-pillars" id="finans-sistemleri" data-experience-stage>
  <div class="home-shell">
    <div class="finance-pillars__head">
      <div><p class="eyebrow">FİNANSAL KARAR SİSTEMLERİ</p><h2>Dört ana finansal karar alanından başlayın.</h2></div>
      <p>Hazır sistem arayan işletmeler için ihtiyaçları finansal karar başlıklarına ayırdık. Daha özel bir veri akışınız varsa doğrudan işletmenize özel sistem hattına geçebilirsiniz.</p>
    </div>
    <div class="finance-pillars__grid">
      <article class="finance-pillar finance-pillar--green"><span class="finance-pillar__no">01</span><h3>Nakit &amp; Likidite Sistemleri</h3><ul><li>13 Haftalık Dinamik Nakit Akışı Modeli</li><li>Günlük Kasa &amp; Banka Likidite Takip Paneli</li><li>Çek &amp; Senet Vade Yaşlandırma Tablosu</li></ul><a href="/sablonlar?q=nakit" data-cta="home_finance_cash" data-location="finance_pillars">Nakit sistemlerini gör →</a></article>
      <article class="finance-pillar finance-pillar--blue"><span class="finance-pillar__no">02</span><h3>Banka, Kredi &amp; Teminat Sistemleri</h3><ul><li>Çoklu Banka Limit-Risk ve Kredi Portföyü</li><li>Rotatif Kredi Faiz ve Finansman Maliyet Simülatörü</li><li>Teminat Mektubu &amp; İpotek Karşılama Matrisi</li></ul><a href="/sablonlar?q=kredi" data-cta="home_finance_credit" data-location="finance_pillars">Banka sistemlerini gör →</a></article>
      <article class="finance-pillar finance-pillar--violet"><span class="finance-pillar__no">03</span><h3>Birim Maliyet &amp; Dinamik Fiyatlama</h3><ul><li>Değişken Maliyet ve Hammadde Endeksli Fiyat Teklif Motoru</li><li>Ürün &amp; Müşteri Bazlı Katkı Payı ve Kârlılık Matrisi</li><li>Şirket &amp; Proje Başabaş (Break-Even) Hesaplayıcı</li></ul><a href="/sablonlar?q=maliyet" data-cta="home_finance_cost" data-location="finance_pillars">Maliyet sistemlerini gör →</a></article>
      <article class="finance-pillar finance-pillar--amber"><span class="finance-pillar__no">04</span><h3>Bütçe, Projeksiyon &amp; Değerleme</h3><ul><li>3 Senaryolu Dinamik Bütçe &amp; Rolling Forecast Modeli</li><li>DCF İskontolanmış Nakit Akımları &amp; Yatırım Fizibilitesi</li><li>Net İşletme Sermayesi (NÖS) &amp; DSO Gösterge Kokpiti</li></ul><a href="/sablonlar?q=butce" data-cta="home_finance_budget" data-location="finance_pillars">Bütçe sistemlerini gör →</a></article>
    </div>
  </div>
</section>`;

if (!html.includes('class="finance-pillars"')) requireReplace('<section class="catalog-proof"', `${financePillars}\n<section class="catalog-proof"`, 'finance pillars');

const bridge = `
<section class="high-ticket-bridge" data-experience-stage aria-labelledby="high-ticket-title">
  <div class="home-shell high-ticket-bridge__inner">
    <div class="high-ticket-bridge__copy"><p class="high-ticket-kicker">KURUMSAL ÇÖZÜMLER</p><h2 id="high-ticket-title">Hazır Şablonlar Sürecinize Dar mı Geliyor?</h2><p>Logo, SAP, Mikro veya Zirve verileriniz dağınıksa; şirketinize has tahsilat, banka, üretim veya kârlılık akışınız varsa, hazır kalıplarla vakit kaybetmeyin. 17 yıllık ticari bankacılık ve saha finans tecrübesiyle şirketinize özel karar motoru kuralım.</p></div>
    <div class="high-ticket-bridge__proof"><ul><li>Yazılımcıya finans anlatmakla uğraşmazsınız.</li><li>İş kuralları gerçek veriniz ve uç senaryolarla test edilir.</li><li>Kullanmayacağınız modül ve aylık lisans yükü eklenmez.</li></ul><a href="/ozel-excel-sistemleri" data-cta="home_high_ticket_bridge" data-location="mid">İhtiyaca Özel Excel Sistemleri Sayfasını İnceleyin →</a></div>
  </div>
</section>`;

if (!html.includes('class="high-ticket-bridge"')) requireReplace('<section class="authority"', `${bridge}\n<section class="authority"`, 'high-ticket bridge');

const bottom = `
<section class="home-finance-close" data-experience-stage aria-labelledby="home-finance-close-title">
  <div class="home-shell home-finance-close__inner"><div><p class="eyebrow">15 DAKİKALIK ÖN TEŞHİS</p><h2 id="home-finance-close-title">Şirketinizin Finansal Tıkanıklığını 15 Dakikada Teşhis Edelim.</h2><p>İster hazır model seçin, ister sürecinizi WhatsApp'tan yazın. Doğru çözüm yolunu ve gerekli kapsamı netleştirelim.</p></div><a class="home-finance-close__cta" href="${bottomWa}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="home_bottom_whatsapp" data-location="bottom">WhatsApp'tan Doğrudan Danışın</a></div>
</section>`;

if (!html.includes('class="home-finance-close"')) requireReplace('<section class="faq-close"', `${bottom}\n<section class="faq-close"`, 'bottom CTA');

const schema = {
  '@context':'https://schema.org',
  '@graph':[
    {'@type':'Organization','@id':'https://excelarsiv.com/#organization',name:'Excel Arşiv',url:'https://excelarsiv.com/',description:'İşletmeler için hazır ve ihtiyaca özel Excel finansal karar sistemleri.'},
    {'@type':'Person','@id':'https://excelarsiv.com/#baris-bagirlar',name:'Barış Bağırlar',jobTitle:'Ticari Bankacılık Uzmanı ve Finansal Sistem Mimarı',worksFor:{'@id':'https://excelarsiv.com/#organization'},url:'https://excelarsiv.com/ozel-excel-sistemleri',description:'17 yıllık ticari bankacılık ve saha finans deneyimiyle nakit akışı, banka limit-risk, maliyet, bütçe, değerleme ve yönetim raporlama süreçlerini Excel karar sistemlerine dönüştürür.'},
    {'@type':'ProfessionalService','@id':'https://excelarsiv.com/#financial-decision-systems',name:'Excel Arşiv Finansal Karar ve Excel Sistemleri',url:'https://excelarsiv.com/',provider:{'@id':'https://excelarsiv.com/#baris-bagirlar'},brand:{'@id':'https://excelarsiv.com/#organization'},areaServed:{'@type':'Country',name:'Türkiye'},description:'Hazır finans modelleri ile işletmeye özel Excel karar sistemlerini aynı çözüm mimarisinde sunar.',knowsAbout:['Nakit ve likidite sistemleri','Banka, kredi ve teminat sistemleri','Birim maliyet ve dinamik fiyatlama','Bütçe, projeksiyon ve değerleme','İşletmeye özel Excel karar sistemleri']}
  ]
};
html = html.replace('</head>', `<script type="application/ld+json" id="dual-funnel-home-schema-v17">${JSON.stringify(schema)}</script>\n</head>`);

const css = `<style id="dual-funnel-home-v17-css">
.finance-pillars{padding:64px 0;border-bottom:1px solid #dfe5e0;background:#fff}.finance-pillars__head{display:grid;grid-template-columns:minmax(0,.9fr) minmax(0,1.1fr);gap:42px;align-items:end;margin-bottom:24px}.finance-pillars__head h2{margin:10px 0 0;color:#0f172a;font-size:clamp(30px,3.6vw,48px);line-height:1.04;letter-spacing:-.04em}.finance-pillars__head>p{margin:0;color:#647068;font-size:15px;line-height:1.7}.finance-pillars__grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}.finance-pillar{position:relative;min-width:0;display:flex;flex-direction:column;padding:22px;border:1px solid #dfe5e0;border-radius:18px;background:#fff;box-shadow:0 10px 28px rgba(18,42,26,.05)}.finance-pillar::before{content:"";position:absolute;left:0;right:0;top:0;height:4px;border-radius:18px 18px 0 0;background:#059669}.finance-pillar--blue::before{background:#2563eb}.finance-pillar--violet::before{background:#7c3aed}.finance-pillar--amber::before{background:#d97706}.finance-pillar__no{color:#64748b;font:800 11px/1 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.08em}.finance-pillar h3{margin:13px 0 0;color:#0f172a;font-size:19px;line-height:1.18;letter-spacing:-.025em}.finance-pillar ul{display:grid;gap:9px;margin:17px 0 22px;padding:0;list-style:none;color:#526176;font-size:13px;line-height:1.5}.finance-pillar li{position:relative;padding-left:15px}.finance-pillar li::before{content:"";position:absolute;left:0;top:.62em;width:6px;height:6px;border-radius:50%;background:#c9d4cc}.finance-pillar>a{margin-top:auto;color:#075f39;font-size:13px;font-weight:800;text-decoration:none}.high-ticket-bridge{padding:68px 0;background:#0f172a;color:#fff}.high-ticket-bridge__inner{display:grid;grid-template-columns:minmax(0,1.15fr) minmax(0,.85fr);gap:68px;align-items:center}.high-ticket-kicker{margin:0;color:#34d399;font:850 11px/1.2 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.14em}.high-ticket-bridge h2{max-width:760px;margin:12px 0 0;color:#fff;font-size:clamp(32px,4vw,50px);line-height:1.02;letter-spacing:-.045em}.high-ticket-bridge__copy>p:last-child{max-width:760px;margin:20px 0 0;color:#cbd5e1;font-size:16px;line-height:1.7}.high-ticket-bridge__proof{padding:26px;border:1px solid #334155;border-radius:18px;background:#111c2f}.high-ticket-bridge__proof ul{display:grid;gap:12px;margin:0 0 22px;padding:0;list-style:none;color:#e2e8f0;font-size:14px;line-height:1.45}.high-ticket-bridge__proof li{position:relative;padding-left:22px}.high-ticket-bridge__proof li::before{content:"✓";position:absolute;left:0;color:#34d399;font-weight:900}.high-ticket-bridge__proof>a{min-height:48px;display:flex;align-items:center;justify-content:center;padding:0 18px;border-radius:12px;background:#10b981;color:#052e1c;font-size:13px;font-weight:850;text-decoration:none}.home-finance-close{padding:52px 0;border-bottom:1px solid #dfe5e0;background:linear-gradient(90deg,#f0fdf4,#eff6ff)}.home-finance-close__inner{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:36px;align-items:center}.home-finance-close h2{margin:9px 0 0;color:#0f172a;font-size:clamp(28px,3.2vw,42px);line-height:1.05;letter-spacing:-.04em}.home-finance-close p:not(.eyebrow){max-width:760px;margin:12px 0 0;color:#526176;font-size:15px;line-height:1.65}.home-finance-close__cta{min-height:50px;display:inline-flex;align-items:center;justify-content:center;padding:0 22px;border-radius:12px;background:#059669;color:#fff;font-size:14px;font-weight:850;text-decoration:none;box-shadow:0 12px 28px rgba(5,150,105,.2)}@media(max-width:980px){.finance-pillars__grid{grid-template-columns:repeat(2,minmax(0,1fr))}.high-ticket-bridge__inner{grid-template-columns:1fr;gap:28px}}@media(max-width:720px){.finance-pillars{padding:38px 0}.finance-pillars__head{grid-template-columns:1fr;gap:12px}.finance-pillars__head h2{font-size:28px}.finance-pillars__grid{grid-template-columns:1fr;gap:10px}.finance-pillar{padding:18px}.high-ticket-bridge{padding:42px 0}.high-ticket-bridge h2{font-size:32px}.high-ticket-bridge__proof{padding:20px}.home-finance-close{padding:38px 0}.home-finance-close__inner{grid-template-columns:1fr;gap:20px}.home-finance-close__cta{width:100%}}
</style>`;
html = html.replace('</head>', `${css}\n</head>`);

for (const token of ['finance-pillars','high-ticket-bridge','home_finance_cash','home_finance_credit','home_finance_cost','home_finance_budget','ProfessionalService']) {
  if (!html.includes(token)) throw new Error(`HOME DUAL FUNNEL V17: required token missing: ${token}`);
}

fs.writeFileSync(file, html, 'utf8');
console.log('HOME DUAL FUNNEL V17 PASS — finance pillars, high-ticket bridge, bottom WhatsApp CTA and entity schema active; homepage hero ownership left untouched.');
