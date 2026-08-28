import fs from 'node:fs';
import path from 'node:path';

const DIST = path.resolve('dist');
const SPECIAL_ENTERPRISE_CSS = path.join(DIST, 'styles', 'ozel-excel-enterprise.css');
const SPECIAL_EDITORIAL_CSS = path.join(DIST, 'styles', 'ozel-excel-editorial.css');

const specialContentReplacements = [
  ['EXCLUSIVE&nbsp;&nbsp;•&nbsp;&nbsp;GÜVENİLİR&nbsp;&nbsp;•&nbsp;&nbsp;ÖLÇÜLEBİLİR','FİNANS&nbsp;&nbsp;•&nbsp;&nbsp;MUHASEBE&nbsp;&nbsp;•&nbsp;&nbsp;YÖNETİM KONTROLÜ'],
  ['<h1>Excel ile Sınırlarınızı Aşın,<span>Gerçek İş Sonuçları Alın.</span></h1>','<h1>Mizan, Nakit, Cari ve Banka Verinizi,<span>Tek Bir Karar Sisteminde Birleştirin.</span></h1>'],
  ['Muhasebe, finans, satış, stok, üretim ve yönetim süreçleriniz için <strong>işleyişinize göre özel tasarlanan Excel sistemleri</strong> geliştiriyoruz. Amaç daha fazla hücre değil; daha görünür nakit, daha hızlı kontrol ve daha güvenilir yönetim kararı.','Mizan, muavin, nakit akışı, cari yaşlandırma, çek/senet, banka limit-risk ve yönetim raporlamasını <strong>işletmenizin gerçek çalışma düzenine göre tek Excel mimarisinde</strong> birleştiriyoruz. Amaç daha fazla hücre değil; kontrol edilebilir finans verisi ve daha hızlı yönetim kararı.'],
  ['<b>%100 Size Özel</b>','<b>İş Akışına Özel</b>'],
  ['<span>İş akışınıza göre kurulan Excel sistemleri</span>','<span>Hazır şablon değil, sürecinize göre kurulan sistem</span>'],
  ['>Satışlar</div>','>13 Haftalık Nakit</div>'],
  ['>Stok</div>','>120/320 Kontrol</div>'],
  ['>Raporlar</div>','>Banka Limit</div>'],
  ['>Analiz</div>','>Valör &amp; Faiz</div>'],
  ['<small>Toplam Ciro</small><strong>₺ 8.750.000</strong><em>▲ %18,6</em>','<small>13 Haftalık Net Nakit</small><strong>₺ 2.375.000</strong><em>13 hafta</em>'],
  ['<small>Brüt Kâr</small><strong>₺ 2.125.000</strong><em>▲ %22,4</em>','<small>Açık Tahsilat</small><strong>₺ 1.480.000</strong><em>yaşlandırılmış</em>'],
  ['<small>Net Kâr</small><strong>₺ 1.125.000</strong><em>▲ %16,8</em>','<small>Banka Limit Boşluğu</small><strong>₺ 3.250.000</strong><em>kullanılabilir</em>'],
  ['<small>Tahsilat Oranı</small><strong>% 92,5</strong><em>▲ %5,3</em>','<small>Kontrol Bekleyen Cari</small><strong>₺ 620.000</strong><em>istisna</em>'],
  ['<div class="panel-title">Aylık Ciro Trendi</div>','<div class="panel-title">13 Haftalık Nakit Eğrisi</div>'],
  ['<div class="panel-title">Ciro Dağılımı</div>','<div class="panel-title">Banka Limit Kullanımı</div>'],
  ['<span>Yurt İçi</span><span>Yurt Dışı</span><span>Diğer</span>','<span>Kullanılan</span><span>Boş Limit</span><span>Risk / Bloke</span>'],
  ['<span>En Çok Satış Yapılan</span><span>Miktar</span><span>Ciro</span>','<span>Kontrol Alanı</span><span>Adet</span><span>Tutar</span>'],
  ['<span>Ürün A</span><span>1.250</span><span>₺2.125.000</span>','<span>120 Alıcılar</span><span>12</span><span>₺2.125.000</span>'],
  ['<span>Ürün B</span><span>980</span><span>₺1.650.000</span>','<span>320 Satıcılar</span><span>8</span><span>₺1.650.000</span>'],
  ['<span>Ürün C</span><span>750</span><span>₺1.250.000</span>','<span>Çek / Senet</span><span>6</span><span>₺1.250.000</span>'],
  ['GÜVENİLEN<br/>ÇÖZÜM ORTAĞINIZ','FİNANSAL<br/>KONTROL ODAKLARI'],
  ["KOBİ'den mali müşavirlik ofisine, finans ekibinden yönetime kadar iş akışına göre yapılandırılmış Excel sistemleri.",'Mizan, nakit, cari, banka, valör ve yönetim raporlamasını aynı kontrol omurgasında birleştiren sistemler.'],
  ['<span class="sector"><span class="sector-icon">₺</span>FİNANS</span><span class="sector"><span class="sector-icon">M</span>MUHASEBE</span><span class="sector"><span class="sector-icon">S</span>SATIŞ</span><span class="sector"><span class="sector-icon">L</span>LOJİSTİK</span><span class="sector"><span class="sector-icon">Ü</span>ÜRETİM</span><span class="sector"><span class="sector-icon">Y</span>YÖNETİM</span>','<span class="sector"><span class="sector-icon">M</span>MİZAN</span><span class="sector"><span class="sector-icon">N</span>NAKİT</span><span class="sector"><span class="sector-icon">C</span>CARİ</span><span class="sector"><span class="sector-icon">B</span>BANKA</span><span class="sector"><span class="sector-icon">V</span>VALÖR</span><span class="sector"><span class="sector-icon">R</span>RAPOR</span>'],
  ['<h2 class="section-title">İşinizi Büyüten Excel Çözümleri</h2>','<h2 class="section-title">Finansal Kararı Taşıyan Excel Sistem Mimarileri</h2>'],
  ['<h3>Finans & Muhasebe</h3><p>Bilanço, nakit akışı, bütçe, mizan kontrolü ve finansal analiz sistemleri.</p>','<h3>13 Haftalık Nakit Akışı & Likidite</h3><p>Günlük/haftalık nakit görünümü, ödeme-tahsilat ufku, erken uyarı ve stres senaryoları.</p>'],
  ['<h3>Satış & Tahsilat</h3><p>Teklif, sipariş, müşteri takibi, yaşlandırma ve tahsilat görünürlüğü.</p>','<h3>Çek / Senet Portföyü & Valör</h3><p>Vade, ortalama valör, portföy yoğunluğu, tahsilat takvimi ve finansman etkisi.</p>'],
  ['<h3>Stok & Depo Yönetimi</h3><p>Stok hareketi, kritik seviye, ürün kârlılığı ve envanter analizi.</p>','<h3>Luca / Logo / Zirve Mizan Entegrasyonu</h3><p>Muavin-mizan aktarımı, 120/320 ters bakiye, kapanış ve istisna kontrolleri.</p>'],
  ['<h3>Üretim & Planlama</h3><p>Üretim planı, kapasite, maliyet ve operasyonel performans görünürlüğü.</p>','<h3>Yönetim & İcra Kurulu Dashboard</h3><p>EBITDA, net işletme sermayesi, tahsilat, likidite ve yönetim KPI görünümü.</p>'],
  ['<h3>Raporlama & Dashboard</h3><p>Yönetici panoları, KPI takibi, senaryo analizi ve karar destek raporları.</p>','<h3>Banka Limit-Risk & Faiz Maliyeti</h3><p>Limit doluluk, risk, rotatif faiz tahakkuku, DSCR ve döviz pozisyonu izleme.</p>'],
  ['<h3>Özel Çözümler</h3><p>İş akışınıza, veri yapınıza ve kullanıcı rollerinize özel Excel uygulamaları.</p>','<h3>Cari Yaşlandırma & 120/320 Kontrol</h3><p>Vade kovaları, ters bakiye, mahsup ve kapanmayan cari istisnalarını görünür kılan kontrol sistemi.</p>']
];

const fixes = {
  '/hakkinda/index.html': `
<style id="final-ui-contrast-fixes">
.hakkinda__hero,.hakkinda__cta{color:#f8fafc!important}.hakkinda__hero h1,.hakkinda__hero h2,.hakkinda__hero h3,.hakkinda__hero strong{color:#fff!important}.hakkinda__hero .hakkinda__kicker{color:#9ef0bd!important}.hakkinda__hero .hakkinda__unvan{color:#c9f7d9!important}.hakkinda__hero .hakkinda__rol{color:#edf8f1!important}.hakkinda__hero .hakkinda__rozet li{color:#f5fff9!important}.hakkinda__cta p{color:#eef8f2!important}.hakkinda__cta a{color:#b8f2cd!important}.hakkinda__cta a:hover{color:#fff!important}
a[class*="bg-green"],button[class*="bg-green"],a[class*="bg-emerald"],button[class*="bg-emerald"],.site-header__cta,.btn-primary{color:#fff!important}a[class*="bg-green"] *,button[class*="bg-green"] *,a[class*="bg-emerald"] *,button[class*="bg-emerald"] *,.site-header__cta *,.btn-primary *{color:inherit!important}
</style>`,
  '/rehber/index.html': `
<style id="final-ui-contrast-fixes">
aside div[class*="from-[#0B192C]"] h3,aside div[class*="from-[#0B192C]"] strong{color:#fff!important}aside div[class*="from-[#0B192C]"] p{color:#d8e5f3!important}aside div[class*="from-[#0B192C]"] a,aside div[class*="from-[#0B192C]"] a span{color:#fff!important}
a[class*="bg-green"],button[class*="bg-green"],a[class*="bg-emerald"],button[class*="bg-emerald"],a[class*="bg-blue"],button[class*="bg-blue"]{color:#fff!important}a[class*="bg-green"] *,button[class*="bg-green"] *,a[class*="bg-emerald"] *,button[class*="bg-emerald"] *{color:inherit!important}
@media(max-width:1023px){main{scroll-margin-top:92px}aside div[class*="from-[#0B192C]"]{margin-bottom:8px}}
</style>`,
  '/ozel-excel-sistemleri/index.html': `
<link id="special-enterprise-css" rel="stylesheet" href="/styles/ozel-excel-enterprise.css" />
<link id="special-editorial-css" rel="stylesheet" href="/styles/ozel-excel-editorial.css" />
<style id="final-ui-contrast-fixes">
.special-v3 section[id]{scroll-margin-top:104px!important}.special-v3 #saha,.special-v3 #karsilastirma,.special-v3 #moduller,.special-v3 #surec,.special-v3 #faq,.special-v3 #iletisim{scroll-margin-top:104px!important}.special-v3 .close-grid{align-items:start!important;grid-template-columns:minmax(0,.88fr) minmax(0,1.12fr)!important}.special-v3 .faq-panel{align-self:start!important;height:auto!important;min-height:0!important}.special-v3 .contact{align-self:start!important}.special-v3 .contact,.special-v3 .contact h1,.special-v3 .contact h2,.special-v3 .contact h3,.special-v3 .contact strong{color:#fff!important}.special-v3 .contact .section-copy,.special-v3 .contact p{color:#dce5ef!important}.special-v3 .contact-points div{color:#eef4fb!important}.special-v3 .btn-primary,.special-v3 .btn-primary *{color:#fff!important}.special-v3 .section{padding-top:60px!important;padding-bottom:60px!important}
body.special-light-v1{--max:1160px;font-family:"Manrope",ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif!important;color-scheme:light!important}.special-light-v1 .wrap{width:min(1160px,calc(100% - 48px))!important}.special-light-v1 .solution-grid{grid-template-columns:repeat(12,1fr)!important}.special-light-v1 .native-info--special{background:#fff!important;color:#10233f!important;border-color:#dbe5ef!important}.special-light-v1 .native-info--special h2{color:#10233f!important}.special-light-v1 .native-info--special .native-info__eyebrow{color:#107c41!important}.special-light-v1 .native-info--special .native-info__lead,.special-light-v1 .native-info--special .native-info__card span,.special-light-v1 .native-info--special .native-info__outcomes span,.special-light-v1 .native-info--special .native-info__footer p{color:#66778e!important}.special-light-v1 .native-info--special .native-info__card,.special-light-v1 .native-info--special .native-info__outcomes article{background:#fff!important;border-color:#dce5ee!important}.special-light-v1 .native-info--special .native-info__card strong,.special-light-v1 .native-info--special .native-info__outcomes strong{color:#10233f!important}.special-light-v1 .native-info--special .native-info__cta{background:#217346!important;color:#fff!important}.special-light-v1 .native-info--special .native-info__cta:hover{background:#185c37!important}
@media(max-width:980px){.special-v3 .close-grid{grid-template-columns:1fr!important;gap:18px!important}}@media(max-width:720px){.special-v3 section[id]{scroll-margin-top:82px!important}.special-v3 .section{padding-top:46px!important;padding-bottom:46px!important}}
</style>`
};

const encodeStandaloneAmpersands = (value) => value.replace(/ & /g, ' &amp; ');
const replaceHtmlVariant = (html, from, to) => {
  if (html.includes(from)) return html.replace(from, to);
  const encodedFrom = encodeStandaloneAmpersands(from);
  if (encodedFrom !== from && html.includes(encodedFrom)) return html.replace(encodedFrom, encodeStandaloneAmpersands(to));
  return html;
};

if (!fs.existsSync(SPECIAL_ENTERPRISE_CSS)) throw new Error('FINAL UI GATE: özel Excel enterprise stylesheet dist içinde bulunamadı');
if (!fs.existsSync(SPECIAL_EDITORIAL_CSS)) throw new Error('FINAL UI GATE: özel Excel editorial stylesheet dist içinde bulunamadı');

let changed = 0;
for (const [relative, css] of Object.entries(fixes)) {
  const file = path.join(DIST, relative.replace(/^\//, ''));
  if (!fs.existsSync(file)) throw new Error(`FINAL UI GATE: hedef HTML bulunamadı: ${relative}`);
  let html = fs.readFileSync(file, 'utf8');
  html = html.replace(/<style id="final-ui-contrast-fixes">[\s\S]*?<\/style>/g, '');
  html = html.replace(/<link id="special-enterprise-css"[^>]*\/>/g, '');
  html = html.replace(/<link id="special-editorial-css"[^>]*\/>/g, '');
  if (relative === '/ozel-excel-sistemleri/index.html') {
    for (const [from, to] of specialContentReplacements) html = replaceHtmlVariant(html, from, to);
  }
  if (!html.includes('</head>')) throw new Error(`FINAL UI GATE: </head> bulunamadı: ${relative}`);
  html = html.replace('</head>', `${css}\n</head>`);
  fs.writeFileSync(file, html);
  changed += 1;
}

const about = fs.readFileSync(path.join(DIST, 'hakkinda/index.html'), 'utf8');
const guide = fs.readFileSync(path.join(DIST, 'rehber/index.html'), 'utf8');
const special = fs.readFileSync(path.join(DIST, 'ozel-excel-sistemleri/index.html'), 'utf8');
const enterpriseCss = fs.readFileSync(SPECIAL_ENTERPRISE_CSS, 'utf8');
const editorialCss = fs.readFileSync(SPECIAL_EDITORIAL_CSS, 'utf8');
const hasEither = (text, raw, encoded) => text.includes(raw) || text.includes(encoded);

if (!about.includes('.hakkinda__hero h1')) throw new Error('FINAL UI GATE: hakkında dark hero kontrast kuralı yok');
if (!guide.includes('from-[#0B192C]')) throw new Error('FINAL UI GATE: rehber dark card kontrast kuralı yok');
if (!special.includes('align-items:start!important')) throw new Error('FINAL UI GATE: özel sistem close-grid düzeltmesi yok');
if (!special.includes('scroll-margin-top:104px')) throw new Error('FINAL UI GATE: sticky anchor düzeltmesi yok');
if (special.includes('data-special-light-v1') && !special.includes('id="special-enterprise-css"')) throw new Error('FINAL UI GATE: enterprise stylesheet link missing');
if (special.includes('data-special-light-v1') && !special.includes('id="special-editorial-css"')) throw new Error('FINAL UI GATE: editorial stylesheet link missing');
if (special.includes('data-special-light-v1') && (!special.includes('--max:1160px') || !special.includes('font-family:"Manrope"') || !special.includes('solution-grid{grid-template-columns:repeat(12,1fr)!important'))) throw new Error('FINAL UI GATE: enterprise light contract missing');
if (special.includes('data-special-light-v1') && !special.includes('.special-light-v1 .native-info--special')) throw new Error('FINAL UI GATE: premium light native infographic override missing');
const financeContentOk = special.includes('Mizan, Nakit, Cari ve Banka Verinizi') && hasEither(special,'13 Haftalık Nakit Akışı & Likidite','13 Haftalık Nakit Akışı &amp; Likidite') && hasEither(special,'Banka Limit-Risk & Faiz Maliyeti','Banka Limit-Risk &amp; Faiz Maliyeti');
if (!financeContentOk) throw new Error('FINAL UI GATE: finance-specific content contract missing');
if (!enterpriseCss.includes('body.special-light-v1') || !enterpriseCss.includes('.special-light-v1 .hero-grid') || !enterpriseCss.includes('.special-light-v1 .solution-card:nth-child(1)') || !enterpriseCss.includes('.special-light-v1 .compare-row .new') || !enterpriseCss.includes('@media(prefers-reduced-motion:reduce)')) throw new Error('FINAL UI GATE: enterprise stylesheet semantic contract incomplete');
if (!editorialCss.includes('AI-free editorial premium layer') || !editorialCss.includes('.special-light-v1 .hero::after{display:none!important}') || !editorialCss.includes('@supports (animation-timeline:view())') || !editorialCss.includes('@media(prefers-reduced-motion:reduce)')) throw new Error('FINAL UI GATE: editorial AI-free design contract incomplete');

console.log(`FINAL UI CONTRAST/LAYOUT GATE PASS — ${changed} sayfa düzeltildi; special=ai-free-enterprise-editorial`);
