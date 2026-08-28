import fs from 'node:fs';
import path from 'node:path';

const DIST = path.resolve('dist');
const SPECIAL_ENTERPRISE_CSS = path.join(DIST, 'styles', 'ozel-excel-enterprise.css');
const SPECIAL_EDITORIAL_CSS = path.join(DIST, 'styles', 'ozel-excel-editorial.css');

const encodeStandaloneAmpersands = (value) => value.replace(/ & /g, ' &amp; ');
const replaceRouteText = (html, from, to) => {
  let next = html.replaceAll(from, to);
  const encodedFrom = encodeStandaloneAmpersands(from);
  if (encodedFrom !== from) next = next.replaceAll(encodedFrom, encodeStandaloneAmpersands(to));
  return next;
};

/*
 * Route-only semantic copy contract.
 * We intentionally replace stable text tokens instead of entire HTML fragments so
 * Astro serialization, whitespace and &amp; encoding cannot break the finance copy.
 */
const specialTextReplacements = [
  ['EXCLUSIVE  •  GÜVENİLİR  •  ÖLÇÜLEBİLİR', 'FİNANS  •  MUHASEBE  •  YÖNETİM KONTROLÜ'],
  ['EXCLUSIVE&nbsp;&nbsp;•&nbsp;&nbsp;GÜVENİLİR&nbsp;&nbsp;•&nbsp;&nbsp;ÖLÇÜLEBİLİR', 'FİNANS&nbsp;&nbsp;•&nbsp;&nbsp;MUHASEBE&nbsp;&nbsp;•&nbsp;&nbsp;YÖNETİM KONTROLÜ'],
  ['Excel ile Sınırlarınızı Aşın,', 'Mizan, Nakit, Cari ve Banka Verinizi,'],
  ['Gerçek İş Sonuçları Alın.', 'Tek Bir Karar Sisteminde Birleştirin.'],
  ['Muhasebe, finans, satış, stok, üretim ve yönetim süreçleriniz için ', 'Mizan, muavin, nakit akışı, cari yaşlandırma, çek/senet, banka limit-risk ve yönetim raporlamasını '],
  ['işleyişinize göre özel tasarlanan Excel sistemleri', 'işletmenizin gerçek çalışma düzenine göre tek Excel mimarisinde'],
  [' geliştiriyoruz. Amaç daha fazla hücre değil; daha görünür nakit, daha hızlı kontrol ve daha güvenilir yönetim kararı.', ' birleştiriyoruz. Amaç daha fazla hücre değil; kontrol edilebilir finans verisi ve daha hızlı yönetim kararı.'],
  ['%100 Size Özel', 'İş Akışına Özel'],
  ['İş akışınıza göre kurulan Excel sistemleri', 'Hazır şablon değil, sürecinize göre kurulan sistem'],
  ['Satışlar', '13 Haftalık Nakit'],
  ['Stok', '120/320 Kontrol'],
  ['Raporlar', 'Banka Limit'],
  ['Analiz', 'Valör & Faiz'],
  ['Toplam Ciro', '13 Haftalık Net Nakit'],
  ['Brüt Kâr', 'Açık Tahsilat'],
  ['Net Kâr', 'Banka Limit Boşluğu'],
  ['Tahsilat Oranı', 'Kontrol Bekleyen Cari'],
  ['₺ 8.750.000', '₺ 2.375.000'],
  ['₺ 2.125.000', '₺ 1.480.000'],
  ['₺ 1.125.000', '₺ 3.250.000'],
  ['% 92,5', '₺ 620.000'],
  ['▲ %18,6', '13 hafta'],
  ['▲ %22,4', 'yaşlandırılmış'],
  ['▲ %16,8', 'kullanılabilir'],
  ['▲ %5,3', 'istisna'],
  ['Aylık Ciro Trendi', '13 Haftalık Nakit Eğrisi'],
  ['Ciro Dağılımı', 'Banka Limit Kullanımı'],
  ['Yurt İçi', 'Kullanılan'],
  ['Yurt Dışı', 'Boş Limit'],
  ['Diğer', 'Risk / Bloke'],
  ['En Çok Satış Yapılan', 'Kontrol Alanı'],
  ['Miktar', 'Adet'],
  ['Ürün A', '120 Alıcılar'],
  ['Ürün B', '320 Satıcılar'],
  ['Ürün C', 'Çek / Senet'],
  ['GÜVENİLEN', 'FİNANSAL'],
  ['ÇÖZÜM ORTAĞINIZ', 'KONTROL ODAKLARI'],
  ["KOBİ'den mali müşavirlik ofisine, finans ekibinden yönetime kadar iş akışına göre yapılandırılmış Excel sistemleri.", 'Mizan, nakit, cari, banka, valör ve yönetim raporlamasını aynı kontrol omurgasında birleştiren sistemler.'],
  ['İşinizi Büyüten Excel Çözümleri', 'Finansal Kararı Taşıyan Excel Sistem Mimarileri'],
  ['Finans & Muhasebe', '13 Haftalık Nakit Akışı & Likidite'],
  ['Bilanço, nakit akışı, bütçe, mizan kontrolü ve finansal analiz sistemleri.', 'Günlük/haftalık nakit görünümü, ödeme-tahsilat ufku, erken uyarı ve stres senaryoları.'],
  ['Satış & Tahsilat', 'Çek / Senet Portföyü & Valör'],
  ['Teklif, sipariş, müşteri takibi, yaşlandırma ve tahsilat görünürlüğü.', 'Vade, ortalama valör, portföy yoğunluğu, tahsilat takvimi ve finansman etkisi.'],
  ['Stok & Depo Yönetimi', 'Luca / Logo / Zirve Mizan Entegrasyonu'],
  ['Stok hareketi, kritik seviye, ürün kârlılığı ve envanter analizi.', 'Muavin-mizan aktarımı, 120/320 ters bakiye, kapanış ve istisna kontrolleri.'],
  ['Üretim & Planlama', 'Yönetim & İcra Kurulu Dashboard'],
  ['Üretim planı, kapasite, maliyet ve operasyonel performans görünürlüğü.', 'EBITDA, net işletme sermayesi, tahsilat, likidite ve yönetim KPI görünümü.'],
  ['Raporlama & Dashboard', 'Banka Limit-Risk & Faiz Maliyeti'],
  ['Yönetici panoları, KPI takibi, senaryo analizi ve karar destek raporları.', 'Limit doluluk, risk, rotatif faiz tahakkuku, DSCR ve döviz pozisyonu izleme.'],
  ['Özel Çözümler', 'Cari Yaşlandırma & 120/320 Kontrol'],
  ['İş akışınıza, veri yapınıza ve kullanıcı rollerinize özel Excel uygulamaları.', 'Vade kovaları, ters bakiye, mahsup ve kapanmayan cari istisnalarını görünür kılan kontrol sistemi.']
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
a[class*="bg-green"],button[class*="bg-green"],a[class*="bg-emerald"],button[class*="bg-emerald"],a[class*="bg-blue"],button[class*="bg-blue"]{color:#fff!important}a[class*="bg-green"] *,button[class*="bg-green"] *,a[class*="bg-emerald"] *,button[class*="bg-emerald"] *,a[class*="bg-blue"] *,button[class*="bg-blue"] *{color:inherit!important}
@media(max-width:1023px){main{scroll-margin-top:92px}aside div[class*="from-[#0B192C]"]{margin-bottom:8px}}
</style>`,

  '/ozel-excel-sistemleri/index.html': `
<link id="special-enterprise-css" rel="stylesheet" href="/styles/ozel-excel-enterprise.css" />
<link id="special-editorial-css" rel="stylesheet" href="/styles/ozel-excel-editorial.css" />
<style id="final-ui-contrast-fixes">
/* Legacy special-v3 safeguards stay namespaced and cannot affect premium-light-v1. */
.special-v3 section[id]{scroll-margin-top:104px!important}.special-v3 #saha,.special-v3 #karsilastirma,.special-v3 #moduller,.special-v3 #surec,.special-v3 #faq,.special-v3 #iletisim{scroll-margin-top:104px!important}.special-v3 .close-grid{align-items:start!important;grid-template-columns:minmax(0,.88fr) minmax(0,1.12fr)!important}.special-v3 .faq-panel{align-self:start!important;height:auto!important;min-height:0!important}.special-v3 .contact{align-self:start!important}.special-v3 .contact,.special-v3 .contact h1,.special-v3 .contact h2,.special-v3 .contact h3,.special-v3 .contact strong{color:#fff!important}.special-v3 .contact .section-copy,.special-v3 .contact p{color:#dce5ef!important}.special-v3 .contact-points div{color:#eef4fb!important}.special-v3 .btn-primary,.special-v3 .btn-primary *{color:#fff!important}.special-v3 .section{padding-top:60px!important;padding-bottom:60px!important}
body.special-light-v1{--max:1160px;font-family:"Manrope",ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif!important;color-scheme:light!important}.special-light-v1 .wrap{width:min(1160px,calc(100% - 48px))!important}.special-light-v1 .solution-grid{grid-template-columns:repeat(12,1fr)!important}.special-light-v1 .native-info--special{background:#fff!important;color:#10233f!important;border-color:#dbe5ef!important}.special-light-v1 .native-info--special h2{color:#10233f!important}.special-light-v1 .native-info--special .native-info__eyebrow{color:#107c41!important}.special-light-v1 .native-info--special .native-info__lead,.special-light-v1 .native-info--special .native-info__card span,.special-light-v1 .native-info--special .native-info__outcomes span,.special-light-v1 .native-info--special .native-info__footer p{color:#66778e!important}.special-light-v1 .native-info--special .native-info__card,.special-light-v1 .native-info--special .native-info__outcomes article{background:#fff!important;border-color:#dce5ee!important}.special-light-v1 .native-info--special .native-info__card strong,.special-light-v1 .native-info--special .native-info__outcomes strong{color:#10233f!important}.special-light-v1 .native-info--special .native-info__cta{background:#217346!important;color:#fff!important}.special-light-v1 .native-info--special .native-info__cta:hover{background:#185c37!important}
@media(max-width:980px){.special-v3 .close-grid{grid-template-columns:1fr!important;gap:18px!important}}@media(max-width:720px){.special-v3 section[id]{scroll-margin-top:82px!important}.special-v3 .section{padding-top:46px!important;padding-bottom:46px!important}}
</style>`
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
    for (const [from, to] of specialTextReplacements) html = replaceRouteText(html, from, to);
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

const financeContentOk =
  special.includes('Mizan, Nakit, Cari ve Banka Verinizi') &&
  hasEither(special, '13 Haftalık Nakit Akışı & Likidite', '13 Haftalık Nakit Akışı &amp; Likidite') &&
  hasEither(special, 'Banka Limit-Risk & Faiz Maliyeti', 'Banka Limit-Risk &amp; Faiz Maliyeti') &&
  special.includes('120/320 Kontrol') &&
  special.includes('Banka Limit Kullanımı');
if (!financeContentOk) throw new Error('FINAL UI GATE: finance-specific content contract missing');

if (!enterpriseCss.includes('body.special-light-v1') || !enterpriseCss.includes('.special-light-v1 .hero-grid') || !enterpriseCss.includes('.special-light-v1 .solution-card:nth-child(1)') || !enterpriseCss.includes('.special-light-v1 .compare-row .new') || !enterpriseCss.includes('@media(prefers-reduced-motion:reduce)')) throw new Error('FINAL UI GATE: enterprise stylesheet semantic contract incomplete');
if (!editorialCss.includes('AI-free editorial premium layer') || !editorialCss.includes('.special-light-v1 .hero::after{display:none!important}') || !editorialCss.includes('@supports (animation-timeline:view())') || !editorialCss.includes('@media(prefers-reduced-motion:reduce)')) throw new Error('FINAL UI GATE: editorial AI-free design contract incomplete');

console.log(`FINAL UI CONTRAST/LAYOUT GATE PASS — ${changed} sayfa düzeltildi; special=ai-free-enterprise-editorial`);
