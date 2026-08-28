import fs from 'node:fs';
import path from 'node:path';

const DIST = path.resolve('dist');

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
<style id="final-ui-contrast-fixes">
/* Legacy special-v3 safeguards remain namespaced and cannot affect premium-light-v1. */
.special-v3 section[id]{scroll-margin-top:104px!important}.special-v3 #saha,.special-v3 #karsilastirma,.special-v3 #moduller,.special-v3 #surec,.special-v3 #faq,.special-v3 #iletisim{scroll-margin-top:104px!important}.special-v3 .close-grid{align-items:start!important;grid-template-columns:minmax(0,.88fr) minmax(0,1.12fr)!important}.special-v3 .faq-panel{align-self:start!important;height:auto!important;min-height:0!important}.special-v3 .contact{align-self:start!important}.special-v3 .contact,.special-v3 .contact h1,.special-v3 .contact h2,.special-v3 .contact h3,.special-v3 .contact strong{color:#fff!important}.special-v3 .contact .section-copy,.special-v3 .contact p{color:#dce5ef!important}.special-v3 .contact-points div{color:#eef4fb!important}.special-v3 .btn-primary,.special-v3 .btn-primary *{color:#fff!important}.special-v3 .section{padding-top:60px!important;padding-bottom:60px!important}
/* Premium light contract: later SEO/native injectors may add semantic blocks, but they may not reintroduce a dark page section. */
.special-light-v1 .native-info--special{background:linear-gradient(180deg,#ffffff,#f8fbff)!important;color:#0b1f41!important;border-color:#dfe7f2!important;box-shadow:0 18px 52px rgba(26,55,101,.08)!important}.special-light-v1 .native-info--special h2{color:#0b1f41!important}.special-light-v1 .native-info--special .native-info__eyebrow{color:#168339!important}.special-light-v1 .native-info--special .native-info__lead{color:#66758f!important}.special-light-v1 .native-info--special .native-info__card{background:#fff!important;border-color:#dfe7f2!important}.special-light-v1 .native-info--special .native-info__card strong{color:#0b1f41!important}.special-light-v1 .native-info--special .native-info__card span{color:#66758f!important}.special-light-v1 .native-info--special .native-info__card small{color:#168339!important}.special-light-v1 .native-info--special .native-info__arrow{color:#1168e8!important}.special-light-v1 .native-info--special .native-info__outcomes article{background:#fff!important;border-color:#dfe7f2!important}.special-light-v1 .native-info--special .native-info__outcomes strong{color:#0b1f41!important}.special-light-v1 .native-info--special .native-info__outcomes span{color:#66758f!important}.special-light-v1 .native-info--special .native-info__footer{border-color:#dfe7f2!important}.special-light-v1 .native-info--special .native-info__footer p{color:#66758f!important}.special-light-v1 .native-info--special .native-info__cta{background:#21a64a!important;color:#fff!important}.special-light-v1 .native-info--special .native-info__cta:hover{background:#168339!important}
@media(max-width:980px){.special-v3 .close-grid{grid-template-columns:1fr!important;gap:18px!important}}@media(max-width:720px){.special-v3 section[id]{scroll-margin-top:82px!important}.special-v3 .section{padding-top:46px!important;padding-bottom:46px!important}}
</style>`
};

let changed = 0;
for (const [relative, css] of Object.entries(fixes)) {
  const file = path.join(DIST, relative.replace(/^\//, ''));
  if (!fs.existsSync(file)) throw new Error(`FINAL UI GATE: hedef HTML bulunamadı: ${relative}`);
  let html = fs.readFileSync(file, 'utf8');
  html = html.replace(/<style id="final-ui-contrast-fixes">[\s\S]*?<\/style>/g, '');
  if (!html.includes('</head>')) throw new Error(`FINAL UI GATE: </head> bulunamadı: ${relative}`);
  html = html.replace('</head>', `${css}\n</head>`);
  fs.writeFileSync(file, html);
  changed += 1;
}

const about = fs.readFileSync(path.join(DIST, 'hakkinda/index.html'), 'utf8');
const guide = fs.readFileSync(path.join(DIST, 'rehber/index.html'), 'utf8');
const special = fs.readFileSync(path.join(DIST, 'ozel-excel-sistemleri/index.html'), 'utf8');
if (!about.includes('.hakkinda__hero h1')) throw new Error('FINAL UI GATE: hakkında dark hero kontrast kuralı yok');
if (!guide.includes('from-[#0B192C]')) throw new Error('FINAL UI GATE: rehber dark card kontrast kuralı yok');
if (!special.includes('align-items:start!important')) throw new Error('FINAL UI GATE: özel sistem close-grid düzeltmesi yok');
if (!special.includes('scroll-margin-top:104px')) throw new Error('FINAL UI GATE: sticky anchor düzeltmesi yok');
if (special.includes('data-special-light-v1') && !special.includes('.special-light-v1 .native-info--special')) throw new Error('FINAL UI GATE: premium light native infographic override missing');

console.log(`FINAL UI CONTRAST/LAYOUT GATE PASS — ${changed} sayfa düzeltildi`);
