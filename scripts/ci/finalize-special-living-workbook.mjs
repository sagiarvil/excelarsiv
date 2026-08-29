#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(file)) throw new Error('LIVING WORKBOOK FINALIZER: special route build output missing');

let html = fs.readFileSync(file, 'utf8');
if (!html.includes('class="workbook"')) {
  console.log('LIVING WORKBOOK FINALIZER SKIP — legacy special page detected.');
  process.exit(0);
}

const stripStyleById = (id) => {
  const pattern = new RegExp(`<style\\b(?=[^>]*\\bid=["']${id}["'])[^>]*>[\\s\\S]*?<\\/style>`, 'gi');
  html = html.replace(pattern, '');
};
const stripLinkById = (id) => {
  const pattern = new RegExp(`<link\\b(?=[^>]*\\bid=["']${id}["'])[^>]*>`, 'gi');
  html = html.replace(pattern, '');
};
const stripScriptById = (id) => {
  const pattern = new RegExp(`<script\\b(?=[^>]*\\bid=["']${id}["'])[^>]*>[\\s\\S]*?<\\/script>`, 'gi');
  html = html.replace(pattern, '');
};

// Remove build-time compatibility layers. They are retained for legacy pages, but
// the Living Workbook source is the SSOT for this route and must own final pixels.
html = html.replace(/<style\b[^>]*data-contextual-imagery[^>]*>[\s\S]*?<\/style>/gi, '');
html = html.replace(/<style\b[^>]*data-special-brand-identity[^>]*>[\s\S]*?<\/style>/gi, '');
html = html.replace(/<style\b[^>]*data-ea-special-light-typography[^>]*>[\s\S]*?<\/style>/gi, '');
html = html.replace(/<style\b[^>]*data-exact-anchor-scroll[^>]*>[\s\S]*?<\/style>/gi, '');
html = html.replace(/<script\b[^>]*data-exact-anchor-scroll[^>]*>[\s\S]*?<\/script>/gi, '');

for (const id of [
  'final-ui-contrast-fixes',
  'special-innovation-css',
  'special-brand-sync-css',
  'special-layout-stabilizer-css',
  'native-seo-infographics',
]) stripStyleById(id);
for (const id of ['special-enterprise-css', 'special-editorial-css', 'special-innovation-css', 'special-brand-sync-css']) stripLinkById(id);
stripScriptById('special-innovation-js');

// Decision Lab is injected only to satisfy the legacy build chain. The new page has
// its own workbook decision console and comparison matrix, so duplicate UI is removed.
html = html.replace(/<section\b(?=[^>]*\bid=["']karar-laboratuvari["'])[^>]*>[\s\S]*?<\/section>/i, '');
html = html.replace(/<section\b(?=[^>]*\bdata-native-info=["']special-decision-map["'])[^>]*>[\s\S]*?<\/section>/i, '');

// After the injected decision lab is gone, the compatibility bridge has no nested
// divs and can be safely removed as a whole. Hidden legacy SEO/copy tokens must never
// ship in the final public HTML.
html = html.replace(/<div\b(?=[^>]*\bdata-special-light-legacy-bridge\b)[^>]*>[\s\S]*?<\/div>/i, '');

const headerMarkup = `<header class="site-nav">
    <div class="wrap-wide nav-inner">
      <a class="brand" href="/" aria-label="Excel Arşiv ana sayfa">
        <img src="/images/excel-logo.png" alt="" width="34" height="34" />
        <span class="brand-copy"><strong>EXCELARŞİV</strong><small>ÖZEL EXCEL SİSTEMLERİ</small></span>
      </a>
      <nav class="nav-links" aria-label="Ana menü">
        <a href="#ihtiyaclar">İhtiyaçlar</a><a href="#karsilastirma">Farkımız</a><a href="#mimariler">Sistemler</a><a href="#surec">Süreç</a><a href="#sss">SSS</a>
      </nav>
      <div class="nav-actions">
        <a class="btn btn-secondary nav-cta" href="/sablonlar">Hazır Sistemler</a>
        <a class="btn btn-primary nav-cta" href="/iletisim">İhtiyacınızı Anlatın</a>
        <details class="mobile-menu">
          <summary aria-label="Menüyü aç"><span class="burger" aria-hidden="true"></span></summary>
          <nav class="mobile-panel" aria-label="Mobil menü">
            <a href="#ihtiyaclar">İhtiyaçlar</a><a href="#karsilastirma">Farkımız</a><a href="#mimariler">Sistemler</a><a href="#surec">Süreç</a><a href="#sss">SSS</a><a href="/sablonlar">Hazır Sistemler</a><a href="/iletisim">İhtiyacınızı Anlatın</a>
          </nav>
        </details>
      </div>
    </div>
  </header>`;
const headerPattern = /<header\b[^>]*>[\s\S]*?<\/header>/i;
if (!headerPattern.test(html)) throw new Error('LIVING WORKBOOK FINALIZER: header shell missing');
html = html.replace(headerPattern, headerMarkup);

// Restore the source-owned SEO and icon contract after legacy brand/CTA mutations.
html = html.replace(/<title>[\s\S]*?<\/title>/i, '<title>Özel Excel Finans & Raporlama Sistemleri | ExcelArşiv</title>');
html = html.replace(/<meta\s+name="description"\s+content="[^"]*"\s*\/?>/i, '<meta name="description" content="İşleyişinize göre kurulan özel Excel finans, raporlama ve yönetim sistemleri. Nakit akışı, tahsilat, banka limitleri, kârlılık, muhasebe ve yönetim raporlarını tek kontrol mimarisinde birleştirin." />');
html = html.replace(/<link\s+rel="icon"\s+type="image\/png"\s+sizes="32x32"\s+href="[^"]*"\s*\/?>/i, '<link rel="icon" type="image/png" sizes="32x32" href="/favicon.png" />');
html = html.replace(/<link\s+rel="icon"\s+type="image\/png"\s+href="[^"]*"\s*\/?>/i, '<link rel="icon" type="image/svg+xml" href="/favicon.svg" />');
html = html.replace(/<link\s+rel="apple-touch-icon"\s+href="[^"]*"\s*\/?>/i, '<link rel="apple-touch-icon" href="/apple-touch-icon.png" />');

// Two legacy semantic-copy replacements collide with the new workbook demo values.
// Restore them deterministically inside their KPI labels.
html = html.replace(/(<small>TAHSİLAT<\/small><strong>)₺ 620\.000(<\/strong>)/, '$1% 92,5$2');
html = html.replace(/(<small>NET KÂR<\/small><strong>)₺ 3\.250\.000(<\/strong>)/, '$1₺ 1.125.000$2');

// Normalize body namespace and keep the site-wide typography/conformance token while
// binding it to the canonical Manrope Living Workbook typeface.
html = html.replace(/<body\b[^>]*>/i, '<body data-living-workbook-v1>');
const finalContract = `
<style data-ea-typography-mandate="chat-readable-v2" data-ea-special-light-typography data-living-workbook-final-contract>
  :root{--ea-font-sans:Manrope,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}
  body[data-living-workbook-v1]{font-family:var(--ea-font-sans)}
  body[data-living-workbook-v1] [id]{scroll-margin-top:84px}
</style>`;
if (!html.includes('</head>')) throw new Error('LIVING WORKBOOK FINALIZER: </head> missing');
html = html.replace('</head>', `${finalContract}\n</head>`);

const required = [
  'data-living-workbook-v1',
  'data-living-workbook-final-contract',
  'class="workbook"',
  'VERİ → HESAP → KONTROL → KARAR',
  'Temsili sistem görünümü',
  'Hata nerede?',
  'Özel Excel Finans & Raporlama Sistemleri | ExcelArşiv',
  '/images/excel-logo.png',
];
for (const token of required) {
  if (!html.includes(token)) throw new Error(`LIVING WORKBOOK FINALIZER: required token missing: ${token}`);
}

const forbidden = [
  'data-special-light-legacy-bridge',
  'id="karar-laboratuvari"',
  'data-native-info="special-decision-map"',
  'id="special-enterprise-css"',
  'id="special-editorial-css"',
  'id="special-innovation-css"',
  'id="special-brand-sync-css"',
  'id="special-layout-stabilizer-css"',
  'id="native-seo-infographics"',
  'id="special-innovation-js"',
  '/images/brand/excelarsiv-header-logo.png',
];
for (const token of forbidden) {
  if (html.includes(token)) throw new Error(`LIVING WORKBOOK FINALIZER: legacy token survived: ${token}`);
}

fs.writeFileSync(file, html);
console.log('LIVING WORKBOOK FINALIZER PASS — source-owned light design restored after legacy compatibility gates; duplicate legacy UI/CSS removed.');
