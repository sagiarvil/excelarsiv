#!/usr/bin/env node
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const routes = {
  home: join('dist', 'index.html'),
  guides: join('dist', 'rehber', 'index.html'),
  special: join('dist', 'ozel-excel-sistemleri', 'index.html'),
  contact: join('dist', 'iletisim', 'index.html'),
  catalog: join('dist', 'sablonlar', 'index.html'),
};
const visual = join('public', 'images', 'site', 'excel-guide-banner.webp');
const realBrandLogo = '/images/brand/excelarsiv-header-logo.png';

function hash(path) { return createHash('sha256').update(readFileSync(path)).digest('hex'); }
function must(path) { if (!existsSync(path)) throw new Error(`PREMIUM UX GATE: missing ${path}`); }
function once(html, from, to, label) {
  const count = html.split(from).length - 1;
  if (count !== 1) throw new Error(`PREMIUM UX GATE: expected 1 ${label}, found ${count}`);
  return html.replace(from, to);
}
Object.values(routes).forEach(must); must(visual);
if (statSync(visual).size < 10000) throw new Error('PREMIUM UX GATE: contextual image is suspiciously small');
const catalogBefore = hash(routes.catalog);

// Until the three contextual visuals are separately re-mastered, use the verified real Excel artwork everywhere instead of blank placeholders.
for (const route of [routes.home, routes.special]) {
  let html = readFileSync(route, 'utf8');
  html = html.replaceAll('/images/site/excel-analytics-section.webp', '/images/site/excel-guide-banner.webp');
  html = html.replaceAll('/images/site/excel-special-systems-hero.webp', '/images/site/excel-guide-banner.webp');
  writeFileSync(route, html, 'utf8');
}

// /ozel-excel-sistemleri is intentionally a standalone premium surface. Keep its brand identity identical to the canonical site header.
const specialBrandStyles = `
<style data-special-brand-identity>
  .special-v3 .brand{min-width:190px;display:flex!important;align-items:center!important}
  .special-v3 .brand-logo{display:block;width:190px;max-width:100%;height:auto;max-height:48px;object-fit:contain;object-position:left center}
  .special-v3 footer .footer-logo-block{display:flex;align-items:center;gap:14px;min-width:260px}
  .special-v3 footer .footer-logo{display:block;width:178px;max-width:100%;height:auto;max-height:50px;object-fit:contain;object-position:left center;filter:brightness(0) invert(1)}
  .special-v3 footer .footer-logo-copy{color:#9fb0c4;line-height:1.5}
  @media(max-width:720px){
    .special-v3 .brand{min-width:150px}
    .special-v3 .brand-logo{width:154px;max-height:42px}
    .special-v3 footer .footer-logo-block{min-width:0;align-items:flex-start;flex-direction:column;gap:8px}
    .special-v3 footer .footer-logo{width:158px;max-height:46px}
  }
</style>`;
let special = readFileSync(routes.special, 'utf8');
special = once(special, 'href="/favicon.png"', `href="${realBrandLogo}"`, 'special png favicon href');
special = once(special, 'type="image/svg+xml" href="/favicon.svg"', `type="image/png" href="${realBrandLogo}"`, 'special svg favicon identity');
special = once(special, 'href="/apple-touch-icon.png"', `href="${realBrandLogo}"`, 'special apple touch identity');
special = once(
  special,
  '<a class="brand" href="/" aria-label="Excel Arşiv ana sayfa"><span class="brand-mark">EA</span><span>excelarsiv.com<small>Finans & Yönetim Mimarisi</small></span></a>',
  `<a class="brand" href="/" aria-label="Excel Arşiv ana sayfa"><img class="brand-logo" src="${realBrandLogo}" alt="Excel Arşiv" width="420" height="120" decoding="async" /></a>`,
  'special real header logo',
);
special = once(
  special,
  '<footer><div class="wrap footer-inner"><div><span class="footer-brand">excelarsiv.com</span><br />Bilanço Odaklı Finansal Modelleme ve Yönetim Mimarisi.</div>',
  `<footer><div class="wrap footer-inner"><div class="footer-logo-block"><img class="footer-logo" src="${realBrandLogo}" alt="Excel Arşiv" width="420" height="120" loading="lazy" decoding="async" /><span class="footer-logo-copy">Bilanço Odaklı Finansal Modelleme ve Yönetim Mimarisi.</span></div>`,
  'special footer logo',
);
special = once(special, '</head>', `${specialBrandStyles}</head>`, 'special brand styles');
writeFileSync(routes.special, special, 'utf8');

const guideStyles = `
<style data-rehber-premium-ux>
  .rehber-premium .hub-kicker{font-size:13px!important;letter-spacing:.12em!important;margin-bottom:18px!important}
  .rehber-premium .cluster-nav{gap:14px!important}
  .rehber-premium .cluster-nav a{min-height:118px!important;padding:20px 22px!important;border-radius:16px!important;box-shadow:0 8px 24px rgba(22,48,36,.04)}
  .rehber-premium .cluster-nav strong{font-size:16px!important;line-height:1.3!important;letter-spacing:-.015em!important}
  .rehber-premium .cluster-nav span{font-size:13px!important;line-height:1.35!important;margin-top:18px!important;color:#708077!important}
  .rehber-premium .hub-stats span{font-size:12px!important;line-height:1.4!important}
  .rehber-premium .cluster-head p,.rehber-premium .catalog-cta>div>p{font-size:11px!important}
  .rehber-premium .cluster-head div>div{font-size:15px!important;line-height:1.7!important}
  .rehber-premium .cluster-head>a{font-size:13px!important}
  .rehber-premium .guide-card>span{font-size:10px!important}
  .rehber-premium .guide-card p{font-size:14px!important;line-height:1.68!important}
  .rehber-premium .guide-card__link{font-size:13px!important}
  @media(max-width:760px){.rehber-premium .cluster-nav a{min-height:104px!important}.rehber-premium .cluster-nav strong{font-size:15px!important}.rehber-premium .cluster-nav span{font-size:12.5px!important}}
</style>`;
let guides = readFileSync(routes.guides, 'utf8');
guides = once(guides, '<body class="', '<body class="rehber-premium ', 'rehber body namespace');
guides = once(guides, '</head>', `${guideStyles}</head>`, 'rehber head');
writeFileSync(routes.guides, guides, 'utf8');

const contactStyles = `
<style data-contact-premium-ux>
  .contact-premium main{background:linear-gradient(180deg,#f7faf8 0,#fff 42%)}
  .contact-premium main>section:first-child{background:radial-gradient(circle at 85% 20%,rgba(22,125,72,.09),transparent 28%),linear-gradient(180deg,#f7faf8,#fff)!important}
  .contact-premium main>section:first-child .container-site{padding-top:72px!important;padding-bottom:66px!important}
  .contact-premium main>section:first-child .display{max-width:760px;font-size:clamp(2.8rem,5.3vw,5rem)!important;line-height:1.02!important;letter-spacing:-.052em!important}
  .contact-premium main>section:first-child .lead{max-width:760px!important;font-size:18px!important;line-height:1.75!important;color:#53635a!important}
  .contact-premium main>section:nth-child(2) .container-site{padding-top:54px!important;padding-bottom:76px!important}
  .contact-premium .grid-12{display:grid!important;grid-template-columns:repeat(6,minmax(0,1fr))!important;gap:18px!important}
  .contact-premium .grid-12>div{grid-column:span 2!important;min-height:245px!important;padding:28px!important;border:1px solid #dce5df!important;border-radius:22px!important;background:#fff!important;box-shadow:0 14px 40px rgba(24,48,36,.055)!important}
  .contact-premium .grid-12>div:nth-child(4),.contact-premium .grid-12>div:nth-child(5){grid-column:span 3!important}
  .contact-premium .grid-12>div:before{content:'';display:block;width:42px;height:4px;border-radius:999px;background:#13824a;margin-bottom:26px}
  .contact-premium .grid-12 .h3{font-size:28px!important;line-height:1.15!important;letter-spacing:-.035em!important;color:#183024!important}
  .contact-premium .grid-12 .body-sm{font-size:15px!important;line-height:1.7!important;color:#66746b!important}
  .contact-premium .grid-12 a.font-mono{font-size:15px!important;color:#0b7440!important}
  .contact-premium .grid-12 .mt-6{margin-top:24px!important}
  .contact-premium .grid-12 a[class*='button'],.contact-premium .grid-12 .btn{min-height:46px}
  @media(max-width:960px){.contact-premium .grid-12{grid-template-columns:repeat(2,minmax(0,1fr))!important}.contact-premium .grid-12>div,.contact-premium .grid-12>div:nth-child(4),.contact-premium .grid-12>div:nth-child(5){grid-column:span 1!important}}
  @media(max-width:640px){.contact-premium main>section:first-child .container-site{padding-top:46px!important;padding-bottom:42px!important}.contact-premium .grid-12{grid-template-columns:1fr!important}.contact-premium .grid-12>div{min-height:0!important;padding:22px!important}.contact-premium .grid-12 .h3{font-size:24px!important}}
</style>`;
let contact = readFileSync(routes.contact, 'utf8');
contact = once(contact, '<body class="', '<body class="contact-premium ', 'contact body namespace');
contact = once(contact, '</head>', `${contactStyles}</head>`, 'contact head');
writeFileSync(routes.contact, contact, 'utf8');

const catalogAfter = hash(routes.catalog);
if (catalogBefore !== catalogAfter) throw new Error('PREMIUM UX GATE: /sablonlar changed');

for (const [name, path, token] of [
  ['rehber', routes.guides, 'data-rehber-premium-ux'],
  ['contact', routes.contact, 'data-contact-premium-ux'],
  ['special-brand', routes.special, 'data-special-brand-identity'],
]) {
  const html = readFileSync(path,'utf8');
  if (!html.includes(token)) throw new Error(`PREMIUM UX GATE: ${name} namespace missing`);
}
const specialFinal = readFileSync(routes.special, 'utf8');
if (specialFinal.includes('<span class="brand-mark">EA</span>')) throw new Error('PREMIUM UX GATE: placeholder EA logo still present');
if ((specialFinal.split(realBrandLogo).length - 1) < 5) throw new Error('PREMIUM UX GATE: real brand logo is not wired to favicon + apple touch + header + footer');
console.log('PREMIUM UX GATE PASS — real Excel visual restored, real brand identity applied to special systems, rehber typography upgraded, contact redesigned, /sablonlar byte-identical');