#!/usr/bin/env node
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync, cpSync, rmSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';

const root = process.cwd();
const premiumSiteDir = resolve(root, 'excelarsiv-premium-tum-sayfalar/site');
const distDir = resolve(root, 'dist');

if (!existsSync(premiumSiteDir)) {
  console.error('HATA: excelarsiv-premium-tum-sayfalar/site bulunamadı!');
  process.exit(1);
}

console.log('=== PREMİUM TASARIM ENTEGRASYONU BAŞLATILIYOR ===');

// 1. Copy assets to dist/assets
const premiumAssetsDir = join(premiumSiteDir, 'assets');
const distAssetsDir = join(distDir, 'assets');
if (existsSync(premiumAssetsDir)) {
  cpSync(premiumAssetsDir, distAssetsDir, { recursive: true });
  console.log('[PASS] assets klasörü dist/assets içine kopyalandı.');
}

// 2. Read sitemap to get canonical URLs
const pagesXml = existsSync(join(distDir, 'sitemap-pages.xml'))
  ? readFileSync(join(distDir, 'sitemap-pages.xml'), 'utf8')
  : readFileSync(join(root, 'public/sitemap-pages.xml'), 'utf8');
const prodsXml = existsSync(join(distDir, 'sitemap-products.xml'))
  ? readFileSync(join(distDir, 'sitemap-products.xml'), 'utf8')
  : readFileSync(join(root, 'public/sitemap-products.xml'), 'utf8');
const sitemap = pagesXml + prodsXml;
const locRegex = /<loc>https:\/\/excelarsiv\.com([^<]*)<\/loc>/g;
const validCanonicalRoutes = new Set();
let m;
while ((m = locRegex.exec(sitemap)) !== null) {
  validCanonicalRoutes.add(m[1] === '' ? '/' : m[1]);
}

const defaultRobotsMeta = '<meta name="robots" content="index, follow, max-image-preview:large, max-snippet:-1, max-video-preview:-1">';

function getJsonLd(html) {
  const match = html.match(/<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/i);
  return match ? match[0] : '';
}

function getFallbackJsonLd(url, title, description) {
  return `<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "WebPage",
  "name": ${JSON.stringify(title || "ExcelArşiv")},
  "description": ${JSON.stringify(description || "")},
  "url": ${JSON.stringify(url)},
  "isPartOf": {
    "@type": "WebSite",
    "name": "ExcelArşiv",
    "url": "https://excelarsiv.com/"
  }
}
</script>`;
}

const llmsMetaTags = `  <link rel="describedby" href="https://excelarsiv.com/llms.txt" type="text/markdown" title="ExcelArsiv.com AI Index">\n  <link rel="alternate" type="text/markdown" href="https://excelarsiv.com/llms.txt" title="llms.txt (Markdown)">\n  <link rel="alternate" type="application/json" href="https://excelarsiv.com/.well-known/agent-card.json" title="Autonomous Agent Protocol">`;

let updatedPagesCount = 0;

for (const route of validCanonicalRoutes) {
  // Never overwrite Astro's /sablon/ product pages and /yerel-erp-crm-moduller because they are dynamically built with latest Astro components and standards
  if (route.startsWith('/sablon/') || route === '/yerel-erp-crm-moduller') {
    continue;
  }

  const relPath = route === '/' ? 'index.html' : join(route.replace(/^\//, ''), 'index.html');
  const premiumFile = join(premiumSiteDir, relPath);
  const distFile = join(distDir, relPath);

  if (existsSync(premiumFile)) {
    let premiumHtml = readFileSync(premiumFile, 'utf8');
    
    // Existing original JSON-LD from dist if available
    let jsonLd = '';
    if (existsSync(distFile)) {
      const originalDistHtml = readFileSync(distFile, 'utf8');
      jsonLd = getJsonLd(originalDistHtml);
    }
    if (!jsonLd) {
      const titleMatch = premiumHtml.match(/<title>([^<]*)<\/title>/i);
      const descMatch = premiumHtml.match(/<meta[^>]*name=["']description["'][^>]*content=["']([^"']*)["']/i);
      jsonLd = getFallbackJsonLd(`https://excelarsiv.com${route}`, titleMatch ? titleMatch[1] : '', descMatch ? descMatch[1] : '');
    }

    // Inject GSC verification meta if not present
    if (!/<meta[^>]*name=["']google-site-verification["']/i.test(premiumHtml)) {
      premiumHtml = premiumHtml.replace("<head>", `<head>\n  <meta name="google-site-verification" content="oEA2wa-CQuq0Al6xQHMJquas_RLnx0KV4eP2Erp1NWE">`);
    }

    // Inject Google tag (gtag.js) if not present
    if (!premiumHtml.includes('googletagmanager.com/gtag/js?id=G-59JM2MQCTF')) {
      const gtagSnippet = `<!-- Google tag (gtag.js) -->\n<script async src="https://www.googletagmanager.com/gtag/js?id=G-59JM2MQCTF"></script>\n<script>\n  window.dataLayer = window.dataLayer || [];\n  function gtag(){dataLayer.push(arguments);}\n  gtag('js', new Date());\n\n  gtag('config', 'G-59JM2MQCTF');\n</script>`;
      premiumHtml = premiumHtml.replace("<head>", `<head>\n${gtagSnippet}`);
    }

    // Inject robots meta if not present
    if (!/<meta[^>]*name=["']robots["']/i.test(premiumHtml)) {
      premiumHtml = premiumHtml.replace('</head>', `  ${defaultRobotsMeta}\n</head>`);
    }

    // Inject mobile and PWA meta tags
    if (!/<meta[^>]*name=["']theme-color["']/i.test(premiumHtml)) {
      premiumHtml = premiumHtml.replace('</head>', `  <meta name="theme-color" content="#107c41">\n</head>`);
    }
    if (!/<meta[^>]*name=["']apple-mobile-web-app-capable["']/i.test(premiumHtml)) {
      premiumHtml = premiumHtml.replace('</head>', `  <meta name="apple-mobile-web-app-capable" content="yes">\n</head>`);
    }

    // Inject LLMS discovery links
    if (!premiumHtml.includes('rel="describedby" href="https://excelarsiv.com/llms.txt"')) {
      premiumHtml = premiumHtml.replace('</head>', `${llmsMetaTags}\n</head>`);
    }

    // Inject Markdown twin link if product
    if (route.startsWith('/sablon/') || route.startsWith('/urun/')) {
      const twinTag = `  <link rel="alternate" type="text/markdown" href="https://excelarsiv.com${route}.md" title="AI Markdown Twin">\n`;
      if (!premiumHtml.includes(`${route}.md`)) {
        premiumHtml = premiumHtml.replace('</head>', `${twinTag}</head>`);
      }
    }

    // Inject JSON-LD if not present
    if (!/<script[^>]*type=["']application\/ld\+json["']/i.test(premiumHtml)) {
      premiumHtml = premiumHtml.replace('</head>', `  ${jsonLd}\n</head>`);
    }

    writeFileSync(distFile, premiumHtml, 'utf8');
    updatedPagesCount++;
  }
}

// 404.html
const premium404 = join(premiumSiteDir, '404.html');
if (existsSync(premium404)) {
  let h404 = readFileSync(premium404, 'utf8');
  if (!/<meta[^>]*name=["']robots["']/i.test(h404)) {
    h404 = h404.replace('</head>', `  <meta name="robots" content="noindex, follow">\n</head>`);
  }
  const jsonLd404 = getFallbackJsonLd('https://excelarsiv.com/404', 'Sayfa Bulunamadı | ExcelArşiv', 'Aradığınız sayfa bulunamadı.');
  if (!/<script[^>]*type=["']application\/ld\+json["']/i.test(h404)) {
    h404 = h404.replace('</head>', `  ${jsonLd404}\n</head>`);
  }
  writeFileSync(join(distDir, '404.html'), h404, 'utf8');
  console.log('[PASS] 404.html güncellendi.');
}

// 3. Homepage (dist/index.html) Corporate Directory Links injection for full graph crawlability
const distIndex = join(distDir, 'index.html');
if (existsSync(distIndex)) {
  let hIndex = readFileSync(distIndex, 'utf8');

  // Ensure hero-answer-engine and hero-answer classes are present for 14KB gate (MG21)
  if (!hIndex.includes('hero-answer-engine')) {
    hIndex = hIndex.replace('class="hero-copy"', 'class="hero-copy hero-answer-engine" data-ai-citation-anchor="true"');
  }
  if (!hIndex.includes('hero-answer')) {
    hIndex = hIndex.replace('<h1>İşiniz ilerlesin.<span>Tablonuz hazır.</span></h1><p>', '<h1>İşiniz ilerlesin.<span>Tablonuz hazır.</span></h1><p class="hero-answer">');
  }

  const extraFooterLinks = `
    <div class="footer-col" style="grid-column: span 2; margin-top: 1rem; border-top: 1px solid rgba(255,255,255,0.08); padding-top: 1rem;">
      <h3 style="font-size: 13px; color: #94A3B8; margin-bottom: 0.5rem;">Karar Sistemleri & Rehberler</h3>
      <div style="display: flex; flex-wrap: wrap; gap: 0.75rem 1.25rem; font-size: 13px;">
        <a href="/karar" style="color: #CBD5E1;">Karar Rehberleri</a>
        <a href="/hesaplayici" style="color: #CBD5E1;">Hesaplayıcılar</a>
        <a href="/basari-hikayeleri" style="color: #CBD5E1;">Başarı Hikayeleri</a>
        <a href="/ortaklik-mali-musavir" style="color: #CBD5E1;">Mali Müşavir Ortaklığı</a>
        <a href="/neden-excel-arsiv" style="color: #CBD5E1;">Neden ExcelArşiv?</a>
        <a href="/referans" style="color: #CBD5E1;">Referanslar</a>
        <a href="/finansal-web" style="color: #CBD5E1;">Finansal Web</a>
        <a href="/kurumsal-web" style="color: #CBD5E1;">Kurumsal Web</a>
        <a href="/yerel-erp-crm-moduller" style="color: #CBD5E1;">Yerel ERP & CRM</a>
        <a href="/demo" style="color: #CBD5E1;">Demo Kataloğu</a>
      </div>
    </div>
  `;
  if (!hIndex.includes('href="/karar"')) {
    hIndex = hIndex.replace('</div><div class="wrap footer-bottom">', `${extraFooterLinks}</div><div class="wrap footer-bottom">`);
    writeFileSync(distIndex, hIndex, 'utf8');
    console.log('[PASS] dist/index.html footer rehber dizini eklendi.');
  }
}

// 4. dist/demo/index.html: Ensure all demo subpages are crawled and indexable
const distDemoIndex = join(distDir, 'demo', 'index.html');
if (existsSync(distDemoIndex)) {
  let hDemo = readFileSync(distDemoIndex, 'utf8');
  const demoDir = join(distDir, 'demo');
  const allDemoSlugs = readdirSync(demoDir).filter((d) => {
    return statSync(join(demoDir, d)).isDirectory() && existsSync(join(demoDir, d, 'index.html'));
  });
  
  const demoLinksHtml = `
  <section class="wrap" style="margin-top: 2rem; padding-top: 2rem; border-top: 1px solid #E2E8F0;">
    <h3 style="font-size: 16px; font-weight: 700; color: #1E293B; margin-bottom: 1rem;">Tüm Excel Karar Sistemi Demoları (${allDemoSlugs.length})</h3>
    <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); gap: 0.5rem; font-size: 13px;">
      ${allDemoSlugs.map((s) => `<a href="/demo/${s}" style="color: #0284C7; text-decoration: none;">${s.replace(/-/g, ' ')}</a>`).join('\n      ')}
    </div>
  </section>
  `;
  if (!hDemo.includes('Tüm Excel Karar Sistemi Demoları')) {
    hDemo = hDemo.replace('</main>', `${demoLinksHtml}\n</main>`);
    writeFileSync(distDemoIndex, hDemo, 'utf8');
    console.log(`[PASS] dist/demo/index.html içine ${allDemoSlugs.length} demo linki entegre edildi.`);
  }
}

// 5. dist/yerel-erp-crm-moduller/index.html: Ensure 10 /urun/ links are present
const distYerelIndex = join(distDir, 'yerel-erp-crm-moduller', 'index.html');
if (existsSync(distYerelIndex)) {
  let hYerel = readFileSync(distYerelIndex, 'utf8');
  const urunSlugs = [
    'crm-musteri-satis-pipeline-excel',
    'satis-yonetim-teklif-prim-dashboard-excel',
    'cari-hesap-borc-alacak-yaslandirma-excel',
    'nakit-akisi-cek-senet-kasa-takip-excel',
    'butce-kar-zarar-maliyet-finansal-analiz-excel',
    'barkodlu-stok-ambar-uretim-recetesi-excel',
    'personel-puantaj-fazla-mesai-izin-excel',
    'pazaryeri-komisyon-karlilik-mutabakat-excel',
    'proje-lojistik-santiye-sektorel-erp-excel',
    'kurumsal-erp-crm-tam-paket',
  ];
  const urunLinksHtml = `
  <section class="wrap" style="margin-top: 2rem; padding: 2rem; background: #F8FAFC; border-radius: 16px; border: 1px solid #E2E8F0;">
    <h3 style="font-size: 18px; font-weight: 800; color: #0F172A; margin-bottom: 0.5rem;">Çekirdek ERP & CRM Karar Sistemleri (10 Modül)</h3>
    <p style="font-size: 14px; color: #475569; margin-bottom: 1.5rem;">Google Cloud Storage üzerinden anında şifreli token teslimatı ve kurumsal e-Fatura güvenceli tak-çalıştır sistemler.</p>
    <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: 1rem;">
      ${urunSlugs.map((s) => `
        <a href="/urun/${s}" style="display: block; padding: 1rem; background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 10px; text-decoration: none; color: #1E293B;">
          <strong style="display: block; font-size: 14px; margin-bottom: 0.25rem;">${s.replace(/-/g, ' ').toUpperCase()}</strong>
          <span style="font-size: 12px; color: #0284C7; font-weight: 600;">İncele & Satın Al →</span>
        </a>
      `).join('\n      ')}
    </div>
  </section>
  `;
  if (!hYerel.includes('Çekirdek ERP & CRM Karar Sistemleri')) {
    hYerel = hYerel.replace('</main>', `${urunLinksHtml}\n</main>`);
    writeFileSync(distYerelIndex, hYerel, 'utf8');
    console.log('[PASS] dist/yerel-erp-crm-moduller/index.html içine 10 ERP/CRM ürün kartı eklendi.');
  }
}

// Clean up any extra folders in dist that violate sitemap parity
if (existsSync(join(distDir, '404'))) {
  rmSync(join(distDir, '404'), { recursive: true, force: true });
}
if (existsSync(join(distDir, 'tek-dosya.html'))) {
  rmSync(join(distDir, 'tek-dosya.html'), { force: true });
}

// Remove any non-canonical subfolders from dist/pratik-programlar
const distPratik = join(distDir, 'pratik-programlar');
if (existsSync(distPratik)) {
  for (const entry of readdirSync(distPratik)) {
    const p = join(distPratik, entry);
    if (statSync(p).isDirectory()) {
      rmSync(p, { recursive: true, force: true });
    }
  }
}

// Remove non-canonical sablonlar/raporlama if present
if (existsSync(join(distDir, 'sablonlar', 'raporlama'))) {
  rmSync(join(distDir, 'sablonlar', 'raporlama'), { recursive: true, force: true });
}

// Strict Assertion: Ensure no legacy design remnants ever exist in dist/index.html
const finalIndexHtml = readFileSync(distIndex, 'utf8');
if (!finalIndexHtml.includes('İşiniz ilerlesin') || !finalIndexHtml.includes('brand-spectrum')) {
  console.error('KRİTİK HATA: dist/index.html içinde güncel premium tasarım bulunamadı!');
  process.exit(1);
}
if (finalIndexHtml.includes("Türkiye'nin En Kapsamlı <span>Excel Çözüm Arşivi</span>")) {
  console.error("KRİTİK HATA: Eski tasarım tespit edildi! Build durduruldu.");
  process.exit(1);
}
console.log('[PASS] Eski tasarım tamamen engellendi, güncel tasarım %100 kilitlendi.');

// 6. Inject Premium WhatsApp button into all dist HTML pages right before </body>
const waPremiumSnippet = `<!-- Premium WhatsApp Butonu -->
<style>
/* Premium WhatsApp butonu — diğer site stillerinden izole */
.wa-premium, .wa-premium * { box-sizing: border-box; }
.wa-premium {
  --wa-lime: #22de49;
  --wa-green: #04be37;
  --wa-deep: #078d32;
  position: fixed;
  z-index: 2147483000;
  right: max(24px, env(safe-area-inset-right));
  bottom: max(24px, env(safe-area-inset-bottom));
  font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  -webkit-font-smoothing: antialiased;
  isolation: isolate;
}
.wa-premium__link {
  position: relative;
  display: flex;
  align-items: center;
  width: max-content;
  min-width: 262px;
  max-width: min(344px, calc(100vw - 32px));
  min-height: 72px;
  padding: 0;
  border: 0;
  color: #fff;
  text-decoration: none;
  cursor: pointer;
  -webkit-tap-highlight-color: transparent;
  filter: drop-shadow(0 10px 17px rgba(10, 80, 30, .22));
  transition: transform .22s ease, filter .22s ease;
}
.wa-premium__body {
  display: flex;
  flex-direction: column;
  justify-content: center;
  gap: 3px;
  min-height: 69px;
  flex: 1;
  min-width: 0;
  padding: 12px 29px 13px 32px;
  margin-left: -15px;
  border: 1px solid rgba(0, 124, 25, .20);
  border-radius: 0 22px 22px 0;
  background: linear-gradient(180deg, #4beb68 0%, #17db45 26%, #09c738 64%, #07ae32 100%);
  box-shadow: inset 0 2px 2px rgba(255,255,255,.34), inset 0 -3px 5px rgba(0,104,27,.16);
  text-align: center;
  white-space: nowrap;
}
.wa-premium__icon {
  position: relative;
  display: grid;
  place-items: center;
  flex: 0 0 79px;
  height: 81px;
  border: 1px solid rgba(0, 118, 30, .18);
  border-radius: 25px;
  background: linear-gradient(148deg, #49f86d 0%, #0cdb42 35%, #08c539 72%, #02a72c 100%);
  box-shadow: inset 2px 3px 3px rgba(255,255,255,.33), inset -3px -5px 8px rgba(0,100,25,.16), 4px 3px 9px rgba(0, 89, 31,.22);
  z-index: 1;
}
.wa-premium__icon::before {
  content: "";
  position: absolute;
  top: 4px;
  right: 10px;
  left: 10px;
  height: 15px;
  border-radius: 40% 40% 50% 50%;
  background: linear-gradient(180deg, rgba(255,255,255,.22), transparent);
  pointer-events: none;
}
.wa-premium__logo {
  width: 47px;
  height: 47px;
  color: #fff;
  filter: drop-shadow(0 2px 2px rgba(0, 95, 26, .23));
}
.wa-premium__title {
  display: block;
  font-weight: 760;
  font-size: 17px;
  letter-spacing: -.30px;
  line-height: 1.18;
  overflow: hidden;
  text-overflow: ellipsis;
  text-shadow: 0 1px 1px rgba(0,77,20,.24);
}
.wa-premium__subtitle {
  display: block;
  font-size: 11px;
  font-weight: 550;
  line-height: 1.25;
  letter-spacing: .1px;
  opacity: .95;
}
.wa-premium__link:hover { transform: translateY(-3px); filter: drop-shadow(0 15px 21px rgba(9, 85, 32, .28)); }
.wa-premium__link:active { transform: translateY(0) scale(.985); }
.wa-premium__link:focus-visible { outline: 3px solid #1453bd; outline-offset: 5px; border-radius: 25px; }
.wa-premium__link[aria-disabled="true"] { cursor: not-allowed; }
.wa-premium__link[aria-disabled="true"]:hover { transform: none; }
@media (max-width: 480px) {
  .wa-premium { right: max(14px, env(safe-area-inset-right)); bottom: max(14px, env(safe-area-inset-bottom)); }
  .wa-premium__link { min-width: 219px; min-height: 63px; }
  .wa-premium__icon { flex-basis: 66px; height: 68px; border-radius: 21px; }
  .wa-premium__logo { width: 41px; height: 41px; }
  .wa-premium__body { min-height: 60px; padding: 10px 18px 11px 28px; border-radius: 0 19px 19px 0; }
  .wa-premium__title { font-size: 15px; }
  .wa-premium__subtitle { font-size: 10px; }
}
@media (prefers-reduced-motion: reduce) { .wa-premium__link { transition: none; } }
</style>
<!-- WHATSAPP BUTONU: Telefonu data-phone içinde ülke koduyla yazın. Örn: 905551112233 -->
<aside class="wa-premium" data-phone="905393333303" data-message="Merhaba, web siteniz üzerinden iletişime geçiyorum." aria-label="WhatsApp üzerinden iletişim">
  <a class="wa-premium__link" href="https://wa.me/905393333303?text=Merhaba%2C%20web%20siteniz%20%C3%BCzerinden%20ileti%C5%9Fime%20ge%C3%A7iyorum." target="_blank" rel="noopener noreferrer" aria-label="WhatsApp üzerinden mesaj gönder">
    <span class="wa-premium__icon"><svg class="wa-premium__logo" viewBox="0 0 24 24" aria-hidden="true" focusable="false" xmlns="http://www.w3.org/2000/svg"><path fill="currentColor" d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.198-.347.223-.644.075-.297-.149-1.255-.462-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.297-.347.446-.52.149-.174.198-.298.297-.497.1-.198.05-.371-.025-.52-.075-.148-.67-1.611-.916-2.206-.242-.579-.487-.5-.67-.51-.173-.009-.371-.011-.57-.011s-.52.074-.792.372c-.272.297-1.04 1.016-1.04 2.479s1.065 2.876 1.213 3.074c.149.198 2.095 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.872.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.003a9.85 9.85 0 0 1-5.031-1.378l-.361-.214-3.741.982.998-3.647-.235-.374a9.86 9.86 0 0 1-1.51-5.26c.002-5.45 4.437-9.886 9.893-9.886a9.83 9.83 0 0 1 6.99 2.895 9.82 9.82 0 0 1 2.893 6.993c-.003 5.45-4.437 9.889-9.893 9.889m8.413-18.297A11.82 11.82 0 0 0 12.05 0C5.495 0 .16 5.333.157 11.893c0 2.096.547 4.14 1.588 5.945L.057 24l6.307-1.654a11.89 11.89 0 0 0 5.681 1.448h.005c6.557 0 11.894-5.334 11.897-11.893a11.82 11.82 0 0 0-3.483-8.413Z"/></svg></span>
    <span class="wa-premium__body">
      <span class="wa-premium__title">WhatsApp'tan Yazın</span>
      <span class="wa-premium__subtitle">Size yardımcı olalım</span>
    </span>
  </a>
</aside>
<script>
/* Numara girilmemişse yanlış bir kişiye mesaj yönlendirmesi yapılmaz. */
(function () {
  "use strict";
  document.querySelectorAll('.wa-premium').forEach(function(widget) {
    var link = widget.querySelector('.wa-premium__link');
    if (!link) return;
    var phone = (widget.getAttribute('data-phone') || '').replace(/\\D/g, '');
    var message = widget.getAttribute('data-message') || '';
    if (/^[1-9]\\d{9,14}$/.test(phone)) {
      link.setAttribute('href', 'https://wa.me/' + phone + (message ? '?text=' + encodeURIComponent(message) : ''));
      link.removeAttribute('aria-disabled');
    } else {
      link.setAttribute('href', '#');
      link.setAttribute('aria-disabled', 'true');
    }
    link.addEventListener('click', function(event) {
      if (link.getAttribute('aria-disabled') === 'true') event.preventDefault();
    });
  });
})();
</script>`;

function injectWaButton(dir) {
  let count = 0;
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const fullPath = join(dir, entry.name);
    if (entry.isDirectory()) {
      count += injectWaButton(fullPath);
    } else if (entry.name.endsWith('.html')) {
      let html = readFileSync(fullPath, 'utf8');
      let changed = false;

      // Remove legacy contact-float if present
      if (html.includes('class="contact-float"')) {
        html = html.replace(/<a class="contact-float"[^>]*>[\s\S]*?<\/a>/gi, '');
        changed = true;
      }

      // Inject waPremiumSnippet right before </body> if not already present
      if (!html.includes('class="wa-premium"') && html.includes('</body>')) {
        html = html.replace('</body>', `${waPremiumSnippet}\n</body>`);
        changed = true;
      }

      if (changed) {
        writeFileSync(fullPath, html, 'utf8');
        count++;
      }
    }
  }
  return count;
}

const waInjectedCount = injectWaButton(distDir);
console.log(`[PASS] ${waInjectedCount} sayfaya Premium WhatsApp butonu entegre edildi.`);

console.log(`[PASS] ${updatedPagesCount} sayfa premium tasarımla başarıyla entegre edildi.`);
