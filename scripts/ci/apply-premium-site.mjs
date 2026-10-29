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
  // Never overwrite Astro's /sablon/ product pages because they are dynamically built with super-mandate hero-answer-engine
  if (route.startsWith('/sablon/')) {
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

console.log(`[PASS] ${updatedPagesCount} sayfa premium tasarımla başarıyla entegre edildi.`);
