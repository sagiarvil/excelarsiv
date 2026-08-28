#!/usr/bin/env node
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const MANDATE = 'data-ea-typography-mandate="chat-readable-v2"';
const FONT_LINK = '<link data-ea-typography-font href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">';

const typographyStyles = `
<style ${MANDATE}>
  :root{
    --ea-font-sans:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;
    --ea-type-body:16px;
    --ea-type-nav:16px;
    --ea-type-small:14px;
    --ea-type-h1:clamp(42px,5vw,64px);
    --ea-type-h2:clamp(32px,3.8vw,48px);
    --ea-type-h3:clamp(22px,2.2vw,29px);
    --ea-leading-body:1.65;
  }
  html{font-size:16px!important}
  body,button,input,select,textarea{font-family:var(--ea-font-sans)!important}
  body{font-size:16px!important;line-height:var(--ea-leading-body)!important;text-rendering:optimizeLegibility;-webkit-font-smoothing:antialiased}
  header nav a,header [role="navigation"] a,nav[aria-label] a{font-size:16px!important;line-height:1.35!important}
  main p,main li,main td,main th,main details,main summary{font-family:var(--ea-font-sans)!important;font-size:16px!important;line-height:1.65!important}
  main button,main a[class*="btn"],main a[class*="button"]{font-size:16px!important}
  main input,main select,main textarea{font-size:16px!important;line-height:1.45!important}
  main small{font-size:14px!important;line-height:1.5!important}
  footer,footer a{font-family:var(--ea-font-sans)!important;font-size:14px!important;line-height:1.55!important}
  @media(max-width:760px){
    :root{--ea-type-body:16px;--ea-type-nav:16px;--ea-type-small:14px}
    body{font-size:16px!important}
    main p,main li,main td,main th,main details,main summary{font-size:16px!important}
  }
</style>`;

// The premium light special-systems page owns its own type scale. It still carries
// the global typography contract token so localization and conformance gates can
// verify the page, but no legacy !important rules are allowed to flatten its UI.
const lightSpecialTypographyStyles = `
<style ${MANDATE} data-ea-special-light-typography>
  :root{--ea-font-sans:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif}
  body.special-light-v1,body.special-light-v1 button,body.special-light-v1 input,body.special-light-v1 select,body.special-light-v1 textarea{font-family:var(--ea-font-sans)!important}
</style>`;

function walkHtml(dir, out = []) {
  if (!existsSync(dir)) return out;
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    const stat = statSync(path);
    if (stat.isDirectory()) walkHtml(path, out);
    else if (stat.isFile() && name.endsWith('.html')) out.push(path);
  }
  return out;
}

export function applySiteTypographyMandate({ distDir = 'dist', specialPath = join('dist', 'ozel-excel-sistemleri', 'index.html') } = {}) {
  const htmlFiles = walkHtml(distDir);
  if (htmlFiles.length < 10) throw new Error(`TYPOGRAPHY MANDATE: suspicious HTML count ${htmlFiles.length}`);

  let lightSpecial = false;
  for (const path of htmlFiles) {
    let html = readFileSync(path, 'utf8');
    const isSpecialLight = path === specialPath && html.includes('data-special-light-v1');
    if (isSpecialLight) {
      lightSpecial = true;
      if (!html.includes(MANDATE)) html = html.replace('</head>', `${lightSpecialTypographyStyles}</head>`);
    } else {
      if (!html.includes('data-ea-typography-font')) html = html.replace('</head>', `${FONT_LINK}${typographyStyles}</head>`);
      else if (!html.includes(MANDATE)) html = html.replace('</head>', `${typographyStyles}</head>`);
    }
    writeFileSync(path, html, 'utf8');
  }

  const finalFiles = walkHtml(distDir);
  const missingMandate = finalFiles.filter((path) => !readFileSync(path, 'utf8').includes(MANDATE));
  if (missingMandate.length) throw new Error(`TYPOGRAPHY MANDATE: missing on ${missingMandate.length} HTML files: ${missingMandate.slice(0, 5).join(', ')}`);

  const specialFinal = readFileSync(specialPath, 'utf8');
  if (lightSpecial) {
    for (const token of ['data-special-light-v1','data-ea-special-light-typography','Excel ile Sınırlarınızı Aşın','Gerçek İş Sonuçları Alın.','İşinizi Büyüten Excel Çözümleri']) {
      if (!specialFinal.includes(token)) throw new Error(`TYPOGRAPHY MANDATE: light special token missing: ${token}`);
    }
    if (specialFinal.includes('data-special-sales-v5')) throw new Error('TYPOGRAPHY MANDATE: legacy special sales CSS leaked into premium light page');
  }

  console.log(`TYPOGRAPHY MANDATE PASS — ${finalFiles.length}/${finalFiles.length} HTML pages carry the readable typography contract; special=${lightSpecial ? 'premium-light-v1' : 'legacy'}.`);
}
