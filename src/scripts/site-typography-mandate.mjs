#!/usr/bin/env node
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const MANDATE = 'data-ea-typography-mandate="chat-readable-v2"';
const FONT_LINK = '<link data-ea-typography-font href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">';

const typographyStyles = `
<style ${MANDATE}>
  :root{
    --ea-font-sans:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;
    --ea-type-body:15px;
    --ea-type-nav:14.5px;
    --ea-type-small:13px;
    --ea-type-h1:clamp(26px,3.5vw,32px);
    --ea-type-h2:clamp(20px,2.4vw,24px);
    --ea-type-h3:clamp(17px,1.8vw,19px);
    --ea-type-h4:15px;
    --ea-leading-body:1.5;
    --ea-color-body:#545454;
    --ea-color-heading:#333333;
    --ea-color-muted:#777777;
  }
  html{font-size:15px!important}
  body,button,input,select,textarea{font-family:var(--ea-font-sans)!important}
  body{font-size:15px!important;line-height:var(--ea-leading-body)!important;color:#545454!important;text-rendering:optimizeLegibility;-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale}
  header nav a,header [role="navigation"] a,nav[aria-label] a{font-size:14.5px!important;font-weight:600!important;line-height:1.4!important}
  h1,h2,h3,h4,h5,h6{color:#333333!important;font-weight:700!important;line-height:1.3!important}
  h1{font-size:clamp(26px,3.5vw,32px)!important}
  h2{font-size:clamp(20px,2.4vw,24px)!important}
  h3{font-size:clamp(17px,1.8vw,19px)!important}
  h4{font-size:15px!important}
  main p,main li,main td,main th,main details,main summary{font-family:var(--ea-font-sans)!important;font-size:15px!important;line-height:1.5!important;color:#545454!important}
  main p{margin-bottom:1.15em}
  main button,main a[class*="btn"],main a[class*="button"]{font-size:14.5px!important;font-weight:600!important}
  main input,main select,main textarea{font-size:14.5px!important;line-height:1.45!important;color:#333333!important}
  main small,.text-muted,.meta{font-size:13px!important;line-height:1.45!important;color:#777777!important}
  footer,footer a{font-family:var(--ea-font-sans)!important;font-size:13px!important;line-height:1.5!important;color:#777777!important}
  footer strong,footer h3,footer h4{color:#333333!important}
  @media(max-width:760px){
    :root{--ea-type-body:15px;--ea-type-nav:14px;--ea-type-small:13px}
    body{font-size:15px!important}
    h1{font-size:24px!important}
    h2{font-size:19px!important}
    h3{font-size:16.5px!important}
    main p,main li,main td,main th,main details,main summary{font-size:15px!important}
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
