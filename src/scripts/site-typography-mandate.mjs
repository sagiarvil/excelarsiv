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
    --ea-type-nav:15px;
    --ea-type-small:13.5px;
    --ea-type-h1:clamp(32px,4vw,44px);
    --ea-type-h2:clamp(24px,2.8vw,32px);
    --ea-type-h3:clamp(19px,2vw,24px);
    --ea-type-h4:17px;
    --ea-leading-body:1.6;
    --ea-color-body:#444444;
    --ea-color-heading:#222222;
    --ea-color-muted:#666666;
    --ea-color-accent:#107c41;
  }
  html{font-size:16px!important}
  body,button,input,select,textarea{font-family:var(--ea-font-sans)!important}
  body{font-size:16px!important;line-height:var(--ea-leading-body)!important;color:#444444!important;background:#ffffff;text-rendering:optimizeLegibility;-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale}
  
  /* Someka Başlık Renkleri & Hiyerarşisi */
  h1,h2,h3,h4,h5,h6,.display,.hero-title,.section-title{color:#222222!important;font-weight:700!important;line-height:1.25!important;letter-spacing:-0.02em}
  h1,.display,.hero-title{font-size:clamp(32px,4vw,44px)!important;font-weight:800!important;line-height:1.15!important;letter-spacing:-0.03em!important}
  h2,.section-title{font-size:clamp(24px,2.8vw,32px)!important;font-weight:700!important}
  h3{font-size:clamp(19px,2vw,24px)!important}
  h4{font-size:17px!important}
  h5,h6{font-size:15px!important}

  /* Someka Gövde Metinleri & Paragraflar */
  main p,main li,main td,main th,main details,main summary,article p,article li{font-family:var(--ea-font-sans)!important;font-size:16px!important;line-height:1.6!important;color:#444444!important}
  main p{margin-bottom:1.2em}
  
  /* Someka Kart Başlıkları & Ürün Başlıkları (Okunaklı & Tok) */
  .single-post-title,.single-post-title a,.card__title,.card__title a,.tcard__title,.tcard__title a,[data-template-card] h3,[data-template-card] a,.mce-quick-card__title,.mce-template-card__title{font-family:var(--ea-font-sans)!important;font-size:16px!important;font-weight:700!important;line-height:1.35!important;color:#222222!important}
  .single-post-title a:hover,.card__title a:hover,.tcard__title a:hover,[data-template-card] a:hover{color:#107c41!important}

  /* Someka Kart Özetleri & Açıklamaları (Asla mikroskobik 11px/12px değil) */
  .vcex-post-excerpt,.card__desc,.tcard__desc,[data-template-card] p,.template-card-desc,.soho-category-desc{font-size:14px!important;line-height:1.55!important;color:#555555!important}

  /* Navigasyon & Butonlar */
  header nav a,header [role="navigation"] a,nav[aria-label] a{font-size:15px!important;font-weight:600!important;line-height:1.4!important;color:#222222!important}
  main button,main a[class*="btn"],main a[class*="button"]{font-size:15px!important;font-weight:600!important}
  main input,main select,main textarea{font-size:15px!important;line-height:1.45!important;color:#222222!important}

  /* Eyebrow / Kicker / Etiketler */
  .eyebrow,.kicker,.hub-kicker,.hero-mobile-copy__eyebrow{font-size:12.5px!important;font-weight:700!important;letter-spacing:.06em!important;color:#107c41!important}

  /* Yardımcı Metinler & Muted Bilgiler */
  main small,.text-muted,.meta,span.text-xs{font-size:13.5px!important;line-height:1.5!important;color:#666666!important}

  /* Footer */
  footer,footer a{font-family:var(--ea-font-sans)!important;font-size:13.5px!important;line-height:1.55!important;color:#666666!important}
  footer strong,footer h3,footer h4{color:#222222!important;font-size:15px!important}

  @media(max-width:760px){
    :root{--ea-type-body:15.5px;--ea-type-nav:14.5px;--ea-type-small:13px}
    body{font-size:15.5px!important}
    h1,.display,.hero-title{font-size:28px!important}
    h2,.section-title{font-size:22px!important}
    h3{font-size:18px!important}
    main p,main li,main td,main th,main details,main summary{font-size:15.5px!important}
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
