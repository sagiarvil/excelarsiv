#!/usr/bin/env node
import { readFileSync, writeFileSync } from 'node:fs';

function once(html, from, to, label) {
  const count = html.split(from).length - 1;
  if (count !== 1) throw new Error(`FUNNEL PREMIUM LAYOUT: expected 1 ${label}, found ${count}`);
  return html.replace(from, to);
}

export function applySpecialFunnelPremiumLayout({ specialPath = 'dist/ozel-excel-sistemleri/index.html' } = {}) {
  let html = readFileSync(specialPath, 'utf8');

  const premiumStyles = `
<style data-special-funnel-premium-layout>
  .special-v3 .funnel-extensions{padding:84px 0 88px!important;background:linear-gradient(180deg,#f7f9f8 0%,#f4f7f5 100%)!important;overflow:hidden}
  .special-v3 .funnel-extensions>.wrap{width:min(1180px,calc(100% - 48px))!important;max-width:1180px!important;margin-inline:auto!important;padding-inline:0!important}
  .special-v3 .funnel-extensions .section-title{max-width:760px!important;font-size:clamp(34px,3.5vw,48px)!important;line-height:1.04!important;letter-spacing:-.045em!important}
  .special-v3 .funnel-extensions .section-copy{max-width:760px!important;margin-top:16px!important;font-size:16px!important;line-height:1.7!important}
  .special-v3 .funnel-grid{display:grid!important;grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:18px!important;margin-top:34px!important;align-items:stretch!important}
  .special-v3 .funnel-card{position:relative!important;min-height:0!important;padding:28px 28px 26px!important;border:1px solid #dfe7e2!important;border-radius:22px!important;background:linear-gradient(180deg,#fff 0%,#fbfcfb 100%)!important;box-shadow:0 14px 36px rgba(20,48,35,.055)!important;overflow:hidden!important}
  .special-v3 .funnel-card::before{content:'';position:absolute;inset:0 auto 0 0;width:4px;background:linear-gradient(180deg,#15945a,#0c6f42);opacity:.9}
  .special-v3 .funnel-card:nth-child(2)::before{background:linear-gradient(180deg,#1d4ed8,#1e40af)}
  .special-v3 .funnel-card:nth-child(3)::before{background:linear-gradient(180deg,#7c3aed,#5b21b6)}
  .special-v3 .funnel-card:nth-child(4)::before{background:linear-gradient(180deg,#d97706,#b45309)}
  .special-v3 .funnel-card .funnel-kicker{display:inline-flex!important;align-items:center!important;width:max-content!important;max-width:100%!important;padding:6px 9px!important;border-radius:999px!important;background:#edf7f1!important;color:#0b6b37!important;font-size:11px!important;line-height:1!important;letter-spacing:.1em!important}
  .special-v3 .funnel-card:nth-child(2) .funnel-kicker{background:#eef4ff!important;color:#1d4ed8!important}
  .special-v3 .funnel-card:nth-child(3) .funnel-kicker{background:#f4f0ff!important;color:#6d28d9!important}
  .special-v3 .funnel-card:nth-child(4) .funnel-kicker{background:#fff7e8!important;color:#b45309!important}
  .special-v3 .funnel-card h3{margin:16px 0 10px!important;font-size:clamp(22px,2.1vw,28px)!important;line-height:1.18!important;letter-spacing:-.035em!important;color:#13231b!important}
  .special-v3 .funnel-card p{font-size:15px!important;line-height:1.72!important;color:#647267!important}
  .special-v3 .funnel-card ul{margin:20px 0 0!important;padding-top:18px!important;border-top:1px solid #e7ece9!important;gap:10px!important}
  .special-v3 .funnel-card li{font-size:14px!important;line-height:1.55!important;color:#425248!important}
  .special-v3 .funnel-card .btn{margin-top:24px!important;align-self:flex-start!important;min-height:44px!important;padding-inline:16px!important;border-radius:10px!important}
  .special-v3 .funnel-card .onboard-note{margin-top:20px!important;padding:15px 16px!important;border-left:0!important;border:1px solid #e2e8e4!important;border-radius:12px!important;background:#f8faf9!important;font-size:13.5px!important;line-height:1.6!important}
  .special-v3 .radar-row{margin-top:20px!important;padding:17px 18px!important;border-radius:13px!important;background:#fff!important;box-shadow:0 4px 14px rgba(20,48,35,.035)!important}
  .special-v3 .radar-row a{font-weight:800!important;text-decoration:none!important;color:#0b6b37!important}
  .special-v3 .radar-row a:hover{text-decoration:underline!important}
  @media(max-width:980px){
    .special-v3 .funnel-extensions>.wrap{width:min(100% - 32px,820px)!important}
    .special-v3 .funnel-grid{grid-template-columns:1fr!important;gap:14px!important}
    .special-v3 .funnel-card{padding:24px!important}
  }
  @media(max-width:560px){
    .special-v3 .funnel-extensions{padding:58px 0 62px!important}
    .special-v3 .funnel-extensions>.wrap{width:calc(100% - 24px)!important}
    .special-v3 .funnel-card{padding:21px 20px!important;border-radius:18px!important}
    .special-v3 .funnel-card h3{font-size:22px!important}
  }
</style>`;

  html = once(html, '</head>', `${premiumStyles}</head>`, 'head');

  for (const token of ['data-special-funnel-premium-layout','width:min(1180px,calc(100% - 48px))','border-radius:22px']) {
    if (!html.includes(token)) throw new Error(`FUNNEL PREMIUM LAYOUT: token missing ${token}`);
  }

  writeFileSync(specialPath, html, 'utf8');
  console.log('FUNNEL PREMIUM LAYOUT PASS — extension section constrained, compact and premium.');
}
