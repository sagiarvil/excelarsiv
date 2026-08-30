#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(file)) throw new Error('SPECIAL LIVE CONTRACT: build output missing');

let html = fs.readFileSync(file, 'utf8');
if (!html.includes('data-living-workbook-v1') || !html.includes('class="workbook"')) {
  console.log('SPECIAL LIVE CONTRACT SKIP — Living Workbook route not detected.');
  process.exit(0);
}

html = html.replace(/<style\b(?=[^>]*\bid=["']living-workbook-live-fix-v2["'])[^>]*>[\s\S]*?<\/style>/gi, '');

const css = `
<style id="living-workbook-live-fix-v2" data-special-live-contract="header-hard-color-v2">
  body[data-living-workbook-v1]{
    --ea-green:#107c41;
    --ea-green-dark:#0b5f31;
    --ea-green-soft:#e8f5ed;
    --ea-blue:#176fe5;
    --ea-blue-soft:#edf4ff;
    --ea-amber:#e79a08;
    --ea-amber-soft:#fff4db;
    --ea-coral:#e45e4f;
    --ea-coral-soft:#fff0ed;
    --ea-teal:#117e87;
    --ea-teal-soft:#e9f7f7;
    --ea-ink:#1b1b1f;
    --ea-muted:#5d6470;
    --ea-line:#d1d5d8;
  }

  /* Header must remain global because the final build replaces its DOM after Astro scoping. */
  body[data-living-workbook-v1] .site-nav{position:sticky;top:0;z-index:120;border-bottom:1px solid #d6d9dc;background:rgba(250,250,248,.97);backdrop-filter:blur(14px);box-shadow:0 6px 18px rgba(24,35,50,.05)}
  body[data-living-workbook-v1] .site-nav .nav-inner{width:min(1320px,calc(100% - 48px));height:72px;margin-inline:auto;display:flex;align-items:center;gap:28px}
  body[data-living-workbook-v1] .site-nav .brand{display:flex;align-items:center;gap:10px;min-width:max-content;color:var(--ea-ink);text-decoration:none}
  body[data-living-workbook-v1] .site-nav .brand img{width:34px;height:34px;display:block}
  body[data-living-workbook-v1] .site-nav .brand-copy{display:grid;line-height:1}
  body[data-living-workbook-v1] .site-nav .brand-copy strong{font:800 17px/1 Manrope,system-ui,sans-serif;letter-spacing:.015em;color:#1b1b1f}
  body[data-living-workbook-v1] .site-nav .brand-copy small{margin-top:6px;font:500 9px/1.2 'IBM Plex Mono',monospace;letter-spacing:.07em;color:#697078}
  body[data-living-workbook-v1] .site-nav .nav-links{display:flex;align-items:center;gap:24px;margin-inline:auto}
  body[data-living-workbook-v1] .site-nav .nav-links a{padding:26px 0 22px;border-bottom:3px solid transparent;color:#3f454d;text-decoration:none;font:750 13px/1 Manrope,system-ui,sans-serif}
  body[data-living-workbook-v1] .site-nav .nav-links a:hover{color:var(--ea-green);border-bottom-color:var(--ea-green)}
  body[data-living-workbook-v1] .site-nav .nav-actions{display:flex;align-items:center;gap:10px}
  body[data-living-workbook-v1] .site-nav .btn{min-height:42px;display:inline-flex;align-items:center;justify-content:center;padding:0 16px;border:1px solid #c9cdd1;border-radius:5px;text-decoration:none;font:800 13px/1 Manrope,system-ui,sans-serif;white-space:nowrap}
  body[data-living-workbook-v1] .site-nav .btn-secondary{background:#fff;color:#25282d}
  body[data-living-workbook-v1] .site-nav .btn-primary{background:var(--ea-green);border-color:var(--ea-green);color:#fff}
  body[data-living-workbook-v1] .site-nav .btn-primary:hover{background:var(--ea-green-dark);border-color:var(--ea-green-dark)}
  body[data-living-workbook-v1] .site-nav .mobile-menu{display:none;position:relative}
  body[data-living-workbook-v1] .site-nav .mobile-menu summary{width:44px;height:44px;display:grid;place-items:center;border:1px solid #c9cdd1;border-radius:5px;background:#fff;cursor:pointer;list-style:none}
  body[data-living-workbook-v1] .site-nav .mobile-menu summary::-webkit-details-marker{display:none}
  body[data-living-workbook-v1] .site-nav .burger,body[data-living-workbook-v1] .site-nav .burger::before,body[data-living-workbook-v1] .site-nav .burger::after{display:block;width:18px;height:2px;background:#202329}
  body[data-living-workbook-v1] .site-nav .burger{position:relative}
  body[data-living-workbook-v1] .site-nav .burger::before,body[data-living-workbook-v1] .site-nav .burger::after{content:"";position:absolute;left:0}
  body[data-living-workbook-v1] .site-nav .burger::before{top:-6px}
  body[data-living-workbook-v1] .site-nav .burger::after{top:6px}
  body[data-living-workbook-v1] .site-nav .mobile-panel{position:absolute;right:0;top:50px;width:min(300px,calc(100vw - 32px));padding:8px;border:1px solid #cfd3d7;border-radius:7px;background:#fff;box-shadow:0 22px 50px rgba(24,35,50,.14)}
  body[data-living-workbook-v1] .site-nav .mobile-panel a{display:block;min-height:44px;padding:12px;border-radius:4px;color:#292d33;text-decoration:none;font:750 14px/1.35 Manrope,system-ui,sans-serif}
  body[data-living-workbook-v1] .site-nav .mobile-panel a:hover{background:var(--ea-green-soft);color:var(--ea-green-dark)}

  /* Hard-color light system: stronger financial signals without turning the page dark. */
  body[data-living-workbook-v1] .hero{border-top:6px solid var(--ea-green)!important}
  body[data-living-workbook-v1] .hero::after{content:"";position:absolute;right:0;top:0;width:31%;height:6px;background:linear-gradient(90deg,var(--ea-blue) 0 34%,var(--ea-amber) 34% 67%,var(--ea-coral) 67% 100%);z-index:3}
  body[data-living-workbook-v1] .workbook{border-color:#bac3cb!important;box-shadow:0 26px 60px rgba(31,47,69,.13)!important}
  body[data-living-workbook-v1] .workbook::before{content:"";display:block;height:6px;background:linear-gradient(90deg,var(--ea-green) 0 28%,var(--ea-blue) 28% 53%,var(--ea-amber) 53% 77%,var(--ea-coral) 77% 100%)}
  body[data-living-workbook-v1] .kpi{background:var(--ea-blue-soft)!important;border-color:#cbdaf2!important}
  body[data-living-workbook-v1] .kpi:nth-child(1){background:var(--ea-green-soft)!important;border-top-color:var(--ea-green)!important;border-color:#c8dfd1!important}
  body[data-living-workbook-v1] .kpi:nth-child(2){background:var(--ea-blue-soft)!important;border-top-color:var(--ea-blue)!important;border-color:#cbdaf2!important}
  body[data-living-workbook-v1] .kpi:nth-child(3){background:var(--ea-amber-soft)!important;border-top-color:var(--ea-amber)!important;border-color:#ead4a8!important}
  body[data-living-workbook-v1] .kpi:nth-child(4){background:var(--ea-coral-soft)!important;border-top-color:var(--ea-coral)!important;border-color:#eccbc4!important}
  body[data-living-workbook-v1] .kpi:nth-child(1) em{color:var(--ea-green-dark)!important}
  body[data-living-workbook-v1] .kpi:nth-child(2) em{color:var(--ea-blue)!important}
  body[data-living-workbook-v1] .kpi:nth-child(3) em{color:#9a6200!important}
  body[data-living-workbook-v1] .kpi:nth-child(4) em{color:#b63b30!important}

  body[data-living-workbook-v1] .decision-item{position:relative;background:#fff!important}
  body[data-living-workbook-v1] .decision-item::before{content:"";position:absolute;left:-1px;right:-1px;top:-1px;height:4px;background:var(--ea-green)}
  body[data-living-workbook-v1] .decision-item:nth-child(2)::before{background:var(--ea-blue)}
  body[data-living-workbook-v1] .decision-item:nth-child(3)::before{background:var(--ea-amber)}
  body[data-living-workbook-v1] .decision-item:nth-child(4)::before{background:var(--ea-coral)}
  body[data-living-workbook-v1] .decision-item:nth-child(1) small{color:var(--ea-green)!important}
  body[data-living-workbook-v1] .decision-item:nth-child(2) small{color:var(--ea-blue)!important}
  body[data-living-workbook-v1] .decision-item:nth-child(3) small{color:#a56800!important}
  body[data-living-workbook-v1] .decision-item:nth-child(4) small{color:#bd4439!important}

  body[data-living-workbook-v1] .intent-card{overflow:hidden!important;border-color:#cbd0d4!important;box-shadow:0 10px 26px rgba(32,45,61,.05)!important}
  body[data-living-workbook-v1] .intent-card::before{content:"";position:absolute;left:0;right:0;top:0;height:5px;background:var(--ea-green)}
  body[data-living-workbook-v1] .intent-card:nth-child(2)::before{background:var(--ea-blue)}
  body[data-living-workbook-v1] .intent-card:nth-child(3)::before{background:var(--ea-amber)}
  body[data-living-workbook-v1] .intent-card:nth-child(4)::before{background:var(--ea-coral)}
  body[data-living-workbook-v1] .intent-card:nth-child(5)::before{background:var(--ea-teal)}
  body[data-living-workbook-v1] .intent-card:nth-child(6)::before{background:#7655c7}
  body[data-living-workbook-v1] .intent-card:nth-child(2) .code{color:var(--ea-blue)!important}
  body[data-living-workbook-v1] .intent-card:nth-child(3) .code{color:#a56800!important}
  body[data-living-workbook-v1] .intent-card:nth-child(4) .code{color:#bd4439!important}
  body[data-living-workbook-v1] .intent-card:nth-child(5) .code{color:var(--ea-teal)!important}
  body[data-living-workbook-v1] .intent-card:nth-child(6) .code{color:#7655c7!important}

  body[data-living-workbook-v1] #karsilastirma{background:linear-gradient(90deg,#fff 0 33%,#f9fbff 33% 66%,#f5fbf7 66% 100%)!important}
  body[data-living-workbook-v1] .comparison-top .good{background:#dff3e7!important;color:var(--ea-green-dark)!important}
  body[data-living-workbook-v1] .compare-old .state::before{background:#c5161d!important}
  body[data-living-workbook-v1] .compare-new{background:#f4fbf6!important}
  body[data-living-workbook-v1] .compare-new .state::before{background:var(--ea-green)!important}

  body[data-living-workbook-v1] .arch-card{border-color:#cbd0d4!important;box-shadow:0 8px 22px rgba(31,45,61,.04)!important}
  body[data-living-workbook-v1] .arch-card:nth-of-type(2) .arch-output{border-left-color:var(--ea-blue)!important;background:var(--ea-blue-soft)!important;color:#2455a2!important}
  body[data-living-workbook-v1] .arch-card:nth-of-type(3) .arch-output{border-left-color:var(--ea-amber)!important;background:var(--ea-amber-soft)!important;color:#855500!important}
  body[data-living-workbook-v1] .arch-card:nth-of-type(4) .arch-output{border-left-color:var(--ea-coral)!important;background:var(--ea-coral-soft)!important;color:#a13f35!important}
  body[data-living-workbook-v1] .arch-card:nth-of-type(5) .arch-output{border-left-color:var(--ea-teal)!important;background:var(--ea-teal-soft)!important;color:#17666b!important}

  body[data-living-workbook-v1] .flow-step{border-top:5px solid var(--ea-green)!important}
  body[data-living-workbook-v1] .flow-step:nth-child(2){border-top-color:var(--ea-blue)!important}
  body[data-living-workbook-v1] .flow-step:nth-child(3){border-top-color:var(--ea-amber)!important}
  body[data-living-workbook-v1] .flow-step:nth-child(4){border-top-color:var(--ea-coral)!important}
  body[data-living-workbook-v1] .flow-step:nth-child(5){border-top-color:var(--ea-teal)!important}
  body[data-living-workbook-v1] .flow-step:nth-child(2) small{color:var(--ea-blue)!important}
  body[data-living-workbook-v1] .flow-step:nth-child(3) small{color:#a56800!important}
  body[data-living-workbook-v1] .flow-step:nth-child(4) small{color:#bd4439!important}
  body[data-living-workbook-v1] .flow-step:nth-child(5) small{color:var(--ea-teal)!important}

  body[data-living-workbook-v1] .proof-card{border-top-width:5px!important;border-top-color:var(--ea-green)!important;box-shadow:0 16px 36px rgba(31,45,61,.08)!important}
  body[data-living-workbook-v1] .cta{border-top:5px solid var(--ea-green)!important;background:linear-gradient(90deg,#f2faf5 0 58%,#f1f6ff 58% 100%)!important}

  @media(max-width:1080px){
    body[data-living-workbook-v1] .site-nav .nav-links{gap:16px}
    body[data-living-workbook-v1] .site-nav .nav-links a{font-size:12px}
    body[data-living-workbook-v1] .site-nav .nav-actions .btn-secondary{display:none}
  }
  @media(max-width:860px){
    body[data-living-workbook-v1] .site-nav .nav-inner{width:min(100% - 32px,1320px);height:64px;gap:14px}
    body[data-living-workbook-v1] .site-nav .nav-links{display:none}
    body[data-living-workbook-v1] .site-nav .mobile-menu{display:block}
    body[data-living-workbook-v1] .site-nav .nav-actions{margin-left:auto}
  }
  @media(max-width:620px){
    body[data-living-workbook-v1] .site-nav .brand-copy small{display:none}
    body[data-living-workbook-v1] .site-nav .nav-actions>.nav-cta{display:none}
    body[data-living-workbook-v1] #karsilastirma{background:#fff!important}
  }
  @media(prefers-reduced-motion:reduce){body[data-living-workbook-v1] *{scroll-behavior:auto!important}}
</style>`;

if (!html.includes('</head>')) throw new Error('SPECIAL LIVE CONTRACT: </head> missing');
html = html.replace('</head>', `${css}\n</head>`);

for (const token of [
  'id="living-workbook-live-fix-v2"',
  'data-special-live-contract="header-hard-color-v2"',
  'body[data-living-workbook-v1] .site-nav',
  '--ea-blue:#176fe5',
  '--ea-amber:#e79a08',
  '--ea-coral:#e45e4f',
]) {
  if (!html.includes(token)) throw new Error(`SPECIAL LIVE CONTRACT: required token missing: ${token}`);
}

fs.writeFileSync(file, html);
console.log('SPECIAL LIVE CONTRACT PASS — scoped-header regression fixed; Living Workbook hard-color finance signals restored.');
