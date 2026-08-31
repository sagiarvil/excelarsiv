#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const HOME = path.resolve('dist/index.html');
const SPECIAL = path.resolve('dist/ozel-excel-sistemleri/index.html');
const WA_PHONE = '905393333303';
const wa = (message) => `https://wa.me/${WA_PHONE}?text=${encodeURIComponent(message)}`;

const specialWa = wa('Merhaba Barış Bey, mobilden ulaşıyorum. İşletmemiz için özel Excel finans karar sistemi ihtiyacını netleştirmek istiyorum.');

const injectMarker = (html, marker) => {
  if (html.includes(marker)) return html;
  return html.replace(/<body\b([^>]*)>/u, `<body$1 ${marker}>`);
};

const mobileCss = `<style id="mobile-premium-v18-css">
@media(max-width:720px){
  body[data-mobile-premium-v18]{overflow-x:clip!important}

  /* HOME — first-screen decision architecture */
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__label-row{margin:0 6px 8px!important;font-size:13px!important}
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__hint{display:none!important}
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__form{grid-template-columns:24px minmax(0,1fr) 50px!important;min-height:62px!important;padding:6px 7px 6px 15px!important;border-radius:18px!important;box-shadow:0 16px 38px rgba(15,23,42,.12)!important}
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__input{height:48px!important;padding:0 10px!important;font-size:16px!important}
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__button{width:50px!important;height:50px!important;border-radius:14px!important}
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__quick-wrap{margin-top:10px!important}
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__quick{gap:7px!important;overflow-x:auto!important;flex-wrap:nowrap!important;padding-bottom:2px;scrollbar-width:none}
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__quick::-webkit-scrollbar{display:none}
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__quick-link{flex:0 0 auto!important;min-height:36px!important;display:inline-flex!important;align-items:center!important;padding:0 11px!important;border-radius:999px!important;font-size:11px!important}
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__panel{top:82px!important;position:fixed!important;left:10px!important;right:10px!important;max-height:calc(100dvh - 170px)!important;border-radius:20px!important}
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__results{max-height:calc(100dvh - 310px)!important}
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__results>li>a{grid-template-columns:38px minmax(0,1fr)!important;padding:12px!important}
  body[data-mobile-premium-v18] .hero-search--premium .hero-search__result-side{display:none!important}

  body[data-mobile-premium-v18] .finance-pillars{padding:48px 0!important;background:linear-gradient(180deg,#fff,#f8fbf9)!important}
  body[data-mobile-premium-v18] .finance-pillars__head{gap:12px!important;margin-bottom:20px!important}
  body[data-mobile-premium-v18] .finance-pillars__head h2{font-size:31px!important;line-height:1.02!important;text-wrap:balance}
  body[data-mobile-premium-v18] .finance-pillars__head>p{font-size:15px!important;line-height:1.62!important}
  body[data-mobile-premium-v18] .finance-pillars__grid{display:grid!important;grid-template-columns:1fr!important;gap:12px!important}
  body[data-mobile-premium-v18] .finance-pillar{padding:21px 19px!important;border-radius:22px!important;box-shadow:0 12px 30px rgba(15,23,42,.065)!important}
  body[data-mobile-premium-v18] .finance-pillar::before{height:5px!important;border-radius:22px 22px 0 0!important}
  body[data-mobile-premium-v18] .finance-pillar__no{font-size:12px!important}
  body[data-mobile-premium-v18] .finance-pillar h3{margin-top:11px!important;font-size:21px!important;line-height:1.16!important}
  body[data-mobile-premium-v18] .finance-pillar ul{gap:8px!important;margin:15px 0 18px!important;font-size:14px!important;line-height:1.55!important}
  body[data-mobile-premium-v18] .finance-pillar>a{min-height:46px!important;display:flex!important;align-items:center!important;justify-content:center!important;padding:0 14px!important;border-radius:13px!important;background:#f1faf5!important;border:1px solid #cfe7d9!important;font-size:13px!important}
  body[data-mobile-premium-v18] .finance-pillar--blue>a{background:#f3f7ff!important;border-color:#d8e4fb!important;color:#1d4ed8!important}
  body[data-mobile-premium-v18] .finance-pillar--violet>a{background:#f8f5ff!important;border-color:#e6ddfb!important;color:#6d28d9!important}
  body[data-mobile-premium-v18] .finance-pillar--amber>a{background:#fff9ef!important;border-color:#f9dfb6!important;color:#a84f06!important}

  body[data-mobile-premium-v18] .high-ticket-bridge{padding:14px 0 18px!important;background:#fff!important;color:#fff!important}
  body[data-mobile-premium-v18] .high-ticket-bridge>.home-shell{width:calc(100% - 20px)!important}
  body[data-mobile-premium-v18] .high-ticket-bridge__inner{display:grid!important;grid-template-columns:1fr!important;gap:20px!important;padding:25px 20px!important;border-radius:26px!important;background:linear-gradient(145deg,#0f172a,#132238)!important;box-shadow:0 18px 46px rgba(15,23,42,.18)!important}
  body[data-mobile-premium-v18] .high-ticket-bridge h2{font-size:32px!important;line-height:1.02!important;text-wrap:balance}
  body[data-mobile-premium-v18] .high-ticket-bridge__copy>p:last-child{font-size:15px!important;line-height:1.65!important}
  body[data-mobile-premium-v18] .high-ticket-bridge__proof{padding:18px!important;border-radius:18px!important}
  body[data-mobile-premium-v18] .high-ticket-bridge__proof ul{font-size:14px!important;line-height:1.5!important}
  body[data-mobile-premium-v18] .high-ticket-bridge__proof>a{min-height:52px!important;border-radius:14px!important;font-size:13px!important;text-align:center!important}

  body[data-mobile-premium-v18] .home-finance-close{padding:14px 0 30px!important;background:#fff!important}
  body[data-mobile-premium-v18] .home-finance-close>.home-shell{width:calc(100% - 20px)!important}
  body[data-mobile-premium-v18] .home-finance-close__inner{padding:22px 18px!important;border:1px solid #d8e8de!important;border-radius:24px!important;background:linear-gradient(135deg,#f0fdf4,#eff6ff)!important;box-shadow:0 12px 32px rgba(15,23,42,.06)!important}
  body[data-mobile-premium-v18] .home-finance-close h2{font-size:30px!important;line-height:1.04!important}
  body[data-mobile-premium-v18] .home-finance-close p:not(.eyebrow){font-size:15px!important;line-height:1.6!important}
  body[data-mobile-premium-v18] .home-finance-close__cta{min-height:52px!important;border-radius:14px!important;font-size:14px!important}

  /* SPECIAL — compact premium mobile system architecture */
  body[data-special-mobile-premium-v18]{padding-bottom:calc(86px + env(safe-area-inset-bottom))!important;overflow-x:clip!important;background:#fff!important}
  body[data-special-mobile-premium-v18] .site-nav{min-height:60px!important;background:rgba(255,255,255,.97)!important;box-shadow:0 5px 18px rgba(15,23,42,.035)!important}
  body[data-special-mobile-premium-v18] .site-nav .wrap-wide{width:calc(100% - 20px)!important}
  body[data-special-mobile-premium-v18] .nav-inner{min-height:60px!important;gap:8px!important}
  body[data-special-mobile-premium-v18] .brand{gap:8px!important}
  body[data-special-mobile-premium-v18] .brand img{width:34px!important;height:34px!important}
  body[data-special-mobile-premium-v18] .brand-copy strong{font-size:16px!important}
  body[data-special-mobile-premium-v18] .brand-copy small{display:none!important}
  body[data-special-mobile-premium-v18] .nav-links{display:none!important}
  body[data-special-mobile-premium-v18] .nav-actions{gap:7px!important;margin-left:auto!important}
  body[data-special-mobile-premium-v18] .nav-cta{min-height:40px!important;padding-inline:12px!important;border-radius:12px!important;font-size:12px!important;white-space:nowrap!important}
  body[data-special-mobile-premium-v18] .mobile-menu{display:block!important}
  body[data-special-mobile-premium-v18] .mobile-menu summary{width:40px!important;height:40px!important;border-radius:12px!important}
  body[data-special-mobile-premium-v18] .mobile-panel{position:fixed!important;top:68px!important;left:10px!important;right:10px!important;width:auto!important;padding:10px!important;border-radius:20px!important;box-shadow:0 24px 64px rgba(15,23,42,.18)!important}
  body[data-special-mobile-premium-v18] .mobile-panel a{min-height:48px!important;display:flex!important;align-items:center!important;padding:0 13px!important;font-size:14px!important;font-weight:650!important}

  body[data-special-mobile-premium-v18] .hero{padding:12px 0 30px!important;background:linear-gradient(180deg,#f5faf7 0%,#fff 100%)!important}
  body[data-special-mobile-premium-v18] .hero>.wrap-wide{width:calc(100% - 20px)!important}
  body[data-special-mobile-premium-v18] .hero-grid{display:grid!important;grid-template-columns:1fr!important;gap:18px!important;padding:22px 18px 18px!important;border:1px solid #dfe8e2!important;border-radius:26px!important;background:linear-gradient(145deg,#fff 0%,#fff 72%,#f3faf6 100%)!important;box-shadow:0 16px 42px rgba(15,23,42,.09)!important}
  body[data-special-mobile-premium-v18] .hero-copy{padding:0!important}
  body[data-special-mobile-premium-v18] .hero .eyebrow{width:100%!important;justify-content:flex-start!important;margin:0 0 16px!important;padding:9px 11px!important;border-radius:14px!important;font-size:12px!important;line-height:1.3!important}
  body[data-special-mobile-premium-v18] .hero h1{max-width:12ch!important;font-size:clamp(34px,9.6vw,40px)!important;line-height:1.01!important;letter-spacing:-.055em!important}
  body[data-special-mobile-premium-v18] .hero-lead{max-width:35ch!important;margin-top:16px!important;font-size:16px!important;line-height:1.6!important;color:#526176!important}
  body[data-special-mobile-premium-v18] .hero-actions{display:grid!important;grid-template-columns:1fr!important;gap:9px!important;margin-top:20px!important}
  body[data-special-mobile-premium-v18] .hero-actions .btn{width:100%!important;min-height:54px!important;padding:0 14px!important;border-radius:15px!important;font-size:14px!important;text-align:center!important}
  body[data-special-mobile-premium-v18] .hero-bullets{display:grid!important;grid-template-columns:1fr 1fr!important;gap:8px!important;margin-top:16px!important}
  body[data-special-mobile-premium-v18] .hero-bullets li{min-height:44px!important;align-items:flex-start!important;padding:9px 10px!important;border:1px solid #e0e7e3!important;border-radius:13px!important;background:#fff!important;font-size:11.5px!important;line-height:1.35!important}
  body[data-special-mobile-premium-v18] .hero-bullets li::before{width:18px!important;height:18px!important;flex:0 0 18px!important}

  body[data-special-mobile-premium-v18] .workbook{display:block!important;margin-top:2px!important;border-radius:18px!important;box-shadow:0 12px 30px rgba(15,23,42,.08)!important}
  body[data-special-mobile-premium-v18] .wb-top{padding:12px!important}
  body[data-special-mobile-premium-v18] .wb-title{font-size:13px!important}.wb-note{font-size:10px!important}
  body[data-special-mobile-premium-v18] .wb-body{padding:10px!important}
  body[data-special-mobile-premium-v18] .metric-grid{display:grid!important;grid-template-columns:1fr 1fr!important;gap:7px!important}
  body[data-special-mobile-premium-v18] .metric{padding:10px!important;border-radius:10px!important}
  body[data-special-mobile-premium-v18] .metric small{font-size:10.5px!important}.metric strong{margin-top:4px!important;font-size:16px!important}.metric span{font-size:10px!important}
  body[data-special-mobile-premium-v18] .wb-chart-grid,body[data-special-mobile-premium-v18] .wb-lists{display:none!important}

  body[data-special-mobile-premium-v18] .hero-proof-premium{padding:0 0 42px!important}
  body[data-special-mobile-premium-v18] .hero-proof-premium__grid{width:calc(100% - 24px)!important;grid-template-columns:1fr!important;gap:10px!important;padding-top:16px!important}
  body[data-special-mobile-premium-v18] .hero-proof-premium__card{grid-template-columns:48px 1fr!important;gap:13px!important;padding:17px!important;border-radius:18px!important;box-shadow:0 9px 24px rgba(15,23,42,.055)!important}
  body[data-special-mobile-premium-v18] .hero-proof-premium__icon{width:48px!important;height:48px!important}
  body[data-special-mobile-premium-v18] .hero-proof-premium__card strong{font-size:17px!important}.hero-proof-premium__card span{font-size:14px!important;line-height:1.5!important}

  body[data-special-mobile-premium-v18] .section{padding:48px 0!important}
  body[data-special-mobile-premium-v18] .section .wrap-wide,body[data-special-mobile-premium-v18] .section .wrap{width:calc(100% - 24px)!important}
  body[data-special-mobile-premium-v18] .section-head{max-width:none!important;margin-bottom:22px!important}
  body[data-special-mobile-premium-v18] .section-kicker{font-size:11px!important;letter-spacing:.1em!important}
  body[data-special-mobile-premium-v18] .section-head h2{font-size:30px!important;line-height:1.04!important;letter-spacing:-.045em!important;text-wrap:balance}
  body[data-special-mobile-premium-v18] .section-head p{max-width:36ch!important;font-size:15px!important;line-height:1.62!important}
  body[data-special-mobile-premium-v18] .problem-grid,body[data-special-mobile-premium-v18] .area-grid,body[data-special-mobile-premium-v18] .process-grid,body[data-special-mobile-premium-v18] .why-grid{grid-template-columns:1fr!important;gap:11px!important}
  body[data-special-mobile-premium-v18] .problem-card,body[data-special-mobile-premium-v18] .area-card,body[data-special-mobile-premium-v18] .process-card,body[data-special-mobile-premium-v18] .why-card{padding:19px!important;border-radius:20px!important;box-shadow:0 9px 24px rgba(15,23,42,.05)!important}
  body[data-special-mobile-premium-v18] .problem-card h3,body[data-special-mobile-premium-v18] .area-card h3,body[data-special-mobile-premium-v18] .process-card h3,body[data-special-mobile-premium-v18] .why-card h3{font-size:20px!important;line-height:1.18!important}
  body[data-special-mobile-premium-v18] .problem-card p,body[data-special-mobile-premium-v18] .area-card p,body[data-special-mobile-premium-v18] .process-card p,body[data-special-mobile-premium-v18] .why-card p{font-size:14.5px!important;line-height:1.58!important}
  body[data-special-mobile-premium-v18] .area-icon{min-width:48px!important;height:48px!important;border-radius:13px!important}

  body[data-special-mobile-premium-v18] .authority-rail{grid-template-columns:1fr!important;gap:8px!important;margin:0 0 22px!important;padding:14px!important;border-radius:18px!important}
  body[data-special-mobile-premium-v18] .authority-item{display:grid!important;grid-template-columns:auto 1fr!important;gap:8px!important;align-items:start!important;padding:10px 11px!important;border-radius:12px!important;background:#f8fafc!important}
  body[data-special-mobile-premium-v18] .authority-item strong{font-size:14px!important}.authority-item span{font-size:13px!important;line-height:1.45!important}

  body[data-special-mobile-premium-v18] .comparison-wrap{overflow:visible!important}
  body[data-special-mobile-premium-v18] .comparison{display:grid!important;gap:10px!important;background:transparent!important;border:0!important}
  body[data-special-mobile-premium-v18] .comparison-head{display:none!important}
  body[data-special-mobile-premium-v18] .comparison-row{display:grid!important;grid-template-columns:1fr!important;gap:0!important;border:1px solid #dfe6eb!important;border-radius:18px!important;background:#fff!important;overflow:hidden!important;box-shadow:0 8px 22px rgba(15,23,42,.045)!important}
  body[data-special-mobile-premium-v18] .comparison-row>div{display:block!important;padding:12px 14px!important;border:0!important;border-bottom:1px solid #edf1f4!important;font-size:13px!important;line-height:1.5!important}
  body[data-special-mobile-premium-v18] .comparison-row>div:first-child{background:#f8fafc!important;color:#0f172a!important;font-size:15px!important;font-weight:800!important}
  body[data-special-mobile-premium-v18] .comparison-row>div:last-child{border-bottom:0!important;background:#eef9f2!important;color:#174a31!important}
  body[data-special-mobile-premium-v18] .comparison-row>div:not(:first-child)::before{content:attr(data-label);display:block;margin-bottom:4px;color:#7a8797;font-size:10px;font-weight:800;letter-spacing:.06em;text-transform:uppercase}

  body[data-special-mobile-premium-v18] .special-mid-cta{padding:12px 0 20px!important;background:#fff!important}
  body[data-special-mobile-premium-v18] .special-mid-cta>.wrap-wide{width:calc(100% - 24px)!important}
  body[data-special-mobile-premium-v18] .special-mid-cta__inner{grid-template-columns:1fr!important;gap:16px!important;padding:20px!important;border-radius:22px!important}
  body[data-special-mobile-premium-v18] .special-mid-cta__inner strong{font-size:20px!important;line-height:1.2!important}.special-mid-cta__inner span{font-size:14px!important;line-height:1.55!important}
  body[data-special-mobile-premium-v18] .special-mid-cta__inner>a{width:100%!important;min-height:52px!important;border-radius:14px!important;text-align:center!important}

  body[data-special-mobile-premium-v18] .faq-grid{grid-template-columns:1fr!important;gap:18px!important}
  body[data-special-mobile-premium-v18] .faq-list{display:grid!important;gap:8px!important}
  body[data-special-mobile-premium-v18] .faq-list details{border-radius:16px!important;overflow:hidden!important}
  body[data-special-mobile-premium-v18] .faq-list summary{min-height:54px!important;display:flex!important;align-items:center!important;padding:0 14px!important;font-size:14px!important}
  body[data-special-mobile-premium-v18] .faq-list details p{font-size:14px!important;line-height:1.58!important}

  body[data-special-mobile-premium-v18] .floating-whatsapp{display:none!important}
  .special-mobile-dock{position:fixed;left:10px;right:10px;bottom:calc(8px + env(safe-area-inset-bottom));z-index:120;display:grid;grid-template-columns:.8fr 1.2fr;gap:7px;padding:7px;border:1px solid rgba(15,23,42,.09);border-radius:20px;background:rgba(255,255,255,.97);box-shadow:0 20px 55px rgba(15,23,42,.18);backdrop-filter:blur(18px) saturate(160%)}
  .special-mobile-dock a{min-height:50px;display:flex;align-items:center;justify-content:center;border-radius:14px;text-decoration:none;font-size:13px;font-weight:850;text-align:center}
  .special-mobile-dock__systems{border:1px solid #dce5df;background:#f7faf8;color:#17452d}.special-mobile-dock__wa{background:#059669;color:#fff;box-shadow:0 9px 20px rgba(5,150,105,.22)}
}
@media(min-width:721px){.special-mobile-dock{display:none!important}}
</style>`;

function processHome() {
  if (!fs.existsSync(HOME)) throw new Error('MOBILE PREMIUM V18: dist/index.html missing');
  let html = fs.readFileSync(HOME, 'utf8');
  html = injectMarker(html, 'data-mobile-premium-v18');
  if (!html.includes('mobile-premium-v18-css')) html = html.replace('</head>', `${mobileCss}\n</head>`);
  for (const token of ['data-mobile-premium-v18','finance-pillars','high-ticket-bridge','mobile-commerce-bar','hero-search--premium']) {
    if (!html.includes(token)) throw new Error(`MOBILE PREMIUM V18 HOME: required token missing: ${token}`);
  }
  fs.writeFileSync(HOME, html, 'utf8');
}

function processSpecial() {
  if (!fs.existsSync(SPECIAL)) throw new Error('MOBILE PREMIUM V18: special route missing');
  let html = fs.readFileSync(SPECIAL, 'utf8');
  html = injectMarker(html, 'data-special-mobile-premium-v18');
  if (!html.includes('mobile-premium-v18-css')) html = html.replace('</head>', `${mobileCss}\n</head>`);
  if (!html.includes('class="special-mobile-dock"')) {
    const dock = `<nav class="special-mobile-dock" aria-label="Mobil hızlı işlemler"><a class="special-mobile-dock__systems" href="/sablonlar" data-cta="special_mobile_systems" data-location="mobile_dock">Hazır Sistemler</a><a class="special-mobile-dock__wa" href="${specialWa}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="special_mobile_whatsapp" data-location="mobile_dock">WhatsApp'tan Yazın</a></nav>`;
    html = html.replace('</body>', `${dock}\n</body>`);
  }
  for (const token of ['data-special-mobile-premium-v18','hero-proof-premium','authority-rail','special-mobile-dock','905393333303']) {
    if (!html.includes(token)) throw new Error(`MOBILE PREMIUM V18 SPECIAL: required token missing: ${token}`);
  }
  fs.writeFileSync(SPECIAL, html, 'utf8');
}

processHome();
processSpecial();
console.log('MOBILE PREMIUM V18 PASS — mobile-first dual funnel, premium card hierarchy, readable special-page cockpit, sticky conversion docks and safe-area behavior applied.');
