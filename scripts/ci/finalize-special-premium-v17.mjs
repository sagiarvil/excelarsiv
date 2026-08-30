#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(file)) throw new Error('SPECIAL PREMIUM V17: dist route missing');
let html = fs.readFileSync(file, 'utf8');

const WA_PHONE = '905393333303';
const wa = (message) => `https://wa.me/${WA_PHONE}?text=${encodeURIComponent(message)}`;
const links = {
  hero: wa('Merhaba Barış Bey, excelarsiv.com üzerinden ulaşıyorum. Şirketimiz için özel Excel finans karar sistemi hakkında görüşmek istiyorum.'),
  mid: wa('Merhaba Barış Bey, şirketimizin mevcut tablosu ve finansal darboğazı hakkında 15 dakikalık bir ön teşhis görüşmesi yapmak istiyorum.'),
  bottom: wa('Merhaba Barış Bey, sürecimizi anlatmak ve özel Excel karar motoru mimarimizi netleştirmek istiyorum.'),
  floating: wa('Merhaba Barış Bey, mobilden ulaşıyorum. Özel sistem ihtiyacımızı 15 dakikada netleştirelim.'),
};

const requireReplace = (pattern, replacement, label) => {
  const next = html.replace(pattern, replacement);
  if (next === html) throw new Error(`SPECIAL PREMIUM V17: replacement missed: ${label}`);
  html = next;
};

if (html.includes('data-special-premium-v17')) {
  console.log('SPECIAL PREMIUM V17: already applied');
  process.exit(0);
}

requireReplace(/<body\b([^>]*)>/u, '<body$1 data-special-premium-v17>', 'body marker');
requireReplace(/<title>[\s\S]*?<\/title>/u, '<title>İhtiyaca Özel Excel Finans ve Karar Sistemleri | Excel Arşiv</title>', 'title');

// Keep the page's approved content; only rebuild conversion paths and presentation.
requireReplace(/<a class="btn btn-primary" href="\/iletisim">İhtiyacımı anlatmak istiyorum <span aria-hidden="true">→<\/span><\/a>/u, `<a class="btn btn-primary btn-whatsapp" href="${links.hero}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="hero_whatsapp" data-location="top">WhatsApp'tan Yazın — 15 Dakikada İhtiyacı Çıkaralım <span aria-hidden="true">→</span></a>`, 'hero WhatsApp CTA');
html = html.replace(/<a class="btn btn-primary nav-cta" href="\/iletisim">[^<]*<\/a>/u, `<a class="btn btn-primary nav-cta btn-whatsapp" href="${links.hero}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="nav_whatsapp" data-location="header">WhatsApp'tan Yazın</a>`);
html = html.replace(/<a href="\/iletisim">İhtiyacımı anlat<\/a>/u, `<a href="${links.hero}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="mobile_menu_whatsapp" data-location="header">WhatsApp'tan Yazın</a>`);
html = html.replace(/<a class="btn btn-primary" href="\/iletisim">Mevcut dosyamı anlatmak istiyorum<\/a>/u, `<a class="btn btn-primary btn-whatsapp" href="${links.mid}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="fit_whatsapp" data-location="problem_solver">WhatsApp'tan Yazın — İhtiyacı Netleştirelim</a>`);
html = html.replace(/<a class="btn btn-primary" href="\/iletisim">İhtiyacımı anlat <span aria-hidden="true">→<\/span><\/a>/u, `<a class="btn btn-primary btn-whatsapp" href="${links.bottom}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="footer_whatsapp" data-location="bottom">WhatsApp'tan Yazın — 15 Dakikada İhtiyacı Çıkaralım <span aria-hidden="true">→</span></a>`);

const premiumProof = `
<section class="hero-proof-premium" aria-label="ExcelArşiv finansal sistem yaklaşımı">
  <div class="wrap-wide hero-proof-premium__grid">
    <article class="hero-proof-premium__card is-green"><div class="hero-proof-premium__icon" aria-hidden="true">↗</div><div><strong>Ticari bankacılık bakışı</strong><span>Nakit, kredi, limit ve risk ilişkisini aynı karar çerçevesinde okuruz.</span></div></article>
    <article class="hero-proof-premium__card is-blue"><div class="hero-proof-premium__icon" aria-hidden="true">◎</div><div><strong>İşletme saha bilgisi</strong><span>Talebin görünen kısmını değil, süreci tıkayan gerçek nedeni ayırırız.</span></div></article>
    <article class="hero-proof-premium__card is-amber"><div class="hero-proof-premium__icon" aria-hidden="true">∑</div><div><strong>Muhasebe gerçekliği</strong><span>Kayıt, mutabakat, dönem akışı ve raporlamayı birbirinden koparmayız.</span></div></article>
    <article class="hero-proof-premium__card is-violet"><div class="hero-proof-premium__icon" aria-hidden="true">◇</div><div><strong>Amaca uygun kapsam</strong><span>Kullanmayacağınız modülü eklemez, gerekli kontrol ve çıktıya odaklanırız.</span></div></article>
  </div>
</section>`;

const heroStart = html.indexOf('<section class="hero">');
const heroEnd = heroStart >= 0 ? html.indexOf('</section>', heroStart) : -1;
if (heroStart < 0 || heroEnd < 0) throw new Error('SPECIAL PREMIUM V17: hero boundary missing');
const heroInsert = heroEnd + '</section>'.length;
html = `${html.slice(0, heroInsert)}\n${premiumProof}${html.slice(heroInsert)}`;

const authorityRail = `<div class="authority-rail" aria-label="ExcelArşiv uzmanlık özeti"><div class="authority-item"><strong>17 Yıl:</strong><span>Ticari bankacılık, kredi, limit ve risk bakışı</span></div><div class="authority-item"><strong>Banka ↔ İşletme:</strong><span>Finans, muhasebe ve saha dilini aynı modelde birleştirme</span></div><div class="authority-item"><strong>ERP → CFO Kararı:</strong><span>Dağınık veriyi rapordan aksiyona taşıyan yönetim sistemi</span></div></div>`;
if (!html.includes('aria-label="ExcelArşiv uzmanlık özeti"')) requireReplace('<div class="comparison-wrap">', `${authorityRail}\n<div class="comparison-wrap">`, 'authority rail');

const midCta = `<section class="special-mid-cta" aria-label="Ön teşhis görüşmesi"><div class="wrap-wide special-mid-cta__inner"><div><strong>Hazır sistem değil, işletmenize göre karar akışı mı gerekiyor?</strong><span>Mevcut tablonuzu veya darboğazı anlatın; gerekli finansal sistemi birlikte netleştirelim.</span></div><a href="${links.mid}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="mid_whatsapp" data-location="problem_solver">WhatsApp'tan Yazın — İhtiyacı Netleştirelim</a></div></section>`;
if (!html.includes('data-cta="mid_whatsapp"')) requireReplace('<section class="section" id="surec">', `${midCta}\n<section class="section" id="surec">`, 'mid WhatsApp CTA');

const floating = `<a class="floating-whatsapp" href="${links.floating}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="floating_whatsapp" data-location="mobile_sticky" aria-label="WhatsApp ile 15 dakikalık ön teşhis görüşmesi başlat"><span aria-hidden="true">◉</span><strong>15 Dk'da Teşhis</strong></a>`;
html = html.replace('</body>', `${floating}\n</body>`);

const schema = {
  '@context':'https://schema.org',
  '@graph':[
    {'@type':'Organization','@id':'https://excelarsiv.com/#organization',name:'Excel Arşiv',url:'https://excelarsiv.com/'},
    {'@type':'Person','@id':'https://excelarsiv.com/#baris-bagirlar',name:'Barış Bağırlar',jobTitle:'Ticari Bankacılık Uzmanı ve Finansal Sistem Mimarı',url:'https://excelarsiv.com/ozel-excel-sistemleri',worksFor:{'@id':'https://excelarsiv.com/#organization'},description:'17 yıllık ticari bankacılık ve saha finans deneyimiyle şirketlerin nakit, kredi, limit-risk, maliyet, bütçe, yatırım ve yönetim raporlama ihtiyaçlarını özel Excel karar sistemlerine dönüştürür.'},
    {'@type':'ProfessionalService','@id':'https://excelarsiv.com/ozel-excel-sistemleri#professional-service',name:'Excel Arşiv - Özel Excel Karar Sistemleri',url:'https://excelarsiv.com/ozel-excel-sistemleri',provider:{'@id':'https://excelarsiv.com/#baris-bagirlar'},brand:{'@id':'https://excelarsiv.com/#organization'},areaServed:{'@type':'Country',name:'Türkiye'},description:'17 yıllık ticari bankacılık ve saha finans deneyimiyle şirketlere özel dinamik nakit akışı, banka limit-risk, maliyet, bütçe, değerleme ve ERP entegre CFO karar sistemleri tasarımı.',knowsAbout:['Dinamik Nakit Akışı (13-Week Cash Flow)','Banka Limit-Risk ve Teminat Havuzu','Birim Maliyet ve Dinamik Fiyatlama','Senaryolu Bütçe ve Rolling Forecast','Yatırım Fizibilitesi ve DCF Değerleme','ERP Entegre CFO Yönetim Kokpiti']}
  ]
};
html = html.replace('</head>', `<script type="application/ld+json" id="special-professional-schema-v17">${JSON.stringify(schema)}</script>\n</head>`);

const css = `<style id="special-premium-v17-css">
body[data-special-premium-v17]{--premium-navy:#0f172a;--premium-green:#059669;--premium-shadow:0 22px 64px rgba(15,23,42,.10)}
body[data-special-premium-v17] .hero{padding:54px 0 50px!important;background:linear-gradient(100deg,#fff 0%,#fff 48%,#f3faf6 100%)!important;overflow:hidden}
body[data-special-premium-v17] .hero>.wrap-wide{width:min(1480px,calc(100% - 72px))!important}
body[data-special-premium-v17] .hero-grid{grid-template-columns:minmax(0,.9fr) minmax(560px,1.1fr)!important;gap:56px!important;align-items:center!important}
body[data-special-premium-v17] .hero-copy{padding-block:14px!important}
body[data-special-premium-v17] .hero .eyebrow{margin-bottom:20px!important;padding:9px 13px!important;font-size:14px!important;line-height:1.2!important;background:#f8fffb!important;border-color:#b9dec8!important}
body[data-special-premium-v17] .hero h1{max-width:730px!important;margin:0!important;font-size:clamp(48px,4.25vw,64px)!important;line-height:1.01!important;font-weight:760!important;letter-spacing:-.052em!important;color:#0f172a!important;text-wrap:balance}
body[data-special-premium-v17] .hero-lead{max-width:720px!important;margin-top:24px!important;font-size:19px!important;line-height:1.62!important;color:#475569!important;letter-spacing:-.012em!important}
body[data-special-premium-v17] .hero-actions{margin-top:30px!important;gap:12px!important}
body[data-special-premium-v17] .hero-actions .btn{min-height:54px!important;padding:0 22px!important;border-radius:13px!important;font-size:15px!important;font-weight:760!important}
body[data-special-premium-v17] .hero-bullets{margin-top:24px!important;gap:12px 22px!important}
body[data-special-premium-v17] .hero-bullets li{font-size:14px!important;color:#425066!important}
body[data-special-premium-v17] .hero-copy>.proof-grid{display:none!important}
body[data-special-premium-v17] .workbook{min-width:0!important;max-width:100%!important;border-radius:22px!important;border-color:#d4e0d8!important;box-shadow:var(--premium-shadow)!important;transform:none!important}
body[data-special-premium-v17] .wb-top{padding:16px 18px!important}.wb-title{font-size:15px!important}.wb-note{font-size:13px!important}
body[data-special-premium-v17] .wb-body{padding:18px!important}.metric-grid{gap:10px!important}.metric{padding:14px!important}.metric small{font-size:13px!important}.metric strong{font-size:21px!important}.metric span{font-size:12.5px!important}.panel h3{font-size:14.5px!important}
.hero-proof-premium{position:relative;padding:0 0 66px;background:#fff}.hero-proof-premium::before{content:"";position:absolute;left:0;right:0;top:0;height:1px;background:linear-gradient(90deg,transparent,#dfe7e1 15%,#dfe7e1 85%,transparent)}.hero-proof-premium__grid{width:min(1320px,calc(100% - 72px));display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin-inline:auto;padding-top:28px}.hero-proof-premium__card{position:relative;min-width:0;display:grid;grid-template-columns:46px minmax(0,1fr);gap:15px;align-items:start;padding:22px 20px;border:1px solid #dde6df;border-radius:18px;background:#fff;box-shadow:0 12px 32px rgba(15,23,42,.06);overflow:hidden}.hero-proof-premium__card::before{content:"";position:absolute;left:0;right:0;top:0;height:4px;background:#059669}.hero-proof-premium__card.is-blue::before{background:#2563eb}.hero-proof-premium__card.is-amber::before{background:#d97706}.hero-proof-premium__card.is-violet::before{background:#7c3aed}.hero-proof-premium__icon{width:46px;height:46px;display:grid;place-items:center;border-radius:13px;background:#ecfdf5;color:#047857;font-size:20px;font-weight:850}.is-blue .hero-proof-premium__icon{background:#eff6ff;color:#1d4ed8}.is-amber .hero-proof-premium__icon{background:#fffbeb;color:#b45309}.is-violet .hero-proof-premium__icon{background:#f5f3ff;color:#6d28d9}.hero-proof-premium__card strong{display:block;color:#0f172a;font-size:17px;line-height:1.25;font-weight:800;letter-spacing:-.02em}.hero-proof-premium__card span{display:block;margin-top:7px;color:#64748b;font-size:15px;line-height:1.5;letter-spacing:-.006em}
.authority-rail{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;max-width:1080px;margin:0 auto 30px;padding:20px;border:1px solid #dbe3eb;border-radius:18px;background:#fff;box-shadow:0 10px 28px rgba(15,23,42,.05)}.authority-item{display:flex;align-items:flex-start;gap:10px;min-width:0;padding:7px 9px;color:#475569;font-size:15px;line-height:1.5}.authority-item strong{flex:0 0 auto;white-space:nowrap;color:#0f172a;font-weight:850}.authority-item span{min-width:0;overflow-wrap:anywhere}.faq-grid{align-items:start!important}.faq-copy,.faq-list{align-self:start!important;min-width:0}.special-mid-cta{padding:0 0 62px;background:#edf9f2}.special-mid-cta__inner{display:flex;align-items:center;justify-content:space-between;gap:32px;padding:27px 30px;border:1px solid #cfe7d8;border-radius:20px;background:#fff;box-shadow:0 14px 36px rgba(15,23,42,.07)}.special-mid-cta__inner>div{display:grid;gap:6px}.special-mid-cta__inner strong{color:#0f172a;font-size:19px;line-height:1.3}.special-mid-cta__inner span{color:#64748b;font-size:15px;line-height:1.5}.special-mid-cta__inner>a{flex:0 0 auto;min-height:50px;display:inline-flex;align-items:center;justify-content:center;padding:0 20px;border-radius:12px;background:#059669;color:#fff;font-size:14px;font-weight:850;text-decoration:none;box-shadow:0 10px 24px rgba(5,150,105,.18)}.btn-whatsapp{background:#059669!important;color:#fff!important;border-color:#059669!important}.btn-whatsapp:hover{background:#047857!important}.floating-whatsapp{display:none;position:fixed;right:16px;bottom:calc(16px + env(safe-area-inset-bottom));z-index:140;align-items:center;gap:8px;min-height:50px;padding:0 16px;border:1px solid rgba(255,255,255,.5);border-radius:999px;background:#059669;color:#fff;font-size:13px;text-decoration:none;box-shadow:0 16px 36px rgba(5,150,105,.3)}.floating-whatsapp span{font-size:18px}.floating-whatsapp strong{font-weight:850}
@media(max-width:1180px){body[data-special-premium-v17] .hero>.wrap-wide,.hero-proof-premium__grid{width:calc(100% - 48px)!important}body[data-special-premium-v17] .hero-grid{grid-template-columns:minmax(0,.94fr) minmax(470px,1.06fr)!important;gap:34px!important}body[data-special-premium-v17] .hero h1{font-size:clamp(42px,4.7vw,54px)!important}.hero-proof-premium__grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:900px){body[data-special-premium-v17] .hero{padding-top:38px!important}body[data-special-premium-v17] .hero-grid{grid-template-columns:1fr!important;gap:32px!important}body[data-special-premium-v17] .hero h1{max-width:820px!important;font-size:clamp(42px,7vw,56px)!important}.workbook{max-width:760px!important;margin-inline:auto!important}.authority-rail{grid-template-columns:1fr}.special-mid-cta__inner{align-items:flex-start;flex-direction:column}.special-mid-cta__inner>a{width:100%}}
@media(max-width:660px){body[data-special-premium-v17] .hero{padding:26px 0 26px!important}body[data-special-premium-v17] .hero>.wrap-wide,.hero-proof-premium__grid{width:calc(100% - 28px)!important}body[data-special-premium-v17] .hero h1{font-size:clamp(34px,9.4vw,43px)!important;line-height:1.02!important}body[data-special-premium-v17] .hero-lead{font-size:16.5px!important;line-height:1.58!important}.hero-actions{display:grid!important}.hero-actions .btn{width:100%!important}.workbook{display:none!important}.hero-proof-premium{padding-bottom:44px}.hero-proof-premium__grid{grid-template-columns:1fr!important;gap:11px!important;padding-top:20px}.hero-proof-premium__card{grid-template-columns:42px minmax(0,1fr);gap:13px;padding:18px 16px}.hero-proof-premium__icon{width:42px;height:42px}.hero-proof-premium__card strong{font-size:16px}.hero-proof-premium__card span{font-size:14px;line-height:1.45}.authority-rail{margin-bottom:22px;padding:15px}.authority-item{display:grid;grid-template-columns:auto minmax(0,1fr);gap:7px}.special-mid-cta{padding-bottom:40px}.special-mid-cta__inner{padding:20px}.floating-whatsapp{display:inline-flex}.footer{padding-bottom:86px!important}}
</style>`;
html = html.replace('</head>', `${css}\n</head>`);

for (const token of ['hero-proof-premium','Ticari bankacılık bakışı','ERP → CFO Kararı:','data-cta="hero_whatsapp"','data-cta="mid_whatsapp"','data-cta="footer_whatsapp"','data-cta="floating_whatsapp"',`https://wa.me/${WA_PHONE}?text=`,'ProfessionalService']) {
  if (!html.includes(token)) throw new Error(`SPECIAL PREMIUM V17: required token missing: ${token}`);
}

const waCount = (html.match(new RegExp(`https://wa\\.me/${WA_PHONE}\\?text=`, 'gu')) || []).length;
if (waCount < 4) throw new Error(`SPECIAL PREMIUM V17: expected at least 4 verified WhatsApp routes, got ${waCount}`);

fs.writeFileSync(file, html, 'utf8');
console.log(`SPECIAL PREMIUM V17 PASS — premium hero typography, spacious full-width proof band, P0 authority rail, aligned FAQ and ${waCount} verified WhatsApp CTAs.`);
