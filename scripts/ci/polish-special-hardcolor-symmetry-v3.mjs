#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(file)) throw new Error('SPECIAL HARDCOLOR V3: build output missing');

let html = fs.readFileSync(file, 'utf8');
if (!html.includes('data-living-workbook-v1') || !html.includes('id="special-symmetry-human-v2"')) {
  throw new Error('SPECIAL HARDCOLOR V3: Living Workbook + symmetry v2 contract missing');
}

html = html.replace(/<style\b(?=[^>]*\bid=["']special-hardcolor-symmetry-v3["'])[^>]*>[\s\S]*?<\/style>/gi, '');

for (const [from, to] of [
  ['Tekrar işi azaltırız', 'Tekrarlanan işleri azaltırız'],
  ['Yönetici gözüyle tasarlarız', 'Yönetim için sade ekranlar kurarız'],
  ['Rapor kalabalığı değil; karar vermeyi hızlandıran ekranlar üretiriz.', 'Gereksiz rapor kalabalığı yerine, karar vermeyi kolaylaştıran ekranlar kurarız.'],
  ['Brief doğruysa iyi sonuç verir; fakat yanlış tanımlanmış iş kuralını doğru kodlamak yine yanlış sistemi hızlandırabilir.', 'İhtiyaç doğru tanımlanmışsa iyi sonuç verir. Ancak yanlış tanımlanan bir iş kuralını eksiksiz kodlamak, yanlış süreci yalnızca hızlandırır.'],
  ["Brief'e uyar", 'Tanımlanan ihtiyaca uyar'],
  ['İstenen yapıyı takip eder.', 'Tanımlanan ekran ve işlevleri takip eder.'],
  ['İhtiyacı ayrıştır', 'İhtiyacı ayrıştırıyoruz'],
  ['Görünen talep ile gerçek karar problemini birbirinden ayır.', 'Görünen talebi, asıl iş ihtiyacından ve karar probleminden ayırıyoruz.'],
  ['İş mantığını yaz', 'İş kurallarını netleştiriyoruz'],
  ['Kim ne giriyor, hangi kayıt ne anlama geliyor, hangi istisna var netleştir.', 'Kimin hangi veriyi girdiğini, kayıtların ne anlama geldiğini ve istisnaları açıkça tanımlıyoruz.'],
  ['Veri ve hesap motorunu kur', 'Veri ve hesap modelini kuruyoruz'],
  ['Ana veri, işlem, formül, kontrol ve rapor zincirini tek yönlü bağla.', 'Ana veriyi, işlemleri, formülleri, kontrolleri ve raporları tek ve tutarlı bir akışta bağlıyoruz.'],
  ['Kullanıcı ekranını tasarla', 'Kullanıcı ekranlarını tasarlıyoruz'],
  ['Veri girişi, rapor, uyarı ve yönetim alanlarını açıkça ayır.', 'Veri girişi, rapor, uyarı ve yönetim alanlarını birbirinden açıkça ayırıyoruz.'],
  ['Sınır senaryolarını doğrula', 'Sınır senaryolarını test ediyoruz'],
  ['Normal kullanım kadar eksik veri, yanlış seçim ve kapanış senaryolarını da kontrol et.', 'Normal kullanımla birlikte eksik veri, yanlış seçim ve dönem kapanışı gibi sınır senaryolarını da test ediyoruz.'],
  ['KULLANICI UX', 'KULLANIM'],
  ['STABİLİTE', 'SAĞLAMLIK'],
  ['SELF-SERVICE', 'KULLANIM REHBERİ'],
  ['Dosya değil; günlük işte kullanılabilecek bir çalışma sistemi.', 'Yalnızca bir dosya değil, günlük işte kullanılabilecek bir çalışma sistemi.'],
  ['Hazır şablon yetmiyor, büyük yazılım gereğinden ağır kalıyorsa doğru yerde olabilirsiniz.', 'Hazır şablon yetersiz kalıyor, büyük yazılım ise gereğinden ağır geliyorsa özel bir Excel sistemi doğru seçenek olabilir.'],
]) html = html.replaceAll(from, to);

html = html.replace(/<section class="section">\s*<div class="wrap deliver-grid">/i, '<section class="section delivery" id="teslim">\n      <div class="wrap deliver-grid">');

const css = `
<style id="special-hardcolor-symmetry-v3" data-special-hardcolor="finance-six-v3">
body[data-living-workbook-v1]{
  --hc-green:#0f7a45;--hc-green-d:#09552f;--hc-green-s:#e8f5ee;
  --hc-blue:#1f66d1;--hc-blue-d:#174b9c;--hc-blue-s:#eaf1ff;
  --hc-amber:#c98212;--hc-amber-d:#8b5908;--hc-amber-s:#fff3dc;
  --hc-coral:#cc574b;--hc-coral-d:#923b33;--hc-coral-s:#fff0ed;
  --hc-violet:#6f55c7;--hc-violet-d:#503a9c;--hc-violet-s:#f1edff;
  --hc-teal:#128087;--hc-teal-d:#0d5e63;--hc-teal-s:#e8f7f7;
  --hc-navy:#173b63;--hc-ink:#17202b;--hc-line:#d4dbe2;--hc-shadow:0 16px 38px rgba(23,42,63,.08);
}
body[data-living-workbook-v1] .section{position:relative!important;padding:82px 0!important;background-color:transparent!important}
body[data-living-workbook-v1] .section-head{grid-template-columns:1fr!important;gap:12px!important;justify-items:center!important;text-align:center!important;max-width:930px!important;margin:0 auto 42px!important}
body[data-living-workbook-v1] .section-head>div:last-child{max-width:930px!important}
body[data-living-workbook-v1] .section-copy{margin-inline:auto!important;max-width:820px!important;text-wrap:pretty}
body[data-living-workbook-v1] .section h2,body[data-living-workbook-v1] .section-title{text-wrap:balance}
body[data-living-workbook-v1] .kicker{display:inline-flex!important;width:max-content!important;align-items:center!important;padding:7px 10px!important;border-radius:6px!important;background:var(--hc-green)!important;color:#fff!important;font-weight:800!important;letter-spacing:.075em!important}
body[data-living-workbook-v1] #ihtiyaclar .kicker{background:var(--hc-blue)!important}
body[data-living-workbook-v1] #karsilastirma .kicker{background:var(--hc-amber)!important}
body[data-living-workbook-v1] #mimariler .kicker{background:var(--hc-violet)!important}
body[data-living-workbook-v1] .value-band .kicker{background:var(--hc-teal)!important}
body[data-living-workbook-v1] .process .kicker{background:var(--hc-coral)!important}
body[data-living-workbook-v1] .delivery .kicker{background:var(--hc-green)!important}
body[data-living-workbook-v1] .faq .kicker{background:var(--hc-blue)!important}
body[data-living-workbook-v1] .cta .kicker{background:var(--hc-navy)!important}

body[data-living-workbook-v1] .hero-grid{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:48px!important}
body[data-living-workbook-v1] .hero::after{display:block!important;content:""!important;position:absolute!important;top:0!important;right:0!important;width:76%!important;height:6px!important;background:linear-gradient(90deg,var(--hc-blue) 0 20%,var(--hc-amber) 20% 40%,var(--hc-coral) 40% 60%,var(--hc-violet) 60% 80%,var(--hc-teal) 80% 100%)!important}
body[data-living-workbook-v1] .workbook::before{height:6px!important;background:linear-gradient(90deg,var(--hc-green) 0 20%,var(--hc-blue) 20% 40%,var(--hc-amber) 40% 60%,var(--hc-coral) 60% 80%,var(--hc-violet) 80% 100%)!important}
body[data-living-workbook-v1] .kpi:nth-child(1){background:var(--hc-green-s)!important;border-top-color:var(--hc-green)!important}
body[data-living-workbook-v1] .kpi:nth-child(2){background:var(--hc-blue-s)!important;border-top-color:var(--hc-blue)!important}
body[data-living-workbook-v1] .kpi:nth-child(3){background:var(--hc-amber-s)!important;border-top-color:var(--hc-amber)!important}
body[data-living-workbook-v1] .kpi:nth-child(4){background:var(--hc-violet-s)!important;border-top-color:var(--hc-violet)!important}

body[data-living-workbook-v1] .decision-grid{grid-template-columns:repeat(4,minmax(0,1fr))!important;gap:0!important;align-items:stretch!important;box-shadow:var(--hc-shadow)!important}
body[data-living-workbook-v1] .decision-item{min-height:174px!important;padding:25px 22px!important}
body[data-living-workbook-v1] .decision-item:nth-child(1){background:var(--hc-green-s)!important}body[data-living-workbook-v1] .decision-item:nth-child(2){background:var(--hc-blue-s)!important}body[data-living-workbook-v1] .decision-item:nth-child(3){background:var(--hc-amber-s)!important}body[data-living-workbook-v1] .decision-item:nth-child(4){background:var(--hc-coral-s)!important}
body[data-living-workbook-v1] .decision-item small{display:inline-flex!important;width:max-content!important;padding:5px 7px!important;border-radius:5px!important;color:#fff!important;background:var(--hc-green)!important}
body[data-living-workbook-v1] .decision-item:nth-child(2) small{background:var(--hc-blue)!important;color:#fff!important}body[data-living-workbook-v1] .decision-item:nth-child(3) small{background:var(--hc-amber)!important;color:#fff!important}body[data-living-workbook-v1] .decision-item:nth-child(4) small{background:var(--hc-coral)!important;color:#fff!important}

body[data-living-workbook-v1] #teshis{box-shadow:inset 0 5px 0 var(--hc-green)!important;background-image:linear-gradient(rgba(232,245,238,.62),rgba(255,255,255,.40))!important}
body[data-living-workbook-v1] .diagnosis-shell{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:42px!important;align-items:start!important}
body[data-living-workbook-v1] .diagnosis-head,body[data-living-workbook-v1] .diagnosis-row{grid-template-columns:repeat(3,minmax(0,1fr))!important}
body[data-living-workbook-v1] .diagnosis-head>div:first-child{background:#37414d!important;color:#fff!important}body[data-living-workbook-v1] .diagnosis-head>div:nth-child(2){background:var(--hc-blue)!important;color:#fff!important}body[data-living-workbook-v1] .diagnosis-head>div:nth-child(3){background:var(--hc-green)!important;color:#fff!important}
body[data-living-workbook-v1] .diagnosis-row>div:nth-child(1){background:#f7f8f9!important}body[data-living-workbook-v1] .diagnosis-row>div:nth-child(2){background:var(--hc-blue-s)!important}body[data-living-workbook-v1] .diagnosis-row>div:nth-child(3){background:var(--hc-green-s)!important}

body[data-living-workbook-v1] #ihtiyaclar{box-shadow:inset 0 5px 0 var(--hc-blue)!important;background-image:linear-gradient(rgba(234,241,255,.50),rgba(255,255,255,.25))!important}
body[data-living-workbook-v1] .intent-grid{grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:18px!important;align-items:stretch!important}
body[data-living-workbook-v1] .intent-card{height:100%!important;min-height:315px!important;padding:30px 26px 25px!important;border-radius:10px!important;box-shadow:var(--hc-shadow)!important;border-color:var(--hc-line)!important}
body[data-living-workbook-v1] .intent-card::before{left:0!important;right:0!important;top:0!important;bottom:auto!important;width:auto!important;height:6px!important}
body[data-living-workbook-v1] .intent-card:nth-child(1)::before{background:var(--hc-green)!important}body[data-living-workbook-v1] .intent-card:nth-child(2)::before{background:var(--hc-blue)!important}body[data-living-workbook-v1] .intent-card:nth-child(3)::before{background:var(--hc-amber)!important}body[data-living-workbook-v1] .intent-card:nth-child(4)::before{background:var(--hc-coral)!important}body[data-living-workbook-v1] .intent-card:nth-child(5)::before{background:var(--hc-violet)!important}body[data-living-workbook-v1] .intent-card:nth-child(6)::before{background:var(--hc-teal)!important}
body[data-living-workbook-v1] .intent-card .code{display:inline-flex!important;width:max-content!important;padding:5px 7px!important;border-radius:5px!important;background:var(--hc-green)!important;color:#fff!important}body[data-living-workbook-v1] .intent-card:nth-child(2) .code{background:var(--hc-blue)!important;color:#fff!important}body[data-living-workbook-v1] .intent-card:nth-child(3) .code{background:var(--hc-amber)!important;color:#fff!important}body[data-living-workbook-v1] .intent-card:nth-child(4) .code{background:var(--hc-coral)!important;color:#fff!important}body[data-living-workbook-v1] .intent-card:nth-child(5) .code{background:var(--hc-violet)!important;color:#fff!important}body[data-living-workbook-v1] .intent-card:nth-child(6) .code{background:var(--hc-teal)!important;color:#fff!important}
body[data-living-workbook-v1] .intent-card h3{min-height:78px!important}body[data-living-workbook-v1] .intent-answer{min-height:98px!important}

body[data-living-workbook-v1] #karsilastirma{box-shadow:inset 0 5px 0 var(--hc-amber)!important;background-image:linear-gradient(rgba(255,243,220,.58),rgba(255,255,255,.30))!important}
body[data-living-workbook-v1] .difference-grid{grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:18px!important;align-items:stretch!important}
body[data-living-workbook-v1] .difference-card{height:100%!important;min-height:238px!important;border-radius:10px!important;box-shadow:var(--hc-shadow)!important}
body[data-living-workbook-v1] .difference-card small{display:inline-flex!important;padding:5px 7px!important;border-radius:5px!important;background:#68717b!important;color:#fff!important}body[data-living-workbook-v1] .difference-card:nth-child(2) small{background:var(--hc-blue)!important}body[data-living-workbook-v1] .difference-card:nth-child(3) small{background:var(--hc-green)!important}
body[data-living-workbook-v1] .comparison-top,body[data-living-workbook-v1] .compare-row{grid-template-columns:repeat(4,minmax(0,1fr))!important}
body[data-living-workbook-v1] .comparison-top>div:first-child{background:#37414d!important;color:#fff!important}body[data-living-workbook-v1] .comparison-top>div:nth-child(2){background:#7a828b!important;color:#fff!important}body[data-living-workbook-v1] .comparison-top>div:nth-child(3){background:var(--hc-blue)!important;color:#fff!important}body[data-living-workbook-v1] .comparison-top .good{background:var(--hc-green)!important;color:#fff!important}

body[data-living-workbook-v1] #mimariler{box-shadow:inset 0 5px 0 var(--hc-violet)!important;background-image:linear-gradient(rgba(241,237,255,.50),rgba(255,255,255,.24))!important}
body[data-living-workbook-v1] .architecture{grid-template-columns:1fr!important;gap:28px!important}
body[data-living-workbook-v1] .arch-index{position:static!important;display:grid!important;grid-template-columns:repeat(5,minmax(0,1fr))!important;border-radius:10px!important}
body[data-living-workbook-v1] .arch-index strong{grid-column:1/-1!important;background:var(--hc-navy)!important;color:#fff!important;text-align:center!important}
body[data-living-workbook-v1] .arch-index a{justify-content:center!important;gap:9px!important;text-align:center!important;background:#fff!important}
body[data-living-workbook-v1] .arch-stack{gap:16px!important}body[data-living-workbook-v1] .arch-card{grid-template-columns:64px minmax(0,1fr) 300px!important;gap:24px!important;min-height:166px!important;border-radius:10px!important;border-left:6px solid var(--hc-green)!important;box-shadow:var(--hc-shadow)!important}body[data-living-workbook-v1] .arch-card:nth-of-type(2){border-left-color:var(--hc-blue)!important}body[data-living-workbook-v1] .arch-card:nth-of-type(3){border-left-color:var(--hc-amber)!important}body[data-living-workbook-v1] .arch-card:nth-of-type(4){border-left-color:var(--hc-coral)!important}body[data-living-workbook-v1] .arch-card:nth-of-type(5){border-left-color:var(--hc-teal)!important}
body[data-living-workbook-v1] .arch-no{width:48px!important;height:48px!important;display:grid!important;place-items:center!important;border-radius:50%!important;background:var(--hc-green)!important;color:#fff!important;font-size:18px!important}body[data-living-workbook-v1] .arch-card:nth-of-type(2) .arch-no{background:var(--hc-blue)!important;color:#fff!important}body[data-living-workbook-v1] .arch-card:nth-of-type(3) .arch-no{background:var(--hc-amber)!important;color:#fff!important}body[data-living-workbook-v1] .arch-card:nth-of-type(4) .arch-no{background:var(--hc-coral)!important;color:#fff!important}body[data-living-workbook-v1] .arch-card:nth-of-type(5) .arch-no{background:var(--hc-teal)!important;color:#fff!important}

body[data-living-workbook-v1] .value-band{box-shadow:inset 0 5px 0 var(--hc-teal)!important;background-image:linear-gradient(rgba(232,247,247,.58),rgba(255,255,255,.28))!important}
body[data-living-workbook-v1] .value-grid{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:42px!important;align-items:stretch!important}body[data-living-workbook-v1] .value-list{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:14px!important;align-items:stretch!important}body[data-living-workbook-v1] .value-item{height:100%!important;min-height:132px!important;border-radius:10px!important;box-shadow:var(--hc-shadow)!important}
body[data-living-workbook-v1] .value-icon{background:var(--hc-green)!important;color:#fff!important}body[data-living-workbook-v1] .value-item:nth-child(2) .value-icon{background:var(--hc-blue)!important;color:#fff!important}body[data-living-workbook-v1] .value-item:nth-child(3) .value-icon{background:var(--hc-amber)!important;color:#fff!important}body[data-living-workbook-v1] .value-item:nth-child(4) .value-icon{background:var(--hc-violet)!important;color:#fff!important}

body[data-living-workbook-v1] .process{box-shadow:inset 0 5px 0 var(--hc-coral)!important;background-image:linear-gradient(rgba(255,240,237,.55),rgba(255,255,255,.24))!important}
body[data-living-workbook-v1] .flow{grid-template-columns:repeat(5,minmax(0,1fr))!important;align-items:stretch!important;box-shadow:var(--hc-shadow)!important}
body[data-living-workbook-v1] .flow-step{height:100%!important;min-height:232px!important;padding:28px 22px!important;background:#fff!important}body[data-living-workbook-v1] .flow-step small{display:inline-flex!important;width:max-content!important;padding:5px 7px!important;border-radius:5px!important;background:var(--hc-green)!important;color:#fff!important}body[data-living-workbook-v1] .flow-step:nth-child(2) small{background:var(--hc-blue)!important;color:#fff!important}body[data-living-workbook-v1] .flow-step:nth-child(3) small{background:var(--hc-amber)!important;color:#fff!important}body[data-living-workbook-v1] .flow-step:nth-child(4) small{background:var(--hc-coral)!important;color:#fff!important}body[data-living-workbook-v1] .flow-step:nth-child(5) small{background:var(--hc-teal)!important;color:#fff!important}body[data-living-workbook-v1] .flow-step strong{min-height:42px!important}

body[data-living-workbook-v1] .delivery{box-shadow:inset 0 5px 0 var(--hc-green)!important;background-image:linear-gradient(rgba(232,245,238,.42),rgba(255,255,255,.20))!important}
body[data-living-workbook-v1] .deliver-grid{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:32px!important;align-items:stretch!important}body[data-living-workbook-v1] .deliver-grid>div,body[data-living-workbook-v1] .proof-card{height:100%!important}body[data-living-workbook-v1] .deliver-list{height:calc(100% - 92px)!important;box-shadow:var(--hc-shadow)!important}body[data-living-workbook-v1] .deliver-row{min-height:78px!important}body[data-living-workbook-v1] .deliver-check{background:var(--hc-green)!important;color:#fff!important}body[data-living-workbook-v1] .deliver-row:nth-child(2) .deliver-check{background:var(--hc-blue)!important;color:#fff!important}body[data-living-workbook-v1] .deliver-row:nth-child(3) .deliver-check{background:var(--hc-amber)!important;color:#fff!important}body[data-living-workbook-v1] .deliver-row:nth-child(4) .deliver-check{background:var(--hc-violet)!important;color:#fff!important}body[data-living-workbook-v1] .deliver-row:nth-child(5) .deliver-check{background:var(--hc-teal)!important;color:#fff!important}

body[data-living-workbook-v1] .faq{box-shadow:inset 0 5px 0 var(--hc-blue)!important;background-image:linear-gradient(rgba(234,241,255,.46),rgba(255,255,255,.22))!important}
body[data-living-workbook-v1] .faq-grid{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:42px!important;align-items:start!important}body[data-living-workbook-v1] .faq-list details{padding-left:14px!important;border-left:4px solid var(--hc-green)!important}body[data-living-workbook-v1] .faq-list details:nth-child(2){border-left-color:var(--hc-blue)!important}body[data-living-workbook-v1] .faq-list details:nth-child(3){border-left-color:var(--hc-amber)!important}body[data-living-workbook-v1] .faq-list details:nth-child(4){border-left-color:var(--hc-violet)!important}body[data-living-workbook-v1] .faq-list details:nth-child(5){border-left-color:var(--hc-teal)!important}

body[data-living-workbook-v1] .cta{border-top:0!important;box-shadow:inset 0 6px 0 var(--hc-green)!important;background-image:linear-gradient(110deg,rgba(232,245,238,.88),rgba(234,241,255,.72))!important}
body[data-living-workbook-v1] .cta::after{display:block!important;content:""!important;position:absolute!important;left:0!important;right:0!important;top:0!important;width:100%!important;height:6px!important;background:linear-gradient(90deg,var(--hc-green) 0 20%,var(--hc-blue) 20% 40%,var(--hc-amber) 40% 60%,var(--hc-coral) 60% 80%,var(--hc-violet) 80% 100%)!important}
body[data-living-workbook-v1] .cta-shell{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:36px!important;align-items:stretch!important}body[data-living-workbook-v1] .cta-actions{justify-self:stretch!important;max-width:none!important;padding:26px!important;border-radius:12px!important;background:var(--hc-navy)!important;align-content:center!important}body[data-living-workbook-v1] .cta-actions .btn-primary{background:var(--hc-green)!important;border-color:var(--hc-green)!important}body[data-living-workbook-v1] .cta-actions .btn-secondary{background:#fff!important;border-color:#fff!important;color:var(--hc-ink)!important}body[data-living-workbook-v1] .cta-micro{color:#dbe6f0!important;text-align:center!important}

@media(max-width:1120px){
  body[data-living-workbook-v1] .section-head{max-width:820px!important}
  body[data-living-workbook-v1] .hero-grid,body[data-living-workbook-v1] .diagnosis-shell,body[data-living-workbook-v1] .value-grid,body[data-living-workbook-v1] .deliver-grid,body[data-living-workbook-v1] .faq-grid,body[data-living-workbook-v1] .cta-shell{grid-template-columns:1fr!important;gap:28px!important}
  body[data-living-workbook-v1] .diagnosis-intro{position:static!important}
  body[data-living-workbook-v1] .arch-index{grid-template-columns:repeat(3,minmax(0,1fr))!important}
  body[data-living-workbook-v1] .arch-card{grid-template-columns:56px minmax(0,1fr)!important}body[data-living-workbook-v1] .arch-output{grid-column:2!important}
  body[data-living-workbook-v1] .cta-actions{max-width:620px!important;justify-self:start!important}
}
@media(max-width:860px){
  body[data-living-workbook-v1] .intent-grid{grid-template-columns:repeat(2,minmax(0,1fr))!important}
  body[data-living-workbook-v1] .flow{grid-template-columns:1fr!important}
  body[data-living-workbook-v1] .flow-step{min-height:auto!important;border-left:5px solid var(--hc-green)!important;border-top:1px solid var(--hc-line)!important}
  body[data-living-workbook-v1] .flow-step:nth-child(2){border-left-color:var(--hc-blue)!important}body[data-living-workbook-v1] .flow-step:nth-child(3){border-left-color:var(--hc-amber)!important}body[data-living-workbook-v1] .flow-step:nth-child(4){border-left-color:var(--hc-coral)!important}body[data-living-workbook-v1] .flow-step:nth-child(5){border-left-color:var(--hc-teal)!important}
  body[data-living-workbook-v1] .arch-index{grid-template-columns:repeat(2,minmax(0,1fr))!important}
}
@media(max-width:620px){
  body[data-living-workbook-v1] .section{padding:60px 0!important}
  body[data-living-workbook-v1] .section-head{text-align:left!important;justify-items:start!important;margin-bottom:30px!important}
  body[data-living-workbook-v1] .intent-grid,body[data-living-workbook-v1] .difference-grid,body[data-living-workbook-v1] .value-list,body[data-living-workbook-v1] .decision-grid{grid-template-columns:1fr!important}
  body[data-living-workbook-v1] .intent-card,body[data-living-workbook-v1] .difference-card{min-height:auto!important}
  body[data-living-workbook-v1] .intent-card h3{min-height:0!important}
  body[data-living-workbook-v1] .diagnosis-head{display:none!important}body[data-living-workbook-v1] .diagnosis-row{grid-template-columns:1fr!important}
  body[data-living-workbook-v1] .arch-index{grid-template-columns:1fr!important}body[data-living-workbook-v1] .arch-card{grid-template-columns:44px minmax(0,1fr)!important;padding:20px 17px!important}body[data-living-workbook-v1] .arch-output{grid-column:1/-1!important}
  body[data-living-workbook-v1] .cta-actions{padding:20px!important}
}
</style>`;

if (!html.includes('</head>')) throw new Error('SPECIAL HARDCOLOR V3: </head> missing');
html = html.replace('</head>', `${css}\n</head>`);

const required = [
  'id="special-hardcolor-symmetry-v3"',
  'data-special-hardcolor="finance-six-v3"',
  'id="teslim"',
  '--hc-coral:#cc574b',
  '--hc-violet:#6f55c7',
  '--hc-teal:#128087',
  'İhtiyacı ayrıştırıyoruz',
  'İş kurallarını netleştiriyoruz',
  'Veri ve hesap modelini kuruyoruz',
  'Kullanıcı ekranlarını tasarlıyoruz',
  'Sınır senaryolarını test ediyoruz',
  'KULLANIM REHBERİ',
];
for (const token of required) if (!html.includes(token)) throw new Error(`SPECIAL HARDCOLOR V3: required token missing: ${token}`);

const forbidden = [
  'Brief doğruysa',
  "Brief'e uyar",
  '>İhtiyacı ayrıştır<',
  '>İş mantığını yaz<',
  '>Veri ve hesap motorunu kur<',
  '>Kullanıcı ekranını tasarla<',
  '>Sınır senaryolarını doğrula<',
  '>SELF-SERVICE<',
];
for (const token of forbidden) if (html.includes(token)) throw new Error(`SPECIAL HARDCOLOR V3: stale copy survived: ${token}`);

fs.writeFileSync(file, html);
console.log('SPECIAL HARDCOLOR V3 PASS — six-color finance palette, balanced geometry and natural Turkish copy verified.');
