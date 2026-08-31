#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/index.html');
if (!fs.existsSync(file)) throw new Error('HOME CONVERSION V20: dist/index.html missing');
let html = fs.readFileSync(file, 'utf8');

for (const marker of ['data-dual-funnel-home-v17','data-mobile-premium-v18','data-desktop-premium-v19']) {
  if (!html.includes(marker)) throw new Error(`HOME CONVERSION V20: prerequisite missing: ${marker}`);
}

const WA_PHONE = '905393333303';
const waText = encodeURIComponent('Merhaba Barış Bey, excelarsiv.com ana sayfasındaki Size Özel bölümünden ulaşıyorum. İşletmemizde hazır şablona sığmayan bir finans veya operasyon süreci var. Özel Excel karar sistemi için kapsamı netleştirmek istiyorum.');
const waHref = `https://wa.me/${WA_PHONE}?text=${waText}`;

if (!html.includes('data-home-conversion-v20')) {
  html = html.replace(/<body\b([^>]*)>/u, '<body$1 data-home-conversion-v20>');
}

if (!html.includes('class="hero-decision-note-v20"')) {
  const anchor = '<nav class="hero-chips"';
  if (!html.includes(anchor)) throw new Error('HOME CONVERSION V20: hero chips anchor missing');
  html = html.replace(anchor, `<div class="hero-decision-note-v20" aria-label="Çözüm seçimi rehberi"><span><b>Hazır sistem</b> · standart ihtiyacı hemen çözün</span><span><b>Size özel</b> · süreci dosyaya değil, sistemi sürecinize uydurun</span></div>\n        ${anchor}`);
}

const customBuild = `
<section class="custom-build custom-build-v20" data-experience-stage aria-labelledby="custom-build-v20-title">
  <div class="home-shell custom-build-v20__shell">
    <div class="custom-build-v20__intro">
      <div class="custom-build-v20__copy">
        <p class="eyebrow">SİZE ÖZEL · KARAR SİSTEMİ</p>
        <h2 id="custom-build-v20-title">Hazır tablo sürecinizi size uydurmaya zorluyorsa, sistemi işletmenize uydurun.</h2>
        <p class="custom-build-v20__lead">Tahsilat, banka, maliyet, üretim veya raporlama akışınız standart bir şablona sığmıyorsa; aynı veriyi tekrar tekrar taşımak yerine işleyişinizi tek karar sisteminde birleştirin.</p>
        <div class="custom-build-v20__signal" role="note"><strong>Özel sistem ihtiyacının en net işareti:</strong><span>Rapor almak için birden fazla dosyayı birleştiriyor, aynı veriyi yeniden giriyor veya yönetici sorusuna cevap vermek için hesabı baştan kuruyorsanız.</span></div>
      </div>
      <div class="custom-build-v20__triggers" aria-label="Özel sistem ihtiyacı göstergeleri">
        <article><span class="trigger-icon trigger-icon--green">01</span><div><strong>Aynı veri iki kez giriliyor</strong><p>Tekrarlı giriş hata ve zaman kaybı yaratıyorsa veri akışını tek yapıda toplayın.</p></div></article>
        <article><span class="trigger-icon trigger-icon--blue">02</span><div><strong>Karar için dosya birleştiriliyor</strong><p>Patron veya finans yöneticisi cevap için ayrı tablolar bekliyorsa tek karar ekranına geçin.</p></div></article>
        <article><span class="trigger-icon trigger-icon--amber">03</span><div><strong>Hazır şablon iş akışını bozuyor</strong><p>İşletme sürecini tabloya uydurmak yerine tabloyu gerçek sürecinize göre kurun.</p></div></article>
      </div>
    </div>

    <div class="custom-build-v20__process">
      <article><span>01</span><div><strong>Süreci anlatın</strong><p>Teknik şartname hazırlamayın. Mevcut dosya, ekran görüntüsü veya işleyiş yeterli.</p></div></article>
      <article><span>02</span><div><strong>Karar noktalarını netleştirelim</strong><p>Hangi verinin girileceği, neyin hesaplanacağı ve yönetimin neyi görmesi gerektiği belirlenir.</p></div></article>
      <article><span>03</span><div><strong>Sistemi işleyişinize göre kuralım</strong><p>Giriş, kontrol, hesap ve yönetici ekranları aynı finansal mantıkta birleştirilir.</p></div></article>
      <article><span>04</span><div><strong>Gerçek veriyle doğrulayın</strong><p>Normal kullanımın yanında uç senaryolar ve hata koşullarıyla doğrulama yapılır.</p></div></article>
    </div>

    <div class="custom-build-v20__close">
      <div class="custom-build-v20__trust">
        <span>17 yıllık ticari bankacılık &amp; saha finans bakışı</span>
        <span>Finans + muhasebe + operasyon dili</span>
        <span>Gerektiği kadar sistem · gereksiz modül yok</span>
      </div>
      <div class="custom-build-v20__close-copy">
        <div><small>ŞİRKETİNİZE ÖZEL</small><h3>Dağınık dosyalardan tek karar sistemine geçin.</h3><p>Önce ihtiyacın gerçekten özel sistem gerektirip gerektirmediğini netleştirin. Hazır sistem yeterliyse onu seçin; yetmiyorsa işletmenize göre kuralım.</p></div>
        <div class="custom-build-v20__cta-group">
          <a class="custom-build-v20__primary" href="/ozel-excel-sistemleri" data-cta="home_custom_system_deep" data-location="custom_build">İşletmenize Özel Sistemi İnceleyin →<small>Kapsamı, yöntemi ve örnek sistemi görün</small></a>
          <a class="custom-build-v20__secondary" href="${waHref}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="home_custom_whatsapp" data-location="custom_build">WhatsApp'tan İhtiyacı Anlatın<small>Hazır mesajla doğrudan başlayın</small></a>
        </div>
      </div>
    </div>
  </div>
</section>`;

const customPattern = /<section class="custom-build"[\s\S]*?<\/section>/u;
if (!customPattern.test(html)) throw new Error('HOME CONVERSION V20: custom-build section missing');
html = html.replace(customPattern, customBuild);

const css = `<style id="home-conversion-v20-css">
.hero-decision-note-v20{display:grid;grid-template-columns:1fr 1fr;gap:8px;max-width:650px;margin-top:12px}.hero-decision-note-v20 span{min-width:0;padding:9px 11px;border:1px solid #e3e8e5;border-radius:11px;background:rgba(255,255,255,.78);color:#6b7787;font:650 10.5px/1.35 ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}.hero-decision-note-v20 b{color:#172033;font-weight:850}.hero-decision-note-v20 span:first-child b{color:#087a46}.hero-decision-note-v20 span:last-child b{color:#2563eb}
.custom-build-v20{position:relative;overflow:clip;background:linear-gradient(180deg,#fbfdfc 0%,#fff 100%);border-bottom:1px solid #dfe5e0}.custom-build-v20::before{content:"";position:absolute;inset:0;pointer-events:none;background:radial-gradient(circle at 7% 10%,rgba(5,150,105,.075),transparent 23%),radial-gradient(circle at 94% 18%,rgba(37,99,235,.06),transparent 22%),radial-gradient(circle at 88% 92%,rgba(245,158,11,.06),transparent 20%)}.custom-build-v20__shell{position:relative;z-index:1;padding-block:84px}.custom-build-v20__intro{display:grid;grid-template-columns:minmax(0,.94fr) minmax(0,1.06fr);gap:74px;align-items:start}.custom-build-v20__copy h2{max-width:720px;margin:11px 0 0;color:#0f172a;font-size:clamp(38px,3.65vw,54px);line-height:1.01;letter-spacing:-.052em;text-wrap:balance}.custom-build-v20__lead{max-width:700px;margin:21px 0 0;color:#596979;font-size:16px;line-height:1.72}.custom-build-v20__signal{display:grid;gap:6px;margin-top:24px;padding:17px 18px;border:1px solid #d8e6de;border-left:4px solid #059669;border-radius:16px;background:#fff;box-shadow:0 10px 26px rgba(15,23,42,.045)}.custom-build-v20__signal strong{color:#14532d;font-size:13px}.custom-build-v20__signal span{color:#657384;font-size:13px;line-height:1.55}.custom-build-v20__triggers{display:grid;gap:11px}.custom-build-v20__triggers article{display:grid;grid-template-columns:48px 1fr;gap:14px;align-items:start;padding:18px;border:1px solid #e0e7e3;border-radius:19px;background:rgba(255,255,255,.92);box-shadow:0 11px 28px rgba(15,23,42,.045)}.trigger-icon{display:grid;place-items:center;width:48px;height:48px;border-radius:15px;font:850 12px/1 ui-monospace,SFMono-Regular,Menlo,monospace}.trigger-icon--green{background:#eaf8ef;color:#087a46}.trigger-icon--blue{background:#eef4ff;color:#2563eb}.trigger-icon--amber{background:#fff6e7;color:#c76a05}.custom-build-v20__triggers strong{display:block;color:#172033;font-size:16px;line-height:1.2}.custom-build-v20__triggers p{margin:6px 0 0;color:#687587;font-size:13px;line-height:1.55}.custom-build-v20__process{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-top:34px}.custom-build-v20__process article{min-width:0;padding:20px;border:1px solid #e0e7e3;border-radius:19px;background:#fff;box-shadow:0 10px 26px rgba(15,23,42,.04)}.custom-build-v20__process article>span{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:11px;background:#0f172a;color:#fff;font:800 11px/1 ui-monospace,SFMono-Regular,Menlo,monospace}.custom-build-v20__process article:nth-child(2)>span{background:#2563eb}.custom-build-v20__process article:nth-child(3)>span{background:#d97706}.custom-build-v20__process article:nth-child(4)>span{background:#dc2626}.custom-build-v20__process strong{display:block;margin-top:14px;color:#172033;font-size:15px;line-height:1.25}.custom-build-v20__process p{margin:7px 0 0;color:#697687;font-size:12.5px;line-height:1.58}.custom-build-v20__close{margin-top:26px;padding:8px;border:1px solid #d9e6df;border-radius:27px;background:#fff;box-shadow:0 18px 44px rgba(15,23,42,.065)}.custom-build-v20__trust{display:grid;grid-template-columns:repeat(3,1fr);gap:1px;overflow:hidden;border-radius:19px 19px 8px 8px;background:#e4ebe7}.custom-build-v20__trust span{padding:13px 16px;background:#f7faf8;color:#496055;font-size:11.5px;font-weight:750;text-align:center}.custom-build-v20__trust span::before{content:"✓";margin-right:7px;color:#059669;font-weight:900}.custom-build-v20__close-copy{display:grid;grid-template-columns:minmax(0,1fr) minmax(460px,.75fr);gap:36px;align-items:center;padding:30px}.custom-build-v20__close-copy small{color:#087a46;font:850 10px/1.2 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.13em}.custom-build-v20__close-copy h3{max-width:650px;margin:9px 0 0;color:#0f172a;font-size:31px;line-height:1.04;letter-spacing:-.04em}.custom-build-v20__close-copy p{max-width:700px;margin:12px 0 0;color:#647384;font-size:14px;line-height:1.62}.custom-build-v20__cta-group{display:grid;gap:9px}.custom-build-v20__cta-group>a{min-height:66px;display:flex;flex-direction:column;justify-content:center;padding:12px 17px;border-radius:15px;text-decoration:none;font-size:14px;font-weight:850;line-height:1.2}.custom-build-v20__cta-group>a small{display:block;margin-top:5px;font:650 10.5px/1.25 ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;letter-spacing:0;opacity:.78}.custom-build-v20__primary{background:linear-gradient(135deg,#087a46,#059669);color:#fff;box-shadow:0 13px 28px rgba(5,150,105,.22)}.custom-build-v20__primary small{color:#fff}.custom-build-v20__secondary{border:1.5px solid #cfd9d3;background:#fff;color:#172033}.custom-build-v20__secondary small{color:#526176}
@media(max-width:1020px){.custom-build-v20__intro{grid-template-columns:1fr;gap:24px}.custom-build-v20__process{grid-template-columns:1fr 1fr}.custom-build-v20__close-copy{grid-template-columns:1fr;gap:20px}.custom-build-v20__trust{grid-template-columns:1fr}}
@media(max-width:720px){.hero-decision-note-v20{grid-template-columns:1fr;gap:7px;max-width:none;margin-top:12px}.hero-decision-note-v20 span{padding:10px 11px;border-radius:12px;font-size:11px}.custom-build-v20__shell{width:calc(100% - 20px)!important;padding-block:50px}.custom-build-v20__copy h2{font-size:31px;line-height:1.03}.custom-build-v20__lead{font-size:15px;line-height:1.64}.custom-build-v20__signal{padding:15px;border-radius:15px}.custom-build-v20__triggers article{grid-template-columns:44px 1fr;padding:15px;border-radius:17px}.trigger-icon{width:44px;height:44px;border-radius:13px}.custom-build-v20__triggers strong{font-size:15px}.custom-build-v20__triggers p{font-size:12.5px}.custom-build-v20__process{grid-template-columns:1fr;gap:9px;margin-top:22px}.custom-build-v20__process article{display:grid;grid-template-columns:38px 1fr;column-gap:12px;padding:15px;border-radius:17px}.custom-build-v20__process article>span{grid-row:1 / span 2;width:38px;height:38px}.custom-build-v20__process strong{margin-top:1px;font-size:14px}.custom-build-v20__process p{margin-top:5px;font-size:12.5px}.custom-build-v20__close{margin-top:18px;border-radius:22px}.custom-build-v20__trust{border-radius:15px 15px 7px 7px}.custom-build-v20__trust span{padding:11px 9px;font-size:10.5px}.custom-build-v20__close-copy{padding:20px 14px 14px}.custom-build-v20__close-copy h3{font-size:27px}.custom-build-v20__close-copy p{font-size:13px}.custom-build-v20__cta-group>a{min-height:62px;border-radius:14px;font-size:13.5px}}
</style>`;

if (!html.includes('id="home-conversion-v20-css"')) html = html.replace('</head>', `${css}\n</head>`);

for (const token of ['data-home-conversion-v20','hero-decision-note-v20','custom-build-v20','home_custom_system_deep','home_custom_whatsapp',WA_PHONE]) {
  if (!html.includes(token)) throw new Error(`HOME CONVERSION V20: required token missing: ${token}`);
}
if (html.includes('905419305372')) throw new Error('HOME CONVERSION V20: retired WhatsApp number resurfaced');

fs.writeFileSync(file, html, 'utf8');
console.log('HOME CONVERSION V20 PASS — transparent decision psychology, premium custom-system persuasion and direct special-system conversion are active.');
