#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/index.html');
const cssFile = path.resolve('src/styles/home-conversion-v20.css');
if (!fs.existsSync(file)) throw new Error('HOME CONVERSION V20: dist/index.html missing');
if (!fs.existsSync(cssFile)) throw new Error('HOME CONVERSION V20: CSS missing');
let html = fs.readFileSync(file, 'utf8');
const css = fs.readFileSync(cssFile, 'utf8').trim();
const heroRestoreCss = `@media(min-width:1021px){
body[data-desktop-premium-v19] .hero-section{position:relative!important;z-index:60!important;margin:28px 0 46px!important;padding:0!important;background:transparent!important}
body[data-desktop-premium-v19] .hero-shell{position:relative!important;z-index:60!important;width:min(1680px,calc(100% - 64px))!important;margin-inline:auto!important;padding-bottom:28px!important}
body[data-desktop-premium-v19] .hero-mobile-copy{position:absolute!important;width:1px!important;height:1px!important;padding:0!important;margin:-1px!important;overflow:hidden!important;clip:rect(0,0,0,0)!important;white-space:nowrap!important;border:0!important}
body[data-desktop-premium-v19] .hero-artwork{position:relative!important;display:block!important;overflow:hidden!important;width:100%!important;margin:0 auto!important;aspect-ratio:3 / 1!important;min-height:0!important;border:1px solid rgba(16,124,65,.12)!important;border-radius:24px!important;background:#eef3f6!important;box-shadow:0 24px 60px rgba(24,55,38,.13)!important}
body[data-desktop-premium-v19] .hero-artwork::before{content:none!important;background:none!important}
body[data-desktop-premium-v19] .hero-artwork img{display:block!important;width:100%!important;height:100%!important;object-fit:cover!important;object-position:center!important}
body[data-desktop-premium-v19] .hero-search-rail{position:absolute!important;left:50%!important;bottom:0!important;z-index:70!important;width:min(690px,calc(100% - 48px))!important;margin:0!important;transform:translateX(-50%)!important;scroll-margin-top:90px!important}
}`;

for (const marker of ['data-dual-funnel-home-v17','data-mobile-premium-v18','data-desktop-premium-v19']) {
  if (!html.includes(marker)) throw new Error(`HOME CONVERSION V20: prerequisite missing: ${marker}`);
}

const WA_PHONE = '905393333303';
const waText = encodeURIComponent('Merhaba Barış Bey, excelarsiv.com ana sayfasındaki Size Özel bölümünden ulaşıyorum. İşletmemizde hazır şablona sığmayan bir finans veya operasyon süreci var. Özel Excel karar sistemi için kapsamı netleştirmek istiyorum.');
const waHref = `https://wa.me/${WA_PHONE}?text=${waText}`;

if (!html.includes('data-home-conversion-v20')) html = html.replace(/<body\b([^>]*)>/u, '<body$1 data-home-conversion-v20>');

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

if (!html.includes('id="home-conversion-v20-css"')) html = html.replace('</head>', `<style id="home-conversion-v20-css">${css}</style>\n</head>`);
if (!html.includes('id="home-original-hero-v21-css"')) html = html.replace('</head>', `<style id="home-original-hero-v21-css">${heroRestoreCss}</style>\n</head>`);

for (const token of ['data-home-conversion-v20','custom-build-v20','home_custom_system_deep','home_custom_whatsapp',WA_PHONE,'home-original-hero-v21-css']) {
  if (!html.includes(token)) throw new Error(`HOME CONVERSION V20: required token missing: ${token}`);
}
if (!html.includes('class="hero-mobile-copy"') || !html.includes('srcset="/images/hero.jpg"')) throw new Error('HOME CONVERSION V20: original homepage hero contract missing');
if (html.includes('class="hero-panel"') || html.includes('class="hero-copy"')) throw new Error('HOME CONVERSION V20: redesigned split hero leaked back into homepage');
if (/stokta son|geri sayım|sadece bugün/iu.test(html)) throw new Error('HOME CONVERSION V20: deceptive urgency language detected');

fs.writeFileSync(file, html, 'utf8');
console.log('HOME CONVERSION V20 PASS — original homepage hero restored; conversion architecture remains outside the hero.');
