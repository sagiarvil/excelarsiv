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
body[data-desktop-premium-v19] .hero-section{position:relative!important;z-index:60!important;margin:16px 0 28px!important;padding:8px 0 0!important;background:transparent!important}
body[data-desktop-premium-v19] .hero-shell{position:relative!important;z-index:60!important;width:min(1150px,calc(100% - 40px))!important;margin-inline:auto!important;padding-bottom:56px!important}
body[data-desktop-premium-v19] .hero-mobile-copy{position:absolute!important;z-index:66!important;left:4%!important;top:10%!important;width:min(48%,520px)!important;height:auto!important;padding:0!important;margin:0!important;overflow:visible!important;clip:auto!important;white-space:normal!important;border:0!important}
body[data-desktop-premium-v19] .hero-mobile-copy__eyebrow{margin:0!important;font:800 12.5px/1.2 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace!important;letter-spacing:.08em!important;color:#107c41!important}
body[data-desktop-premium-v19] .hero-title{max-width:520px!important;margin:8px 0 0!important;font:800 clamp(32px,2.6vw,40px)/1.12 ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif!important;letter-spacing:-.03em!important;color:#222222!important;text-wrap:balance!important}
body[data-desktop-premium-v19] .hero-mobile-copy__summary{max-width:500px!important;margin:10px 0 0!important;font:500 16px/1.55 ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif!important;color:#444444!important}
body[data-desktop-premium-v19] .hero-artwork{position:relative!important;display:block!important;overflow:hidden!important;width:100%!important;margin:0 auto!important;aspect-ratio:3 / 1!important;min-height:0!important;border:1px solid rgba(16,124,65,.12)!important;border-radius:16px!important;background:#eef3f6!important;box-shadow:0 12px 32px rgba(24,55,38,.07)!important}
body[data-desktop-premium-v19] .hero-artwork::before{content:none!important;background:none!important}
body[data-desktop-premium-v19] .hero-panel::after{content:none!important;background:none!important}
body[data-desktop-premium-v19] .hero-artwork::after{content:""!important;position:absolute!important;inset:0!important;pointer-events:none!important;background:linear-gradient(90deg,#f7faf8 0%,#f7faf8 48%,rgba(247,250,248,.96) 54%,rgba(247,250,248,.72) 60%,rgba(247,250,248,0) 68%)!important}
body[data-desktop-premium-v19] .hero-artwork img{display:block!important;width:100%!important;height:100%!important;object-fit:cover!important;object-position:center!important}
body[data-desktop-premium-v19] .hero-tech-visual{position:absolute!important;right:0!important;top:0!important;bottom:0!important;width:55%!important;z-index:3!important;pointer-events:none!important;display:flex!important;align-items:center!important;justify-content:center!important}
body[data-desktop-premium-v19] .hero-search-rail{position:absolute!important;left:4%!important;bottom:0!important;z-index:70!important;width:min(520px,calc(50% - 20px))!important;margin:0!important;transform:none!important;scroll-margin-top:90px!important}
body[data-desktop-premium-v19] .home-shell{width:min(1150px,calc(100% - 40px))!important}
body[data-desktop-premium-v19] .soho-reviews-bar{background:#f8fafc!important;border:1px solid #e2e8f0!important;border-radius:10px!important;padding:10px 16px!important;margin-top:0!important;margin-bottom:24px!important}
body[data-desktop-premium-v19] .soho-category-headband{margin-top:32px!important;border-bottom:2px solid #eeeeee!important;padding-bottom:2px!important;display:flex!important;align-items:baseline!important;justify-content:space-between!important}
body[data-desktop-premium-v19] .soho-category-title{font-size:22px!important;font-weight:700!important;color:#222222!important;border-bottom:2px solid #107c41!important;padding-bottom:3px!important;margin:0!important}
body[data-desktop-premium-v19] .soho-category-title a{color:#222222!important;text-decoration:none!important}
body[data-desktop-premium-v19] .soho-category-more a{color:#107c41!important;font-size:14.5px!important;font-weight:600!important;text-decoration:none!important}
body[data-desktop-premium-v19] .soho-category-more a:hover{text-decoration:underline!important}
body[data-desktop-premium-v19] .soho-category-desc{margin:8px 0 20px!important;font-size:14.5px!important;line-height:1.55!important;color:#555555!important;font-weight:400!important}

/* SOMEKA KART STİLİ %100 BİREBİR */
body[data-desktop-premium-v19] .soho-product-grid{gap:16px!important;grid-template-columns:repeat(4,minmax(0,1fr))!important;margin-bottom:28px!important}
body[data-desktop-premium-v19] .someka-product-custom-card{border:1px solid #d6dce4!important;border-radius:4px!important;background:#ffffff!important;box-shadow:0 1px 3px rgba(0,0,0,0.05)!important;transition:box-shadow .2s ease,transform .18s ease!important;overflow:hidden!important}
body[data-desktop-premium-v19] .someka-product-custom-card:hover{box-shadow:0 8px 24px rgba(15,23,42,.09)!important;transform:translateY(-2px)!important;border-color:#cbd5e1!important}
body[data-desktop-premium-v19] .someka-card-thumb-wrap{position:relative!important;display:block!important;aspect-ratio:16 / 10!important;background:#f8fafc!important;border-bottom:1px solid #eee!important;overflow:hidden!important}
body[data-desktop-premium-v19] .someka-thumb-img{width:100%!important;height:100%!important;object-fit:cover!important;object-position:top center!important;margin-top:0!important;display:block!important;transition:transform .2s ease!important}
body[data-desktop-premium-v19] .someka-card-thumb-wrap:hover .someka-thumb-img{transform:scale(1.03)!important}
body[data-desktop-premium-v19] .someka-category-tag{display:none!important}
body[data-desktop-premium-v19] .someka-card-details{padding:12px 10px 14px!important;text-align:center!important;display:flex!important;flex-direction:column!important;flex:1!important}
body[data-desktop-premium-v19] .single-post-title{margin:2px 0 4px!important;min-height:38px!important;line-height:1.35!important;text-align:center!important}
body[data-desktop-premium-v19] .single-post-title a{color:#222222!important;font-size:16px!important;font-weight:700!important;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif!important;text-decoration:none!important;transition:color .15s ease!important}
body[data-desktop-premium-v19] .single-post-title a:hover{color:#107c41!important}
body[data-desktop-premium-v19] .vcex-star-rating{display:none!important}
body[data-desktop-premium-v19] .vcex-post-excerpt{display:-webkit-box!important;-webkit-line-clamp:2!important;-webkit-box-orient:vertical!important;overflow:hidden!important;font-size:14px!important;line-height:1.5!important;color:#555555!important;margin:4px 0 0!important;min-height:35px!important;text-align:center!important}
body[data-desktop-premium-v19] .someka-card-bottom{display:none!important}
}
/* SOMEKA BOLD TİPOGRAFİ STANDARDI — TÜM SİTE VE TÜM SAYFALAR */
.single-post-title,
.single-post-title a,
.card__title,
.card__title a,
.tcard__title,
.tcard__title a,
.mce-quick-card__title,
.mce-template-card__title,
[data-template-card] .card__title,
[data-template-card] .card__title a,
.soho-product-grid .single-post-title,
.soho-product-grid .single-post-title a {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif !important;
  font-size: 16px !important;
  font-weight: 700 !important;
  line-height: 1.35 !important;
  color: #222222 !important;
  text-decoration: none !important;
  transition: color .15s ease !important;
  letter-spacing: -0.01em !important;
}
.single-post-title a:hover,
.card__title a:hover,
.tcard__title a:hover,
[data-template-card] .card__title a:hover,
.soho-product-grid .single-post-title a:hover {
  color: #107c41 !important;
}`;

for (const marker of ['data-dual-funnel-home-v17','data-mobile-premium-v18','data-desktop-premium-v19']) {
  if (!html.includes(marker)) throw new Error(`HOME CONVERSION V20: prerequisite missing: ${marker}`);
}

const WA_PHONE = '905393333303';
const waText = encodeURIComponent('Merhaba Barış Bey, excelarsiv.com ana sayfasındaki Size Özel bölümünden ulaşıyorum. İşletmemizde hazır şablona sığmayan bir finans veya operasyon süreci var. Özel Excel karar sistemi için kapsamı netleştirmek istiyorum.');
if (html.includes('data-appstore-mode="today"')) {
  if (!html.includes('data-home-conversion-v20')) html = html.replace(/<body\b([^>]*)>/u, '<body$1 data-home-conversion-v20>');
  if (!html.includes('id="home-conversion-v20-css"')) html = html.replace('</head>', `<style id="home-conversion-v20-css">${css}</style>\n</head>`);
  fs.writeFileSync(file, html, 'utf8');
  console.log('HOME CONVERSION V20 PASS — App Store Today mode verified.');
  process.exit(0);
}

if (!html.includes('data-home-conversion-v20')) html = html.replace(/<body\b([^>]*)>/u, '<body$1 data-home-conversion-v20>');

const customBuild = `
<section class="custom-build custom-build-v20" data-experience-stage aria-labelledby="custom-build-v20-title">
  <div class="home-shell custom-build-v20__shell">
    <div class="custom-build-v20__intro">
      <div class="custom-build-v20__copy">
        <p class="eyebrow">SİZE ÖZEL · KARAR SİSTEMİ</p>
        <h2 id="custom-build-v20-title">Dağınık Tablo Yok, Formül Hatası Yok.</h2>
        <p class="custom-build-v20__lead">Hazır şablonlar iş akışınızı sınırlandırıyorsa; süreci tabloya uydurmak yerine, sistemi işletmenize uyduralım. Tahsilat, banka, maliyet veya raporlama akışınızı tek karar motorunda birleştirin.</p>
        <div class="custom-build-v20__signal" role="note"><strong>Özel sistem ihtiyacının en net işareti:</strong><span>Rapor almak için birden fazla dosyayı birleştiriyor, aynı veriyi yeniden giriyor veya yönetici sorusuna cevap vermek için hesabı baştan kuruyorsanız.</span></div>
      </div>
      <div class="custom-build-v20__triggers" aria-label="Özel sistem ihtiyacı göstergeleri">
        <article><span class="trigger-icon trigger-icon--green">01</span><div><strong>Aynı veri iki kez giriliyor</strong><p>Tekrarlı giriş hata ve zaman kaybı yaratıyorsa veri akışını tek yapıda toplayın.</p></div></article>
        <article><span class="trigger-icon trigger-icon--blue">02</span><div><strong>Karar için dosya birleştiriliyor</strong><p>Patron veya finans yöneticisi cevap için ayrı tablolar bekliyorsa tek karar ekranına geçin.</p></div></article>
        <article><span class="trigger-icon trigger-icon--amber">03</span><div><strong>Hazır şablon iş akışını bozuyor</strong><p>İşletme sürecini tabloya uydurmak yerine tabloyu gerçek sürecinize göre kurun.</p></div></article>
      </div>
    </div>

    <!-- 5 Temel Direk Bento Izgarası -->
    <div class="custom-build-v20__five-pillars" aria-label="ExcelArşiv 5 Temel Direk">
      <article class="pillar-mini-box">
        <div class="pillar-mini-top"><span class="p-no">01</span><span class="p-tag">MÜŞTERİ ODAĞI</span></div>
        <strong>Saha Dili &amp; Aynı Gün Çözüm</strong>
        <p>Aylarca ERP beklemek yerine, aynı gün çalışan hazır finansal karar motoru.</p>
      </article>
      <article class="pillar-mini-box">
        <div class="pillar-mini-top"><span class="p-no">02</span><span class="p-tag">ALTIN KALKAN</span></div>
        <strong>17 Yıl Bankacılık &amp; Sıfır Makro</strong>
        <p>%100 açık formül. Mali verileriniz asla buluta gitmez, bilgisayarınızda kalır.</p>
      </article>
      <article class="pillar-mini-box">
        <div class="pillar-mini-top"><span class="p-no">03</span><span class="p-tag">10 SN DAĞITIM</span></div>
        <strong>Anında Teslim &amp; Sıfır Kurulum</strong>
        <p>Windows, Mac ve iPad tam uyumlu. Tek seferlik ödeme, sıfır abonelik tuzağı.</p>
      </article>
      <article class="pillar-mini-box">
        <div class="pillar-mini-top"><span class="p-no">04</span><span class="p-tag">24 SEKTÖR</span></div>
        <strong>1.000+ İşletme Referansı</strong>
        <p>Kafeden fabrikaya, inşaattan e-ticarete onaylanmış 51+ karar sistemi.</p>
      </article>
      <article class="pillar-mini-box">
        <div class="pillar-mini-top"><span class="p-no">05</span><span class="p-tag">GÜVENLİ SATIŞ</span></div>
        <strong>Hızlı Teklif &amp; Kurumsal E-Fatura</strong>
        <p>WhatsApp hattımız üzerinden hızlı teklif ve resmi kurumsal e-arşiv fatura.</p>
      </article>
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
        <div>
          <small>ŞİRKETİNİZE ÖZEL</small>
          <h3>Dağınık dosyalardan tek karar sistemine geçin.</h3>
          <p>Önce ihtiyacın gerçekten özel sistem gerektirip gerektirmediğini netleştirin. Hazır sistem yeterliyse onu seçin; yetmiyorsa işletmenize göre kuralım.</p>
          <div class="home-card-installments">
            <span class="hci-title">Güvenli Ödeme &amp; Kurumsal Fatura:</span>
            <div class="hci-badges">
              <span>%100 Açık Formül</span>
              <span>Sıfır Makro</span>
              <span>Anında Teslim</span>
              <span>E-Arşiv Fatura</span>
            </div>
          </div>
        </div>
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
if (!html.includes('home_ready_systems') || !html.includes('home_custom_system') || !html.includes('hero-route-actions')) throw new Error('HOME CONVERSION V20: dual hero CTA contract missing');
if (html.includes('class="hero-panel"') || html.includes('class="hero-copy"')) throw new Error('HOME CONVERSION V20: redesigned split hero leaked back into homepage');
if (/stokta son|geri sayım|sadece bugün/iu.test(html)) throw new Error('HOME CONVERSION V20: deceptive urgency language detected');

// Soho Kategori Sekmeleri Aktifleştirme (Alpine.js olmadan hatasız vanilya JS)
html = html.replace(
  '<button type="button" class="soho-tab-btn" :class="{ \'is-active\': tab === \'nakit\' }"',
  '<button type="button" class="soho-tab-btn is-active" :class="{ \'is-active\': tab === \'nakit\' }"'
);

const tabsScript = `
<script>
(function() {
  function setupSohoTabs() {
    var nav = document.querySelector('.soho-tabs-nav');
    var section = document.querySelector('.soho-filter-section');
    if (!nav || !section) return;
    var btns = Array.prototype.slice.call(nav.querySelectorAll('.soho-tab-btn'));
    var panels = Array.prototype.slice.call(section.querySelectorAll('.soho-tab-panel'));
    if (!btns.length || !panels.length) return;

    btns.forEach(function(btn, index) {
      btn.addEventListener('click', function(e) {
        e.preventDefault();
        btns.forEach(function(b) { b.classList.remove('is-active'); });
        btn.classList.add('is-active');
        panels.forEach(function(panel, pIdx) {
          if (pIdx === index) {
            panel.style.display = 'block';
          } else {
            panel.style.display = 'none';
          }
        });
      });
    });
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', setupSohoTabs);
  } else {
    setupSohoTabs();
  }
})();
</script>
`;

if (!html.includes('setupSohoTabs')) {
  html = html.replace('<!-- 3. YENİ ÇIKAN', `${tabsScript}\n<!-- 3. YENİ ÇIKAN`);
}

fs.writeFileSync(file, html, 'utf8');
console.log('HOME CONVERSION V20 PASS — original image-first hero preserved; dual CTA routes exposed without restoring the split hero.');
