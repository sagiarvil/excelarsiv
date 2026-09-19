#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/index.html');
if (!fs.existsSync(file)) throw new Error('HOME DUAL FUNNEL V17: dist/index.html missing');
let html = fs.readFileSync(file, 'utf8');

const WA_PHONE = '905393333303';
const wa = (message) => `https://wa.me/${WA_PHONE}?text=${encodeURIComponent(message)}`;
const bottomWa = wa('Merhaba Barış Bey, excelarsiv.com ana sayfasından ulaşıyorum. Şirketimizin finansal tıkanıklığını ve uygun Excel karar sistemini netleştirmek istiyorum.');

const requireReplace = (pattern, replacement, label) => {
  const next = html.replace(pattern, replacement);
  if (next === html) throw new Error(`HOME DUAL FUNNEL V17: replacement missed: ${label}`);
  html = next;
};

if (html.includes('data-dual-funnel-home-v17')) {
  console.log('HOME DUAL FUNNEL V17: already applied');
  process.exit(0);
}

requireReplace(/<body\b([^>]*)>/u, '<body$1 data-dual-funnel-home-v17>', 'body marker');
requireReplace(/<title>[\s\S]*?<\/title>/u, '<title>Finansal Karar ve Excel Sistemleri | Excel Arşiv</title>', 'title');

const financePillars = `
<section class="finance-pillars" id="finans-sistemleri" data-experience-stage>
  <div class="home-shell">
    <div class="finance-pillars__head">
      <div>
        <p class="eyebrow">TÜM FİNANSAL İHTİYAÇLARINIZ TEK ÇATIDA — EXCEL ARŞİV</p>
        <h2>Nakit açığından maliyet hesabına, işletmenizin tüm finansal kontrolleri.</h2>
      </div>
      <p>Dağınık tablolarla vakit kaybetmeyin. 17 yıllık ticari bankacılık saha tecrübesiyle kurgulanmış, %100 açık formüllü Excel karar sistemlerini inceleyin veya şirketinize özel uyarlayın.</p>
    </div>

    <!-- PDF Slayt 3 Şeması: Merkezi Excel Dashboard Mockup Sahnesi + 4 Altın Mühür -->
    <div class="pdf-showcase-stage" aria-label="Excel Arşiv Finansal Kokpit Sahnesi">
      <div class="showcase-glow" aria-hidden="true"></div>
      <div class="laptop-mockup-wrapper">
        <div class="laptop-frame">
          <div class="laptop-screen">
            <!-- Excel Ribbon & Başlık Çubuğu -->
            <div class="excel-ui-bar">
              <div class="excel-dots"><span class="dot-red"></span><span class="dot-yellow"></span><span class="dot-green"></span></div>
              <div class="excel-title-tab">📊 13_Haftalik_Nakit_Akisi_ve_Patron_Paneli.xlsx</div>
              <div class="excel-status-tag">SAF .XLSX · SIFIR MAKRO</div>
            </div>
            <!-- Excel Menü Çubuğu -->
            <div class="excel-menu-strip">
              <span>Dosya</span><span class="active">Giriş</span><span>Ekle</span><span>Formüller</span><span>Veri</span><span>Gözden Geçir</span><span>Görünüm</span>
            </div>
            <!-- Excel Formül Çubuğu -->
            <div class="excel-formula-bar">
              <span class="fx-label">fx</span>
              <span class="formula-text">=EĞER(HAFTALIK_NET_NAKİT&lt;0; "⚠️ DÖNEMSEL NAKİT AÇIĞI UYARISI"; "✅ LİKİDİTE GÜVENLİ BÖLGEDE")</span>
            </div>
            <!-- Excel Tablo ve KPI Ekranı -->
            <div class="excel-grid-preview">
              <div class="excel-kpi-row">
                <div class="excel-kpi-card kpi-green">
                  <small>Haftalık Net Nakit Akışı</small>
                  <strong>+₺1.845.000</strong>
                  <span>Dinamik 13 Hafta Projeksiyonu</span>
                </div>
                <div class="excel-kpi-card kpi-blue">
                  <small>Çoklu Banka Rotatif Kredi Limiti</small>
                  <strong>₺4.500.000</strong>
                  <span>Faiz &amp; Taksit Vade Simülasyonu</span>
                </div>
                <div class="excel-kpi-card kpi-amber">
                  <small>Cari Tahsilat Yaşlandırma &amp; Risk</small>
                  <strong>%94.2 Güvenli</strong>
                  <span>Vadesi Geçen Alacak Alarmı</span>
                </div>
              </div>
              <div class="excel-table-mock">
                <div class="table-mock-head">
                  <span class="cell-id">Hafta</span>
                  <span>Açılış Kasası</span>
                  <span>Tahsilat Projeksiyonu</span>
                  <span>Tediye &amp; Tedarikçi</span>
                  <span>Kredi Taksiti</span>
                  <span class="cell-status">Net Likidite</span>
                </div>
                <div class="table-mock-row">
                  <span class="cell-id">H1</span>
                  <span>₺620.000</span>
                  <span>₺1.150.000</span>
                  <span>(₺780.000)</span>
                  <span>(₺120.000)</span>
                  <span class="status-pos">+₺870.000</span>
                </div>
                <div class="table-mock-row">
                  <span class="cell-id">H2</span>
                  <span>₺870.000</span>
                  <span>₺980.000</span>
                  <span>(₺650.000)</span>
                  <span>(₺140.000)</span>
                  <span class="status-pos">+₺1.060.000</span>
                </div>
                <div class="table-mock-row is-alert">
                  <span class="cell-id">H3</span>
                  <span>₺1.060.000</span>
                  <span>₺450.000</span>
                  <span>(₺1.820.000)</span>
                  <span>(₺120.000)</span>
                  <span class="status-neg">⚠️ -₺430.000 (Erken Uyarı)</span>
                </div>
              </div>
            </div>
          </div>
          <div class="laptop-base"><div class="laptop-notch"></div></div>
        </div>

        <!-- Mockup Çevresindeki 4 Altın Mühür Rozeti (Floating Medallions) -->
        <div class="floating-medallion medallion-tl">
          <span class="medallion-icon">🔒</span>
          <div><strong>%100 Açık Formül</strong><small>Gizli sayfa veya şifre yok</small></div>
        </div>
        <div class="floating-medallion medallion-tr">
          <span class="medallion-icon">⚡</span>
          <div><strong>Sıfır Makro Riski</strong><small>Saf .xlsx, virüs uyarısı yok</small></div>
        </div>
        <div class="floating-medallion medallion-bl">
          <span class="medallion-icon">🏦</span>
          <div><strong>17 Yıl Bankacılık</strong><small>Mali tahlil ve nakit refleksi</small></div>
        </div>
        <div class="floating-medallion medallion-br">
          <span class="medallion-icon">💾</span>
          <div><strong>Yerel Cihazda Gizlilik</strong><small>Mali veriniz cihazınızda kalır</small></div>
        </div>
      </div>
    </div>

    <!-- PDF Slayt 3 Şeması: 15 Numaralı Excel Karar Sistemi Matrisi -->
    <div class="finance-needs-grid" aria-label="15 Temel Çözüm İhtiyacı">
      <a href="/sablon/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi" class="finance-need-item"><span class="need-badge">01</span><span class="need-name">13 Haftalık Dinamik Nakit Akışı</span></a>
      <a href="/sablon/banka-kredi-ve-taksit-takip-sistemi" class="finance-need-item"><span class="need-badge">02</span><span class="need-name">Çoklu Banka &amp; Kredi Taksit Takibi</span></a>
      <a href="/sablon/akilli-kasa-defteri-ve-nakit-kontrol-sistemi" class="finance-need-item"><span class="need-badge">03</span><span class="need-name">Kasa Defteri &amp; Günlük Likidite</span></a>
      <a href="/sablon/cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi" class="finance-need-item"><span class="need-badge">04</span><span class="need-name">Müşteri Risk Skoru &amp; Cari Yaşlandırma</span></a>
      <a href="/sablonlar?q=maliyet" class="finance-need-item"><span class="need-badge">05</span><span class="need-name">Birim Maliyet &amp; Dinamik Fiyatlama</span></a>
      <a href="/sablon/aylik-patron-finans-paneli" class="finance-need-item"><span class="need-badge">06</span><span class="need-name">Aylık Patron &amp; Yönetici Finans Paneli</span></a>
      <a href="/sablon/cek-senet-ve-vade-risk-sistemi" class="finance-need-item"><span class="need-badge">07</span><span class="need-name">Çek - Senet &amp; Vade Risk Portföyü</span></a>
      <a href="/sablonlar?q=pazaryeri" class="finance-need-item"><span class="need-badge">08</span><span class="need-name">Pazaryeri Net Kâr &amp; Komisyon Hesabı</span></a>
      <a href="/sablonlar?q=butce" class="finance-need-item"><span class="need-badge">09</span><span class="need-name">3 Senaryolu Dinamik Bütçe &amp; Tahmin</span></a>
      <a href="/sablon/ithalat-depo-teslim-rafa-gelen-net-birim-maliyet" class="finance-need-item"><span class="need-badge">10</span><span class="need-name">İthalat Depo Teslim Birim Maliyet</span></a>
      <a href="/sablon/kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici" class="finance-need-item"><span class="need-badge">11</span><span class="need-name">Kıdem, İhbar &amp; Personel Maliyeti</span></a>
      <a href="/sablon/sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli" class="finance-need-item"><span class="need-badge">12</span><span class="need-name">Şirket Öz Kaynağı &amp; TTK 376 Cetveli</span></a>
      <a href="/sablon/pos-komisyon-ve-net-tahsilat-kontrol-sistemi" class="finance-need-item"><span class="need-badge">13</span><span class="need-name">POS Komisyonu &amp; Net Tahsilat Kontrolü</span></a>
      <a href="/sablon/uretim-recetesi-ve-zam-yansitma-hesaplayici" class="finance-need-item"><span class="need-badge">14</span><span class="need-name">Üretim Reçetesi &amp; Zam Yansıtma</span></a>
      <a href="/sablon/doviz-acik-pozisyonu-ve-kur-riski-stres-testi" class="finance-need-item"><span class="need-badge">15</span><span class="need-name">Döviz Pozisyonu &amp; Kur Riski Stres Testi</span></a>
    </div>

    <!-- PDF Slayt 5 Şeması: 6'lı Kurumsal Güvence Bento Izgarası -->
    <div class="pdf-security-bento-section" aria-label="Kurumsal Güvenlik ve Güvence Standartları">
      <div class="bento-section-head">
        <span class="eyebrow">KURUMSAL GÜVENLİK VE STANDARTLAR</span>
        <h3>Neden Finansal Yönetimde Excel Arşiv Altyapısı?</h3>
        <p>İşletmenizin finansal mahremiyetini, veri güvenliğini ve operasyonel hızını koruyan 6 demir kural.</p>
      </div>
      <div class="pdf-six-bento-grid">
        <article class="bento-item">
          <div class="bento-num">01</div>
          <h4>Tam Yerel Veri Mahremiyeti</h4>
          <p>Mali tablolarınız, müşteri bakiyeleriniz ve cironuz asla bulut sunuculara veya yabancı veritabanlarına aktarılmaz. Dosyalar %100 şirket bilgisayarınızda çalışır.</p>
        </article>
        <article class="bento-item">
          <div class="bento-num">02</div>
          <h4>Sıfır Makro &amp; Sıfır Virüs Güvencesi</h4>
          <p>Şablonlarımızda tek bir makro (.xlsm / VBA) kodu dahi bulunmaz. Saf .xlsx formatında olup antivirüs ve kurumsal IT güvenlik duvarlarına asla takılmaz.</p>
        </article>
        <article class="bento-item">
          <div class="bento-num">03</div>
          <h4>%100 Açık ve Düzenlenebilir Formül</h4>
          <p>Kilitli sayfa, gizli hücre veya şifreli formül yoktur. Şirketinizin değişen ihtiyaçlarına göre istediğiniz satırı, sütunu ve formülü özgürce genişletebilirsiniz.</p>
        </article>
        <article class="bento-item">
          <div class="bento-num">04</div>
          <h4>Tek Seferlik Ödeme (Sıfır Abonelik)</h4>
          <p>Aylık veya yıllık zorunlu lisans yenileme ücreti yoktur. Bir kez satın alır, ömür boyu şirketiniz bünyesinde sınırsız olarak kullanırsınız.</p>
        </article>
        <article class="bento-item">
          <div class="bento-num">05</div>
          <h4>10 Saniyede Anında Dijital Teslimat</h4>
          <p>Siparişiniz tamamlandığı anda Excel dosyanız, kullanım kılavuzunuz ve kurumsal e-arşiv faturanız anında e-posta adresinize otomatik iletilir.</p>
        </article>
        <article class="bento-item">
          <div class="bento-num">06</div>
          <h4>17 Yıllık Ticari Bankacılık Metodolojisi</h4>
          <p>Yazılımcı mantığıyla değil; yüzlerce reel sektör şirketinin mali tahlilini ve nakit akışını yönetmiş bankacı saha tecrübesiyle kurgulanmış karar mimarisi.</p>
        </article>
      </div>
    </div>

    <!-- PDF Slayt 21 Şeması: 7 Adımlı Neden ExcelArşiv Standartları -->
    <div class="pdf-reasons-timeline-card" aria-label="Neden Excel Arşiv 7 Demir Standart">
      <div class="timeline-head">
        <span class="eyebrow">NEDEN EXCEL ARŞİV?</span>
        <h4>Şirketlerin Karar Alırken Bizi Tercih Etmesinin 7 Nedeni</h4>
      </div>
      <div class="pdf-reasons-list">
        <div class="reason-row"><span class="r-no">1</span><div class="r-content"><strong>17 Yıllık Saha ve Ticari Bankacılık Disiplini</strong><p>Banka kredi komitelerinin ve tecrübeli CFO'ların baktığı rasyolarla modellenmiş hazır finansal refleks.</p></div></div>
        <div class="reason-row"><span class="r-no">2</span><div class="r-content"><strong>Ağır ERP'lere Göre %95 Daha Düşük Maliyet</strong><p>Yüz binlerce liralık hantal ERP kurulumları ve aylar süren uyarlama süreçleri yerine aynı gün çalışan net çözüm.</p></div></div>
        <div class="reason-row"><span class="r-no">3</span><div class="r-content"><strong>Sıfır Kurulum ve Sıfır Eğitim Maliyeti</strong><p>Ekibinizin halihazırda bildiği Excel ortamında çalışır. Personel eğitim maliyeti ve adaptasyon kaybı yaşanmaz.</p></div></div>
        <div class="reason-row"><span class="r-no">4</span><div class="r-content"><strong>Logo, SAP, Mikro ve Zirve Verileriyle Tam Uyum</strong><p>Muhasebe veya ERP programınızdan aldığınız standart mizan ve ekstreleri kolayca besleyebilir, tek ekranda karar üretebilirsiniz.</p></div></div>
        <div class="reason-row"><span class="r-no">5</span><div class="r-content"><strong>24 Farklı Sektörde 1.000+ İşletmede Kanıtlandı</strong><p>İnşaattan e-ticarete, üretimden toptan ticarete kadar piyasa gerçekleriyle test edilmiş 51+ karar sistemi.</p></div></div>
        <div class="reason-row"><span class="r-no">6</span><div class="r-content"><strong>Doğrudan Finansal Sistem Mimarı Desteği</strong><p>Standart şablonların ötesinde, şirketinize özel karmaşık darboğazlar için doğrudan Barış Bağırlar mimari danışmanlığı.</p></div></div>
        <div class="reason-row"><span class="r-no">7</span><div class="r-content"><strong>Sürekli Güncellenen Finans ve Mevzuat Standartları</strong><p>TTK 376, güncel kıdem tavanı, asgari ücret parametreleri ve bankacılık mevzuatıyla daima güncel kalan model yapısı.</p></div></div>
      </div>
    </div>

    <!-- PDF Slayt 7 Şeması: 15 Dakikada Hızlı Teşhis ve Karar Diyaloğu -->
    <div class="pdf-chat-dialog-card" aria-label="WhatsApp Ön Teşhis ve Çözüm Diyaloğu">
      <div class="pdf-chat-header">
        <span class="pdf-chat-kicker">15 DAKİKADA HIZLI TEŞHİS VE ÇÖZÜM</span>
        <h4>İşletmeler Nasıl Sorun Yaşıyor, Excel Arşiv Nasıl Çözüyor?</h4>
      </div>
      <div class="pdf-chat-bubbles">
        <div class="chat-bubble chat-bubble--client">
          <div class="bubble-meta"><span class="avatar">🏢</span><strong>İşletme Sahibi / Finans Yöneticisi</strong><small>10:42</small><span class="b-tag b-tag--alert">DARBOĞAZ</span></div>
          <p>“4 depomuz ve 3 pazaryerimiz var. Cari hesaplar ve banka hareketleri farklı tablolarda tutuluyor. Kasa bir türlü tutmuyor, vadeli alacakları göremiyoruz. Ağır bir ERP için 400.000 TL ve 6 ay kurulum süresi istediler. Çalışma düzenimizi bozmadan bunu çözebilir miyiz?”</p>
        </div>
        <div class="chat-divider">
          <span class="divider-line"></span>
          <span class="divider-badge">3 GÜNDE ANALİZ &amp; TESLİMAT</span>
          <span class="divider-line"></span>
        </div>
        <div class="chat-bubble chat-bubble--advisor">
          <div class="bubble-meta"><span class="avatar">⚡</span><strong>Barış Bağırlar · Finansal Sistem Mimarı</strong><small>11:05</small><span class="b-tag b-tag--success">NET ÇÖZÜM</span></div>
          <p>“6 ay beklemenize gerek yok. Mevcut Logo ve banka ekstrelerinizi tek bir standart veri tablosuna bağlayalım. 13 haftalık dinamik nakit akışı ve otomatik cari yaşlandırma kokpitinizi 3 gün içinde teslim ederiz. Sıfır makro, %100 açık formül; ekibiniz anında kullanmaya başlar.”</p>
        </div>
      </div>
      <div class="chat-outcome-bar">
        <div class="outcome-stat"><strong>4 Günden 15 Dk'ya</strong><span>Aylık mutabakat süresi</span></div>
        <div class="outcome-stat"><strong>2 Hafta Önceden</strong><span>Nakit açığı erken uyarısı</span></div>
        <div class="outcome-stat"><strong>%0 Veri Sızıntısı</strong><span>Tamamen yerel cihazda çalışma</span></div>
      </div>
    </div>

    <!-- PDF Slayt 23 Şeması: 24 Sektörde 1.000+ İşletme Referansı -->
    <div class="pdf-sector-ecosystem-card" aria-label="Sektörel Referanslar">
      <div class="sector-eco-head">
        <div>
          <span class="pdf-card-badge">24 SEKTÖRDE 1.000+ İŞLETME REFERANSI</span>
          <h4>Türkiye Çapında Onaylanmış Karar Sistemleri</h4>
        </div>
        <a href="/referans" class="sector-all-link">Tüm 24 Sektörü İncele →</a>
      </div>
      <div class="pdf-sectors-pill-cloud">
        <span class="sector-pill">🏗️ İnşaat &amp; Hakediş</span>
        <span class="sector-pill">🛒 E-Ticaret &amp; Pazaryeri</span>
        <span class="sector-pill">☕ Kafe &amp; Restoran</span>
        <span class="sector-pill">⚙️ Üretim &amp; İmalat</span>
        <span class="sector-pill">🚚 Lojistik &amp; Nakliye</span>
        <span class="sector-pill">📦 Toptan &amp; Dağıtım</span>
        <span class="sector-pill">🏥 Sağlık &amp; Klinik</span>
        <span class="sector-pill">🚗 Otomotiv &amp; Filo</span>
        <span class="sector-pill">🏨 Otel &amp; Turizm</span>
        <span class="sector-pill">🌍 İthalat &amp; İhracat</span>
        <span class="sector-pill">📊 Mali Müşavir &amp; YMM</span>
        <span class="sector-pill">💻 Yazılım &amp; Ajans</span>
      </div>
    </div>

    <!-- PDF Slayt 24 Şeması: Tüm Finansal Karar Sistemleri Tek Çatıda (4 Sütun) -->
    <div class="finance-pillars__grid">
      <article class="finance-pillar finance-pillar--green"><span class="finance-pillar__no">01</span><h3>Nakit &amp; Likidite Sistemleri</h3><ul><li>13 Haftalık Dinamik Nakit Akışı Modeli</li><li>Günlük Kasa &amp; Banka Likidite Takip Paneli</li><li>Çek &amp; Senet Vade Yaşlandırma Tablosu</li></ul><a href="/sablonlar?q=nakit" data-cta="home_finance_cash" data-location="finance_pillars">Nakit sistemlerini gör →</a></article>
      <article class="finance-pillar finance-pillar--blue"><span class="finance-pillar__no">02</span><h3>Banka, Kredi &amp; Teminat Sistemleri</h3><ul><li>Çoklu Banka Limit-Risk ve Kredi Portföyü</li><li>Rotatif Kredi Faiz ve Finansman Maliyet Simülatörü</li><li>Teminat Mektubu &amp; İpotek Karşılama Matrisi</li></ul><a href="/sablonlar?q=kredi" data-cta="home_finance_credit" data-location="finance_pillars">Banka sistemlerini gör →</a></article>
      <article class="finance-pillar finance-pillar--violet"><span class="finance-pillar__no">03</span><h3>Birim Maliyet &amp; Dinamik Fiyatlama</h3><ul><li>Değişken Maliyet ve Hammadde Endeksli Fiyat Teklif Motoru</li><li>Ürün &amp; Müşteri Bazlı Katkı Payı ve Kârlılık Matrisi</li><li>Şirket &amp; Proje Başabaş (Break-Even) Hesaplayıcı</li></ul><a href="/sablonlar?q=maliyet" data-cta="home_finance_cost" data-location="finance_pillars">Maliyet sistemlerini gör →</a></article>
      <article class="finance-pillar finance-pillar--amber"><span class="finance-pillar__no">04</span><h3>Bütçe, Projeksiyon &amp; Değerleme</h3><ul><li>3 Senaryolu Dinamik Bütçe &amp; Rolling Forecast Modeli</li><li>DCF İskontolanmış Nakit Akımları &amp; Yatırım Fizibilitesi</li><li>Net İşletme Sermayesi (NÖS) &amp; DSO Gösterge Kokpiti</li></ul><a href="/sablonlar?q=butce" data-cta="home_finance_budget" data-location="finance_pillars">Bütçe sistemlerini gör →</a></article>
    </div>
  </div>
</section>`;

if (!html.includes('class="finance-pillars"')) requireReplace('<section class="catalog-proof"', `${financePillars}\n<section class="catalog-proof"`, 'finance pillars');

const bridge = `
<section class="high-ticket-bridge" data-experience-stage aria-labelledby="high-ticket-title">
  <div class="home-shell high-ticket-bridge__inner">
    <div class="high-ticket-bridge__copy"><p class="high-ticket-kicker">KURUMSAL ÇÖZÜMLER</p><h2 id="high-ticket-title">Hazır Şablonlar Sürecinize Dar mı Geliyor?</h2><p>Logo, SAP, Mikro veya Zirve verileriniz dağınıksa; şirketinize has tahsilat, banka, üretim veya kârlılık akışınız varsa, hazır kalıplarla vakit kaybetmeyin. 17 yıllık ticari bankacılık ve saha finans tecrübesiyle şirketinize özel karar motoru kuralım.</p></div>
    <div class="high-ticket-bridge__proof"><ul><li>Yazılımcıya finans anlatmakla uğraşmazsınız.</li><li>İş kuralları gerçek veriniz ve uç senaryolarla test edilir.</li><li>Kullanmayacağınız modül ve aylık lisans yükü eklenmez.</li></ul><a href="/ozel-excel-sistemleri" data-cta="home_high_ticket_bridge" data-location="mid">İhtiyaca Özel Excel Sistemleri Sayfasını İnceleyin →</a></div>
  </div>
</section>`;

if (!html.includes('class="high-ticket-bridge"')) requireReplace('<section class="authority"', `${bridge}\n<section class="authority"`, 'high-ticket bridge');

const bottom = `
<section class="home-finance-close" data-experience-stage aria-labelledby="home-finance-close-title">
  <div class="home-shell home-finance-close__inner"><div><p class="eyebrow">15 DAKİKALIK ÖN TEŞHİS</p><h2 id="home-finance-close-title">Şirketinizin Finansal Tıkanıklığını 15 Dakikada Teşhis Edelim.</h2><p>İster hazır model seçin, ister sürecinizi WhatsApp'tan yazın. Doğru çözüm yolunu ve gerekli kapsamı netleştirelim.</p></div><a class="home-finance-close__cta" href="${bottomWa}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="home_bottom_whatsapp" data-location="bottom">WhatsApp'tan Doğrudan Danışın</a></div>
</section>`;

if (!html.includes('class="home-finance-close"')) requireReplace('<section class="faq-close"', `${bottom}\n<section class="faq-close"`, 'bottom CTA');

const schema = {
  '@context':'https://schema.org',
  '@graph':[
    {'@type':'Organization','@id':'https://excelarsiv.com/#organization',name:'Excel Arşiv',url:'https://excelarsiv.com/',description:'İşletmeler için hazır ve ihtiyaca özel Excel finansal karar sistemleri.'},
    {'@type':'Person','@id':'https://excelarsiv.com/#baris-bagirlar',name:'Barış Bağırlar',jobTitle:'Ticari Bankacılık Uzmanı ve Finansal Sistem Mimarı',worksFor:{'@id':'https://excelarsiv.com/#organization'},url:'https://excelarsiv.com/ozel-excel-sistemleri',description:'17 yıllık ticari bankacılık ve saha finans deneyimiyle nakit akışı, banka limit-risk, maliyet, bütçe, değerleme ve yönetim raporlama süreçlerini Excel karar sistemlerine dönüştürür.'},
    {'@type':'ProfessionalService','@id':'https://excelarsiv.com/#financial-decision-systems',name:'Excel Arşiv Finansal Karar ve Excel Sistemleri',url:'https://excelarsiv.com/',provider:{'@id':'https://excelarsiv.com/#baris-bagirlar'},brand:{'@id':'https://excelarsiv.com/#organization'},areaServed:{'@type':'Country',name:'Türkiye'},description:'Hazır finans modelleri ile işletmeye özel Excel karar sistemlerini aynı çözüm mimarisinde sunar.',knowsAbout:['Nakit ve likidite sistemleri','Banka, kredi ve teminat sistemleri','Birim maliyet ve dinamik fiyatlama','Bütçe, projeksiyon ve değerleme','İşletmeye özel Excel karar sistemleri']}
  ]
};
html = html.replace('</head>', `<script type="application/ld+json" id="dual-funnel-home-schema-v17">${JSON.stringify(schema)}</script>\n</head>`);

const css = `<style id="dual-funnel-home-v17-css">
.finance-pillars{padding:64px 0;border-bottom:1px solid #dfe5e0;background:#fff}.finance-pillars__head{display:grid;grid-template-columns:minmax(0,.9fr) minmax(0,1.1fr);gap:42px;align-items:end;margin-bottom:28px}.finance-pillars__head h2{margin:10px 0 0;color:#0f172a;font-size:clamp(30px,3.6vw,48px);line-height:1.04;letter-spacing:-.04em}.finance-pillars__head>p{margin:0;color:#647068;font-size:15px;line-height:1.7}

/* Slayt 3 Mockup Sahnesi */
.pdf-showcase-stage{position:relative;margin:12px 0 38px;padding:36px 20px 44px;border-radius:28px;background:linear-gradient(150deg,#0a192f 0%,#0f172a 60%,#064e3b 100%);box-shadow:0 24px 60px rgba(10,25,47,.22);overflow:hidden;border:1px solid #1e3a5f}
.showcase-glow{position:absolute;left:50%;top:40%;transform:translate(-50%,-50%);width:680px;height:340px;background:radial-gradient(ellipse at center,rgba(16,185,129,.18) 0%,rgba(37,99,235,.12) 45%,transparent 70%);pointer-events:none}
.laptop-mockup-wrapper{position:relative;max-width:880px;margin:0 auto;z-index:2}
.laptop-frame{background:#1e293b;border-radius:18px 18px 0 0;padding:12px 12px 0;box-shadow:0 20px 50px rgba(0,0,0,.45);border:2px solid #334155;border-bottom:0}
.laptop-screen{background:#ffffff;border-radius:10px 10px 0 0;overflow:hidden;border:1px solid #cbd5e1}
.excel-ui-bar{display:flex;align-items:center;justify-content:space-between;padding:8px 14px;background:#107c41;color:#fff;font-size:12px;font-weight:700}
.excel-dots{display:flex;gap:6px}.excel-dots span{width:9px;height:9px;border-radius:50%}
.dot-red{background:#ef4444}.dot-yellow{background:#f59e0b}.dot-green{background:#10b981}
.excel-title-tab{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:-.01em}
.excel-status-tag{padding:3px 8px;border-radius:6px;background:rgba(255,255,255,.2);font:800 10px/1 ui-monospace,monospace;letter-spacing:.06em}
.excel-menu-strip{display:flex;gap:14px;padding:6px 14px;background:#f1f5f9;border-bottom:1px solid #cbd5e1;font-size:11.5px;color:#475569;font-weight:600}
.excel-menu-strip .active{color:#107c41;font-weight:800;border-bottom:2px solid #107c41}
.excel-formula-bar{display:flex;align-items:center;gap:8px;padding:6px 14px;background:#ffffff;border-bottom:1px solid #e2e8f0;font-size:12px;font-family:ui-monospace,monospace}
.fx-label{color:#64748b;font-weight:850;font-style:italic}
.formula-text{color:#0f172a;font-weight:700;letter-spacing:-.01em}
.excel-grid-preview{padding:16px 18px 20px;background:#f8fafc}
.excel-kpi-row{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin-bottom:14px}
.excel-kpi-card{padding:12px 14px;border-radius:12px;border:1px solid #e2e8f0;background:#fff;display:grid;gap:4px}
.excel-kpi-card small{color:#64748b;font-size:11px;font-weight:700}
.excel-kpi-card strong{font-size:18px;font-weight:850;line-height:1}
.excel-kpi-card span{color:#475569;font-size:10.5px}
.kpi-green strong{color:#059669}.kpi-blue strong{color:#2563eb}.kpi-amber strong{color:#d97706}
.excel-table-mock{border:1px solid #cbd5e1;border-radius:8px;background:#fff;overflow:hidden;font-size:12px}
.table-mock-head{display:grid;grid-template-columns:50px repeat(4,minmax(0,1fr)) 1.2fr;gap:6px;padding:8px 10px;background:#f1f5f9;font-weight:750;color:#334155;border-bottom:1px solid #cbd5e1}
.table-mock-row{display:grid;grid-template-columns:50px repeat(4,minmax(0,1fr)) 1.2fr;gap:6px;padding:7px 10px;border-bottom:1px solid #f1f5f9;color:#1e293b}
.table-mock-row.is-alert{background:#fffbeb;font-weight:700}
.cell-id{color:#64748b;font-family:ui-monospace,monospace;font-weight:750}
.status-pos{color:#059669;font-weight:800}.status-neg{color:#b45309;font-weight:800}
.laptop-base{height:16px;background:linear-gradient(180deg,#64748b 0%,#475569 100%);border-radius:0 0 16px 16px;position:relative;box-shadow:0 8px 18px rgba(0,0,0,.3)}
.laptop-notch{width:90px;height:5px;background:#334155;border-radius:0 0 5px 5px;margin:0 auto}

/* Floating Medallions */
.floating-medallion{position:absolute;z-index:10;display:flex;align-items:center;gap:10px;padding:10px 14px;border-radius:14px;background:rgba(15,23,42,.88);backdrop-filter:blur(12px);border:1px solid rgba(255,255,255,.2);box-shadow:0 12px 30px rgba(0,0,0,.35);color:#fff;width:fit-content;max-width:210px}
.floating-medallion .medallion-icon{font-size:20px;line-height:1}
.floating-medallion strong{display:block;font-size:12px;font-weight:850;color:#fff;line-height:1.2}
.floating-medallion small{display:block;font-size:10px;color:#94a3b8;line-height:1.2;margin-top:2px}
.medallion-tl{top:20px;left:-18px}
.medallion-tr{top:20px;right:-18px}
.medallion-bl{bottom:40px;left:-18px}
.medallion-br{bottom:40px;right:-18px}

/* Slayt 3 15 Numaralı Matris */
.finance-needs-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-bottom:44px}
.finance-need-item{display:flex;align-items:center;gap:12px;padding:12px 16px;border:1px solid #e1e8e3;border-radius:14px;background:#f8faf9;color:#18221b;text-decoration:none;font-size:13.5px;font-weight:700;transition:all .15s ease;box-shadow:0 2px 6px rgba(15,23,42,.02)}
.finance-need-item:hover{background:#fff;border-color:#059669;color:#059669;box-shadow:0 8px 18px rgba(5,150,105,.1);transform:translateY(-1.5px)}
.need-badge{display:inline-flex;align-items:center;justify-content:center;flex-shrink:0;width:28px;height:28px;border-radius:9px;background:#eaf6ee;color:#087a46;font:850 11px/1 ui-monospace,SFMono-Regular,Menlo,monospace}
.need-name{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}

/* Slayt 5 6'lı Bento Grid */
.pdf-security-bento-section{margin:0 0 44px;padding:36px;border-radius:24px;background:#f8fafc;border:1px solid #e2e8f0;box-shadow:0 10px 30px rgba(15,23,42,.03)}
.bento-section-head{margin-bottom:28px}
.bento-section-head h3{margin:8px 0 6px;color:#0f172a;font-size:clamp(22px,2.4vw,32px);letter-spacing:-.03em}
.bento-section-head p{margin:0;color:#64748b;font-size:15px}
.pdf-six-bento-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px}
.bento-item{padding:22px;border-radius:18px;background:#fff;border:1px solid #e2e8f0;box-shadow:0 4px 14px rgba(15,23,42,.03);display:flex;flex-direction:column}
.bento-num{display:inline-block;color:#059669;font:850 16px/1 ui-monospace,monospace;margin-bottom:12px}
.bento-item h4{margin:0 0 8px;color:#0f172a;font-size:16px;font-weight:800;letter-spacing:-.02em}
.bento-item p{margin:0;color:#475569;font-size:13.5px;line-height:1.55}

/* Slayt 21 7 Numaralı Standart */
.pdf-reasons-timeline-card{margin:0 0 44px;padding:34px;border-radius:24px;background:#fff;border:1px solid #dfe5e0;box-shadow:0 10px 32px rgba(15,23,42,.04)}
.timeline-head{margin-bottom:24px}
.timeline-head h4{margin:8px 0 0;color:#0f172a;font-size:clamp(22px,2.2vw,30px);letter-spacing:-.03em}
.pdf-reasons-list{display:grid;gap:12px}
.reason-row{display:grid;grid-template-columns:40px 1fr;gap:16px;align-items:start;padding:14px 18px;border-radius:14px;background:#f8faf9;border:1px solid #e6ede8}
.r-no{width:36px;height:36px;border-radius:10px;background:#059669;color:#fff;display:grid;place-items:center;font:850 14px/1 ui-monospace,monospace}
.r-content strong{display:block;color:#0f172a;font-size:14.5px;font-weight:800;margin-bottom:3px}
.r-content p{margin:0;color:#526176;font-size:13px;line-height:1.5}

/* Slayt 7 Chat Dialog */
.pdf-chat-dialog-card{margin:0 0 44px;padding:26px 30px;border-radius:24px;background:#f8fafc;border:1px solid #e2e8f0;box-shadow:0 8px 26px rgba(15,23,42,.03)}
.pdf-chat-header{margin-bottom:18px}
.pdf-chat-kicker{color:#059669;font:800 10.5px/1 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.12em}
.pdf-chat-header h4{margin:6px 0 0;color:#0f172a;font-size:19px;letter-spacing:-.03em}
.pdf-chat-bubbles{display:grid;gap:14px;margin-bottom:18px}
.chat-bubble{padding:16px 20px;border-radius:18px;border:1px solid #e2e8f0}
.chat-bubble--client{background:#fff;border-left:4px solid #f59e0b}
.chat-bubble--advisor{background:#ecfdf5;border-left:4px solid #059669}
.bubble-meta{display:flex;align-items:center;gap:10px;margin-bottom:8px}
.bubble-meta .avatar{font-size:16px}
.bubble-meta strong{color:#0f172a;font-size:13.5px}
.bubble-meta small{color:#64748b;font-size:11px}
.b-tag{margin-left:auto;padding:3px 8px;border-radius:6px;font:800 10px/1 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.05em}
.b-tag--alert{background:#fef3c7;color:#b45309}
.b-tag--success{background:#d1fae5;color:#047857}
.chat-bubble p{margin:0;color:#334155;font-size:13.5px;line-height:1.6}
.chat-divider{display:flex;align-items:center;gap:14px;margin:2px 0}
.divider-line{flex:1;height:1px;background:#cbd5e1}
.divider-badge{padding:4px 12px;border-radius:999px;background:#0f172a;color:#fff;font:800 10px/1 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.08em}
.chat-outcome-bar{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;padding:16px 20px;border-radius:16px;background:#fff;border:1px solid #e2e8f0}
.outcome-stat strong{display:block;color:#059669;font-size:16px;font-weight:850}
.outcome-stat span{color:#64748b;font-size:12px}

/* Slayt 23 Sektör Bulutu */
.pdf-sector-ecosystem-card{margin:0 0 44px;padding:24px 28px;border-radius:24px;background:#fff;border:1px solid #e2e8f0;box-shadow:0 8px 24px rgba(15,23,42,.03)}
.sector-eco-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}
.sector-eco-head h4{margin:6px 0 0;color:#0f172a;font-size:18px;letter-spacing:-.03em}
.pdf-card-badge{display:inline-block;padding:5px 11px;border-radius:8px;background:#e2e8f0;color:#0f172a;font:800 10.5px/1 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.08em;width:fit-content}
.sector-all-link{color:#059669;font-size:13px;font-weight:800;text-decoration:none}
.pdf-sectors-pill-cloud{display:flex;flex-wrap:wrap;gap:8px}
.sector-pill{padding:8px 14px;border-radius:11px;background:#f8fafc;color:#1e293b;font-size:12.5px;font-weight:750;border:1px solid #e2e8f0;box-shadow:0 2px 5px rgba(15,23,42,.02)}

/* 4 Sütun */
.finance-pillars__grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}.finance-pillar{position:relative;min-width:0;display:flex;flex-direction:column;padding:22px;border:1px solid #dfe5e0;border-radius:18px;background:#fff;box-shadow:0 10px 28px rgba(18,42,26,.05)}.finance-pillar::before{content:"";position:absolute;left:0;right:0;top:0;height:4px;border-radius:18px 18px 0 0;background:#059669}.finance-pillar--blue::before{background:#2563eb}.finance-pillar--violet::before{background:#7c3aed}.finance-pillar--amber::before{background:#d97706}.finance-pillar__no{color:#64748b;font:800 11px/1 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.08em}.finance-pillar h3{margin:13px 0 0;color:#0f172a;font-size:19px;line-height:1.18;letter-spacing:-.025em}.finance-pillar ul{display:grid;gap:9px;margin:17px 0 22px;padding:0;list-style:none;color:#526176;font-size:13px;line-height:1.5}.finance-pillar li{position:relative;padding-left:15px}.finance-pillar li::before{content:"";position:absolute;left:0;top:.62em;width:6px;height:6px;border-radius:50%;background:#c9d4cc}.finance-pillar>a{margin-top:auto;color:#075f39;font-size:13px;font-weight:800;text-decoration:none}.high-ticket-bridge{padding:68px 0;background:#fff;color:#0f172a}.high-ticket-bridge__inner{display:grid;grid-template-columns:minmax(0,1.15fr) minmax(0,.85fr);gap:68px;align-items:center;padding:48px 44px;border:1px solid #dfe5e0;border-radius:24px;background:#fff;box-shadow:0 14px 36px rgba(18,42,26,.05)}.high-ticket-kicker{margin:0;color:#059669;font:850 11px/1.2 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.14em}.high-ticket-bridge h2{max-width:760px;margin:12px 0 0;color:#0f172a;font-size:clamp(32px,4vw,50px);line-height:1.02;letter-spacing:-.045em}.high-ticket-bridge__copy>p:last-child{max-width:760px;margin:20px 0 0;color:#526176;font-size:16px;line-height:1.7}.high-ticket-bridge__proof{padding:26px;border:1px solid #dfe5e0;border-radius:18px;background:#f8faf9;box-shadow:0 8px 24px rgba(18,42,26,.04)}.high-ticket-bridge__proof ul{display:grid;gap:12px;margin:0 0 22px;padding:0;list-style:none;color:#334155;font-size:14px;line-height:1.45}.high-ticket-bridge__proof li{position:relative;padding-left:22px}.high-ticket-bridge__proof li::before{content:"✓";position:absolute;left:0;color:#059669;font-weight:900}.high-ticket-bridge__proof>a{min-height:48px;display:flex;align-items:center;justify-content:center;padding:0 18px;border-radius:12px;background:#059669;color:#fff;font-size:13px;font-weight:850;text-decoration:none;box-shadow:0 8px 20px rgba(5,150,105,.2)}.home-finance-close{padding:52px 0;border-bottom:1px solid #dfe5e0;background:linear-gradient(90deg,#f0fdf4,#eff6ff)}.home-finance-close__inner{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:36px;align-items:center}.home-finance-close h2{margin:9px 0 0;color:#0f172a;font-size:clamp(28px,3.2vw,42px);line-height:1.05;letter-spacing:-.04em}.home-finance-close p:not(.eyebrow){max-width:760px;margin:12px 0 0;color:#526176;font-size:15px;line-height:1.65}.home-finance-close__cta{min-height:50px;display:inline-flex;align-items:center;justify-content:center;padding:0 22px;border-radius:12px;background:#059669;color:#fff;font-size:14px;font-weight:850;text-decoration:none;box-shadow:0 12px 28px rgba(5,150,105,.2)}
@media(max-width:980px){.pdf-six-bento-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.finance-needs-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.chat-outcome-bar{grid-template-columns:1fr}.finance-pillars__grid{grid-template-columns:repeat(2,minmax(0,1fr))}.high-ticket-bridge__inner{grid-template-columns:1fr;gap:28px}.floating-medallion{position:static;max-width:100%;margin-top:10px}}
@media(max-width:720px){.pdf-showcase-stage{padding:20px 14px}.excel-kpi-row{grid-template-columns:1fr}.table-mock-head,.table-mock-row{grid-template-columns:40px repeat(2,1fr) 1fr}.table-mock-head span:nth-child(4),.table-mock-head span:nth-child(5),.table-mock-row span:nth-child(4),.table-mock-row span:nth-child(5){display:none}.pdf-six-bento-grid{grid-template-columns:1fr}.finance-needs-grid{grid-template-columns:1fr;gap:8px}.finance-pillars{padding:38px 0}.finance-pillars__head{grid-template-columns:1fr;gap:12px}.finance-pillars__head h2{font-size:28px}.finance-pillars__grid{grid-template-columns:1fr;gap:10px}.finance-pillar{padding:18px}.high-ticket-bridge{padding:42px 0}.high-ticket-bridge h2{font-size:32px}.high-ticket-bridge__proof{padding:20px}.home-finance-close{padding:38px 0}.home-finance-close__inner{grid-template-columns:1fr;gap:20px}.home-finance-close__cta{width:100%}}
</style>`;
html = html.replace('</head>', `${css}\n</head>`);

for (const token of ['finance-pillars','high-ticket-bridge','home_finance_cash','home_finance_credit','home_finance_cost','home_finance_budget','ProfessionalService']) {
  if (!html.includes(token)) throw new Error(`HOME DUAL FUNNEL V17: required token missing: ${token}`);
}

fs.writeFileSync(file, html, 'utf8');
console.log('HOME DUAL FUNNEL V17 PASS — finance pillars, high-ticket bridge, bottom WhatsApp CTA and entity schema active; homepage hero ownership left untouched.');
