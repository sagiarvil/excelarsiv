#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/index.html');
if (!fs.existsSync(file)) throw new Error('HOME DUAL FUNNEL V17: dist/index.html missing');
let html = fs.readFileSync(file, 'utf8');

const WA_PHONE = '905393333303';
const wa = (message) => `https://wa.me/${WA_PHONE}?text=${encodeURIComponent(message)}`;
const bottomWa = wa('Merhaba Barış Bey, excelarsiv.com üzerinden ulaşıyorum. Şirketimizin nakit açığı, banka kredi engelleri veya finansal tıkanıklığı için 90 dakikalık karar teşhisini netleştirmek istiyorum.');

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

const painCardsSection = '';


const financePillars = `
<section class="finance-pillars" id="finans-sistemleri" data-experience-stage>
  <div class="home-shell">
    
    <!-- PDF Slayt 2 & 24 Stili: Kurumsal Güven ve Entegrasyon Ekosistemi -->
    <div class="pdf-ecosystem-hero-card" aria-label="Kurumsal Finans Ekosistemi">
      <div class="eco-hero-left">
        <span class="decision-eyebrow">KURUMSAL GÜVENCE VE ALTYAPI</span>
        <h3 class="eco-hero-title">17 Yıllık Ticari Bankacılık Gücüyle Reel Sektör Karar Mimarisi</h3>
        <p class="eco-hero-desc">Mali tablolarınızı, nakit akışınızı ve kredi riskinizi ağır ve maliyetli yazılımlara bağımlı kalmadan; güvenli, açık formüllü ve yerel çalışan Excel sistemleriyle yönetin.</p>
        
        <div class="eco-pill-stack">
          <div class="eco-pill-box">
            <span class="eco-pill-icon">🏛️</span>
            <div>
              <strong>17 Yıllık Ticari Bankacılık Metodolojisi</strong>
              <small>Kredi komitelerinin ve tecrübeli CFO'ların baktığı rasyolarla modellenmiş hazır finansal refleks.</small>
            </div>
          </div>
          <div class="eco-pill-box">
            <span class="eco-pill-icon">⚡</span>
            <div>
              <strong>51 Hazır Karar Sistemi &amp; Özel Çözüm Mimarisi</strong>
              <small>İster 10 saniyede hazır sistemi indirin, ister şirketinize has darboğazı 3 günde karar motoruna dönüştürelim.</small>
            </div>
          </div>
          <div class="eco-pill-box">
            <span class="eco-pill-icon">🔒</span>
            <div>
              <strong>%100 Yerel Veri Mahremiyeti &amp; Sıfır Makro</strong>
              <small>Cironuz ve müşteri veriniz asla buluta çıkmaz; saf .xlsx formatında virüssüz ve engelsiz çalışır.</small>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Slayt 3 15 Numaralı Matris -->
    <div class="section-divider-title">
      <span class="decision-eyebrow">TEK TIKLA ÇALIŞAN FİNANSAL SİSTEMLER</span>
      <h3>15 Temel İşletme İhtiyacı İçin Hazır Karar Sistemleri</h3>
      <p>Tek tıkla ihtiyacınız olan karar motorunu seçin, hemen kullanmaya başlayın.</p>
    </div>
    <div class="finance-needs-grid" aria-label="15 Temel Çözüm İhtiyacı">
      <a href="/sablon/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi" class="finance-need-item" data-cta="home_finance_cash"><span class="need-badge">01</span><span class="need-name">13 Haftalık Dinamik Nakit Akışı</span></a>
      <a href="/sablon/banka-kredi-ve-taksit-takip-sistemi" class="finance-need-item" data-cta="home_finance_credit"><span class="need-badge">02</span><span class="need-name">Çoklu Banka &amp; Kredi Taksit Takibi</span></a>
      <a href="/sablon/akilli-kasa-defteri-ve-nakit-kontrol-sistemi" class="finance-need-item"><span class="need-badge">03</span><span class="need-name">Kasa Defteri &amp; Günlük Likidite</span></a>
      <a href="/sablon/cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi" class="finance-need-item"><span class="need-badge">04</span><span class="need-name">Müşteri Risk Skoru &amp; Cari Yaşlandırma</span></a>
      <a href="/sablonlar?q=maliyet" class="finance-need-item" data-cta="home_finance_cost"><span class="need-badge">05</span><span class="need-name">Birim Maliyet &amp; Dinamik Fiyatlama</span></a>
      <a href="/sablon/aylik-patron-finans-paneli" class="finance-need-item"><span class="need-badge">06</span><span class="need-name">Aylık Patron &amp; Yönetici Finans Paneli</span></a>
      <a href="/sablon/cek-senet-ve-vade-risk-sistemi" class="finance-need-item"><span class="need-badge">07</span><span class="need-name">Çek - Senet &amp; Vade Risk Portföyü</span></a>
      <a href="/sablonlar?q=pazaryeri" class="finance-need-item"><span class="need-badge">08</span><span class="need-name">Pazaryeri Net Kâr &amp; Komisyon Hesabı</span></a>
      <a href="/sablonlar?q=butce" class="finance-need-item" data-cta="home_finance_budget"><span class="need-badge">09</span><span class="need-name">3 Senaryolu Dinamik Bütçe &amp; Tahmin</span></a>
      <a href="/sablon/ithalat-depo-teslim-rafa-gelen-net-birim-maliyet" class="finance-need-item"><span class="need-badge">10</span><span class="need-name">İthalat Depo Teslim Birim Maliyet</span></a>
      <a href="/sablon/kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici" class="finance-need-item"><span class="need-badge">11</span><span class="need-name">Kıdem, İhbar &amp; Personel Maliyeti</span></a>
      <a href="/sablon/sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli" class="finance-need-item"><span class="need-badge">12</span><span class="need-name">Şirket Öz Kaynağı &amp; TTK 376 Cetveli</span></a>
      <a href="/sablon/pos-komisyon-ve-net-tahsilat-kontrol-sistemi" class="finance-need-item"><span class="need-badge">13</span><span class="need-name">POS Komisyonu &amp; Net Tahsilat Kontrolü</span></a>
      <a href="/sablon/uretim-recetesi-ve-zam-yansitma-hesaplayici" class="finance-need-item"><span class="need-badge">14</span><span class="need-name">Üretim Reçetesi &amp; Zam Yansıtma</span></a>
      <a href="/sablon/doviz-acik-pozisyonu-ve-kur-riski-stres-testi" class="finance-need-item"><span class="need-badge">15</span><span class="need-name">Döviz Pozisyonu &amp; Kur Riski Stres Testi</span></a>
    </div>

    <!-- MANDATE BÖLÜM 3: PROBLEM MAĞAZASI (Kullanıcı ürün adı bilmeden doğrudan problemle buluşur) -->
    <div class="problem-store-section" id="problemler" aria-label="Problem Mağazası">
      <div class="section-divider-title">
        <p class="decision-eyebrow">PROBLEM MAĞAZASI · İŞLETME DARBOĞAZLARI</p>
        <h3>İşletmenizde Hangi Karar Tıkanıyor?</h3>
        <p>Ürün adı ezberlemek zorunda değilsiniz. Şirketinizdeki soruya tıklayın, doğrudan çözen karar sistemine gidin.</p>
      </div>

      <div class="problem-store-grid">
        <a href="/sablon/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi" class="problem-store-card" style="background-color: #f0f7ff !important; border-color: #bfdbfe !important;">
          <div class="psc-top">
            <span class="psc-badge" style="background: #e0f2fe; color: #0369a1; border-color: #bae6fd;">01 · NAKİT DARBOĞAZI</span>
            <span class="psc-icon" style="background: #e0f2fe; border-color: #bae6fd;">💸</span>
          </div>
          <h4>“Önümüzdeki haftalarda para yetişecek mi?”</h4>
          <div class="psc-solution-box" style="background: rgba(255,255,255,0.85); border-color: #bfdbfe;">
            <small style="color: #0284c7;">Hazır Karar Çözümü</small>
            <p>13 Haftalık Dinamik Nakit Akışı ve Likidite Karar Sistemi</p>
          </div>
          <span class="psc-cta" style="background: #0284c7; color: #ffffff; border-color: #0284c7;">Problemi Çöz →</span>
        </a>

        <a href="/sablon/cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi" class="problem-store-card" style="background-color: #f0fdf4 !important; border-color: #bbf7d0 !important;">
          <div class="psc-top">
            <span class="psc-badge" style="background: #dcfce7; color: #15803d; border-color: #bbf7d0;">02 · TAHSİLAT RİSKİ</span>
            <span class="psc-icon" style="background: #dcfce7; border-color: #bbf7d0;">👥</span>
          </div>
          <h4>“Hangi müşteri nakdimi kilitliyor?”</h4>
          <div class="psc-solution-box" style="background: rgba(255,255,255,0.85); border-color: #bbf7d0;">
            <small style="color: #16a34a;">Hazır Karar Çözümü</small>
            <p>Cari Yaşlandırma, Müşteri Risk Skoru ve Tahsilat Takip Sistemi</p>
          </div>
          <span class="psc-cta" style="background: #16a34a; color: #ffffff; border-color: #16a34a;">Problemi Çöz →</span>
        </a>

        <a href="/sablon/banka-kredi-ve-taksit-takip-sistemi" class="problem-store-card" style="background-color: #fff7ed !important; border-color: #fed7aa !important;">
          <div class="psc-top">
            <span class="psc-badge" style="background: #ffedd5; color: #c2410c; border-color: #fed7aa;">03 · KREDİ &amp; FAİZ</span>
            <span class="psc-icon" style="background: #ffedd5; border-color: #fed7aa;">🏛️</span>
          </div>
          <h4>“Kredilerim hangi ay sıkıştıracak?”</h4>
          <div class="psc-solution-box" style="background: rgba(255,255,255,0.85); border-color: #fed7aa;">
            <small style="color: #ea580c;">Hazır Karar Çözümü</small>
            <p>Çoklu Banka Kredi Portföyü, Rotatif Faiz ve Taksit Takip Sistemi</p>
          </div>
          <span class="psc-cta" style="background: #ea580c; color: #ffffff; border-color: #ea580c;">Problemi Çöz →</span>
        </a>

        <a href="/sablon/proje-ve-is-bazinda-gercek-karlilik-sistemi" class="problem-store-card" style="background-color: #faf5ff !important; border-color: #e9d5ff !important;">
          <div class="psc-top">
            <span class="psc-badge" style="background: #f3e8ff; color: #7e22ce; border-color: #e9d5ff;">04 · KÂR SIZINTISI</span>
            <span class="psc-icon" style="background: #f3e8ff; border-color: #e9d5ff;">📈</span>
          </div>
          <h4>“Fiyat artırmazsam nerede zarar ediyorum?”</h4>
          <div class="psc-solution-box" style="background: rgba(255,255,255,0.85); border-color: #e9d5ff;">
            <small style="color: #9333ea;">Hazır Karar Çözümü</small>
            <p>Değişken Birim Maliyet, Katkı Payı ve Dinamik Fiyatlama Motoru</p>
          </div>
          <span class="psc-cta" style="background: #9333ea; color: #ffffff; border-color: #9333ea;">Problemi Çöz →</span>
        </a>

        <a href="/sablon/sube-karlilik-ve-nakit-hesaplayici" class="problem-store-card" style="background-color: #fff1f2 !important; border-color: #fecdd3 !important;">
          <div class="psc-top">
            <span class="psc-badge" style="background: #ffe4e6; color: #be123c; border-color: #fecdd3;">05 · ŞUBE VERİMLİLİĞİ</span>
            <span class="psc-icon" style="background: #ffe4e6; border-color: #fecdd3;">🏢</span>
          </div>
          <h4>“Şube gerçekten para kazanıyor mu?”</h4>
          <div class="psc-solution-box" style="background: rgba(255,255,255,0.85); border-color: #fecdd3;">
            <small style="color: #e11d48;">Hazır Karar Çözümü</small>
            <p>Şube Kârlılık, Metrekare Verimliliği ve Başabaş Noktası Sistemi</p>
          </div>
          <span class="psc-cta" style="background: #e11d48; color: #ffffff; border-color: #e11d48;">Problemi Çöz →</span>
        </a>

        <a href="/sablon/stok-satis-ve-nakit-baglanma-sistemi" class="problem-store-card" style="background-color: #fffbeb !important; border-color: #fef08a !important;">
          <div class="psc-top">
            <span class="psc-badge" style="background: #fef3c7; color: #b45309; border-color: #fde68a;">06 · ÖLÜ STOK</span>
            <span class="psc-icon" style="background: #fef3c7; border-color: #fde68a;">📦</span>
          </div>
          <h4>“Stokta ne kadar para bekliyor?”</h4>
          <div class="psc-solution-box" style="background: rgba(255,255,255,0.85); border-color: #fde68a;">
            <small style="color: #d97706;">Hazır Karar Çözümü</small>
            <p>Stok Devir Hızı, Kilitlenen Nakit ve Tasfiye Analiz Sistemi</p>
          </div>
          <span class="psc-cta" style="background: #d97706; color: #ffffff; border-color: #d97706;">Problemi Çöz →</span>
        </a>

        <a href="/sablon/sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli" class="problem-store-card" style="background-color: #f0fdfa !important; border-color: #99f6e4 !important;">
          <div class="psc-top">
            <span class="psc-badge" style="background: #ccfbf1; color: #0f766e; border-color: #99f6e4;">07 · VERGİ &amp; DENETİM</span>
            <span class="psc-icon" style="background: #ccfbf1; border-color: #99f6e4;">⚖️</span>
          </div>
          <h4>“Muhasebe hesabında hata yapıyor olabilir miyim?”</h4>
          <div class="psc-solution-box" style="background: rgba(255,255,255,0.85); border-color: #99f6e4;">
            <small style="color: #0d9488;">Hazır Karar Çözümü</small>
            <p>TTK 376 Özkaynak Koruma, Yeniden Değerleme ve Vergi Risk Sistemleri</p>
          </div>
          <span class="psc-cta" style="background: #0d9488; color: #ffffff; border-color: #0d9488;">Problemi Çöz →</span>
        </a>

        <a href="/sablon/asgari-ucret-zam-etkisi-fiyat-ayarlama-cetveli" class="problem-store-card" style="background-color: #fdf4ff !important; border-color: #f5d0fe !important;">
          <div class="psc-top">
            <span class="psc-badge" style="background: #fae8ff; color: #86198f; border-color: #f5d0fe;">08 · PERSONEL &amp; MALİYET</span>
            <span class="psc-icon" style="background: #fae8ff; border-color: #f5d0fe;">💼</span>
          </div>
          <h4>“Personel maliyetim kârımı ne kadar eritiyor?”</h4>
          <div class="psc-solution-box" style="background: rgba(255,255,255,0.85); border-color: #f5d0fe;">
            <small style="color: #a21caf;">Hazır Karar Çözümü</small>
            <p>Asgari Ücret Zam Etkisi, Kıdem Yükü ve İşçilik Maliyet Paneli</p>
          </div>
          <span class="psc-cta" style="background: #a21caf; color: #ffffff; border-color: #a21caf;">Problemi Çöz →</span>
        </a>
      </div>
    </div>

    <!-- PDF Slayt 18 & 21 Stili: Neden Excel Arşiv? Ortada Görsel ve 6 Simetrik Karar Standardı -->
    <div class="pdf-reasons-center-stage" aria-label="Neden Excel Arşiv Karar Standartları">
      <div class="reasons-stage-head">
        <span class="decision-eyebrow">NEDEN EXCEL ARŞİV?</span>
        <h3>Finansal Yönetimde 6 Katı Karar Standardı</h3>
        <p>Yazılımcı mantığıyla değil, reel sektör şirketlerinin mali tahlilini yönetmiş bankacı refleksiyle kurgulandı.</p>
      </div>

      <div class="reasons-three-col-layout">
        <!-- Sol 3 Madde (1, 2, 3) -->
        <div class="reasons-col">
          <div class="reason-pill-box">
            <span class="r-circle-num">1</span>
            <div>
              <strong>17 Yıllık Saha ve Bankacılık Disiplini</strong>
              <p>Kredi komitelerinin ve tecrübeli CFO'ların baktığı rasyolarla modellenmiş hazır finansal refleks.</p>
            </div>
          </div>
          <div class="reason-pill-box">
            <span class="r-circle-num">2</span>
            <div>
              <strong>Aynı Gün İndirme &amp; Anında Kullanım</strong>
              <p>Sipariş sonrası 10 saniyede indirin. Aylar süren eğitim veya kurulum beklemeden doğrudan kullanmaya başlayın.</p>
            </div>
          </div>
          <div class="reason-pill-box">
            <span class="r-circle-num">3</span>
            <div>
              <strong>%100 Açık Formül &amp; Sıfır Makro</strong>
              <p>Kilitli sayfa veya gizli hücre yok. Şirketinizin değişen ihtiyaçlarına göre istediğiniz gibi genişletin.</p>
            </div>
          </div>
        </div>

        <!-- Orta Görsel Sahnesi -->
        <div class="reasons-center-visual">
          <div class="center-img-wrapper">
            <img src="/images/kapak/kobi-finans-yonetim-paketi.webp" alt="Excel Arşiv KOBİ Finans Karar Sistemi" loading="lazy" class="center-kobi-img" />
            <div class="center-floating-tag">
              <span class="tag-icon">⭐</span>
              <div>
                <strong>Ticari &amp; KOBİ'lerin 1 Numaralı Tercihi</strong>
                <small>24 Sektörde 1.000+ Aktif İşletme</small>
              </div>
            </div>
          </div>
        </div>

        <!-- Sağ 3 Madde (4, 5, 6) -->
        <div class="reasons-col">
          <div class="reason-pill-box">
            <span class="r-circle-num">4</span>
            <div>
              <strong>Yazılımcıya Finans Anlatma Derdi Yok</strong>
              <p>İş kuralları gerçek şirket bilançoları, çek-senet döngüsü ve banka limit-risk gerçekleriyle hazır kodlandı.</p>
            </div>
          </div>
          <div class="reason-pill-box">
            <span class="r-circle-num">5</span>
            <div>
              <strong>Logo, SAP, Mikro ve Ekstrelerle Uyumlu</strong>
              <p>Verilerinizi tek bir standart veri tablosuna bağlayarak mutabakat süresini günlerden dakikalara indirin.</p>
            </div>
          </div>
          <div class="reason-pill-box">
            <span class="r-circle-num">6</span>
            <div>
              <strong>Tek Seferlik Ödeme (Sıfır Lisans Yükü)</strong>
              <p>Aylık veya yıllık abonelik dayatması yok. Teklifinizi alın, ömür boyu sınırsız şirket içi kullanın.</p>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- 15 Dakikalık WhatsApp Ön Teşhis ve Çözüm Diyaloğu -->
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

    <!-- PDF Slayt 24 Stili: Bankacılık Gücüyle Dijital Karar Çözümleri Tek Çatı Altında -->
    <div class="pdf-umbrella-vision-section" aria-label="Finansal Karar Ekosistemi">
      <div class="umbrella-copy">
        <span class="decision-eyebrow">BÜTÜNCÜL KARAR EKOSİSTEMİ</span>
        <h3>Piyasa Değişir, Şirketinizin Karar Gücü Baki Kalır</h3>
        <p>17 yıllık ticari bankacılık disiplini, güçlü finansal rasyolar ve pratik Excel sistemleriyle işletmenizi geleceğe taşıyın. Şirketinizin finansal sağlığını korumak bir söylem değil, karar sistemlerimizin temelidir.</p>
        <div class="umbrella-cta-row">
          <a href="/sablonlar" class="umbrella-primary-btn">51 Karar Sistemini Keşfedin →</a>
          <a href="/ozel-excel-sistemleri" class="umbrella-secondary-btn">Özel Mimari Talep Edin</a>
        </div>
      </div>

      <div class="umbrella-cluster-visual">
        <div class="cluster-bubble bubble-main">
          <span class="bubble-icon">🏛️</span>
          <strong>Excel Arşiv</strong>
          <small>Karar Mimarisi</small>
        </div>
        <div class="cluster-bubble bubble-1">
          <span class="bubble-icon">💸</span>
          <strong>Nakit Akışı</strong>
          <small>13 Hafta</small>
        </div>
        <div class="cluster-bubble bubble-2">
          <span class="bubble-icon">🏦</span>
          <strong>Banka &amp; Kredi</strong>
          <small>Limit Risk</small>
        </div>
        <div class="cluster-bubble bubble-3">
          <span class="bubble-icon">📈</span>
          <strong>Birim Maliyet</strong>
          <small>Fiyatlama</small>
        </div>
        <div class="cluster-bubble bubble-4">
          <span class="bubble-icon">📊</span>
          <strong>Senaryolu Bütçe</strong>
          <small>Projeksiyon</small>
        </div>
        <div class="cluster-bubble bubble-5">
          <span class="bubble-icon">⚡</span>
          <strong>Özel Sistemler</strong>
          <small>Kurumsal</small>
        </div>
      </div>
    </div>


  </div>
</section>`;

if (!html.includes('class="finance-pillars"')) requireReplace('<section class="difference"', `${financePillars}\n<section class="difference"`, 'finance pillars');

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
  <div class="home-shell home-finance-close__inner"><div><p class="decision-eyebrow">90 DAKİKALIK KARAR TEŞHİSİ</p><h2 id="home-finance-close-title">Şirketinizin Kredi Engellerini ve Nakit Darboğazını 90 Dakikada Çözelim.</h2><p>İster hazır sistemi seçin, ister mali tablonuzu WhatsApp'tan iletin. Bankacının gözüyle engelleri ve net karar adımlarını hemen belirleyelim.</p></div><a class="home-finance-close__cta" href="${bottomWa}" target="_blank" rel="noopener noreferrer" data-event="cta_whatsapp_click" data-cta="home_bottom_whatsapp" data-location="bottom">WhatsApp'tan Doğrudan Danışın (Barış Bağırlar)</a></div>
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
.finance-pillars{padding:60px 0;border-bottom:1px solid #dfe5e0;background:#fff}

/* 5 Saniyede Anlaşılan Değer Vaadi */
.decision-clarity-head{margin-bottom:34px}
.decision-eyebrow{color:#059669;font:850 11px/1.2 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.12em;margin-bottom:10px}
.decision-clarity-head h2{margin:0 0 14px;color:#0f172a;font-size:clamp(28px,3.5vw,46px);line-height:1.06;letter-spacing:-.04em}
.decision-lead{max-width:820px;margin:0 0 24px;color:#475569;font-size:16px;line-height:1.6}

.decision-quick-steps{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px;margin-bottom:38px}
.quick-step-box{display:flex;align-items:flex-start;gap:14px;padding:18px 20px;border-radius:16px;background:#f8fafc;border:1px solid #e2e8f0;box-shadow:0 4px 12px rgba(15,23,42,.02)}
.step-num{display:grid;place-items:center;flex-shrink:0;width:34px;height:34px;border-radius:10px;background:#059669;color:#fff;font:850 13px/1 ui-monospace,monospace}
.quick-step-box strong{display:block;color:#0f172a;font-size:14.5px;font-weight:800;margin-bottom:4px}
.quick-step-box small{display:block;color:#64748b;font-size:12.5px;line-height:1.45}

/* 4 Bento Acı & Karar Kartı (iyzico tarzı) */
.pain-decision-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:20px;margin-bottom:46px}
.decision-card{padding:28px 26px;border-radius:22px;border:1px solid #e2e8f0;background:#fff;box-shadow:0 10px 30px rgba(15,23,42,.04);display:flex;flex-direction:column;position:relative;overflow:hidden}
.decision-card::before{content:"";position:absolute;left:0;right:0;top:0;height:4px}
.decision-card--navy::before{background:#1e3a8a}
.decision-card--emerald::before{background:#059669}
.decision-card--violet::before{background:#7c3aed}
.decision-card--amber::before{background:#d97706}

.decision-card__top{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}
.tag-pill{padding:4px 10px;border-radius:8px;font:800 10.5px/1 ui-monospace,monospace;letter-spacing:.06em}
.tag-pill--alert{background:#fef2f2;color:#dc2626;border:1px solid #fecaca}
.tag-pill--success{background:#ecfdf5;color:#059669;border:1px solid #a7f3d0}
.tag-pill--purple{background:#f5f3ff;color:#7c3aed;border:1px solid #ddd6fe}
.tag-pill--warn{background:#fffbeb;color:#d97706;border:1px solid #fde68a}
.card-icon{font-size:24px}

.decision-card h3{margin:0 0 16px;color:#0f172a;font-size:21px;letter-spacing:-.03em;font-weight:800}
.pain-point{padding:12px 14px;border-radius:12px;background:#fef2f2;border-left:3px solid #ef4444;margin-bottom:12px}
.pain-point strong{display:block;color:#991b1b;font-size:11.5px;font-weight:850;margin-bottom:3px}
.pain-point p{margin:0;color:#7f1d1d;font-size:13px;line-height:1.5}

.solution-point{padding:12px 14px;border-radius:12px;background:#f0fdf4;border-left:3px solid #10b981;margin-bottom:20px}
.solution-point strong{display:block;color:#065f46;font-size:11.5px;font-weight:850;margin-bottom:3px}
.solution-point p{margin:0;color:#064e3b;font-size:13px;line-height:1.5}

.decision-link{margin-top:auto;display:inline-flex;align-items:center;color:#059669;font-size:13.5px;font-weight:800;text-decoration:none}
.decision-link:hover{text-decoration:underline}

/* Slayt 3 Mockup Sahnesi */

/* Problem Mağazası Stilleri (PDF Şablon Kutuları) */
.problem-store-section{margin:36px 0 44px;padding:36px 32px;border-radius:26px;background:linear-gradient(180deg,#f8faf9 0%,#f0fdf4 100%);border:1px solid #d1fae5}
.problem-store-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:20px;margin-top:24px}
.problem-store-card{position:relative;padding:24px 26px;border-radius:20px;background:#ffffff;border:1px solid #dbe5de;box-shadow:0 4px 16px rgba(15,23,42,.03);display:flex;flex-direction:column;text-decoration:none;transition:all .2s ease;overflow:hidden}
.problem-store-card:hover{transform:translateY(-4px);border-color:#059669;box-shadow:0 16px 36px rgba(5,150,105,.12)}
.psc-top{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:12px}
.psc-badge{display:inline-flex;align-items:center;padding:5px 11px;border-radius:8px;background:#ecfdf5;color:#047857;font:850 11px/1 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.06em;border:1px solid #bbf7d0}
.psc-icon{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:10px;background:#f0fdf4;border:1px solid #dcfce7;font-size:17px;flex-shrink:0}
.problem-store-card h4{margin:0 0 14px;color:#0f172a;font-size:18px;font-weight:800;letter-spacing:-.025em;line-height:1.3;min-height:46px}
.psc-solution-box{padding:12px 14px;border-radius:12px;background:#f8fafc;border:1px solid #e2e8f0;margin-bottom:18px;display:grid;gap:4px}
.psc-solution-box small{color:#059669;font-weight:850;font-size:10.5px;text-transform:uppercase;letter-spacing:.08em;font-family:ui-monospace,monospace}
.psc-solution-box p{margin:0;color:#334155;font-size:13px;font-weight:700;line-height:1.45}
.psc-cta{margin-top:auto;display:inline-flex;align-items:center;gap:6px;padding:8px 16px;border-radius:10px;background:#f0fdf4;color:#087a46;font-size:12.5px;font-weight:800;border:1px solid #bbf7d0;width:fit-content;transition:all .15s ease}
.problem-store-card:hover .psc-cta{background:#059669;color:#ffffff;border-color:#059669}
.problem-store-card:nth-child(1){background:#f0f7ff;border-color:#dbeafe}
.problem-store-card:nth-child(2){background:#f0fdf4;border-color:#dcfce7}
.problem-store-card:nth-child(3){background:#fff7ed;border-color:#ffedd5}
.problem-store-card:nth-child(4){background:#f5f3ff;border-color:#ede9fe}
.problem-store-card:nth-child(5){background:#fff1f2;border-color:#ffe4e6}
.problem-store-card:nth-child(6){background:#fffbeb;border-color:#fef3c7}
.problem-store-card .psc-solution-box{background:rgba(255,255,255,0.65)}
@media(max-width:820px){.problem-store-grid{grid-template-columns:1fr}}

.section-divider-title{margin:36px 0 18px}
.section-divider-title h3{margin:0 0 6px;color:#0f172a;font-size:22px;letter-spacing:-.03em}
.section-divider-title p{margin:0;color:#64748b;font-size:14.5px}

/* Slayt 3 15 Numaralı Matris */
.finance-needs-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-bottom:44px}
.finance-need-item{display:flex;align-items:center;gap:12px;padding:12px 16px;border:1px solid #e1e8e3;border-radius:14px;background:#f8faf9;color:#18221b;text-decoration:none;font-size:13.5px;font-weight:700;transition:all .15s ease;box-shadow:0 2px 6px rgba(15,23,42,.02)}
.finance-need-item:hover{background:#fff;border-color:#059669;color:#059669;box-shadow:0 8px 18px rgba(5,150,105,.1);transform:translateY(-1.5px)}
.need-badge{display:inline-flex;align-items:center;justify-content:center;flex-shrink:0;width:28px;height:28px;border-radius:9px;background:#eaf6ee;color:#087a46;font:850 11px/1 ui-monospace,SFMono-Regular,Menlo,monospace}
.need-name{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}


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


.high-ticket-bridge{padding:64px 0;background:#fff;color:#0f172a}.high-ticket-bridge__inner{display:grid;grid-template-columns:minmax(0,1.15fr) minmax(0,.85fr);gap:68px;align-items:center;padding:48px 44px;border:1px solid #dfe5e0;border-radius:24px;background:#fff;box-shadow:0 14px 36px rgba(18,42,26,.05)}.high-ticket-kicker{margin:0;color:#059669;font:850 11px/1.2 ui-monospace,SFMono-Regular,Menlo,monospace;letter-spacing:.14em}.high-ticket-bridge h2{max-width:760px;margin:12px 0 0;color:#0f172a;font-size:clamp(32px,4vw,50px);line-height:1.02;letter-spacing:-.045em}.high-ticket-bridge__copy>p:last-child{max-width:760px;margin:20px 0 0;color:#526176;font-size:16px;line-height:1.7}.high-ticket-bridge__proof{padding:26px;border:1px solid #dfe5e0;border-radius:18px;background:#f8faf9;box-shadow:0 8px 24px rgba(18,42,26,.04)}.high-ticket-bridge__proof ul{display:grid;gap:12px;margin:0 0 22px;padding:0;list-style:none;color:#334155;font-size:14px;line-height:1.45}.high-ticket-bridge__proof li{position:relative;padding-left:22px}.high-ticket-bridge__proof li::before{content:"✓";position:absolute;left:0;color:#059669;font-weight:900}.high-ticket-bridge__proof>a{min-height:48px;display:flex;align-items:center;justify-content:center;padding:0 18px;border-radius:12px;background:#059669;color:#fff;font-size:13px;font-weight:850;text-decoration:none;box-shadow:0 8px 20px rgba(5,150,105,.2)}.home-finance-close{padding:52px 0;border-bottom:1px solid #dfe5e0;background:linear-gradient(90deg,#f0fdf4,#eff6ff)}.home-finance-close__inner{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:36px;align-items:center}.home-finance-close h2{margin:9px 0 0;color:#0f172a;font-size:clamp(28px,3.2vw,42px);line-height:1.05;letter-spacing:-.04em}.home-finance-close p:not(.decision-eyebrow){max-width:760px;margin:12px 0 0;color:#526176;font-size:15px;line-height:1.65}.home-finance-close__cta{min-height:50px;display:inline-flex;align-items:center;justify-content:center;padding:0 22px;border-radius:12px;background:#059669;color:#fff;font-size:14px;font-weight:850;text-decoration:none;box-shadow:0 12px 28px rgba(5,150,105,.2)}

/* PDF Slayt 2 & 24: Kurumsal Güven ve Finans Mimarisi */
.pdf-ecosystem-hero-card{padding:38px 36px;border-radius:28px;background:linear-gradient(140deg,#f8fafc 0%,#eff6ff 60%,#ecfdf5 100%);border:1px solid #dbeafe;box-shadow:0 12px 34px rgba(15,23,42,.04);margin-bottom:44px}
.eco-hero-title{margin:8px 0 12px;color:#0f172a;font-size:clamp(22px,2.4vw,32px);line-height:1.15;letter-spacing:-.03em}
.eco-hero-desc{margin:0 0 24px;color:#475569;font-size:15px;line-height:1.6;max-width:860px}
.eco-pill-stack{display:grid;grid-template-columns:repeat(auto-fit, minmax(280px, 1fr));gap:16px}
.eco-pill-box{display:flex;align-items:flex-start;gap:14px;padding:16px 20px;border-radius:16px;background:#ffffff;border:1px solid #e2e8f0;box-shadow:0 4px 14px rgba(15,23,42,.025)}
.eco-pill-icon{font-size:24px;line-height:1;flex-shrink:0}
.eco-pill-box strong{display:block;color:#0f172a;font-size:14px;font-weight:800;margin-bottom:3px}
.eco-pill-box small{display:block;color:#64748b;font-size:12px;line-height:1.45}


/* PDF Slayt 18 & 21: Neden Excel Arşiv? (Ortada Görsel + 6 Simetrik Standart) */
.pdf-reasons-center-stage{margin:40px 0;padding:38px 34px;border-radius:28px;background:#ffffff;border:1px solid #dfe5e0;box-shadow:0 12px 36px rgba(15,23,42,.04)}
.reasons-stage-head{margin-bottom:28px}
.reasons-stage-head h3{margin:8px 0 6px;color:#0f172a;font-size:clamp(22px,2.4vw,32px);letter-spacing:-.03em}
.reasons-stage-head p{margin:0;color:#64748b;font-size:15px}
.reasons-three-col-layout{display:grid;grid-template-columns:1fr 1.05fr 1fr;gap:24px;align-items:center}
.reasons-col{display:grid;gap:14px}
.reason-pill-box{display:flex;align-items:flex-start;gap:14px;padding:16px 18px;border-radius:18px;background:#f8fafc;border:1px solid #e2e8f0;box-shadow:0 4px 14px rgba(15,23,42,.02)}
.r-circle-num{display:grid;place-items:center;width:34px;height:34px;border-radius:50%;background:#059669;color:#fff;font:850 14px/1 ui-monospace,monospace;flex-shrink:0;box-shadow:0 4px 10px rgba(5,150,105,.25)}
.reason-pill-box strong{display:block;color:#0f172a;font-size:14px;font-weight:800;margin-bottom:3px}
.reason-pill-box p{margin:0;color:#526176;font-size:12.5px;line-height:1.5}
.center-img-wrapper{position:relative;border-radius:20px;overflow:hidden;border:2px solid #e2e8f0;box-shadow:0 18px 44px rgba(15,23,42,.12)}
.center-kobi-img{width:100%;height:auto;display:block}
.center-floating-tag{position:absolute;bottom:12px;left:12px;right:12px;background:rgba(15,23,42,.92);backdrop-filter:blur(8px);padding:10px 14px;border-radius:12px;display:flex;align-items:center;gap:10px;color:#fff;border:1px solid rgba(255,255,255,.2)}
.center-floating-tag .tag-icon{font-size:20px}
.center-floating-tag strong{display:block;font-size:12.5px;font-weight:850;color:#fff}
.center-floating-tag small{display:block;font-size:10.5px;color:#93c5fd}

/* PDF Slayt 24: Ekosistem Küreleri */
.pdf-umbrella-vision-section{display:grid;grid-template-columns:1.1fr 1fr;gap:36px;align-items:center;margin:40px 0;padding:42px 38px;border-radius:28px;background:linear-gradient(135deg,#0a192f 0%,#0f172a 60%,#1e3a8a 100%);color:#fff;border:1px solid #1e3a5f;box-shadow:0 24px 60px rgba(10,25,47,.25)}
.pdf-umbrella-vision-section h3{margin:8px 0 12px;color:#fff;font-size:clamp(22px,2.4vw,34px);line-height:1.12;letter-spacing:-.03em}
.pdf-umbrella-vision-section p{margin:0;color:#cbd5e1;font-size:15px;line-height:1.65}
.umbrella-cta-row{display:flex;gap:14px;margin-top:24px;flex-wrap:wrap}
.umbrella-primary-btn{display:inline-flex;align-items:center;justify-content:center;min-height:48px;padding:0 22px;border-radius:12px;background:#059669;color:#fff;font-size:13.5px;font-weight:850;text-decoration:none;box-shadow:0 8px 20px rgba(5,150,105,.3)}
.umbrella-secondary-btn{display:inline-flex;align-items:center;justify-content:center;min-height:48px;padding:0 20px;border-radius:12px;background:rgba(255,255,255,.1);color:#fff;border:1px solid rgba(255,255,255,.25);font-size:13.5px;font-weight:750;text-decoration:none}
.umbrella-cluster-visual{display:flex;flex-wrap:wrap;gap:12px;align-items:center;justify-content:center;padding:10px}
.cluster-bubble{padding:14px 18px;border-radius:20px;background:rgba(255,255,255,.12);backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,.2);text-align:center;color:#fff;box-shadow:0 8px 20px rgba(0,0,0,.25);display:flex;flex-direction:column;align-items:center;gap:3px;transition:all .15s ease}
.cluster-bubble .bubble-icon{font-size:22px;line-height:1}
.cluster-bubble strong{font-size:13px;font-weight:850;color:#fff}
.cluster-bubble small{font-size:10.5px;color:#93c5fd}
.cluster-bubble.bubble-main{background:linear-gradient(135deg,#059669,#10b981);border-color:#34d399;transform:scale(1.08);box-shadow:0 12px 28px rgba(5,150,105,.35)}
@media(max-width:980px){
  .decision-quick-steps{grid-template-columns:1fr}
  .pain-decision-grid{grid-template-columns:1fr}

  .finance-needs-grid{grid-template-columns:repeat(2,minmax(0,1fr))}
  .chat-outcome-bar{grid-template-columns:1fr}
  .high-ticket-bridge__inner{grid-template-columns:1fr;gap:28px}

  .reasons-three-col-layout{grid-template-columns:1fr;gap:24px}
  .pdf-umbrella-vision-section{grid-template-columns:1fr;gap:28px}
}
@media(max-width:768px){
  .finance-pillars{padding:36px 0;overflow-x:clip}
  .decision-clarity-head{margin-bottom:24px}
  .decision-clarity-head h2{font-size:clamp(23px,5.8vw,34px);line-height:1.15;letter-spacing:-.03em}
  .decision-lead{font-size:14.5px;line-height:1.55;margin-bottom:18px}
  .decision-quick-steps{gap:10px;margin-bottom:28px}
  .quick-step-box{padding:14px 16px;border-radius:14px}
  .pain-decision-grid{gap:14px;margin-bottom:32px}
  .decision-card{padding:20px 18px;border-radius:18px}
  .decision-card h3{font-size:18px;margin-bottom:12px}
  .pain-point,.solution-point{padding:10px 12px;margin-bottom:10px}
  .pain-point p,.solution-point p{font-size:12.5px}

  /* PDF Yeni Blok Mobil Kuralları */
  .pdf-ecosystem-hero-card{padding:20px 16px;border-radius:20px;margin-bottom:30px}
  .pdf-reasons-center-stage{padding:20px 16px;border-radius:20px;margin:28px 0}
  .pdf-umbrella-vision-section{padding:24px 18px;border-radius:20px;margin:28px 0}
  .umbrella-cta-row a{width:100%;box-sizing:border-box;text-align:center}

  /* Problem Mağazası Mobil Standardı */
  .problem-store-section{margin:24px 0 32px;padding:20px 14px;border-radius:20px}
  .problem-store-grid{grid-template-columns:1fr;gap:12px;margin-top:16px}
  .problem-store-card{padding:18px 16px;border-radius:16px}
  .problem-store-card h4{font-size:16px;min-height:auto;margin-bottom:12px;line-height:1.35}
  .psc-solution-box{padding:10px 12px;border-radius:10px;margin-bottom:14px}
  .psc-solution-box p{font-size:12.5px;line-height:1.45}
  .psc-cta{width:100%;box-sizing:border-box;justify-content:center;min-height:44px;font-size:13px;border-radius:10px;touch-action:manipulation}

  /* 15 İhtiyaç Matrisi */
  .finance-needs-grid{grid-template-columns:1fr;gap:8px;margin-bottom:32px}
  .finance-need-item{padding:11px 14px;font-size:13px;min-height:44px;border-radius:12px}
  .need-name{white-space:normal;line-height:1.3}


  /* Chat Teşhis */
  .pdf-chat-dialog-card{margin:0 0 32px;padding:18px 14px;border-radius:20px}
  .pdf-chat-header h4{font-size:17px;line-height:1.3}
  .pdf-chat-bubbles{gap:12px;margin-bottom:14px}
  .chat-bubble{padding:14px 14px;border-radius:14px}
  .bubble-meta{flex-wrap:wrap;gap:6px}
  .bubble-meta strong{font-size:12.5px}
  .bubble-meta small{font-size:10px}
  .chat-bubble p{font-size:12.5px;line-height:1.55}
  .chat-divider{margin:0}
  .divider-badge{font-size:9.5px;padding:3px 10px}
  .chat-outcome-bar{grid-template-columns:1fr;gap:10px;padding:14px 16px;border-radius:14px}
  .outcome-stat strong{font-size:15px}
  .outcome-stat span{font-size:11.5px}
  /* Alt Dönüşüm Köprüleri */
  .high-ticket-bridge{padding:36px 0;overflow-x:clip}
  .high-ticket-bridge__inner{padding:26px 18px;border-radius:20px;gap:22px}
  .high-ticket-bridge h2{font-size:clamp(22px,5.5vw,32px);line-height:1.1}
  .high-ticket-bridge__copy>p:last-child{font-size:14px;margin-top:14px;line-height:1.6}
  .high-ticket-bridge__proof{padding:18px 16px;border-radius:16px}
  .high-ticket-bridge__proof ul{font-size:13px;gap:10px;margin-bottom:18px}
  .high-ticket-bridge__proof>a{min-height:46px;font-size:12.5px;text-align:center;padding:0 14px}
  .home-finance-close{padding:32px 0;overflow-x:clip}
  .home-finance-close__inner{grid-template-columns:1fr;gap:18px}
  .home-finance-close h2{font-size:clamp(21px,5vw,30px);line-height:1.2}
  .home-finance-close p:not(.decision-eyebrow){font-size:13.5px;line-height:1.55}
  .home-finance-close__cta{min-height:48px;font-size:13.5px;width:100%;box-sizing:border-box;text-align:center;padding:0 16px}
}
@media(max-width:480px){
  .quick-step-box{flex-direction:column;gap:8px}
  .step-num{width:28px;height:28px;font-size:12px;border-radius:8px}
  .table-mock-head,.table-mock-row{grid-template-columns:30px 1fr 1fr;font-size:10px}
  .table-mock-head span:nth-child(3),.table-mock-row span:nth-child(3){display:none}
  .r-content strong{font-size:13.5px}
  .bento-item h4{font-size:15px}
  .eco-logos-grid{grid-template-columns:1fr}
}
</style>`;
html = html.replace('</head>', `${css}\n</head>`);

for (const token of ['finance-pillars','high-ticket-bridge','home_finance_cash','home_finance_credit','home_finance_cost','home_finance_budget','ProfessionalService']) {
  if (!html.includes(token)) throw new Error(`HOME DUAL FUNNEL V17: required token missing: ${token}`);
}

fs.writeFileSync(file, html, 'utf8');
console.log('HOME DUAL FUNNEL V17 PASS — 5-second clarity, pain-decision grid and high-ticket decision pipeline active.');
