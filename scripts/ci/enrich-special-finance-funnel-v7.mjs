#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const distFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(distFile)) throw new Error('SPECIAL FINANCE FUNNEL V7: dist route missing');

let html = fs.readFileSync(distFile, 'utf8');

const WA_URL = 'https://wa.me/905419305372?text=Merhaba%2C%20ExcelAr%C5%9Fiv%20%C3%B6zel%20Excel%20sistemi%20i%C3%A7in%20i%C5%9F%20ak%C4%B1%C5%9F%C4%B1m%C4%B1%20anlatmak%20istiyorum.';

const replaceOne = (pattern, replacement, label) => {
  const before = html;
  html = html.replace(pattern, replacement);
  if (html === before) throw new Error(`SPECIAL FINANCE FUNNEL V7: replacement missed: ${label}`);
};

const replaceSection = (id, replacement) => {
  replaceOne(new RegExp(`<section\\b[^>]*\\bid="${id}"[^>]*>[\\s\\S]*?<\\/section>`, 'u'), replacement, `section#${id}`);
};

if (html.includes('data-finance-funnel-v7')) {
  console.log('SPECIAL FINANCE FUNNEL V7: already applied');
  process.exit(0);
}

replaceOne(/<title>[\s\S]*?<\/title>/u, '<title>Özel Excel Karar Sistemleri | 17 Yıllık Ticari Bankacılık Tecrübesi | ExcelArşiv</title>', 'title');
replaceOne(/<meta name="description" content="[^"]*"\s*\/>/u, '<meta name="description" content="13 haftalık nakit akışı, banka limit-risk, dinamik fiyatlama, rolling forecast, DCF/DSCR ve ERP verisinden CFO kokpiti için işletmeye özel Excel karar sistemleri." />', 'meta description');
replaceOne(/<meta property="og:title" content="[^"]*"\s*\/>/u, '<meta property="og:title" content="Özel Excel Karar Sistemleri | ExcelArşiv" />', 'og title');
replaceOne(/<meta property="og:description" content="[^"]*"\s*\/>/u, '<meta property="og:description" content="17 yıllık ticari bankacılık bakışıyla nakit, kredi, kârlılık ve ERP verisini yönetici karar ekranına dönüştüren özel Excel sistemleri." />', 'og description');

const schema = {
  '@context': 'https://schema.org',
  '@graph': [
    {
      '@type': 'BreadcrumbList',
      itemListElement: [
        {'@type': 'ListItem', position: 1, name: 'Ana Sayfa', item: 'https://excelarsiv.com'},
        {'@type': 'ListItem', position: 2, name: 'Özel Excel Karar Sistemleri', item: 'https://excelarsiv.com/ozel-excel-sistemleri'},
      ],
    },
    {
      '@type': 'Person',
      '@id': 'https://excelarsiv.com/#baris-bagirlar',
      name: 'Barış Bağırlar',
      jobTitle: 'Finansal Danışman ve Excel Sistem Tasarımcısı',
      url: 'https://excelarsiv.com/ozel-excel-sistemleri',
      sameAs: ['https://www.linkedin.com/in/barisbagirlar/'],
      worksFor: {'@type': 'Organization', '@id': 'https://excelarsiv.com/#organization', name: 'ExcelArşiv', url: 'https://excelarsiv.com'},
      description: '17 yıllık ticari bankacılık geçmişiyle işletmelerin nakit, kredi, tahsilat, kârlılık ve yönetim raporlama ihtiyaçlarını özel Excel karar sistemlerine dönüştürür.',
      knowsAbout: [
        '13 haftalık nakit akışı',
        'ticari kredi ve banka limit-risk yönetimi',
        'DSCR borç servis kapasitesi',
        'DCF şirket değerleme',
        'dinamik fiyatlama ve başabaş analizi',
        'ERP yönetim raporlaması',
      ],
    },
    {
      '@type': 'Organization',
      '@id': 'https://excelarsiv.com/#organization',
      name: 'ExcelArşiv',
      url: 'https://excelarsiv.com',
    },
    {
      '@type': 'Service',
      '@id': 'https://excelarsiv.com/ozel-excel-sistemleri#service',
      name: 'Özel Excel Finans ve Karar Sistemleri',
      serviceType: 'Finansal modelleme, risk kontrolü ve yönetim raporlama sistemi tasarımı',
      provider: {'@id': 'https://excelarsiv.com/#baris-bagirlar'},
      brand: {'@id': 'https://excelarsiv.com/#organization'},
      areaServed: {'@type': 'Country', name: 'Türkiye'},
      url: 'https://excelarsiv.com/ozel-excel-sistemleri',
      description: 'Dağınık işletme ve ERP verisini ticari-finansal kurallardan geçirerek patron ve CFO için karar ekranına dönüştüren işletmeye özel Excel sistemleri.',
      hasOfferCatalog: {
        '@type': 'OfferCatalog',
        name: 'Özel Finansal Karar Sistemleri',
        itemListElement: [
          {'@type': 'Offer', itemOffered: {'@type': 'Service', name: '13 Haftalık Dinamik Nakit Akışı Modeli (13-Week Cash Flow)'}},
          {'@type': 'Offer', itemOffered: {'@type': 'Service', name: 'Çoklu Banka Limit-Risk, Kredi ve Teminat Matrisi'}},
          {'@type': 'Offer', itemOffered: {'@type': 'Service', name: 'Birim Maliyet, Dinamik Fiyatlama ve Başabaş Simülatörü'}},
          {'@type': 'Offer', itemOffered: {'@type': 'Service', name: 'Senaryolu Bütçe ve Dinamik Projeksiyon (Rolling Forecast)'}},
          {'@type': 'Offer', itemOffered: {'@type': 'Service', name: 'Yatırım Fizibilitesi ve DCF Şirket Değerleme Modeli'}},
          {'@type': 'Offer', itemOffered: {'@type': 'Service', name: 'ERP Entegre CFO Yönetim Kokpiti (Executive Dashboard)'}},
        ],
      },
    },
    {
      '@type': 'FAQPage',
      mainEntity: [
        {'@type': 'Question', name: 'Özel Excel sistemi hazır şablondan nasıl ayrılır?', acceptedAnswer: {'@type': 'Answer', text: 'Hazır şablon önceden belirlenmiş bir düzene göre çalışır. Özel sistemde veri girişi, hesaplama, kontrol ve raporlama işletmenin gerçek iş akışına ve karar ihtiyacına göre kurulur.'}},
        {'@type': 'Question', name: 'ERP veya muhasebe programı verileri kullanılabilir mi?', acceptedAnswer: {'@type': 'Answer', text: 'Uygun veri dışa aktarımı bulunduğunda Logo, SAP, Mikro, Zirve ve benzeri sistemlerden alınan Excel veya CSV verileri karar ve yönetim raporlaması için kullanılabilir.'}},
        {'@type': 'Question', name: 'Makro kullanmak zorunlu mu?', acceptedAnswer: {'@type': 'Answer', text: 'Hayır. İhtiyaç standart Excel işlevleriyle güvenli biçimde çözülebiliyorsa makro kullanılmaz; otomasyon gerçekten gerekiyorsa kapsam ayrıca belirlenir.'}},
        {'@type': 'Question', name: 'Fiyat nasıl belirlenir?', acceptedAnswer: {'@type': 'Answer', text: 'Fiyat; veri yapısı, finansal mantık, kontrol noktaları, rapor sayısı, otomasyon düzeyi ve teslim kapsamına göre belirlenir. Gereksiz modül eklenmez.'}},
        {'@type': 'Question', name: 'Veri güvenliği nasıl ele alınır?', acceptedAnswer: {'@type': 'Answer', text: 'İnceleme öncesinde hassas alanlar maskelenebilir. Projenin gerektirmediği durumlarda veriyi harici bir sisteme taşıma zorunluluğu oluşturulmaz.'}},
      ],
    },
  ],
};

replaceOne(/<script type="application\/ld\+json">[\s\S]*?<\/script>/u, `<script type="application/ld+json">${JSON.stringify(schema)}</script>`, 'json-ld');

const css = `
  <style id="finance-funnel-v7-css">
    body[data-finance-funnel-v7]{--ink:#0f172a;--ink-2:#1e293b;--green:#059669;--green-dark:#047857;--green-soft:#ecfdf5;--green-mid:#d1fae5;--blue:#2563eb;--blue-soft:#eff6ff;--amber:#d97706;--amber-soft:#fffbeb}
    body[data-finance-funnel-v7] .hero .eyebrow{border-color:#bfdbfe;background:#eff6ff;color:#1d4ed8}
    body[data-finance-funnel-v7] .hero .eyebrow::before{background:#2563eb}
    body[data-finance-funnel-v7] .hero h1{font-size:clamp(38px,4.2vw,56px);line-height:1.055;letter-spacing:-.047em;color:#0f172a}
    body[data-finance-funnel-v7] .hero-lead{font-size:18px;line-height:1.62;color:#334155}
    body[data-finance-funnel-v7] .hero-microcopy{margin:12px 0 0;color:#64748b;font-size:14px;line-height:1.45;font-weight:550}
    body[data-finance-funnel-v7] .proof-chip strong{font-size:14px}
    body[data-finance-funnel-v7] .proof-chip span{font-size:14px;line-height:1.42}
    body[data-finance-funnel-v7] .node-flow{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:28px;align-items:stretch}
    body[data-finance-funnel-v7] .node-card{position:relative;min-width:0;height:100%;padding:28px;border:1px solid #dbe3eb;border-radius:18px;background:#fff;box-shadow:0 10px 24px rgba(15,23,42,.06)}
    body[data-finance-funnel-v7] .node-card::before{content:"";position:absolute;left:0;right:0;top:0;height:5px;border-radius:18px 18px 0 0;background:#2563eb}
    body[data-finance-funnel-v7] .node-card:nth-child(2)::before{background:#d97706}
    body[data-finance-funnel-v7] .node-card:nth-child(3)::before{background:#059669}
    body[data-finance-funnel-v7] .node-card:not(:last-child)::after{content:"→";position:absolute;right:-22px;top:50%;transform:translateY(-50%);z-index:2;width:28px;height:28px;display:grid;place-items:center;border:1px solid #dbe3eb;border-radius:50%;background:#fff;color:#64748b;font-size:16px;font-weight:800}
    body[data-finance-funnel-v7] .node-label{display:inline-flex;align-items:center;gap:8px;margin-bottom:18px;padding:7px 10px;border-radius:999px;background:#eff6ff;color:#1d4ed8;font-size:13px;line-height:1.2;font-weight:800;letter-spacing:.03em}
    body[data-finance-funnel-v7] .node-card:nth-child(2) .node-label{background:#fffbeb;color:#b45309}
    body[data-finance-funnel-v7] .node-card:nth-child(3) .node-label{background:#ecfdf5;color:#047857}
    body[data-finance-funnel-v7] .node-card h3{margin:0;font-size:23px;line-height:1.2;font-weight:720;letter-spacing:-.03em}
    body[data-finance-funnel-v7] .node-card p{margin:12px 0 0;color:#526176;font-size:16px;line-height:1.55}
    body[data-finance-funnel-v7] .node-tags{display:flex;flex-wrap:wrap;gap:7px;margin-top:18px}
    body[data-finance-funnel-v7] .node-tags span{padding:6px 9px;border:1px solid #e2e8f0;border-radius:999px;background:#f8fafc;color:#475569;font-size:13px;line-height:1.2;font-weight:600}
    body[data-finance-funnel-v7] .area-card h3{font-size:20px}
    body[data-finance-funnel-v7] .area-en{display:block;margin-top:4px;color:#64748b;font-size:14px;font-weight:650;letter-spacing:-.01em}
    body[data-finance-funnel-v7] .area-benefit{margin-top:14px!important;padding-top:13px;border-top:1px solid #e6ebf0;color:#334155!important;font-size:15px!important}
    body[data-finance-funnel-v7] .area-benefit strong{color:#0f172a}
    body[data-finance-funnel-v7] .authority-summary{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin:0 0 18px}
    body[data-finance-funnel-v7] .authority-stat{padding:18px;border:1px solid #dbe3eb;border-radius:16px;background:#fff}
    body[data-finance-funnel-v7] .authority-stat strong{display:block;color:#0f172a;font-size:23px;line-height:1.15;font-weight:760;letter-spacing:-.03em}
    body[data-finance-funnel-v7] .authority-stat span{display:block;margin-top:6px;color:#64748b;font-size:14px;line-height:1.4}
    body[data-finance-funnel-v7] .why-card p{font-size:16px}
    body[data-finance-funnel-v7] .comparison-row>div,body[data-finance-funnel-v7] .comparison-head>div{font-size:15px;line-height:1.5}
    body[data-finance-funnel-v7] .cta-email{display:flex;justify-content:center;margin-top:3px;color:#334155;font-size:14px;font-weight:650;text-decoration:underline;text-decoration-color:#a7b3c2;text-underline-offset:3px}
    body[data-finance-funnel-v7] .btn-whatsapp{background:#059669;color:#fff;box-shadow:0 10px 22px rgba(5,150,105,.2)}
    body[data-finance-funnel-v7] .btn-whatsapp:hover{background:#047857}
    @media(max-width:980px){body[data-finance-funnel-v7] .node-flow{grid-template-columns:1fr}body[data-finance-funnel-v7] .node-card:not(:last-child)::after{content:"↓";right:auto;left:50%;top:auto;bottom:-29px;transform:translateX(-50%)}}
    @media(max-width:660px){body[data-finance-funnel-v7] .hero h1{font-size:38px}body[data-finance-funnel-v7] .hero-lead{font-size:17px}body[data-finance-funnel-v7] .authority-summary{grid-template-columns:1fr}body[data-finance-funnel-v7] .node-card{padding:23px}}
    @media(max-width:430px){body[data-finance-funnel-v7] .hero h1{font-size:34px}}
  </style>`;
replaceOne('</head>', `${css}\n</head>`, 'finance CSS');

replaceOne('<body data-special-source-v5>', '<body data-special-source-v5 data-finance-funnel-v7>', 'body data attribute');

html = html.replaceAll('Ne zaman gerekir?', 'Sistem nasıl kurulur?');
html = html.replaceAll('Neden ExcelArşiv?', 'Neden Barış Bağırlar?');

replaceOne(/<a class="btn btn-primary nav-cta" href="\/iletisim">İhtiyacımı anlat<\/a>/u, `<a class="btn btn-primary nav-cta btn-whatsapp" href="${WA_URL}" target="_blank" rel="noopener noreferrer">WhatsApp'tan anlat</a>`, 'desktop nav CTA');
replaceOne(/<a href="\/iletisim">İhtiyacımı anlat<\/a>/u, `<a href="${WA_URL}" target="_blank" rel="noopener noreferrer">WhatsApp'tan anlat</a>`, 'mobile nav CTA');

const heroCopy = `<div class="hero-copy">
          <p class="eyebrow">17 yıllık ticari bankacılık &amp; saha finans tecrübesi</p>
          <h1>Yazılımcıya finans öğretmekle uğraşmayın. <em>Şirketinizin dilini bilen özel Excel karar sistemleri.</em></h1>
          <p class="hero-lead">Esnafın tahsilat sancısını da, CFO'nun bilanço ve likidite riskini de aynı finansal çerçevede okuyoruz. Renkli tablo değil; Logo, SAP, Mikro, Zirve, banka ve mevcut Excel verisini ticari kurallardan geçirip paranızı, kârınızı ve riskinizi görünür kılan karar modelleri kuruyoruz.</p>
          <div class="hero-actions">
            <a class="btn btn-primary btn-whatsapp" href="${WA_URL}" target="_blank" rel="noopener noreferrer">WhatsApp'tan Sürecinizi Anlatın <span aria-hidden="true">→</span></a>
            <a class="btn btn-secondary" href="#alanlar">Örnek Karar Modellerini İnceleyin <span aria-hidden="true">↓</span></a>
          </div>
          <p class="hero-microcopy">Kısa ön teşhis • Dosya maskeleme desteği • Teknik brifing gerekmez</p>
          <ul class="hero-bullets" aria-label="Hizmet özellikleri"><li>Nakit &amp; likidite</li><li>Kredi &amp; limit</li><li>Kârlılık &amp; fiyatlama</li><li>ERP verisi → yönetici kararı</li></ul>
          <div class="proof-grid" aria-label="ExcelArşiv yaklaşımı">
            <div class="proof-chip"><strong>Bankanın baktığı yerden</strong><span>Limit, vade, borç servisi ve riski birlikte okur.</span></div>
            <div class="proof-chip"><strong>Sahanın konuştuğu dilden</strong><span>Tahsilat, kasa, ödeme ve fiyat sorununu gerçek akışa bağlar.</span></div>
            <div class="proof-chip"><strong>ERP'den karar ekranına</strong><span>Dağınık veriyi tek yönetim mantığında birleştirir.</span></div>
            <div class="proof-chip"><strong>Gerektiği kadar sistem</strong><span>Kullanmayacağınız modülü kapsam ve maliyete yüklemez.</span></div>
          </div>
        </div>`;
replaceOne(/<div class="hero-copy">[\s\S]*?<\/div>\s*\n\s*<div class="workbook"/u, `${heroCopy}\n\n        <div class="workbook"`, 'hero copy');

replaceSection('ne-zaman', `<section class="section section-soft" id="ne-zaman">
      <div class="wrap-wide">
        <div class="section-head"><span class="section-kicker">NODE TABANLI TERSİNE MÜHENDİSLİK</span><h2>Girdi → Bankacı Mantığı → Karar Ekranı</h2><p>İhtiyacınız formül değil; dağınık veriyi ticari-finansal kurallardan geçirip patronun ve CFO'nun karar vereceği tek ekrana dönüştüren bir iş akışıdır.</p></div>
        <div class="node-flow" aria-label="Özel Excel karar sistemi veri hattı">
          <article class="node-card"><span class="node-label">NODE 01 · GİRDİ</span><h3>Dağınık veriyi temizleyip tek merkeze alırız.</h3><p>Her ay elle toplanan faturalar, banka listeleri, ERP dışa aktarımları, Excel/CSV dosyaları ve ekip notları aynı veri sözlüğünde buluşur.</p><div class="node-tags"><span>Logo</span><span>SAP</span><span>Mikro</span><span>Zirve</span><span>Banka Ekstresi</span><span>Excel / CSV</span></div></article>
          <article class="node-card"><span class="node-label">NODE 02 · MANTIK</span><h3>17 yıllık bankacı ve saha finans mantığını çalıştırırız.</h3><p>Vade yaşlandırma, kredi maliyeti, nakit açığı, kur-faiz etkisi, kâr marjı, limit kullanımı ve borç servis kapasitesi aynı ticari kurallara bağlanır.</p><div class="node-tags"><span>Vade</span><span>Kredi Maliyeti</span><span>Nakit Açığı</span><span>DSCR</span><span>Marj</span><span>Risk</span></div></article>
          <article class="node-card"><span class="node-label">NODE 03 · KARAR</span><h3>Formül değil, yönetici kokpiti üretiriz.</h3><p>“Kime ne zaman ödeyeceğim?”, “Hangi ürün gerçekten kârlı?”, “Hangi bankada limit sıkışıyor?” ve “Nakit açığı ne zaman başlıyor?” soruları tek ekranda cevaplanır.</p><div class="node-tags"><span>Patron</span><span>CFO</span><span>Finans</span><span>Muhasebe</span><span>Aksiyon</span></div></article>
        </div>
      </div>
    </section>`);

replaceSection('alanlar', `<section class="section section-green" id="alanlar">
      <div class="wrap-wide">
        <div class="section-head"><span class="section-kicker">6 STRATEJİK FİNANS SİSTEMİ</span><h2>Genel Excel değil, adı ve kararı belli finans sistemleri.</h2><p>Arama, raporlama ve yönetim açısından ne kurduğumuz nettir: nakit, banka riski, fiyatlama, bütçe, değerleme ve CFO kokpiti.</p></div>
        <div class="area-grid">
          <article class="area-card"><div class="area-icon">13W</div><h3>13 Haftalık Dinamik Nakit Akışı Modeli <span class="area-en">13-Week Cash Flow</span></h3><p>Haftalık nakit giriş-çıkış projeksiyonu, kasa açığı erken uyarısı, senaryolu tahsilat ve tedarikçi ödeme optimizasyonu.</p><p class="area-benefit"><strong>Karar:</strong> Maaş, çek, vergi ve kritik ödemelerin hangi haftada baskı yaratacağını önceden görün.</p></article>
          <article class="area-card"><div class="area-icon">₺</div><h3>Çoklu Banka Limit-Risk, Kredi ve Teminat Matrisi</h3><p>Nakdi/gayrinakdi limit kullanımı, rotatif kredi faiz yükü, spot vadeler, çek teminat havuzu ve bloke riskleri.</p><p class="area-benefit"><strong>Karar:</strong> Hangi bankada kullanılabilir alan kaldığını, hangi limitin pahalı veya sıkışık olduğunu tek tabloda görün.</p></article>
          <article class="area-card"><div class="area-icon">%</div><h3>Birim Maliyet, Dinamik Fiyatlama ve Başabaş Simülatörü</h3><p>Hammadde, navlun, işçilik ve finansman değişimlerini fiyat teklifine bağlayan ürün/müşteri bazlı katkı payı ve break-even hesabı.</p><p class="area-benefit"><strong>Karar:</strong> Fiyat vermeden önce gerçek marjı ve başabaş seviyesini görün; gizli zararı satıştan önce yakalayın.</p></article>
          <article class="area-card"><div class="area-icon">Σ</div><h3>Senaryolu Bütçe ve Dinamik Projeksiyon <span class="area-en">Rolling Forecast</span></h3><p>Kur, faiz ve enflasyon değişimlerine duyarlı iyimser / baz / kötü senaryo bütçesi ve fiili sapma analizi.</p><p class="area-benefit"><strong>Karar:</strong> Yıl başında donan bütçe yerine her ay güncellenen, sapmayı ve yeni ihtiyacı gösteren yönetim görünümü kullanın.</p></article>
          <article class="area-card"><div class="area-icon">DCF</div><h3>Yatırım Fizibilitesi ve DCF Şirket Değerleme Modeli</h3><p>İskontolanmış nakit akımları (DCF), IRR, geri ödeme süresi, EBITDA normalizasyonu ve DSCR borç servis kapasitesi.</p><p class="area-benefit"><strong>Karar:</strong> Makine, şube, yatırım veya şirket alımında getiriyi, borç taşıma gücünü ve değer aralığını aynı modelde test edin.</p></article>
          <article class="area-card"><div class="area-icon">KPI</div><h3>ERP Entegre CFO Yönetim Kokpiti <span class="area-en">Executive Dashboard</span></h3><p>Logo, SAP, Mikro veya Zirve verisinden beslenen net işletme sermayesi, DSO, nakit, kredi ve yönetim KPI paneli.</p><p class="area-benefit"><strong>Karar:</strong> Günler süren manuel rapor birleştirmeyi azaltıp kritik finans göstergelerini tek yönetici ekranına taşıyın.</p></article>
        </div>
      </div>
    </section>`);

replaceSection('surec', `<section class="section" id="surec">
      <div class="wrap-wide">
        <div class="section-head"><span class="section-kicker">SIFIR ZAMAN KAYBI</span><h2>Süreç nasıl işler?</h2><p>Teknik brifing hazırlamanız gerekmez. İşin nerede tıkandığını anlatmanız, mevcut ekranı göstermeniz veya maskelenmiş dosyayı paylaşmanız başlangıç için yeterlidir.</p></div>
        <div class="process-grid">
          <article class="process-card"><div class="process-step"><span class="step-no">1</span><strong>WhatsApp'tan anlatın</strong></div><h3>Sorunu veya mevcut dosyayı gösterin.</h3><p>Karışık tabloyu, ekran görüntüsünü ya da yalnızca “nakit yetmiyor”, “limit dağınık”, “kârı göremiyorum” gibi tıkanıklığı tarif edin.</p><div class="process-art"><span>⌕</span></div></article>
          <article class="process-card"><div class="process-step"><span class="step-no">2</span><strong>15 dakikalık ön teşhis</strong></div><h3>Gerçek ihtiyacı ve karar mantığını ayırırız.</h3><p>İlk görüşmede semptomu, veri kaynağını, finansal etkiyi ve hangi yönetim kararının üretilmesi gerektiğini netleştiririz.</p><div class="process-art"><span>✓</span></div></article>
          <article class="process-card"><div class="process-step"><span class="step-no">3</span><strong>Kurulum &amp; stres testi</strong></div><h3>Sistemi gerçek ve sınır senaryolarıyla test ederiz.</h3><p>Boş veri, yanlış seçim, yüksek kayıt sayısı, dönem geçişi, sıra dışı tutar ve hata koşullarını çalıştırıp sonuçların tutarlılığını kontrol ederiz.</p><div class="process-art"><span>▥</span></div></article>
          <article class="process-card"><div class="process-step"><span class="step-no">4</span><strong>Teslim &amp; kullanım</strong></div><h3>Ekibinizin hangi ekranda ne yapacağı nettir.</h3><p>Veri giriş alanları, kontroller ve yönetim raporları ayrıştırılır; günlük kullanım için dosyanın içinde anlaşılır yönlendirme bırakılır.</p><div class="process-art"><span>↗</span></div></article>
        </div>
      </div>
    </section>`);

replaceSection('karsilastirma', `<section class="section section-soft" id="karsilastirma">
      <div class="wrap-wide">
        <div class="section-head"><span class="section-kicker">YAKLAŞIM KARŞILAŞTIRMASI</span><h2>Aynı bütçe kalemi değil: yaklaşım ve karar çıktısı farklı.</h2><p>Kurumsal yazılım, hazır şablon ve işletmeye özel Excel sistemi farklı ihtiyaçlara hizmet eder. Buradaki farkı ölçülebilir kriterlerle görün.</p></div>
        <div class="comparison-wrap">
          <div class="comparison" role="table" aria-label="Çözüm yaklaşımları karşılaştırması">
            <div class="comparison-head" role="row"><div role="columnheader">Kriter</div><div role="columnheader">Kurumsal yazılım / danışmanlık</div><div role="columnheader">Hazır Excel şablonu</div><div role="columnheader">Barış Bağırlar / ExcelArşiv</div></div>
            <div class="comparison-row" role="row"><div role="rowheader">Başlangıç noktası</div><div role="cell" data-label="Kurumsal yazılım / danışmanlık">Geniş süreç standardizasyonu, entegrasyon ve proje kapsamıyla başlar.</div><div role="cell" data-label="Hazır Excel şablonu">Önceden tanımlanmış dosya yapısına uyum bekler.</div><div role="cell" data-label="Barış Bağırlar / ExcelArşiv">Önce işletmenin nakit, kredi, tahsilat, fiyatlama veya raporlama tıkanıklığı teşhis edilir.</div></div>
            <div class="comparison-row" role="row"><div role="rowheader">Finans &amp; saha bakışı</div><div role="cell" data-label="Kurumsal yazılım / danışmanlık">Proje ekibinin finansal ve sektörel uzmanlığına göre değişir.</div><div role="cell" data-label="Hazır Excel şablonu">Standart kullanım senaryosu ve önceden yazılmış mantık sunar.</div><div role="cell" data-label="Barış Bağırlar / ExcelArşiv"><strong>17 yıllık ticari bankacılık</strong> birikimi, saha-finans ve yönetim ihtiyacıyla birlikte modele girer.</div></div>
            <div class="comparison-row" role="row"><div role="rowheader">Kapsam &amp; maliyet</div><div role="cell" data-label="Kurumsal yazılım / danışmanlık">Lisans, entegrasyon ve proje yönetimi modele göre ek maliyetler yaratabilir.</div><div role="cell" data-label="Hazır Excel şablonu">Düşük başlangıç maliyeti; uyarlama sorumluluğu çoğunlukla kullanıcıdadır.</div><div role="cell" data-label="Barış Bağırlar / ExcelArşiv">İhtiyaca göre tek seferlik proje kapsamı oluşturulur; kullanılmayacak modül projeye eklenmez.</div></div>
            <div class="comparison-row" role="row"><div role="rowheader">Veri kontrolü</div><div role="cell" data-label="Kurumsal yazılım / danışmanlık">Bulut, sunucu veya entegrasyon mimarisine göre değişir.</div><div role="cell" data-label="Hazır Excel şablonu">Dosyanın tasarımına ve kullanıcının disiplinine bağlıdır.</div><div role="cell" data-label="Barış Bağırlar / ExcelArşiv">Hassas veri maskelenebilir; proje gerektirmiyorsa veriyi harici sisteme taşıma zorunluluğu oluşturulmaz.</div></div>
            <div class="comparison-row" role="row"><div role="rowheader">Karar çıktısı</div><div role="cell" data-label="Kurumsal yazılım / danışmanlık">Geniş modül ve raporlama ekosistemi sağlar.</div><div role="cell" data-label="Hazır Excel şablonu">Dosyada önceden tanımlanmış raporları sunar.</div><div role="cell" data-label="Barış Bağırlar / ExcelArşiv">Patron / CFO için “ne oldu?”dan çok “şimdi ne yapmalıyız?” sorusuna odaklanan karar ekranı kurulur.</div></div>
          </div>
        </div>
      </div>
    </section>`);

replaceSection('neden', `<section class="section" id="neden">
      <div class="wrap-wide">
        <div class="section-head"><span class="section-kicker">NEDEN BARIŞ BAĞIRLAR?</span><h2>Excel yalnız araç. Asıl değer, işletmenin finansal dilini doğru okumakta.</h2><p>17 yıllık ticari bankacılık geçmişini; işletmenin günlük nakit, kredi, tahsilat, maliyet ve kârlılık gerçekliğiyle birleştirip ölçülebilir karar ekranına çeviriyoruz.</p></div>
        <div class="authority-summary" aria-label="Uzmanlık özeti">
          <div class="authority-stat"><strong>17 yıl</strong><span>Ticari bankacılık, kredi, limit ve risk bakışı</span></div>
          <div class="authority-stat"><strong>Banka ↔ Saha</strong><span>Finans, muhasebe ve işletme dilini aynı modelde buluşturma</span></div>
          <div class="authority-stat"><strong>ERP → Karar</strong><span>Dağınık veriyi patron ve CFO için aksiyon ekranına dönüştürme</span></div>
        </div>
        <div class="why-grid">
          <article class="why-card"><div class="why-icon">₺</div><h3>Bankanın baktığı yerden bakarız.</h3><p>Limit, kredi maliyeti, vade, borç servis kapasitesi ve likiditeyi yalnız hücre olarak değil, aynı finansal tablonun parçaları olarak okuruz.</p></article>
          <article class="why-card"><div class="why-icon">◎</div><h3>Esnafın ve CFO'nun cümlesini ayırırız.</h3><p>“Kasa yetmiyor”, “tahsilat gecikiyor”, “limit dolu”, “satış var ama kâr yok” gibi ifadelerin arkasındaki gerçek finansal nedeni modele çeviririz.</p></article>
          <article class="why-card"><div class="why-icon">▦</div><h3>ERP verisini rapora değil karara çeviririz.</h3><p>Logo, SAP, Mikro, Zirve veya mevcut Excel verisini yalnız görselleştirmek yerine kontrol, risk ve yönetim kararına bağlarız.</p></article>
          <article class="why-card"><div class="why-icon">◇</div><h3>Kapsamı şişirmeyiz.</h3><p>Hazır sistem yeterliyse hazır sistemi, küçük düzeltme yeterliyse küçük düzeltmeyi öneririz; özel geliştirmeyi yalnız ihtiyaç gerçekten gerektiriyorsa kurarız.</p></article>
        </div>
      </div>
    </section>`);

replaceOne(/<a class="btn btn-primary" href="\/iletisim">Mevcut dosyamı anlatmak istiyorum<\/a>/u, `<a class="btn btn-primary btn-whatsapp" href="${WA_URL}" target="_blank" rel="noopener noreferrer">WhatsApp'tan dosyamı anlat</a>`, 'delivery CTA');

replaceOne(/<section class="cta">[\s\S]*?<\/section>/u, `<section class="cta">
      <div class="wrap cta-shell">
        <div><h2>“Bizim sorun çok karışık, Excel çözer mi?” diye düşünmeyin.</h2><p>Nakit, stok, kredi, tahsilat veya kârlılık düğümünü anlatın. Excel doğru çözümse sistemi kuralım; değilse sizi daha doğru araca yönlendirelim.</p><div class="cta-points"><span>Nakit &amp; likidite</span><span>Kredi &amp; limit</span><span>Kârlılık &amp; fiyatlama</span><span>ERP verisi</span></div></div>
        <div class="cta-actions"><a class="btn btn-whatsapp" href="${WA_URL}" target="_blank" rel="noopener noreferrer">WhatsApp Üzerinden Görüşelim <span aria-hidden="true">→</span></a><a class="btn btn-secondary" href="mailto:baris@excelarsiv.com">E-posta ile dosya ilet</a><p class="cta-note">Dosya göndermeden de başlayabilirsiniz; önce tıkanıklığı yazmanız yeterli.</p><a class="cta-email" href="mailto:baris@excelarsiv.com">baris@excelarsiv.com</a></div>
      </div>
    </section>`, 'closing CTA');

html = html.replaceAll('href="#ne-zaman">Ne zaman gerekir?</a>', 'href="#ne-zaman">Sistem nasıl kurulur?</a>');
html = html.replaceAll('href="#neden">Neden ExcelArşiv?</a>', 'href="#neden">Neden Barış Bağırlar?</a>');

const required = [
  'data-finance-funnel-v7',
  'Yazılımcıya finans öğretmekle uğraşmayın.',
  '17 yıllık ticari bankacılık',
  'NODE 01 · GİRDİ',
  'NODE 02 · MANTIK',
  'NODE 03 · KARAR',
  '13 Haftalık Dinamik Nakit Akışı Modeli',
  'Çoklu Banka Limit-Risk, Kredi ve Teminat Matrisi',
  'Senaryolu Bütçe ve Dinamik Projeksiyon',
  'Yatırım Fizibilitesi ve DCF Şirket Değerleme Modeli',
  'ERP Entegre CFO Yönetim Kokpiti',
  'Neden Barış Bağırlar?',
  'https://wa.me/905419305372',
  'baris@excelarsiv.com',
  '"@type":"Person"',
  '"@type":"Service"',
  'finance-funnel-v7-css',
];
for (const token of required) {
  if (!html.includes(token)) throw new Error(`SPECIAL FINANCE FUNNEL V7: required token missing: ${token}`);
}

const forbidden = [
  'Hazır dosya işinize uymuyorsa,',
  'Özel bir sisteme ne zaman ihtiyaç olur?',
  'Neden ExcelArşiv ile çalışmalısınız?',
  'href="/hakkimizda"',
  'color-scheme:dark',
  'overflow-x:hidden',
];
for (const token of forbidden) {
  if (html.includes(token)) throw new Error(`SPECIAL FINANCE FUNNEL V7: forbidden legacy token present: ${token}`);
}

fs.writeFileSync(distFile, html, 'utf8');
console.log('SPECIAL FINANCE FUNNEL V7 PASS — finance-led hero, 3-node decision flow, six semantic finance systems, authority positioning, WhatsApp conversion path and Person/Service schema applied.');
