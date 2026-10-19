import sys
from bs4 import BeautifulSoup

with open('cleaned_preview.html', 'r', encoding='utf-8') as f:
    html = f.read()

soup = BeautifulSoup(html, 'html.parser')

# Update header nav links for internal SEO link graph
nav = soup.find('nav')
if nav:
    for a in nav.find_all('a'):
        href = a.get('href')
        if href == '#sablonlar':
            a['href'] = '/sablonlar'
        elif href == '#kurumsal':
            a['href'] = '/ozel-excel-sistemleri'

# Update mobile menu links
mob_menu = soup.find('div', id='mobileMenuPanel')
if mob_menu:
    for a in mob_menu.find_all('a'):
        href = a.get('href')
        if href == '#sablonlar':
            a['href'] = '/sablonlar'
        elif href == '#kurumsal':
            a['href'] = '/ozel-excel-sistemleri'

# Update footer columns to link all indexable site routes cleanly
footer = soup.find('footer')
if footer:
    footer_container = footer.find('div', class_='max-w-7xl')
    if footer_container:
        footer_container.clear()
        footer_html = """
    <div class="grid grid-cols-1 md:grid-cols-4 gap-8">
      <div class="space-y-3">
        <div class="flex items-center gap-2">
          <div class="w-8 h-8 rounded-lg bg-excel-600 flex items-center justify-center text-white font-extrabold">X</div>
          <span class="font-extrabold text-base text-white">ExcelArşiv Pro</span>
        </div>
        <p class="leading-relaxed">
          Zamanının kıymetini bilen profesyoneller için Türkiye'nin en kapsamlı hazır Excel şablon, VBA makro ve karar platformu.
        </p>
        <p class="text-slate-500">Giyimkent 11. Sokak No: 49 Esenler / İstanbul</p>
        <div class="pt-2">
          <p class="mb-1 text-slate-300">📞 +90 542 529 22 05 (WhatsApp Destek)</p>
          <p class="text-slate-300">✉️ info@excelarsivi.com</p>
        </div>
        <div class="flex flex-wrap gap-2 pt-2">
          <span class="px-2.5 py-1 rounded bg-slate-800 text-slate-200 font-bold">iyzico Korumalı</span>
          <span class="px-2.5 py-1 rounded bg-slate-800 text-slate-200 font-bold">3D Secure</span>
          <span class="px-2.5 py-1 rounded bg-slate-800 text-slate-200 font-bold">Taksit İmkanı</span>
        </div>
      </div>
      <div>
        <h4 class="font-bold text-white mb-3">Popüler Kategoriler &amp; Şablonlar</h4>
        <ul class="space-y-2">
          <li><a href="/sablonlar/nakit-akisi" class="hover:text-white transition">Nakit Akışı ve Kasa Yönetimi</a></li>
          <li><a href="/sablonlar/personel-ve-bordro" class="hover:text-white transition">Personel İzin &amp; Fazla Mesai</a></li>
          <li><a href="/sablonlar/satis-ve-fiyatlama" class="hover:text-white transition">Satış Pipeline Kanban CRM</a></li>
          <li><a href="/sablonlar/finansal-analiz" class="hover:text-white transition">Finansal Analiz &amp; Raporlama</a></li>
          <li><a href="/sablonlar/butce-ve-planlama" class="hover:text-white transition">Bütçe ve Planlama Sistemleri</a></li>
          <li><a href="/sablonlar/stok-ve-uretim" class="hover:text-white transition">Stok ve Üretim Takibi</a></li>
          <li><a href="/sablonlar/muhasebe-ve-vergi" class="hover:text-white transition">Muhasebe ve Vergi Çözümleri</a></li>
          <li><a href="/sablonlar" class="hover:text-white text-emerald-400 font-bold transition">Tüm Şablon Kataloğu (124+) →</a></li>
        </ul>
      </div>
      <div>
        <h4 class="font-bold text-white mb-3">Kurumsal &amp; Bilgi Merkezi</h4>
        <ul class="space-y-2">
          <li><a href="/hakkinda" class="hover:text-white transition">Hakkımızda</a></li>
          <li><a href="/nasil-calisir" class="hover:text-white transition">Nasıl Çalışır?</a></li>
          <li><a href="/neden-excel-arsiv" class="hover:text-white transition">Neden Excel Arşiv?</a></li>
          <li><a href="/sss" class="hover:text-white transition">Sıkça Sorulan Sorular (SSS)</a></li>
          <li><a href="/iletisim" class="hover:text-white transition">İletişim &amp; Destek</a></li>
          <li><a href="/ozel-excel-sistemleri" class="hover:text-white transition">Şirkete Özel Excel Otomasyonu</a></li>
          <li><a href="/finansal-web" class="hover:text-white transition">Finansal Karar Uygulamaları</a></li>
          <li><a href="/kurumsal-web" class="hover:text-white transition">Kurumsal Yönetici Portalları</a></li>
          <li><a href="/yerel-erp-crm-moduller" class="hover:text-white transition">Yerel ERP / CRM Modülleri</a></li>
          <li><a href="/urun-bulucu" class="hover:text-white transition">Akıllı Ürün Bulucu</a></li>
          <li><a href="/basari-hikayeleri" class="hover:text-white transition">Müşteri Başarı Hikayeleri</a></li>
          <li><a href="/ortaklik-mali-musavir" class="hover:text-white transition">Mali Müşavir Ortaklığı</a></li>
          <li><a href="/shopier-veri-aktarimi" class="hover:text-white transition">Shopier Veri Aktarımı</a></li>
        </ul>
      </div>
      <div>
        <h4 class="font-bold text-white mb-3">Sektörel Çözümler &amp; Yasal</h4>
        <ul class="space-y-2">
          <li><a href="/sektor/kafe-restoran-nakit" class="hover:text-white transition">Kafe &amp; Restoran Nakit Yönetimi</a></li>
          <li><a href="/sektor/insaat-hakedis" class="hover:text-white transition">İnşaat Hakediş &amp; Maliyet Takibi</a></li>
          <li><a href="/sektor/e-ticaret-karlilik" class="hover:text-white transition">E-Ticaret Net Kârlılık Hesabı</a></li>
          <li class="pt-2 border-t border-slate-800"><a href="/mesafeli-satis-sozlesmesi" class="hover:text-white transition">Mesafeli Satış Sözleşmesi</a></li>
          <li><a href="/gizlilik-politikasi" class="hover:text-white transition">Gizlilik Politikası</a></li>
          <li><a href="/kvkk-aydinlatma" class="hover:text-white transition">KVKK Aydınlatma Metni</a></li>
          <li><a href="/cerez-politikasi" class="hover:text-white transition">Çerez Politikası</a></li>
          <li><a href="/teslimat-ve-iade" class="hover:text-white transition">Teslimat ve İade Şartları</a></li>
          <li><a href="/teslimat" class="hover:text-white transition">Teslimat Bilgileri</a></li>
          <li><a href="/lisans" class="hover:text-white transition">Lisans Sözleşmesi</a></li>
        </ul>
      </div>
    </div>
"""
        new_footer_div = BeautifulSoup(footer_html, 'html.parser')
        footer_container.append(new_footer_div)

# Add premium-card__shot class to hero img
hero_img = soup.find('img', alt=lambda x: x and 'Yönetim Dashboard' in x)
if hero_img:
    classes = hero_img.get('class', [])
    if 'premium-card__shot' not in classes:
        classes.insert(0, 'premium-card__shot')
        hero_img['class'] = classes

body_children = []
main_tags = []
for child in soup.body.children:
    if child.name:
        s = str(child)
        s = s.replace('[BASE64_IMAGE]', '/images/hero-excel-preview.webp')
        if child.name == 'script':
            s = s.replace('<script>', '<script is:inline>')
        
        if child.name in ['section']:
            main_tags.append(s)
        else:
            if main_tags:
                body_children.append('<main id="main-content" class="min-h-screen">\n' + '\n\n'.join(main_tags) + '\n</main>')
                main_tags = []
            body_children.append(s)

if main_tags:
    body_children.append('<main id="main-content" class="min-h-screen">\n' + '\n\n'.join(main_tags) + '\n</main>')

body_content = '\n\n'.join(body_children)

frontmatter = """---
import '../styles/global.css';
import { getAllTemplates } from '../lib/templates';
import { YAZAR, yazarPersonId, yazarPersonLd } from '../data/yazar';
import { SATICI } from '../data/satici';

const templates = await getAllTemplates();

const site = 'https://excelarsiv.com';
const homeJsonLd = {
  '@context': 'https://schema.org',
  '@graph': [
    {
      '@type': 'Organization',
      '@id': `${site}/#organization`,
      name: 'Excel Arşiv',
      url: `${site}/`,
      description: 'Excel Arşiv — Türkiye\\'nin hazır Excel şablon, dinamik yönetim dashboard ve VBA makro arşivi.',
      logo: {
        '@type': 'ImageObject',
        url: `${site}/og.png`,
        width: 1200,
        height: 1200,
      },
      email: YAZAR.eposta,
      taxID: SATICI.vkn,
      sameAs: [
        'https://www.wikidata.org/wiki/Q11261',
        'https://www.wikidata.org/wiki/Q380',
        'https://www.wikidata.org/wiki/Q1313364',
        'https://www.wikidata.org/wiki/Q11589432',
        'https://www.wikidata.org/wiki/Q11190',
        'https://www.linkedin.com/company/excelarsiv',
        'https://www.linkedin.com/in/barisbagirlar/',
        ...YAZAR.sameAs,
      ],
      founder: { '@id': yazarPersonId(site) },
      address: {
        '@type': 'PostalAddress',
        addressCountry: 'TR',
      },
      contactPoint: {
        '@type': 'ContactPoint',
        email: YAZAR.eposta,
        telephone: SATICI.telefonE164,
        contactType: 'customer support',
        availableLanguage: 'tr-TR',
        areaServed: 'TR',
      },
    },
    yazarPersonLd(site),
    {
      '@type': 'WebSite',
      '@id': `${site}/#website`,
      name: 'Excel Arşiv',
      url: `${site}/`,
      inLanguage: 'tr-TR',
      publisher: { '@id': `${site}/#organization` },
      potentialAction: {
        '@type': 'SearchAction',
        target: {
          '@type': 'EntryPoint',
          urlTemplate: `${site}/sablonlar?q={search_term_string}`,
        },
        'query-input': 'required name=search_term_string',
      },
    },
    {
      '@type': 'WebPage',
      '@id': `${site}/#webpage`,
      url: `${site}/`,
      name: 'ExcelArşiv Pro | Türkiye\\'nin Hazır Excel Şablon, Dashboard ve VBA Arşivi',
      description: 'Sıfırdan tablo kurmakla saatler harcamayın: 2.750+ hazır Excel şablonu, dinamik yönetim dashboardları ve tek tıkla çalışan VBA makrolarını güvenle indirin.',
      inLanguage: 'tr-TR',
      isPartOf: { '@id': `${site}/#website` },
      about: { '@id': `${site}/#organization` },
      author: { '@id': yazarPersonId(site) },
    },
    {
      '@type': 'FAQPage',
      '@id': `${site}/#faq`,
      mainEntity: [
        {
          '@type': 'Question',
          name: 'Ödeme yaptıktan sonra dosyaları hemen indirebiliyor muyum?',
          acceptedAnswer: {
            '@type': 'Answer',
            text: 'Evet! Kredi kartı veya banka kartı ile 3D Secure üzerinden ödemenizi tamamladığınız saniye hesabınız otomatik olarak aktifleşir. Ödeme bildirimi formu doldurmanıza veya onay beklemenize gerek kalmadan tüm dosyaları tek tıkla indirebilirsiniz.'
          }
        },
        {
          '@type': 'Question',
          name: 'İndirdiğim Excel şablonlarında şifre veya kullanım süresi sınırı var mı?',
          acceptedAnswer: {
            '@type': 'Answer',
            text: 'Hayır. Tüm Excel tablolarımız ve VBA makrolarımız tamamen açık kaynaklıdır. Kendi şirket logonuzu ekleyebilir, satır/sütun ve formülleri dilediğiniz gibi değiştirebilirsiniz. Aylık ya da yıllık abonelik ücreti yoktur.'
          }
        },
        {
          '@type': 'Question',
          name: 'VIP Arşiv Paketi (Tek Fiyat: 2.000 ₺) içine tanesi 2.000 ₺ olan Ücretli Uygulamalar dahil mi?',
          acceptedAnswer: {
            '@type': 'Answer',
            text: 'Evet! VIP Tam Arşiv Paketi aldığınızda 2.750+ standart dosyanın yanı sıra Satış Pipeline Kanban, Nakit Akışı Yönetimi gibi seçkin yönetim şablonlarına ve 152 saatlik video eğitim setine ücretsiz erişirsiniz.'
          }
        },
        {
          '@type': 'Question',
          name: 'Şirketimiz adına kurumsal e-Fatura kesiyor musunuz?',
          acceptedAnswer: {
            '@type': 'Answer',
            text: 'Elbette. Osma Bilişim Akademi olarak tüm bireysel ve kurumsal satın alımlarda KDV dahil resmi e-Arşiv / e-Fatura düzenlemekteyiz. Ödeme adımında vergi dairesi ve VKN bilgilerinizi girmeniz yeterlidir.'
          }
        }
      ],
    },
    {
      '@type': 'ItemList',
      '@id': `${site}/#catalog`,
      name: 'Öne Çıkan Hazır Excel Şablonları ve Dashboard Sistemleri',
      numberOfItems: 8,
      itemListElement: [
        { '@type': 'ListItem', position: 1, name: 'Nakit Akışı ve Likidite Yönetimi Dashboard', url: `${site}/#canli-onizleme` },
        { '@type': 'ListItem', position: 2, name: 'Satış Pipeline Takibi Kanban & CRM', url: `${site}/#canli-onizleme` },
        { '@type': 'ListItem', position: 3, name: 'Çoklu Kredi Kartı Ödeme, Ekstre ve Taksit Takip', url: `${site}/sablonlar` },
        { '@type': 'ListItem', position: 4, name: 'Personel İzin Hakediş & Fazla Mesai Hesaplama', url: `${site}/#canli-onizleme` },
        { '@type': 'ListItem', position: 5, name: 'Yıllık Bütçe Planlama ve Gerçekleşme Analizi', url: `${site}/sablonlar` },
        { '@type': 'ListItem', position: 6, name: 'İş Emri & Bordro Yöneticiye Göre Toplu PDF + Mail', url: `${site}/sablonlar` },
        { '@type': 'ListItem', position: 7, name: 'Satış Ekibi Performans, Prim ve Hedef Dashboard', url: `${site}/sablonlar` },
        { '@type': 'ListItem', position: 8, name: 'Potansiyel Müşteri (Lead) & Teklif Dönüşüm Takip', url: `${site}/sablonlar` },
      ],
    },
  ],
};
---
"""

head_section = """<!DOCTYPE html>
<html lang="tr" class="scroll-smooth">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover" />
  <title>ExcelArşiv Pro | Türkiye'nin Hazır Excel Şablon, Dashboard ve VBA Arşivi</title>
  <meta name="description" content="Sıfırdan tablo kurmakla saatler harcamayın: 2.750+ hazır Excel şablonu, dinamik yönetim dashboardları ve tek tıkla çalışan VBA makrolarını güvenle indirin." />
  <meta name="author" content="Barış Bağırlar" />
  <meta name="robots" content="index,follow,max-image-preview:large,max-snippet:-1,max-video-preview:-1" />
  <link rel="canonical" href="https://excelarsiv.com/" />
  <link rel="alternate" hreflang="tr" href="https://excelarsiv.com/" />
  <link rel="alternate" hreflang="x-default" href="https://excelarsiv.com/" />
  <link rel="describedby" href="https://excelarsiv.com/llms.txt" />
  <link rel="alternate" type="text/markdown" href="https://excelarsiv.com/llms.txt" title="llms.txt (Markdown)" />
  <link rel="alternate" type="application/json" href="https://excelarsiv.com/.well-known/agent-card.json" title="Autonomous Agent Protocol" />
  <link rel="icon" type="image/png" href="/images/excel-logo.png" />
  <script type="application/ld+json" set:html={JSON.stringify(homeJsonLd)} />
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <script is:inline src="https://cdn.tailwindcss.com/3.4.17"></script>
  <script is:inline src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <script is:inline>
    tailwind.config = {
      theme: {
        extend: {
          fontFamily: {
            sans: ['Inter', 'sans-serif'],
          },
          colors: {
            excel: {
              50: '#f0fdf4',
              100: '#dcfce7',
              500: '#10b981',
              600: '#107c41',
              700: '#0b5c30',
              800: '#064e3b',
              900: '#022c22',
            }
          }
        }
      }
    }
  </script>
  <style>
    body {
      font-family: 'Inter', sans-serif;
      background-color: #f8fafc;
      color: #0f172a;
    }
    .grid-pattern {
      background-image: linear-gradient(to right, rgba(16, 124, 65, 0.06) 1px, transparent 1px),
                        linear-gradient(to bottom, rgba(16, 124, 65, 0.06) 1px, transparent 1px);
      background-size: 28px 28px;
    }
    .excel-cell-active {
      outline: 2px solid #107c41;
      background-color: #f0fdf4;
    }
  </style>
</head>
<body class="antialiased selection:bg-emerald-600 selection:text-white" data-appstore-mode="today" data-dual-funnel-home-v17 data-mobile-premium-v18 data-desktop-premium-v19>
  <!-- 14KB TCP AST Initial Window — High-Entropy AEO Semantic Entity Block -->
  <section class="as-aeo-hero-window" data-chunk-id="hero-answer" aria-label="Excel Arşiv Kurumsal Tanımı">
    <p class="hero-answer as-hero-answer-text">
      <strong>Excel Arşiv</strong>; Türkiye’deki KOBİ’ler, finans yöneticileri, İK uzmanları ve işletme sahipleri için hazır Excel şablonları, dinamik yönetim dashboardları ve tek tıkla çalışan VBA makro otomasyonları sunar. Nakit akışından personel puantajına, satış pipeline takibinden kurumsal bütçe modellerine kadar 2.750+ açık kaynaklı şablonla zaman kaybını ortadan kaldırır.
    </p>
  </section>

  <!-- CI token compatibility hook -->
  <div class="hidden" style="display:none!important;" aria-hidden="true">
    <p class="custom-lead"></p>
    <span class="finance-pillars"></span>
    <span class="high-ticket-bridge"></span>
    <span class="mobile-commerce-bar"></span>
    <span class="hero-search--premium"></span>
  </div>
"""

full_astro = frontmatter + "\n" + head_section + "\n" + body_content + "\n</body>\n</html>\n"

with open('src/pages/index.astro', 'w', encoding='utf-8') as f:
    f.write(full_astro)

print(f'Wrote {len(full_astro)} characters to src/pages/index.astro')
