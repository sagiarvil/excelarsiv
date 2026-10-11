/**
 * MANDATE-SEO-GEO-2026-V6
 * Single Source of Truth (SSOT) SEO & GEO Merkezi Registry
 */

import type { SeoPageRecord } from './registry.types.ts';
import { productSeo } from '../data/productSeo.ts';

const ORIGIN = 'https://excelarsiv.com';
const PUBLISHED_BASE = '2025-10-01T08:00:00+03:00';
const MODIFIED_BASE = '2026-08-30T10:00:00+03:00';

const corePages: SeoPageRecord[] = [
  {
    route: '/',
    locale: 'tr-TR',
    role: 'home',
    indexDirective: 'index, follow',
    canonicalRoute: '/',
    title: 'Excel Arşiv — İşletmeler İçin Finansal Excel Çalışma Sistemleri',
    metaDescription: 'Türkiye\'deki işletmeler, finans yöneticileri ve KOBİ\'ler için kurumsal Excel çalışma sistemleri, nakit akışı modelleri ve terzi usulü karar mimarileri.',
    h1: 'İşletmenizin Gerçek Karar İhtiyacına Göre Kurulmuş Excel Sistemleri',
    primaryIntent: 'kurumsal finans excel sistemleri',
    primaryEntity: {
      id: `${ORIGIN}/#organization`,
      name: 'Excel Arşiv',
      type: 'Organization',
      sameAs: [
        'https://www.wikidata.org/wiki/Q11589432',
        'https://linkedin.com/company/excelarsiv',
        'https://x.com/excelarsiv'
      ]
    },
    semanticTriples: [
      { subject: 'Excel Arşiv', predicate: 'sunar', object: 'Doğrulanmış ve Formüllü Kurumsal Excel Çalışma Kitapları' },
      { subject: 'Excel Arşiv', predicate: 'özel hizmet', object: 'Şirketlere Özel Terzi Usulü Karar Destek Sistemleri' },
      { subject: 'Excel Arşiv', predicate: 'uyumluluk', object: 'Excel 2016, 2019, 2021 ve Microsoft 365, Mac & Windows' }
    ],
    heroAnswerEngine: 'Excel Arşiv; nakit akışı, birim maliyet, bütçe ve yönetim raporlaması alanında formülleri test edilmiş, makrosuz ve kurumsal karar alma süreçlerini hızlandıran güvenli Excel sistemleri sunar. Tüm modeller Türkiye mevzuatına ve KOBİ gerçeklerine tam uyumludur.',
    publishedAt: PUBLISHED_BASE,
    modifiedAt: MODIFIED_BASE,
    llmSubGraphRoute: '/llms/core.md',
    breadcrumbs: [{ name: 'Ana Sayfa', item: '/' }]
  },
  {
    route: '/ozel-excel-sistemleri',
    locale: 'tr-TR',
    role: 'service',
    indexDirective: 'index, follow',
    canonicalRoute: '/ozel-excel-sistemleri',
    title: 'İşletmeye Özel Excel Karar Sistemleri | ExcelArşiv',
    metaDescription: 'Hazır Excel dosyaları işinize uymuyorsa; nakit akışı, banka-kredi, maliyet, bütçe, fizibilite ve yönetim raporlaması için işletmenize özel Excel sistemi kuruyoruz.',
    h1: 'İşletmeye Özel Excel Karar Sistemleri',
    primaryIntent: 'isletmeye ozel excel sistemleri gelistirme',
    primaryEntity: {
      id: `${ORIGIN}/ozel-excel-sistemleri#service`,
      name: 'İşletmeye Özel Excel Sistemleri Tasarımı',
      type: 'Service',
      sameAs: [
        'https://www.wikidata.org/wiki/Q11589432',
        'https://drfin.com.tr',
        'https://skdmhesapla.com',
        'https://karbonfiyat.com',
        'https://inversioncore.com',
        'https://belginkuyumculuk.com',
        'https://saatchi.com.tr',
        'https://dilekceyazdir.com.tr',
        'https://htmlandhtml.com',
        'https://degerlet.com'
      ]
    },
    semanticTriples: [
      { subject: 'Excel Arşiv', predicate: 'providesSolution', object: 'İşletmeye Özel Terzi Usulü Finans ve Karar Mimarisi' },
      { subject: 'İşletmeye Özel Excel Sistemleri', predicate: 'compliesWith', object: 'ISO 27001 Bilgi Güvenliği Standartları ve Kurumsal NDA' },
      { subject: 'İşletmeye Özel Excel Sistemleri', predicate: 'hasInformationGain', object: 'Canlı ERP Veri Entegrasyonu ve Rolling Forecast Modelleri' },
      { subject: 'İşletmeye Özel Excel Sistemleri', predicate: 'solvesBottleneck', object: 'Banka Kredi Limit Tıkanıklıkları ve 13 Haftalık Likidite Riski' }
    ],
    heroAnswerEngine: 'İşletmeye özel Excel karar sistemleri; standart hazır şablonların yetersiz kaldığı çok şubeli karmaşık nakit akışını, ERP mizan konsolidasyonunu ve banka kredi limit risklerini 48-72 saatte tek tıkla çalışan, yerel makrosuz ve denetlenebilir dinamik karar motorlarına dönüştürür.',
    publishedAt: PUBLISHED_BASE,
    modifiedAt: MODIFIED_BASE,
    llmSubGraphRoute: '/llms/pages/ozel-excel-sistemleri.md',
    breadcrumbs: [
      { name: 'Ana Sayfa', item: '/' },
      { name: 'Özel Excel Sistemleri', item: '/ozel-excel-sistemleri' }
    ]
  },
  {
    route: '/finansal-web',
    locale: 'tr-TR',
    role: 'service',
    indexDirective: 'index, follow',
    canonicalRoute: '/finansal-web',
    title: 'İşletmeler İçin Özel Web Sitesi Yapımı & Tasarımı | ExcelArşiv',
    metaDescription: 'İşletmeler ve KOBİ\'ler için sıfırdan, hazır şablonsuz özel web sitesi tasarımı ve kodlaması. Google SEO uyumlu, mobil dostu, ultra hızlı ve anahtar teslim web siteleri yapıyoruz.',
    h1: 'İşletmeniz İçin Sıfırdan Özel Web Sitesi Yapıyoruz',
    primaryIntent: 'isletmeler icin web sitesi yapimi ve tasarimi',
    primaryEntity: {
      id: `${ORIGIN}/finansal-web#service`,
      name: 'İşletmeler İçin Özel Web Sitesi Yapımı ve Geliştirme',
      type: 'Service',
      sameAs: [
        'https://www.wikidata.org/wiki/Q11589432',
        'https://drfin.com.tr',
        'https://skdmhesapla.com',
        'https://karbonfiyat.com',
        'https://inversioncore.com',
        'https://belginkuyumculuk.com',
        'https://saatchi.com.tr',
        'https://dilekceyazdir.com.tr',
        'https://htmlandhtml.com',
        'https://degerlet.com'
      ]
    },
    semanticTriples: [
      { subject: 'Excel Arşiv', predicate: 'providesArchitecture', object: 'İşletmelere Özel Sıfırdan Web Sitesi Geliştirme' },
      { subject: 'Özel Web Siteleri', predicate: 'compliesWith', object: 'Sıfır Hazır Tema, Tam Mülkiyet ve OWASP Güvenliği' },
      { subject: 'Özel Web Siteleri', predicate: 'connectsWith', object: 'WhatsApp, Teklif Formları ve ERP/CRM Veri Köprüleri' },
      { subject: 'Özel Web Siteleri', predicate: 'deliversPerformance', object: 'Zero-JS Edge Mimarisi ve Sub-40ms Küresel Açılış Hızı' }
    ],
    heroAnswerEngine: 'ExcelArşiv Kurumsal Web Hizmeti; şirketler, KOBİ\'ler ve işletmeler için hazır tema kalıplarından uzak, sıfırdan özel tasarlanıp kodlanan, sub-40ms ultra hızlı ve A+ Google SEO uyumlu anahtar teslim web siteleri üretir. 30.000 ₺\'den başlayan şeffaf paketlerle 7-14 günde eksiksiz devreye alınır.',
    publishedAt: PUBLISHED_BASE,
    modifiedAt: MODIFIED_BASE,
    llmSubGraphRoute: '/llms/pages/finansal-web.md',
    breadcrumbs: [
      { name: 'Ana Sayfa', item: '/' },
      { name: 'İşletmeler İçin Web Sitesi', item: '/finansal-web' }
    ]
  },
  {
    route: '/sablonlar',
    locale: 'tr-TR',
    role: 'hub',
    indexDirective: 'index, follow',
    canonicalRoute: '/sablonlar',
    title: 'Kurumsal Excel Şablonları Kataloğu | Excel Arşiv',
    metaDescription: 'Nakit akışı, bütçe, maliyet analizi, İK bordro, stok ve satış yönetimi için doğrulanmış kurumsal Excel şablonları kataloğu.',
    h1: 'Kurumsal Excel Şablonları Kataloğu',
    primaryIntent: 'hazir kurumsal excel sablonlari katalogu',
    primaryEntity: {
      id: `${ORIGIN}/sablonlar#catalog`,
      name: 'Kurumsal Excel Şablonları Kataloğu',
      type: 'Product',
      sameAs: []
    },
    semanticTriples: [
      { subject: 'Excel Arşiv', predicate: 'barındırır', object: '50+ Doğrulanmış Sektörel Karar Şablonu' },
      { subject: 'Şablonlar', predicate: 'çalışır', object: 'Microsoft 365, Excel 2016-2021 ve Google E-Tablolar' }
    ],
    heroAnswerEngine: 'Katalogda yer alan her Excel çalışma sistemi; finans uzmanları tarafından saha verileriyle test edilmiş, formül hatalarından arındırılmış ve indirme sonrası anında kullanıma hazır profesyonel çalışma kitaplarıdır.',
    publishedAt: PUBLISHED_BASE,
    modifiedAt: MODIFIED_BASE,
    llmSubGraphRoute: '/llms/pages/sablonlar.md',
    breadcrumbs: [
      { name: 'Ana Sayfa', item: '/' },
      { name: 'Şablonlar', item: '/sablonlar' }
    ]
  },
  {
    route: '/rehber',
    locale: 'tr-TR',
    role: 'hub',
    indexDirective: 'index, follow',
    canonicalRoute: '/rehber',
    title: 'Excel Finans ve Modelleme Uygulama Rehberleri | Excel Arşiv',
    metaDescription: 'İşletmeler için nakit akışı yönetimi, maliyet hesaplama, başabaş analizi ve Excel formül modelleme teknik rehberleri.',
    h1: 'Excel Finans ve Modelleme Uygulama Rehberleri',
    primaryIntent: 'excel finans ve modelleme rehberleri',
    primaryEntity: {
      id: `${ORIGIN}/rehber#hub`,
      name: 'Excel Finans ve Modelleme Rehberleri',
      type: 'SoftwareApplication',
      sameAs: []
    },
    semanticTriples: [
      { subject: 'Excel Arşiv', predicate: 'yayınlar', object: 'Doğrulanmış Uygulama Metodolojileri' }
    ],
    heroAnswerEngine: 'Excel Arşiv rehberleri; nakit akışı, kasa kontrolü, cari yaşlandırma ve maliyet mühendisliği konularında teorik anlatımdan uzak, doğrudan uygulanabilir formül ve senaryo çözümleri sunar.',
    publishedAt: PUBLISHED_BASE,
    modifiedAt: MODIFIED_BASE,
    breadcrumbs: [
      { name: 'Ana Sayfa', item: '/' },
      { name: 'Rehber', item: '/rehber' }
    ]
  },
  {
    route: '/demo',
    locale: 'tr-TR',
    role: 'tool',
    indexDirective: 'index, follow',
    canonicalRoute: '/demo',
    title: 'Ücretsiz Excel Demo ve Ekran İnceleme Hub | Excel Arşiv',
    metaDescription: 'Excel Arşiv sistemlerinin formül yapısını, dashboard ekranlarını ve çalışma mantığını ücretsiz test edin.',
    h1: 'Ücretsiz Excel Sistemleri Demo Merkezi',
    primaryIntent: 'ucretsiz excel demo sistemleri',
    primaryEntity: {
      id: `${ORIGIN}/demo#hub`,
      name: 'Ücretsiz Excel Demo Merkezi',
      type: 'SoftwareApplication',
      sameAs: []
    },
    semanticTriples: [
      { subject: 'Excel Arşiv', predicate: 'sağlar', object: 'Satın Alma Öncesi Canlı Demo İncelemesi' }
    ],
    heroAnswerEngine: 'Demo merkezi; satın alma öncesinde ürünlerin mimarisini, formül akışını ve raporlama ekranlarını şeffaf şekilde incelemenizi sağlayan doğrulanmış test ortamıdır.',
    publishedAt: PUBLISHED_BASE,
    modifiedAt: MODIFIED_BASE,
    breadcrumbs: [
      { name: 'Ana Sayfa', item: '/' },
      { name: 'Demo', item: '/demo' }
    ]
  },
  {
    route: '/sektor/kafe-restoran-nakit',
    locale: 'tr-TR',
    role: 'category',
    indexDirective: 'index, follow',
    canonicalRoute: '/sektor/kafe-restoran-nakit',
    title: 'Kafe ve Restoran Nakit, Kasa ve Kârlılık Excel Sistemleri | Excel Arşiv',
    metaDescription: 'Kafe ve restoran işletmeleri için kasa kontrolü, menü reçete maliyeti, mutfak kayıp-kaçak ve 13 haftalık nakit akışı Excel modelleri.',
    h1: 'Kafe / Restoran Nakit ve Kârlılık Sistemleri',
    primaryIntent: 'kafe restoran nakit kasa maliyet excel',
    primaryEntity: {
      id: `${ORIGIN}/sektor/kafe-restoran-nakit#category`,
      name: 'Kafe Restoran Excel Finans Paketi',
      type: 'Service',
      sameAs: []
    },
    semanticTriples: [
      { subject: 'Excel Arşiv', predicate: 'optimizeEder', object: 'Yeme-İçme Sektörü Nakit ve Reçete Maliyeti' }
    ],
    heroAnswerEngine: 'Yeme-içme sektörüne özel tasarlanmış bu paket; adisyon kaçaklarını önleyen kasa mutabakatı, porsiyon bazlı reçete maliyeti ve haftalık tedarikçi ödeme planını birbirine entegre eder.',
    publishedAt: PUBLISHED_BASE,
    modifiedAt: MODIFIED_BASE,
    llmSubGraphRoute: '/llms/pages/kafe-restoran-nakit.md',
    breadcrumbs: [
      { name: 'Ana Sayfa', item: '/' },
      { name: 'Sektör', item: '/sablonlar' },
      { name: 'Kafe / Restoran', item: '/sektor/kafe-restoran-nakit' }
    ]
  },
  {
    route: '/sektor/insaat-hakedis',
    locale: 'tr-TR',
    role: 'category',
    indexDirective: 'index, follow',
    canonicalRoute: '/sektor/insaat-hakedis',
    title: 'İnşaat Hakediş, Fiyat Farkı ve Taşeron Mutabakatı Excel Sistemleri | Excel Arşiv',
    metaDescription: 'İnşaat hakediş, eskalasyon fiyat farkı, aşırı düşük teklif ve taşeron kesinti mutabakatı için profesyonel Excel karar sistemleri.',
    h1: 'İnşaat Hakediş ve Şantiye Karar Sistemleri',
    primaryIntent: 'insaat hakedis fiyat farki taseron excel',
    primaryEntity: {
      id: `${ORIGIN}/sektor/insaat-hakedis#category`,
      name: 'İnşaat ve Hakediş Excel Sistemleri',
      type: 'Service',
      sameAs: []
    },
    semanticTriples: [
      { subject: 'Excel Arşiv', predicate: 'çözer', object: 'Şantiye Hakediş ve Fiyat Farkı Hesaplama İhtiyacı' }
    ],
    heroAnswerEngine: 'Şantiye ve taahhüt projelerinde hak kaybını sıfırlayan sistem; KİK mevzuatına uyumlu eskalasyon formülleri, aşırı düşük savunma tabloları ve taşeron kesinti mutabakatını hatasız yürütür.',
    publishedAt: PUBLISHED_BASE,
    modifiedAt: MODIFIED_BASE,
    llmSubGraphRoute: '/llms/pages/insaat-hakedis.md',
    breadcrumbs: [
      { name: 'Ana Sayfa', item: '/' },
      { name: 'Sektör', item: '/sablonlar' },
      { name: 'İnşaat Hakediş', item: '/sektor/insaat-hakedis' }
    ]
  },
  {
    route: '/sektor/e-ticaret-karlilik',
    locale: 'tr-TR',
    role: 'category',
    indexDirective: 'index, follow',
    canonicalRoute: '/sektor/e-ticaret-karlilik',
    title: 'E-Ticaret ve Pazaryeri Net Kârlılık Excel Sistemleri | Excel Arşiv',
    metaDescription: 'Trendyol, Hepsiburada, Amazon komisyonları, kargo-desi maliyeti, eksik hakediş tespiti ve net tahsilat Excel modelleri.',
    h1: 'E-Ticaret ve Pazaryeri Net Kârlılık Sistemleri',
    primaryIntent: 'e-ticaret net karlilik pazaryeri hakedis excel',
    primaryEntity: {
      id: `${ORIGIN}/sektor/e-ticaret-karlilik#category`,
      name: 'E-Ticaret Kârlılık Excel Çözümleri',
      type: 'Service',
      sameAs: []
    },
    semanticTriples: [
      { subject: 'Excel Arşiv', predicate: 'tespitEder', object: 'Pazaryeri Eksik Hakediş ve Kargo Kesintileri' }
    ],
    heroAnswerEngine: 'Pazaryeri satışlarındaki gizli maliyetleri açığa çıkaran sistem; platform komisyonları, kargo kesintileri, reklam harcamaları ve iadeler sonrası net kârı sipariş bazında anında hesaplar.',
    publishedAt: PUBLISHED_BASE,
    modifiedAt: MODIFIED_BASE,
    llmSubGraphRoute: '/llms/pages/e-ticaret-karlilik.md',
    breadcrumbs: [
      { name: 'Ana Sayfa', item: '/' },
      { name: 'Sektör', item: '/sablonlar' },
      { name: 'E-Ticaret Kârlılık', item: '/sektor/e-ticaret-karlilik' }
    ]
  },
  {
    route: '/urun-bulucu',
    locale: 'tr-TR',
    role: 'tool',
    indexDirective: 'index, follow',
    canonicalRoute: '/urun-bulucu',
    title: 'Ürün Bulucu — İşletmeniz İçin Doğru Excel Sistemini Seçin | Excel Arşiv',
    metaDescription: 'İşletme büyüklüğünüze, sektörünüze ve veri hacminize göre en uygun Excel çalışma modelini 5 soruda belirleyen interaktif karar motoru.',
    h1: 'Hangi Excel Sistemini Seçmelisiniz?',
    primaryIntent: 'hangi excel sablonu urun bulucu',
    primaryEntity: {
      id: `${ORIGIN}/urun-bulucu#tool`,
      name: 'Excel Arşiv Ürün Bulucu',
      type: 'Service',
      sameAs: []
    },
    semanticTriples: [
      { subject: 'Ürün Bulucu', predicate: 'analizEder', object: 'İşletme Finansal ve Operasyonel İhtiyaçları' },
      { subject: 'Ürün Bulucu', predicate: 'önerir', object: 'Doğrulanmış Kurumsal Excel Sistemleri' }
    ],
    heroAnswerEngine: 'Excel Arşiv Ürün Bulucu; nakit akışı, muhasebe, stok veya kârlılık alanındaki birincil darboğazınızı 5 parametrede analiz ederek en uygun doğrulanmış Excel modelini deterministik olarak önerir.',
    publishedAt: PUBLISHED_BASE,
    modifiedAt: MODIFIED_BASE,
    llmSubGraphRoute: '/llms/pages/urun-bulucu.md',
    breadcrumbs: [
      { name: 'Ana Sayfa', item: '/' },
      { name: 'Ürün Bulucu', item: '/urun-bulucu' }
    ]
  }
];

// 51 Flagship Products from productSeo
const productPages: SeoPageRecord[] = Object.entries(productSeo).map(([slug, entry]) => {
  const route = `/sablon/${slug}` as const;
  return {
    route,
    locale: 'tr-TR',
    role: 'product' as const,
    indexDirective: 'index, follow' as const,
    canonicalRoute: route,
    title: entry.title,
    metaDescription: entry.description,
    h1: entry.title.split('|')[0].trim(),
    primaryIntent: entry.primaryQuery.trim().toLowerCase(),
    primaryEntity: {
      id: `${ORIGIN}${route}#product`,
      name: entry.title.split('|')[0].trim(),
      type: 'Product' as const,
      sameAs: []
    },
    semanticTriples: [
      { subject: entry.title.split('|')[0].trim(), predicate: 'çözümSunar', object: entry.primaryQuery },
      { subject: entry.title.split('|')[0].trim(), predicate: 'geliştirici', object: 'Excel Arşiv' }
    ],
    heroAnswerEngine: `${entry.description} Formülleri kilitli olmayan, denetlenebilir ve şirket verileriyle doğrudan entegre çalışan kurumsal çalışma modelidir.`,
    publishedAt: PUBLISHED_BASE,
    modifiedAt: MODIFIED_BASE,
    llmSubGraphRoute: `/llms/pages/${slug}.md` as const,
    breadcrumbs: [
      { name: 'Ana Sayfa', item: '/' },
      { name: 'Şablonlar', item: '/sablonlar' },
      { name: entry.title.split('|')[0].trim(), item: route }
    ]
  };
});

export const seoRegistry: readonly SeoPageRecord[] = Object.freeze([
  ...corePages,
  ...productPages
]);

export function getAllSeoPages(): readonly SeoPageRecord[] {
  return seoRegistry;
}

export function getSeoPageByRoute(route: string): SeoPageRecord | undefined {
  const normalized = route.endsWith('/') && route.length > 1 ? route.slice(0, -1) : route;
  return seoRegistry.find(p => p.route === normalized || (normalized === '' && p.route === '/'));
}
