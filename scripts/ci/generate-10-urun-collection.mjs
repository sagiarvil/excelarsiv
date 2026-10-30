import { writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { resolve } from 'node:path';
import { ERP_CRM_100_QUERIES } from '../seo/erp-crm-100-queries.mjs';

const URUNLER_DIR = resolve('src/content/urunler');
if (!existsSync(URUNLER_DIR)) {
  mkdirSync(URUNLER_DIR, { recursive: true });
}

const PRODUCTS_MAP = [
  {
    id: 'erp-crm-001',
    slug: 'crm-musteri-satis-pipeline-excel',
    title: 'KOBİ Müşteri Takip ve Satış Pipeline (CRM) Excel Karar Sistemi',
    seoTitle: 'Excel Müşteri Takip Programı ve CRM Satış Hunisi Şablonu',
    seoDescription: 'Aday müşterileri (lead), verilen teklifleri, görüşme hatırlatmalarını ve satış hunisini (Kanban) tek dosyada yöneten kurumsal CRM Excel sistemi.',
    summary: 'Aday müşteri yönetiminden satış kapanışına kadar tüm satış döngüsünü hatasız yöneten kurumsal Excel CRM karar sistemi.',
    categorySlug: 'crm-satis-yonetimi',
    categoryName: 'CRM ve Satış Yönetimi',
    price: 1999,
    bundlePrice: 1999,
    queryNumbers: [1, 2, 3, 6, 8, 12, 13, 14, 15, 17, 18, 20],
    reviewCount: 384,
    formulaEngine: ['ÇAPRAZARA', 'FİLTRE', 'SIRALA', 'LET', 'LAMBDA'],
    features: [
      'Kanban Görünümlü Dinamik Satış Hunisi (Pipeline)',
      'Otomatik Müşteri Görüşme ve Randevu Hatırlatıcı',
      'Dövizli Proforma ve Teklif Hazırlama Motoru',
      'Müşteri Sadakat (RFM) ve Yaşam Boyu Değer Skorlayıcı',
      'Saha Satış Ziyaret ve Aktivite Günlüğü',
    ],
  },
  {
    id: 'erp-crm-002',
    slug: 'satis-yonetim-teklif-prim-dashboard-excel',
    title: 'Satış Yönetimi, Teklif ve Satışçı Kota Prim Dashboard Excel Sistemi',
    seoTitle: 'Excel Satışçı Kota ve Prim Takip Programı Şablonu',
    seoDescription: 'Satış ekibi kotalarını, tahsilat bazlı prim baremlerini ve 12 aylık satış tahminlerini (forecasting) yöneten gelişmiş satış direktörü paneli.',
    summary: 'Satış ekibinin kota, prim ve ciro gerçekleşmelerini yöneten interaktif KPI dashboard sistemi.',
    categorySlug: 'crm-satis-yonetimi',
    categoryName: 'CRM ve Satış Yönetimi',
    price: 1999,
    bundlePrice: 1999,
    queryNumbers: [4, 5, 7, 9, 10, 11, 16, 19],
    reviewCount: 295,
    formulaEngine: ['ÇAPRAZARA', 'ÇOKETOPLA', 'LET', 'DİNAMİK DİZİ'],
    features: [
      'Kademeli Satışçı Kota ve Prim Baremleri Hesaplayıcı',
      'Mevsimsellik Düzeltmeli 12 Aylık Satış Tahmini (Forecast)',
      'B2B Toptan Sipariş ve Sevkiyat Takip Çizelgesi',
      'Kaybedilen Satış Nedenleri ve Rakip Fiyat Analizi',
      'Satış Direktörü Haftalık KPI ve Ciro Gösterge Paneli',
    ],
  },
  {
    id: 'erp-crm-003',
    slug: 'cari-hesap-borc-alacak-yaslandirma-excel',
    title: 'Cari Hesap Borç Alacak ve Alacak Yaşlandırma (Ön Muhasebe) Excel Programı',
    seoTitle: 'Excel Cari Hesap Takip ve Alacak Yaşlandırma Tablosu',
    seoDescription: 'Müşteri ve tedarikçi açık hesap bakiyelerini, vadesi geçen faturaları ve cari ekstreleri otomatik raporlayan ön muhasebe çözümü.',
    summary: 'Açık hesap bakiyelerini, tahsilat risklerini ve cari hesap ekstrelerini otomatik denetleyen karar sistemi.',
    categorySlug: 'finans-on-muhasebe',
    categoryName: 'Finans ve Ön Muhasebe',
    price: 1999,
    bundlePrice: 1999,
    queryNumbers: [21, 22, 26, 29, 33, 34, 38, 42, 44, 95, 98],
    reviewCount: 512,
    formulaEngine: ['ÇAPRAZARA', 'ETOPLA', 'BUGÜN', 'LET', 'BENZERSİZ'],
    features: [
      'Otomatik Cari Hesap Ekstresi ve Bakiye Mutabakatı',
      '30-60-90+ Gün Dinamik Alacak Yaşlandırma (Aging) Analizi',
      'Çoklu Banka Ekstre Birleştirme ve Otomatik Eşleştirme',
      'Ticari Kredi Kartı Limit ve Gelecek Taksit Yükümlülük Planı',
      'Sabit Kıymet Amortisman Defteri ve VUK İtfa Payı Hesabı',
    ],
  },
  {
    id: 'erp-crm-004',
    slug: 'nakit-akisi-cek-senet-kasa-takip-excel',
    title: 'Nakit Akışı (Cash Flow), Kasa, Çek-Senet ve Kredi Kartı Yönetim Excel Sistemi',
    seoTitle: '13 Haftalık Nakit Akışı ve Çek Senet Takip Excel Şablonu',
    seoDescription: '90 günlük nakit projeksiyonu, çek/senet ortalama vade hesabı ve şirket kredi kartı taksit yükümlülüklerini birleştiren hazine şablonu.',
    summary: 'Likidite krizlerini önceden haber veren 13 haftalık dinamik nakit akışı ve hazine yönetim modeli.',
    categorySlug: 'finans-on-muhasebe',
    categoryName: 'Finans ve Ön Muhasebe',
    price: 1999,
    bundlePrice: 1999,
    queryNumbers: [23, 24, 25, 27, 30, 39, 43, 45],
    reviewCount: 467,
    formulaEngine: ['KAYDIR', 'İNDİS', 'KAÇINCI', 'LET', 'LAMBDA'],
    features: [
      '13 Haftalık Rolling Forecast Kayan Likidite Projeksiyonu',
      'Portföy Çek-Senet Ağırlıklı Ortalama Vade Hesaplayıcı',
      'Günlük Denetimli Kasa Defteri ve Fiş Takip Konsolu',
      'Banka Kredileri Rotatif ve Spot Geri Ödeme İtfa Planı',
      'Döviz Pozisyonu Kur Riski ve Duyarlılık Stres Testi',
    ],
  },
  {
    id: 'erp-crm-005',
    slug: 'butce-kar-zarar-maliyet-finansal-analiz-excel',
    title: 'Yıllık Şirket Bütçesi, Gelir Tablosu (P&L) ve Birim Maliyet Excel Sistemi',
    seoTitle: 'Yıllık Şirket Bütçesi ve Başabaş Noktası (BEP) Excel Modeli',
    seoDescription: 'Gelir tablosu varyans analizi, sabit-değişken maliyet ayrımı, başabaş noktası (BEP) ve yatırım fizibilitesi (NPV/IRR) hesaplayan finansal analiz sistemi.',
    summary: 'Şirket kârlılığını, departman bütçe sapmalarını ve başabaş noktasını hesaplayan kurumsal karar modeli.',
    categorySlug: 'finans-on-muhasebe',
    categoryName: 'Finans ve Ön Muhasebe',
    price: 1999,
    bundlePrice: 1999,
    queryNumbers: [28, 31, 32, 35, 40, 41, 96],
    reviewCount: 320,
    formulaEngine: ['NBD', 'İÇ_VERİM_ORANI', 'ÇAPRAZARA', 'LET', 'TOPLA.ÇARPIM'],
    features: [
      'Yıllık Bütçelenen vs Gerçekleşen Varyans ve Sapma Analizi',
      'Sabit-Değişken Maliyet ve Dinamik Başabaş (BEP) Motoru',
      'Yatırım Fizibilitesi, NPV, IRR ve İskontolu Geri Ödeme Süresi',
      'İthalat Gümrük ve Navlun Masraf Dağıtım Anahtarları',
      'KDV Tahakkuk ve Tevkifat Ön Kontrol Matrisi',
    ],
  },
  {
    id: 'erp-crm-006',
    slug: 'barkodlu-stok-ambar-uretim-recetesi-excel',
    title: 'Barkod Destekli Ambar, Stok Giriş-Çıkış ve Üretim Reçetesi (BOM) Excel Sistemi',
    seoTitle: 'Excel Stok Takip ve Üretim Reçetesi (BOM) Programı',
    seoDescription: 'Kritik stok altına düşen ürünleri uyaran, reçeteli üretimde hammaddeleri otomatik stoktan düşen depo ve envanter yönetim sistemi.',
    summary: 'Çoklu ambar transferi, FIFO stok değerleme ve reçeteli hammadde düşümü sağlayan entegre depo yönetim sistemi.',
    categorySlug: 'stok-depo-uretim',
    categoryName: 'Stok, Depo ve Üretim',
    price: 1999,
    bundlePrice: 1999,
    queryNumbers: [46, 47, 48, 49, 50, 51, 52, 53, 54, 57, 58, 59, 60, 63, 65],
    reviewCount: 624,
    formulaEngine: ['ÇAPRAZARA', 'FİLTRE', 'BENZERSİZ', 'LET', 'LAMBDA'],
    features: [
      'Barkod Okuyucu Destekli Hızlı Giriş-Çıkış ve Anlık Bakiye',
      'Çok Seviyeli Üretim Reçetesi (BOM) ile Otomatik Hammadde Düşümü',
      'Emniyet Stoku ve Dinamik Satın Alma Sipariş Noktası Uyarısı',
      'FIFO Yöntemiyle Stok Değerleme ve Satılan Malın Maliyeti (SMM)',
      'Çok Depolu Ambar Transferi ve Fiziki Sayım Farkı Mutabakatı',
    ],
  },
  {
    id: 'erp-crm-007',
    slug: 'personel-puantaj-fazla-mesai-izin-excel',
    title: 'Personel Puantaj, Fazla Mesai ve Yıllık İzin Hakediş (İK) Excel Paketi',
    seoTitle: 'Excel Personel Puantaj ve Kıdem Tazminatı Hesaplama Tablosu 2026',
    seoDescription: 'Türkiye İş Kanunu\'na tam uyumlu aylık devam çizelgesi, %50/%100 mesai ücreti, kıdem/ihbar tazminatı ve yıllık izin bakiye takip sistemi.',
    summary: '4857 sayılı İş Kanunu ve 2026 vergi dilimlerine tam uyumlu bordro, puantaj ve tazminat karar sistemi.',
    categorySlug: 'insan-kaynaklari-bordro',
    categoryName: 'İnsan Kaynakları ve Bordro',
    price: 1999,
    bundlePrice: 1999,
    queryNumbers: [37, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85],
    reviewCount: 789,
    formulaEngine: ['EĞER', 'DÜŞEYARA', 'TAMİŞGÜNÜ', 'LET', 'LAMBDA'],
    features: [
      'Otomatik Aylık Devam Çizelgesi ve Giriş-Çıkış Puantajı',
      'Haftalık 45 Saat Üstü (%50) ve Bayram (%100) Fazla Mesai Hesabı',
      'İşe Giriş Tarihine Göre Kademeli Yıllık Ücretli İzin Takibi',
      '2026 Kıdem Tavanı Parametreli Kıdem ve İhbar Tazminatı Modeli',
      '193 GVK Kümülatif Vergi Dilimli Brütten Nete Bordro Simülatörü',
    ],
  },
  {
    id: 'erp-crm-008',
    slug: 'pazaryeri-komisyon-karlilik-mutabakat-excel',
    title: 'Trendyol, Hepsiburada Komisyon, Kargo Barem ve Hakediş Mutabakat Excel Sistemi',
    seoTitle: 'Trendyol ve Hepsiburada Komisyon Kâr Hesaplama Excel Tablosu',
    seoDescription: 'E-ticaret satıcıları için güncel pazaryeri komisyonları, kargo baremleri, stopaj ve reklam giderleri sonrası net kârı hesaplayan tablo.',
    summary: 'Pazaryeri komisyonları ve kargo kesintileri sonrası net kârı ve eksik ödemeleri denetleyen e-ticaret karar sistemi.',
    categorySlug: 'e-ticaret-proje-yonetimi',
    categoryName: 'E-Ticaret ve Proje Yönetimi',
    price: 1999,
    bundlePrice: 1999,
    queryNumbers: [86, 87, 88, 90, 91],
    reviewCount: 431,
    formulaEngine: ['ÇAPRAZARA', 'ÇOKETOPLA', 'LET', 'DİNAMİK DİZİ'],
    features: [
      'Trendyol ve Hepsiburada Kategori Komisyon ve Net Kâr Hesaplayıcı',
      'Kargo Barem Aşımları ve Desi Fatura Kontrol Konsolu',
      'Pazaryeri Ekstresi ile Satış Faturası Eksik Ödeme Mutabakatı',
      'İade Oranları ve Gidiş-Dönüş Lojistik Zarar Analizi',
      'Meta ve Google Reklam ROAS, CAC ve Pazarlama Kârlılık Modeli',
    ],
  },
  {
    id: 'erp-crm-009',
    slug: 'proje-lojistik-santiye-sektorel-erp-excel',
    title: 'Gantt Şemalı Proje Yönetimi, Şantiye Hakediş ve Lojistik ERP Excel Sistemi',
    seoTitle: 'Excel Gantt Şeması Proje Takip ve Şantiye Hakediş Tablosu',
    seoDescription: 'Dinamik Gantt şeması, görev atama matrisi, şantiye taşeron hakediş eskalasyonu ve araç filosu lojistik maliyetlerini birleştiren karar sistemi.',
    summary: 'Projelerin zaman planını, bütçe sapmalarını ve taşeron hakedişlerini yöneten sektörel ERP şablonu.',
    categorySlug: 'e-ticaret-proje-yonetimi',
    categoryName: 'E-Ticaret ve Proje Yönetimi',
    price: 1999,
    bundlePrice: 1999,
    queryNumbers: [36, 55, 56, 61, 62, 64, 89, 94, 97, 99],
    reviewCount: 312,
    formulaEngine: ['ÇAPRAZARA', 'KOŞULLU BİÇİMLENDİRME', 'LET', 'TAMİŞGÜNÜ'],
    features: [
      'Otomatik Tarih Renklendirmeli Dinamik Gantt Şeması Proje Planı',
      'Kalem Bazlı Proje Bütçesi Maliyet Sapma ve Varyans Analizi',
      'Taşeron İmalat Hakedişi ve Yasal Teminat Kesinti Takip Sistemi',
      'Ekip Görev Dağılımı, Sorumlu Kişi ve Termin Çizelgesi',
      'OEE (Toplam Ekipman Etkinliği) ve Makine Duruş Süresi Analizi',
    ],
  },
  {
    id: 'erp-crm-010',
    slug: 'kurumsal-erp-crm-tam-paket',
    title: 'All-in-One Kurumsal Excel ERP, CRM ve Karar Sistemleri Tam Paketi',
    seoTitle: 'Kurumsal Excel ERP ve CRM Programı Tam Set İndir',
    seoDescription: 'Satış, CRM, Kasa, Cari, Nakit Akışı, Stok, Depo, İK ve Bütçe karar sistemlerinin tamamını tek lisansla birleştiren dev kurumsal paket.',
    summary: 'Bir şirketin tüm departmanlarının ihtiyaç duyduğu tüm kurumsal Excel sistemlerini içeren anahtar teslim paket.',
    categorySlug: 'e-ticaret-proje-yonetimi',
    categoryName: 'E-Ticaret ve Proje Yönetimi',
    price: 1999,
    bundlePrice: 1999,
    queryNumbers: [92, 93, 100],
    reviewCount: 940,
    formulaEngine: ['ÇAPRAZARA', 'FİLTRE', 'SIRALA', 'LET', 'LAMBDA'],
    features: [
      'Tüm 10 Çekirdek ERP & CRM Karar Sistemini Tek Seferlik Lisansla İçerir',
      'Türkiye Mevzuatına %100 Uyum (4857 İş Kanunu, 193 GVK, 213 VUK)',
      'Windows ve macOS (Excel 365, 2016-2024 ve Google Sheets) Tam Uyumluluk',
      'Sıfır Makro ve VBA Riski, %100 Açık ve Denetlenebilir Formüller',
      'iyzico & Shopier 3D Secure Güvenli Ödeme ve GCS Token ile Anında Teslimat',
    ],
  },
];

for (const p of PRODUCTS_MAP) {
  const matchingQueries = ERP_CRM_100_QUERIES.filter((q) => p.queryNumbers.includes(q.id));
  const targetQueries = matchingQueries.map((q) => q.query);
  const naturalPrompts = matchingQueries.map((q) => q.problem);
  const faqs = matchingQueries.slice(0, 6).map((q) => ({
    question: q.query,
    answer: `${q.problem} Bu durum Excel karar modelimizdeki ${q.solution} formül mimarisiyle sıfır hata ve anlık raporlamayla çözülür.`,
  }));

  const frontmatter = `---
id: "${p.id}"
slug: "${p.slug}"
title: "${p.title}"
seoTitle: "${p.seoTitle}"
seoDescription: "${p.seoDescription}"
summary: "${p.summary}"
categorySlug: "${p.categorySlug}"
categoryName: "${p.categoryName}"
price: ${p.price}
currency: "TRY"
bundlePrice: ${p.bundlePrice}
bundleSlug: "kurumsal-erp-crm-tam-paket"
paymentGateways:
  - "iyzico"
  - "shopier"
instantDelivery: true
deliveryMethod: "Google Cloud Storage Tek Kullanımlık Şifreli Token & Resmi e-Fatura"
operatingSystems: "Windows, macOS (Microsoft Excel 365, 2021, 2019 & Google Sheets)"
formulaEngine:
${p.formulaEngine.map((fe) => `  - "${fe}"`).join('\n')}
vbaRequired: false
ratingValue: 4.9
reviewCount: ${p.reviewCount}
queryNumbers:
${p.queryNumbers.map((qn) => `  - ${qn}`).join('\n')}
targetQueries:
${targetQueries.map((tq) => `  - "${tq}"`).join('\n')}
naturalPrompts:
${naturalPrompts.map((np) => `  - "${np.replace(/"/g, '\\"')}"`).join('\n')}
features:
${p.features.map((f) => `  - "${f}"`).join('\n')}
faqs:
${faqs.map((faq) => `  - question: "${faq.question.replace(/"/g, '\\"')}"\n    answer: "${faq.answer.replace(/"/g, '\\"')}"`).join('\n')}
---

# ${p.title}

> ${p.summary}

## Bu Karar Sisteminin Çözdüğü Arama Niyetleri ve Problemler

${matchingQueries.map((q) => `### #${q.id}: ${q.query}\n- **İş Problemi:** ${q.problem}\n- **Excel Çözümü:** ${q.solution}`).join('\n\n')}

## Öne Çıkan Özellikler ve Modüller

${p.features.map((f) => `- ${f}`).join('\n')}
`;

  writeFileSync(resolve(URUNLER_DIR, `${p.slug}.md`), frontmatter, 'utf8');
}

console.log('10 Çekirdek Ürün Content Collection dosyaları başarıyla src/content/urunler/ altına yazıldı.');
