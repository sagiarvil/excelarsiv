// src/components/mobile/template-data.ts

export type TabType = 'home' | 'sablonlar' | 'ozel-sistemler';

export interface TemplateItem {
  id: string;
  title: string;
  category: string;
  badgeColor: string;
  authorityBadge: string;
  painPoint: string;
  solutionMetric: string;
  price: number;
  originalPrice: number;
  rating: number;
  salesCount: number;
  tags: string[];
  shopierUrl: string;
  slug: string;
}

export const TEMPLATE_DATA: TemplateItem[] = [
  {
    id: 'nakit-akim-13-hafta',
    slug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi',
    title: '13 Haftalık Dinamik Nakit Akım & Likidite Modeli',
    category: 'Finans & Nakit Yönetimi',
    badgeColor: 'bg-blue-600 text-white',
    authorityBadge: '17Y Banka Kredi Standardı',
    painPoint: 'Nakit tıkanıklığını ve çek açıklarını geç fark etme riski',
    solutionMetric: '13 Hafta Önceden Kesintisiz Kasa & Çek Hakimiyeti',
    price: 649,
    originalPrice: 1299,
    rating: 4.9,
    salesCount: 342,
    tags: ['Makrosuz & Hızlı', 'Excel + Sheets', 'DSCR Banka Rasyolu'],
    shopierUrl: 'https://www.shopier.com/49653399'
  },
  {
    id: 'kobi-gelir-gider-dashboard',
    slug: 'akilli-kasa-defteri-ve-nakit-kontrol-sistemi',
    title: 'KOBİ Otomatik Finansal Kontrol & Dashboard',
    category: 'Yönetim Raporlaması',
    badgeColor: 'bg-emerald-600 text-white',
    authorityBadge: 'Patron & Yönetici Formatı',
    painPoint: 'Ay sonu görünmeyen gider kaçakları ve kâr belirsizliği',
    solutionMetric: 'Tek Ekranda Anlık Kâr/Zarar ve Net Durum Raporu',
    price: 499,
    originalPrice: 950,
    rating: 4.8,
    salesCount: 518,
    tags: ['Formülleri Kilitli', 'Otomatik Grafik', 'Sıfır Hata'],
    shopierUrl: 'https://www.shopier.com/49652321'
  },
  {
    id: 'uretim-maliyet-hesaplama',
    slug: 'uretim-recetesi-ve-zam-yansitma-hesaplayici',
    title: 'Birim Maliyet & SKDM Karbon Uyumlu Üretim Matrisi',
    category: 'Maliyet & Operasyon',
    badgeColor: 'bg-amber-600 text-white',
    authorityBadge: 'Sanayi & İhracat Standardı',
    painPoint: 'Hammadde dalgalanmasında eksik fiyat verip zararına satış',
    solutionMetric: 'Kuruş Kuruş Net Maliyet ve Başabaş Fiyatlama',
    price: 799,
    originalPrice: 1500,
    rating: 5.0,
    salesCount: 189,
    tags: ['Endüstriyel Reçete', 'Fire Hesaplama', 'SKDM Uyumlu'],
    shopierUrl: 'https://www.shopier.com/49652403'
  }
];
