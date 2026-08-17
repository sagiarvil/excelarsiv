import type { CategorySlug } from '../lib/categories';

export interface KategoriSecim {
  problem: string;
  slug: string;
}

export interface KategoriFaq {
  question: string;
  answer: string;
}

export const kategoriKarar: Record<CategorySlug, {
  secim: KategoriSecim[];
  faq: KategoriFaq[];
  related: CategorySlug[];
}> = {
  'nakit-akisi': {
    secim: [
      { problem: '13 haftalık nakit ve ödeme planı', slug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi' },
      { problem: 'Günlük kasa ve nakit kontrolü', slug: 'akilli-kasa-defteri-ve-nakit-kontrol-sistemi' },
      { problem: 'Çek, senet ve vade riski', slug: 'cek-senet-ve-vade-risk-sistemi' },
    ],
    faq: [
      { question: 'Kasa defteri ile 13 haftalık nakit akışı aynı işi mi görür?', answer: 'Hayır. Kasa defteri günlük giriş-çıkışı kontrol eder; 13 haftalık sistem ödeme planını ve nakit açığını öne çeker.' },
      { question: 'Hangi dosyadan başlamalıyım?', answer: 'Günlük kasa dağınıksa kasa defterinden; ileriye dönük ödeme baskısı varsa 13 haftalık nakit akışından başlayın.' },
    ],
    related: ['muhasebe-ve-vergi', 'butce-ve-planlama'],
  },
  'muhasebe-ve-vergi': {
    secim: [
      { problem: 'Cari, tahsilat ve müşteri riski', slug: 'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi' },
      { problem: 'KDV tevkifat, mahsup ve iade listesi', slug: 'kdv-tevkifat-mahsup-iade-listesi' },
      { problem: 'Amortisman ve yeniden değerleme', slug: 'amortisman-2026-yeniden-degerleme' },
    ],
    faq: [
      { question: 'Cari takip ile KDV dosyası birbirinin yerine geçer mi?', answer: 'Geçmez. Cari dosya tahsilat ve risk içindir; KDV dosyası vergi hesabı ve iade listesi içindir.' },
      { question: 'Muhasebe programım varken bu sistemler ne işe yarar?', answer: 'Karar ekranı ve kontrol listesi olarak çalışırlar; resmi defter yerine geçmezler.' },
    ],
    related: ['nakit-akisi', 'personel-ve-bordro'],
  },
  'butce-ve-planlama': {
    secim: [
      { problem: 'Aylık konsolide finans görünümü', slug: 'kobi-finans-yonetim-paketi' },
      { problem: '13 haftalık nakit ve ödeme planı', slug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi' },
      { problem: 'Vergi/SGK borcu tecil kararı', slug: 'vergi-sgk-borcunu-tecil-etmeli-miyim-kredi-mi-tecil-mi' },
    ],
    faq: [
      { question: 'Bütçe dosyası nakit akışının yerini tutar mı?', answer: 'Tutmaz. Bütçe dönem hedefini koyar; nakit akışı haftalık ödemeyi görünür kılar.' },
    ],
    related: ['nakit-akisi', 'finansal-analiz'],
  },
  'stok-ve-uretim': {
    secim: [
      { problem: 'Stok, satış ve nakit bağlanması', slug: 'stok-satis-ve-nakit-baglanma-sistemi' },
      { problem: 'Üretim reçetesi ve zam yansıtma', slug: 'uretim-recetesi-ve-zam-yansitma-hesaplayici' },
      { problem: 'Restoran reçete, maliyet ve fire', slug: 'restoran-recete-maliyet-fire' },
    ],
    faq: [
      { question: 'Stok takip ile kârlılık dosyası aynı mıdır?', answer: 'Değildir. Stok dosyası bağlı nakdi ve devri gösterir; kârlılık dosyası satış sonrası marjı gösterir.' },
    ],
    related: ['satis-ve-fiyatlama', 'nakit-akisi'],
  },
  'satis-ve-fiyatlama': {
    secim: [
      { problem: 'POS komisyonu ve net tahsilat', slug: 'pos-komisyon-ve-net-tahsilat-kontrol-sistemi' },
      { problem: 'Trendyol komisyon sonrası net kâr', slug: 'trendyol-komisyon-sonrasi-net-kar' },
      { problem: 'Proje ve iş bazında gerçek kârlılık', slug: 'proje-ve-is-bazinda-gercek-karlilik-sistemi' },
    ],
    faq: [
      { question: 'POS kontrolü e-ticaret kârlılığının yerini tutar mı?', answer: 'Tutmaz. POS dosyası fiziksel tahsilat kırılımı içindir; pazaryeri dosyası komisyon ve hakediş içindir.' },
    ],
    related: ['stok-ve-uretim', 'finansal-analiz'],
  },
  'personel-ve-bordro': {
    secim: [
      { problem: 'Kıdem, ihbar ve çıkış maliyeti', slug: 'kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici' },
      { problem: 'SGK teşviki ve gerçek işçilik maliyeti', slug: 'kacirilan-sgk-tesvikleri-ve-gercek-iscilik-maliyeti-analizi' },
      { problem: 'Asgari ücret zam etkisi', slug: 'asgari-ucret-zam-etkisi-fiyat-ayarlama-cetveli' },
    ],
    faq: [
      { question: 'Bu dosyalar bordro programının yerini tutar mı?', answer: 'Tutmaz. Karar ve maliyet simülasyonu içindir; resmi bordro üretmez.' },
    ],
    related: ['muhasebe-ve-vergi', 'butce-ve-planlama'],
  },
  'finansal-analiz': {
    secim: [
      { problem: 'Aylık patron finans paneli', slug: 'aylik-patron-finans-paneli' },
      { problem: 'Şube kârlılığı ve nakit', slug: 'sube-karlilik-ve-nakit-hesaplayici' },
      { problem: 'TTK 376 öz kaynak kontrolü', slug: 'sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli' },
    ],
    faq: [
      { question: 'Yönetim paneli tek ürün mü yoksa paket mi?', answer: 'Aylık patron paneli tek üründür. Konsolide kasa-banka-cari görünümü için KOBİ finans ürününe bakın.' },
    ],
    related: ['butce-ve-planlama', 'nakit-akisi'],
  },
};
