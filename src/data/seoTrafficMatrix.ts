/**
 * EXCEL ARŞİV — TÜM ÜRÜNLER İÇİN ORGANİK TRAFİK VE MÜŞTERİ YAKALAMA MATRİSİ (V3.0)
 * 51+ sistem için arama hacmi yüksek sorgular, tıklama tetikleyicileri ve dönüşüm kancaları.
 */

export interface ProductTrafficStrategy {
  slug: string;
  category: string;
  highVolumeSearchQueries: string[];
  serpTitleHook: string;
  userSearchIntent: 'nasıl_yapılır' | 'hazır_şablon' | 'mevzuat_hesaplama' | 'risk_önleme';
  conversionLeadHook: string;
  crossSellSlug: string;
  decisionGuideSlug: string;
}

export const SEO_TRAFFIC_MATRIX: Record<string, ProductTrafficStrategy> = {
  '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi': {
    slug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi',
    category: 'Nakit Akışı',
    highVolumeSearchQueries: [
      '13 haftalık nakit akışı excel tablosu',
      'nakit akış tablosu nasıl hazırlanır',
      'kasa nakit projeksiyonu excel indir'
    ],
    serpTitleHook: '13 Haftalık Nakit Akışı Excel Şablonu | Ödeme Planlama Tablosu',
    userSearchIntent: 'hazır_şablon',
    conversionLeadHook: 'Kritik nakit açıklarını 13 hafta önceden görün, sürpriz ödeme krizlerini engelleyin.',
    crossSellSlug: 'akilli-kasa-defteri-ve-nakit-kontrol-sistemi',
    decisionGuideSlug: 'kobi-nakit-akisi-excel'
  },
  'akilli-kasa-defteri-ve-nakit-kontrol-sistemi': {
    slug: 'akilli-kasa-defteri-ve-nakit-kontrol-sistemi',
    category: 'Nakit Akışı',
    highVolumeSearchQueries: [
      'kasa defteri excel şablonu indir',
      'günlük kasa takip tablosu excel',
      'gelir gider kasa defteri formülü'
    ],
    serpTitleHook: 'Kasa Defteri Excel Şablonu | Günlük Gelir Gider ve Nakit Takibi',
    userSearchIntent: 'hazır_şablon',
    conversionLeadHook: 'Fiili kasa sayımı ile defter bakiyesi arasındaki açıkları anında yakalayın.',
    crossSellSlug: 'gunluk-gelir-gider-ve-gercek-karlilik-sistemi',
    decisionGuideSlug: 'kobi-nakit-akisi-excel'
  },
  'aylik-patron-finans-paneli': {
    slug: 'aylik-patron-finans-paneli',
    category: 'Finansal Analiz',
    highVolumeSearchQueries: [
      'şirket finans takip tablosu excel',
      'patron finans raporu excel şablonu',
      'yönetici dashboard finans paneli'
    ],
    serpTitleHook: 'Şirket Finans Takip Tablosu Excel | Aylık Patron Rapor Paneli',
    userSearchIntent: 'hazır_şablon',
    conversionLeadHook: 'Ciro, net kâr, nakit pozisyonu ve alacak gününü tek bakışta patron ekranında görün.',
    crossSellSlug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi',
    decisionGuideSlug: 'kobi-nakit-akisi-excel'
  },
  'banka-kredi-ve-taksit-takip-sistemi': {
    slug: 'banka-kredi-ve-taksit-takip-sistemi',
    category: 'Nakit Akışı',
    highVolumeSearchQueries: [
      'kredi taksit takip excel tablosu',
      'banka kredileri ödeme planı excel',
      'ticari kredi faiz anapara hesaplama'
    ],
    serpTitleHook: 'Kredi Taksit Takip Excel Tablosu | Banka Kredileri ve Ödeme Takvimi',
    userSearchIntent: 'mevzuat_hesaplama',
    conversionLeadHook: 'Tüm bankalardaki ticari kredileri ve yaklaşan taksitleri tek takvimde birleştirin.',
    crossSellSlug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi',
    decisionGuideSlug: 'kobi-nakit-akisi-excel'
  },
  'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi': {
    slug: 'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi',
    category: 'Muhasebe ve Vergi',
    highVolumeSearchQueries: [
      'cari hesap takip excel şablonu',
      'müşteri tahsilat takip tablosu',
      'vadesi geçen alacaklar listesi excel'
    ],
    serpTitleHook: 'Cari Hesap Takip Excel Şablonu | Müşteri Tahsilat ve Risk Tablosu',
    userSearchIntent: 'hazır_şablon',
    conversionLeadHook: 'Vadesi geçen alacakları müşteri risk puanıyla önceliklendirip tahsilatı hızlandırın.',
    crossSellSlug: 'cek-senet-ve-vade-risk-sistemi',
    decisionGuideSlug: 'mali-musavir-cari-takip-excel'
  },
  'cek-senet-ve-vade-risk-sistemi': {
    slug: 'cek-senet-ve-vade-risk-sistemi',
    category: 'Nakit Akışı',
    highVolumeSearchQueries: [
      'çek senet takip excel tablosu',
      'alınan verilen çekler takip programı',
      'çek vade bordrosu excel şablonu'
    ],
    serpTitleHook: 'Çek Senet Takip Excel Tablosu | Vade ve Risk Kontrol Sistemi',
    userSearchIntent: 'hazır_şablon',
    conversionLeadHook: 'Yaklaşan çek ödemelerinin nakit akışında oluşturacağı tıkanıklığı günler öncesinden görün.',
    crossSellSlug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi',
    decisionGuideSlug: 'kobi-nakit-akisi-excel'
  },
  'stok-satis-ve-nakit-baglanma-sistemi': {
    slug: 'stok-satis-ve-nakit-baglanma-sistemi',
    category: 'Stok ve Üretim',
    highVolumeSearchQueries: [
      'stok takip excel şablonu indir',
      'depo giriş çıkış takip tablosu',
      'stok maliyet ve devir hızı hesaplama'
    ],
    serpTitleHook: 'Stok Takip Excel Şablonu | Giriş Çıkış, Satış ve Maliyet Tablosu',
    userSearchIntent: 'hazır_şablon',
    conversionLeadHook: 'Depoda atıl bekleyen ölü stoğu ve stoğa bağlanan nakit tutarını netleştirin.',
    crossSellSlug: 'proje-ve-is-bazinda-gercek-karlilik-sistemi',
    decisionGuideSlug: 'trendyol-pazaryeri-net-kar-excel'
  },
  'pos-komisyon-ve-net-tahsilat-kontrol-sistemi': {
    slug: 'pos-komisyon-ve-net-tahsilat-kontrol-sistemi',
    category: 'Satış ve Fiyatlama',
    highVolumeSearchQueries: [
      'pos komisyon hesaplama excel',
      'banka pos kesintisi hesaplama tablosu',
      'pos net tahsilat kontrol şablonu'
    ],
    serpTitleHook: 'POS Komisyon Hesaplama Excel | Banka Kesintisi ve Net Tahsilat',
    userSearchIntent: 'mevzuat_hesaplama',
    conversionLeadHook: 'Banka POS kesintilerinin ve erken bloke maliyetlerinin kârınızı eritmesini engelleyin.',
    crossSellSlug: 'gunluk-gelir-gider-ve-gercek-karlilik-sistemi',
    decisionGuideSlug: 'pos-komisyon-kontrol-excel'
  },
  'proje-ve-is-bazinda-gercek-karlilik-sistemi': {
    slug: 'proje-ve-is-bazinda-gercek-karlilik-sistemi',
    category: 'Finansal Analiz',
    highVolumeSearchQueries: [
      'proje maliyet takip excel şablonu',
      'iş bazında kârlılık analizi tablosu',
      'şantiye proje gelir gider takip'
    ],
    serpTitleHook: 'Proje Maliyet Takip Excel Şablonu | İş Bazında Gerçek Kârlılık',
    userSearchIntent: 'hazır_şablon',
    conversionLeadHook: 'Hangi projenin kazandırdığını, hangi işin gizli maliyetle para yaktığını kesinleştirin.',
    crossSellSlug: 'aylik-patron-finans-paneli',
    decisionGuideSlug: 'kobi-nakit-akisi-excel'
  },
  'kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici': {
    slug: 'kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici',
    category: 'Personel ve Bordro',
    highVolumeSearchQueries: [
      'kıdem ihbar tazminatı hesaplama excel 2026',
      'personel çıkarma maliyeti hesaplayıcı',
      'kıdem tazminatı tavanı formüllü excel'
    ],
    serpTitleHook: 'Kıdem ve İhbar Tazminatı Hesaplama Excel | Personel Maliyet Cetveli',
    userSearchIntent: 'mevzuat_hesaplama',
    conversionLeadHook: 'İşten ayrılmalarda güncel tavan ve giydirilmiş ücretle hatasız tazminat hesabı yapın.',
    crossSellSlug: 'fazla-mesai-ve-isci-dava-riski-tespit-dosyasi',
    decisionGuideSlug: 'mali-musavir-cari-takip-excel'
  },
  'trendyol-komisyon-sonrasi-net-kar': {
    slug: 'trendyol-komisyon-sonrasi-net-kar',
    category: 'Satış ve Fiyatlama',
    highVolumeSearchQueries: [
      'trendyol komisyon hesaplama excel',
      'pazaryeri net kâr hesaplayıcı tablosu',
      'trendyol kargo komisyon düşüldükten sonra kâr'
    ],
    serpTitleHook: 'Trendyol Komisyon ve Net Kâr Hesaplama Excel | Pazaryeri Kârlılık Motoru',
    userSearchIntent: 'mevzuat_hesaplama',
    conversionLeadHook: 'Kargo, reklam ve komisyon sonrası zarar eden ürünleri tespit edip fiyatınızı koruyun.',
    crossSellSlug: 'stok-satis-ve-nakit-baglanma-sistemi',
    decisionGuideSlug: 'trendyol-pazaryeri-net-kar-excel'
  },
  'ithalat-depo-teslim-rafa-gelen-net-birim-maliyet': {
    slug: 'ithalat-depo-teslim-rafa-gelen-net-birim-maliyet',
    category: 'Maliyet ve Kârlılık',
    highVolumeSearchQueries: [
      'ithalat maliyet hesaplama excel tablosu',
      'gümrük navlun ithalat birim maliyet',
      'ithalat depo teslim maliyet şablonu'
    ],
    serpTitleHook: 'İthalat Maliyeti Hesaplama Excel | Depo Teslim Net Birim Maliyet',
    userSearchIntent: 'mevzuat_hesaplama',
    conversionLeadHook: 'Navlun, gümrük vergisi ve ardiye dahil depoya giren net birim maliyeti eksiksiz bulun.',
    crossSellSlug: 'doviz-acik-pozisyonu-ve-kur-riski-stres-testi',
    decisionGuideSlug: 'kobi-nakit-akisi-excel'
  },
  'sube-karlilik-ve-nakit-hesaplayici': {
    slug: 'sube-karlilik-ve-nakit-hesaplayici',
    category: 'Finansal Analiz',
    highVolumeSearchQueries: [
      'şube kârlılık analizi excel',
      'mağaza kâr zarar hesaplama tablosu',
      'şube kapatma kararı nakit akışı'
    ],
    serpTitleHook: 'Şube Kârlılık ve Nakit Yakış Analizi Excel | Mağaza Karar Motoru',
    userSearchIntent: 'risk_önleme',
    conversionLeadHook: 'Zarar eden şubelerin ana şirketten ne kadar nakit yaktığını matematiksel olarak görün.',
    crossSellSlug: 'aylik-patron-finans-paneli',
    decisionGuideSlug: 'kobi-nakit-akisi-excel'
  }
};

/**
 * Ürün slug'ına göre trafik ve yakalama stratejisini getirir.
 */
export function getProductTrafficStrategy(slug: string): ProductTrafficStrategy | undefined {
  return SEO_TRAFFIC_MATRIX[slug];
}
