// Ürün slug'ından kapak görseli yolunu döndürür.
// Kapak yalnızca kapak görseli üretilmiş ürünlerde mevcuttur; diğerlerinde undefined.
const KAPAKLAR: Record<string, string> = {
  '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi': '/images/kapak/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi.svg',
  'asiri-dusuk-teklif-savunma-robotu': '/images/kapak/asiri-dusuk-teklif-savunma-robotu.svg',
  'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi': '/images/kapak/cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi.svg',
  'cek-senet-ve-vade-risk-sistemi': '/images/kapak/cek-senet-ve-vade-risk-sistemi.svg',
  'kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici': '/images/kapak/kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici.svg',
  'pos-komisyon-ve-net-tahsilat-kontrol-sistemi': '/images/kapak/pos-komisyon-ve-net-tahsilat-kontrol-sistemi.svg',
  'sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli': '/images/kapak/sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli.svg',
  'uretim-recetesi-ve-zam-yansitma-hesaplayici': '/images/kapak/uretim-recetesi-ve-zam-yansitma-hesaplayici.svg',
  'vergi-sgk-borcunu-tecil-etmeli-miyim-kredi-mi-tecil-mi': '/images/kapak/vergi-sgk-borcunu-tecil-etmeli-miyim-kredi-mi-tecil-mi.svg',
};

export function kapakYolu(slug: string): string | undefined {
  return KAPAKLAR[slug];
}

// 50 ürünlük premium katalog kapağı katmanı.
// Mevcut resolver davranışını bozmaz; yalnız exact commerce slug eşleşmesinde öncelik alır.
const PREMIUM_KAPAK_SLUGS = new Set<string>([
  "13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi",
  "akilli-kasa-defteri-ve-nakit-kontrol-sistemi",
  "amortisman-2026-yeniden-degerleme",
  "amortisman-ve-sabit-kiymet-satis-zamanlama-stratejisti",
  "arsa-ve-gayrimenkul-proje-gelistirme-fizibilite",
  "asgari-kurumlar-vergisi-simulasyon-motoru",
  "asgari-ucret-zam-etkisi-fiyat-ayarlama-cetveli",
  "asiri-dusuk-teklif-savunma-robotu",
  "aylik-patron-finans-paneli",
  "banka-kredi-covenant-erken-uyari",
  "banka-kredi-ve-taksit-takip-sistemi",
  "borcla-sirket-satin-alma-ve-yatirim-getirisi",
  "cari-ba-bs-toplu-mutabakat",
  "cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi",
  "cek-senet-ve-vade-risk-sistemi",
  "dava-masraf-harc-hesaplama",
  "defter-beyan-e-arsiv-aktarim",
  "depo-doluluk-lokasyon",
  "dernek-kooperatif-mali-yonetim",
  "doviz-acik-pozisyonu-ve-kur-riski-stres-testi",
  "e-belge-zorunluluk-radari",
  "e-fatura-satir-defteri-pdf-kaniti",
  "e-fatura-toplu-donusturucu",
  "e-ticaret-gercek-karlilik-fiyatlama",
  "elektrik-faturasi-dogrulama",
  "emlak-pipeline-komisyon",
  "emsal-sirket-carpanlariyla-degerleme",
  "fason-uretim-takip-sistemi",
  "fazla-mesai-ve-isci-dava-riski-tespit-dosyasi",
  "fesih-maliyeti-simulatoru",
  "filo-arac-maliyet-komutasi",
  "filo-yatirim-ve-arac-yenileme-fizibilite",
  "franchise-yatirim-fizibilite",
  "gayrimenkul-tut-sat-karar",
  "gayrimenkul-yatirim-fizibilite-ve-karlilik",
  "ges-uretim-performans",
  "ges-yatirim-fizibilite-ve-proje-finansmani",
  "ges-yatirim-fizibilite",
  "gida-lot-maliyet-izlenebilirlik",
  "gunluk-gelir-gider-ve-gercek-karlilik-sistemi",
  "hakedis-fiyat-farki-hak-kaybi-cetveli",
  "hizmet-ihracati-100-indirim-transfer-takip",
  "hukuk-burosu-dosya-vekalet",
  "ihale-fiyat-farki-eskalasyon-pro",
  "ihaleye-kac-tl-teklif-vermeliyim",
  "ihracat-siparis-karlilik-kur",
  "insaat-hakedis-santiye-maliyet",
  "insaat-hakedis-yonetim-sistemi",
  "isg-risk-degerlendirme-pro-6331",
  "ithalat-depo-teslim-rafa-gelen-net-birim-maliyet",
  "ithalat-landed-cost-motoru",
  "kacirilan-sgk-tesvikleri-ve-gercek-iscilik-maliyeti-analizi",
  "kargo-desi-maliyet-optimizasyonu",
  "kat-karsiligi-hasilat-paylasimi-simulator",
  "kdv-iade-listesi-robotu-gib7",
  "kdv-iadesi-azami-alacak-hesabi-dosya-hazirlayici",
  "kdv-iadesi-tutar-surec-simulasyonu",
  "kdv-tevkifat-mahsup-iade-listesi",
  "kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici",
  "kik-asiri-dusuk-savunma-sinir-deger",
  "kira-avans-takip-dekont",
  "kira-portfoyu-getiri-komutasi",
  "kkeg-ve-finansman-gider-kisitlamasi-vergi-savunma-seti",
  "klinik-saglik-merkezi-fizibilite-ve-karlilik",
  "kobi-finans-yonetim-paketi",
  "konkordato-nakit-akis-on-projesi",
  "konkordato-ttk376-kriz-paketi",
  "kvkk-veri-envanteri-verbis-uyum",
  "logo-sql-cari-yaslandirma-tahsilat-karar-motoru",
  "makine-bakim-kalibrasyon-durus-maliyeti",
  "mini-mrp-bom-malzeme-kapasite",
  "mizan-anomali-denetim-oncesi-kontrol",
  "mutfak-kayip-kacak-hesaplayici",
  "nakliye-maliyeti-hesaplayici",
  "oee-durus-kok-neden-komutasi",
  "on-uc-haftalik-nakit-odeme-onceligi",
  "ortaklar-cari-ve-kasa-adat-faiz-faturasi-hesaplayici",
  "ortakli-gayrimenkul-yatirimi-kar-dagitim",
  "ortulu-sermaye-emsal-faiz-tf-paketi",
  "otel-fizibilite-ve-karlilik",
  "pazaryeri-hakedis-mutabakat-motoru",
  "pazaryeri-net-kar-ve-eksik-hakedis-yakalayici",
  "perakende-magaza-acilis-fizibilite",
  "pos-komisyon-ve-net-tahsilat-kontrol-sistemi",
  "profesyonel-hizmet-sirketi-karlilik-ve-kapasite",
  "proje-finansmani-dscr-llcr-paketi",
  "proje-ve-is-bazinda-gercek-karlilik-sistemi",
  "puantaj-vardiya-fazla-mesai-kanit-sistemi",
  "recete-maliyeti-menu-muhendisligi",
  "restoran-kafe-yatirim-fizibilite-ve-karlilik",
  "restoran-recete-maliyet-fire",
  "saas-finansal-planlama-runway-ve-yatirim",
  "sarj-istasyonu-yatirim",
  "sekiz-d-duzeltici-faaliyet-dosya",
  "sera-kurulum-fizibilite",
  "sevkiyat-fiyatlama-navlun",
  "sgk-prim-tesvikleri-optimizasyon-motoru",
  "sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli",
  "sirket-satin-alma-ve-ortaklik-devir-analizi",
  "site-apartman-yonetim-sistemi",
  "spc-proses-yetenek-analizi",
  "startup-yatirim-alma-ve-nakit-runway",
  "stok-optimizasyon-abc-olu-stok-nakit",
  "stok-satis-ve-nakit-baglanma-sistemi",
  "sube-karlilik-ve-nakit-hesaplayici",
  "sut-surusu-yonetim",
  "tahsilat-riski-vade-komuta-paneli",
  "tarimsal-destek-uygunluk",
  "taseron-hakedis-kesinti-mutabakati",
  "tesvikli-bordro-avantajli-tesvik",
  "tesvikli-bordro-optimizasyon",
  "trendyol-komisyon-sonrasi-net-kar",
  "uretim-kapasite-yatirim-ve-karlilik",
  "uretim-kari-125-kv-optimizasyonu",
  "uretim-recetesi-ve-zam-yansitma-hesaplayici",
  "vergi-sgk-borcunu-tecil-etmeli-miyim-kredi-mi-tecil-mi",
  "vergi-sgk-ve-maas-karsilik-ayirma-sistemi",
  "wacc-hesaplama-ve-sermaye-maliyeti",
  "yem-rasyonu-maliyet",
  "yeniden-degerleme-komuta-merkezi",
  "yeniden-degerleme-yapmali-miyim-vergi-tasarruf-analizi",
  "yillara-sari-insaat-stopaj-nakit-akis-planlayici",
  "ymm-tasdik-kontrol-robotu",
  "yurt-disi-yapilanma-vergi-simulatoru",
]);

export function premiumKapakUrl(input: unknown): string | undefined {
  const __premiumValue = arguments[0] as any;
  const __premiumSlug = typeof __premiumValue === 'string'
    ? __premiumValue
    : typeof __premiumValue?.slug === 'string'
      ? __premiumValue.slug
      : typeof __premiumValue?.id === 'string'
        ? __premiumValue.id
        : undefined;
  if (__premiumSlug && PREMIUM_KAPAK_SLUGS.has(__premiumSlug)) {
    return `/images/kapak/${__premiumSlug}.svg`;
  }

  const value = input as { slug?: unknown; id?: unknown } | string | null | undefined;
  const slug = typeof value === 'string'
    ? value
    : typeof value?.slug === 'string'
      ? value.slug
      : typeof value?.id === 'string'
        ? value.id
        : undefined;
  return slug && PREMIUM_KAPAK_SLUGS.has(slug) ? `/images/kapak/${slug}.svg` : undefined;
}
