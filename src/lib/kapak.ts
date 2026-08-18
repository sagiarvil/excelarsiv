// Ürün slug'ından kapak görseli yolunu döndürür.
// Kapak yalnızca kapak görseli üretilmiş ürünlerde mevcuttur; diğerlerinde undefined.
const KAPAKLAR: Record<string, string> = {
  '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi': '/images/kapak/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi.webp',
  'asiri-dusuk-teklif-savunma-robotu': '/images/kapak/asiri-dusuk-teklif-savunma-robotu.webp',
  'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi': '/images/kapak/cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi.webp',
  'cek-senet-ve-vade-risk-sistemi': '/images/kapak/cek-senet-ve-vade-risk-sistemi.webp',
  'kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici': '/images/kapak/kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici.webp',
  'pos-komisyon-ve-net-tahsilat-kontrol-sistemi': '/images/kapak/pos-komisyon-ve-net-tahsilat-kontrol-sistemi.webp',
  'sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli': '/images/kapak/sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli.webp',
  'uretim-recetesi-ve-zam-yansitma-hesaplayici': '/images/kapak/uretim-recetesi-ve-zam-yansitma-hesaplayici.webp',
  'vergi-sgk-borcunu-tecil-etmeli-miyim-kredi-mi-tecil-mi': '/images/kapak/vergi-sgk-borcunu-tecil-etmeli-miyim-kredi-mi-tecil-mi.webp',
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
  "asgari-ucret-zam-etkisi-fiyat-ayarlama-cetveli",
  "asiri-dusuk-teklif-savunma-robotu",
  "aylik-patron-finans-paneli",
  "banka-kredi-ve-taksit-takip-sistemi",
  "cari-ba-bs-toplu-mutabakat",
  "cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi",
  "cek-senet-ve-vade-risk-sistemi",
  "defter-beyan-e-arsiv-aktarim",
  "doviz-acik-pozisyonu-ve-kur-riski-stres-testi",
  "e-fatura-satir-defteri-pdf-kaniti",
  "e-fatura-toplu-donusturucu",
  "fazla-mesai-ve-isci-dava-riski-tespit-dosyasi",
  "gunluk-gelir-gider-ve-gercek-karlilik-sistemi",
  "hakedis-fiyat-farki-hak-kaybi-cetveli",
  "ihaleye-kac-tl-teklif-vermeliyim",
  "insaat-hakedis-santiye-maliyet",
  "ithalat-depo-teslim-rafa-gelen-net-birim-maliyet",
  "kacirilan-sgk-tesvikleri-ve-gercek-iscilik-maliyeti-analizi",
  "kdv-iade-listesi-robotu-gib7",
  "kdv-iadesi-azami-alacak-hesabi-dosya-hazirlayici",
  "kdv-tevkifat-mahsup-iade-listesi",
  "kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici",
  "kira-avans-takip-dekont",
  "kkeg-ve-finansman-gider-kisitlamasi-vergi-savunma-seti",
  "kobi-finans-yonetim-paketi",
  "konkordato-nakit-akis-on-projesi",
  "mutfak-kayip-kacak-hesaplayici",
  "nakliye-maliyeti-hesaplayici",
  "ortaklar-cari-ve-kasa-adat-faiz-faturasi-hesaplayici",
  "pazaryeri-net-kar-ve-eksik-hakedis-yakalayici",
  "pos-komisyon-ve-net-tahsilat-kontrol-sistemi",
  "proje-ve-is-bazinda-gercek-karlilik-sistemi",
  "restoran-recete-maliyet-fire",
  "sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli",
  "stok-satis-ve-nakit-baglanma-sistemi",
  "sube-karlilik-ve-nakit-hesaplayici",
  "taseron-hakedis-kesinti-mutabakati",
  "tesvikli-bordro-avantajli-tesvik",
  "tesvikli-bordro-optimizasyon",
  "trendyol-komisyon-sonrasi-net-kar",
  "uretim-recetesi-ve-zam-yansitma-hesaplayici",
  "vergi-sgk-borcunu-tecil-etmeli-miyim-kredi-mi-tecil-mi",
  "vergi-sgk-ve-maas-karsilik-ayirma-sistemi",
  "yeniden-degerleme-yapmali-miyim-vergi-tasarruf-analizi",
  "yillara-sari-insaat-stopaj-nakit-akis-planlayici",
  "ymm-tasdik-kontrol-robotu",
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
    return `/images/kapak/${__premiumSlug}.webp`;
  }

  const value = input as { slug?: unknown; id?: unknown } | string | null | undefined;
  const slug = typeof value === 'string'
    ? value
    : typeof value?.slug === 'string'
      ? value.slug
      : typeof value?.id === 'string'
        ? value.id
        : undefined;
  return slug && PREMIUM_KAPAK_SLUGS.has(slug) ? `/images/kapak/${slug}.webp` : undefined;
}
