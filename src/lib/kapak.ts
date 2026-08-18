// Ürün slug'ından kapak görseli yolunu döndürür.
// Kapak yalnızca kapak görseli üretilmiş ürünlerde mevcuttur; diğerlerinde undefined.
const KAPAKLAR: Record<string, string> = {
  '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi': '/images/kapak/13-haftalik-nakit-akisi.webp',
  'asiri-dusuk-teklif-savunma-robotu': '/images/kapak/asiri-dusuk-teklif-savunma-robotu.webp',
  'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi': '/images/kapak/cari-hesap-tahsilat-musteri-risk.webp',
  'cek-senet-ve-vade-risk-sistemi': '/images/kapak/cek-senet-vade-risk.webp',
  'kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici': '/images/kapak/kidem-ihbar-yuku.webp',
  'pos-komisyon-ve-net-tahsilat-kontrol-sistemi': '/images/kapak/pos-komisyon-net-tahsilat.webp',
  'sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli': '/images/kapak/ttk-376-sermaye-tamamlama.webp',
  'uretim-recetesi-ve-zam-yansitma-hesaplayici': '/images/kapak/uretim-recetesi-zam-yansitma.webp',
  'vergi-sgk-borcunu-tecil-etmeli-miyim-kredi-mi-tecil-mi': '/images/kapak/vergi-sgk-tecil-mi-kredi-mi.webp',
};

export function kapakYolu(slug: string): string | undefined {
  return KAPAKLAR[slug];
}
