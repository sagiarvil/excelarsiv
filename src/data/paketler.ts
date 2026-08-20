import paketlerData from '../../veri/paketler.json';

export interface PaketItem {
  slug: string;
  ad: string;
  fiyat_tl: number;
  shopier_id: string;
  icerik_slugs: string[];
  aciklama: string;
  tekil_toplam_tl: number;
  toplam_sayfa_sayisi: number;
  tasarruf_tl: number;
  urun_adedi: number;
}

export const PAKETLER: PaketItem[] = paketlerData as PaketItem[];

export function getPaketBySlug(slug: string): PaketItem | undefined {
  return PAKETLER.find((p) => p.slug === slug);
}
