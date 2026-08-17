// Product Story Card System için görsel manifest tipleri.
// Tüm alan adları ASCII kalır (kaynak dil guard'ı gereği); değerler Türkçe serbesttir.

export type DegerEkseni = 'P' | 'Z' | 'R' | 'K' | 'I';

export type YerlesimTipi =
  | 'sol-metin-sag-sahne'
  | 'ust-baslik-alt-sahne'
  | 'diyagonal'
  | 'sag-agirlikli-ui'
  | 'merkez-kahraman';

export type BakisAcisi = 'dikey-2b' | 'izometrik-hafif' | 'on-ui' | 'karisik';

export type UiModulTipi =
  | 'kpi'
  | 'tablo'
  | 'grafik'
  | 'huni'
  | 'takvim'
  | 'yaslandirma'
  | 'nakit-akisi'
  | 'donusum'
  | 'liste';

export type SahneAnahtari =
  | 'nakit-akisi'
  | 'tahsilat-takibi'
  | 'stok-yonetimi'
  | 'karlilik-takibi'
  | 'pos-mutabakat'
  | 'vade-takibi'
  | 'bordro-yuku'
  | 'kredi-taksit'
  | 'premium-pano'
  | 'fonksiyon-paketi';

export type Ton = 'olumlu' | 'notr' | 'uyari';

export interface UrunOlcutu {
  label: string;
  value: string;
  tone?: Ton;
}

export interface UrunSatiri {
  left: string;
  right: string;
  tone?: Ton;
}

export interface UiYapilandirmasi {
  type: UiModulTipi;
  heading?: string;
  metrics?: UrunOlcutu[];
  rows?: UrunSatiri[];
}

export interface GorselParmakIzi {
  layout: YerlesimTipi;
  heroObject: string;
  uiModule: UiModulTipi;
  perspective: BakisAcisi;
  accent: string;
  resultSignal: string;
}

export interface UrunGorselManifest {
  id: string;
  slug: string;
  title: string;
  story: string;
  primaryPain: string;
  valueAxis: DegerEkseni[];
  resultSignal: string;
  sceneKey: SahneAnahtari;
  prohibited: string[];
  fingerprint: GorselParmakIzi;
  ui: UiYapilandirmasi;
}

export interface UrunHikayeKatalogOgesi {
  id: string;
  slug: string;
  title: string;
  href: string;
  priceText?: string;
  badge?: string;
  categorySlug: string;
  categoryLabel: string;
  visual: UrunGorselManifest;
}
