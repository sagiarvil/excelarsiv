import type { UrunGorselManifest } from './turler.ts';

// Manifesti henüz yazılmamış ürün için sistemin kırılmadan çalışmasını sağlar.
// Canlıda yalnızca geçici kullanım içindir; pilot set dışındaki ürünler buraya düşer.

export function yedekGorselOlustur(slug: string, title: string): UrunGorselManifest {
  return {
    id: `yedek-${slug}`,
    slug,
    title,
    story: 'Şablonun temel çıktısını hızlıca gösterir.',
    primaryPain: 'Veri dağınık, görünürlük düşük.',
    valueAxis: ['K'],
    resultSignal: 'Kontrol Sağlandı',
    sceneKey: 'premium-pano',
    prohibited: [],
    fingerprint: {
      layout: 'sol-metin-sag-sahne',
      heroObject: 'jenerik-pano',
      uiModule: 'kpi',
      perspective: 'on-ui',
      accent: 'green',
      resultSignal: 'Kontrol Sağlandı',
    },
    ui: {
      type: 'kpi',
      heading: 'Özet',
      metrics: [
        { label: 'Durum', value: 'Hazır', tone: 'olumlu' },
        { label: 'Kontrol', value: 'Aktif', tone: 'notr' },
      ],
    },
  };
}
