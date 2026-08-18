import type { TemplateViewModel } from '../lib/templates.ts';
import type { UrunGorselManifest, UrunHikayeKatalogOgesi } from '../lib/urun-hikayesi/turler.ts';
import { yedekGorselOlustur } from '../lib/urun-hikayesi/yedek-gorsel.ts';

// Pilot set: 8 ürün için küratörlü görsel manifest. Ölçüt değerleri ürünün kendi
// görsel kimliğiyle (CatalogProductVisual temaları) tutarlıdır; uydurma veri değildir.

export const urunHikayeManifestleri: UrunGorselManifest[] = [
  {
    id: 'visual-13-haftalik-nakit-akisi',
    slug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi',
    title: '13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi',
    story: '13 haftalık nakit modeliyle giriş-çıkış tahminlerini planlayın; kritik haftaları önceden görüp ödeme önceliği verin.',
    primaryPain: 'Nakit açığı sürpriz yakalar, ödeme planı geç kurulur.',
    valueAxis: ['K', 'Z'],
    resultSignal: 'Nakit Açığı: Önceden Görünür',
    sceneKey: 'nakit-akisi',
    prohibited: ['huni', 'staff-list'],
    fingerprint: {
      layout: 'sol-metin-sag-sahne',
      heroObject: 'haftalik-nakit-grafigi',
      uiModule: 'nakit-akisi',
      perspective: 'on-ui',
      accent: 'green',
      resultSignal: 'Nakit Açığı: Önceden Görünür',
    },
    ui: {
      type: 'nakit-akisi',
      heading: '13 Haftalık Görünüm',
      metrics: [
        { label: 'Giriş', value: '₺3,45M' },
        { label: 'Çıkış', value: '₺2,43M' },
        { label: 'Net', value: '₺1,02M' },
        { label: 'Min. Nakit', value: '₺210K' },
      ],
    },
  },
  {
    id: 'visual-cari-hesap-tahsilat',
    slug: 'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi',
    title: 'Cari Hesap, Tahsilat ve Müşteri Risk Takip Sistemi',
    story: 'Fatura, tahsilat ve vade takibini tek yapıda toplayın; vadesi geçen bakiyeleri müşteri risk skoruyla önceliklendirin.',
    primaryPain: 'Vadesi geçen bakiye ve riskli müşteri gözden kaçar.',
    valueAxis: ['R', 'K'],
    resultSignal: 'Tahsilat %78,5',
    sceneKey: 'tahsilat-takibi',
    prohibited: ['huni', 'takvim'],
    fingerprint: {
      layout: 'sol-metin-sag-sahne',
      heroObject: 'alacak-yaslandirma-tablosu',
      uiModule: 'yaslandirma',
      perspective: 'dikey-2b',
      accent: 'navy',
      resultSignal: 'Tahsilat %78,5',
    },
    ui: {
      type: 'yaslandirma',
      heading: 'Alacak Yaşlandırma',
      metrics: [
        { label: 'Alacak', value: '₺3,12M' },
        { label: 'Vadesi Geçen', value: '₺620K', tone: 'uyari' },
        { label: 'Tahsilat', value: '%78,5' },
        { label: 'Vade', value: '45 gün' },
      ],
      rows: [
        { left: '0-30 Gün', right: '₺1,05M', tone: 'olumlu' },
        { left: '31-60 Gün', right: '₺420K', tone: 'notr' },
        { left: '61+ Gün', right: '₺620K', tone: 'uyari' },
      ],
    },
  },
  {
    id: 'visual-stok-satis-nakit',
    slug: 'stok-satis-ve-nakit-baglanma-sistemi',
    title: 'Stok, Satış ve Nakit Bağlanma Sistemi',
    story: 'Stok giriş-çıkışını ortalama maliyetle işleyin; stoğa bağlanan nakdi ve devir hızını dönem sonunda görün.',
    primaryPain: 'Stokta bağlı nakit ve kritik kalemler görünmez.',
    valueAxis: ['K', 'R'],
    resultSignal: 'Kritik Stok: 18',
    sceneKey: 'stok-yonetimi',
    prohibited: ['huni', 'takvim'],
    fingerprint: {
      layout: 'sag-agirlikli-ui',
      heroObject: 'stok-sku-raf',
      uiModule: 'tablo',
      perspective: 'on-ui',
      accent: 'orange',
      resultSignal: 'Kritik Stok: 18',
    },
    ui: {
      type: 'tablo',
      heading: 'Stok Durumu',
      metrics: [
        { label: 'SKU', value: '2.450' },
        { label: 'Stok Değeri', value: '₺1,24M' },
        { label: 'Devir', value: '3,2' },
        { label: 'Kritik', value: '18', tone: 'uyari' },
      ],
      rows: [
        { left: 'A102', right: 'Kritik', tone: 'uyari' },
        { left: 'B205', right: 'Aktif', tone: 'olumlu' },
        { left: 'C314', right: 'Aktif', tone: 'olumlu' },
      ],
    },
  },
  {
    id: 'visual-gunluk-gelir-gider',
    slug: 'gunluk-gelir-gider-ve-gercek-karlilik-sistemi',
    title: 'Günlük Gelir Gider ve Gerçek Kârlılık Sistemi',
    story: 'Günlük gelir ve giderleri kategorilerle işleyin; aylık net kârı ve kategori marjını gerçek veriyle görün.',
    primaryPain: 'Ciro kâr sanılır, giderler köprüde kaybolur.',
    valueAxis: ['P', 'K'],
    resultSignal: 'Net Kâr: Ayrışık',
    sceneKey: 'karlilik-takibi',
    prohibited: ['huni', 'takvim'],
    fingerprint: {
      layout: 'sol-metin-sag-sahne',
      heroObject: 'kar-koprusu-cubuk',
      uiModule: 'grafik',
      perspective: 'izometrik-hafif',
      accent: 'emerald',
      resultSignal: 'Net Kâr: Ayrışık',
    },
    ui: {
      type: 'grafik',
      heading: 'Kâr Köprüsü',
      metrics: [
        { label: 'Gelir', value: '₺985K' },
        { label: 'Gider', value: '₺742K' },
        { label: 'Net Kâr', value: '₺243K', tone: 'olumlu' },
        { label: 'Marj', value: '%24,7' },
      ],
      rows: [
        { left: 'Gelir', right: '₺985K', tone: 'olumlu' },
        { left: 'Gider', right: '₺742K', tone: 'notr' },
        { left: 'Net Kâr', right: '₺243K', tone: 'olumlu' },
      ],
    },
  },
  {
    id: 'visual-pos-komisyon',
    slug: 'pos-komisyon-ve-net-tahsilat-kontrol-sistemi',
    title: 'POS Komisyon ve Net Tahsilat Kontrol Sistemi',
    story: 'POS kanallarının brüt satışını, komisyon ve iade düşüldükten sonraki net tahsilatını karşılaştırın; kanal maliyetini görün.',
    primaryPain: 'Komisyon ve kesintiler net tahsilatı ezer.',
    valueAxis: ['R', 'K'],
    resultSignal: 'Net: ₺666K',
    sceneKey: 'pos-mutabakat',
    prohibited: ['takvim', 'staff-list'],
    fingerprint: {
      layout: 'diyagonal',
      heroObject: 'pos-komisyon-hunisi',
      uiModule: 'huni',
      perspective: 'on-ui',
      accent: 'blue',
      resultSignal: 'Net: ₺666K',
    },
    ui: {
      type: 'huni',
      heading: 'POS Mutabakat',
      metrics: [
        { label: 'Brüt', value: '₺684K' },
        { label: 'Komisyon', value: '₺17,8K', tone: 'uyari' },
        { label: 'Net', value: '₺666K', tone: 'olumlu' },
        { label: 'Fark', value: '₺2,4K' },
      ],
      rows: [
        { left: 'Brüt', right: '₺684K', tone: 'notr' },
        { left: 'Komisyon', right: '-₺17,8K', tone: 'uyari' },
        { left: 'Net', right: '₺666K', tone: 'olumlu' },
      ],
    },
  },
  {
    id: 'visual-cek-senet-vade',
    slug: 'cek-senet-ve-vade-risk-sistemi',
    title: 'Çek Senet ve Vade Risk Sistemi',
    story: 'Çek ve senet portföyünü vade aralıklarıyla izleyin; tahsilat beklentisini ve riskli kalemleri tek bakışta görün.',
    primaryPain: 'Çek-senet vadeleri dağınık, risk günü sürpriz olur.',
    valueAxis: ['R', 'K'],
    resultSignal: 'Geçen Evrak: ₺85K',
    sceneKey: 'vade-takibi',
    prohibited: ['huni', 'staff-list'],
    fingerprint: {
      layout: 'ust-baslik-alt-sahne',
      heroObject: 'vade-takvimi',
      uiModule: 'takvim',
      perspective: 'dikey-2b',
      accent: 'red',
      resultSignal: 'Geçen Evrak: ₺85K',
    },
    ui: {
      type: 'takvim',
      heading: 'Vade Takvimi',
      metrics: [
        { label: 'Evrak', value: '125' },
        { label: 'Toplam', value: '₺1,85M' },
        { label: 'Yaklaşan', value: '₺320K' },
        { label: 'Geçen', value: '₺85K', tone: 'uyari' },
      ],
    },
  },
  {
    id: 'visual-vergi-sgk-maas',
    slug: 'vergi-sgk-ve-maas-karsilik-ayirma-sistemi',
    title: 'Vergi SGK ve Maaş Karşılık Ayırma Sistemi',
    story: 'Dönem gelirine göre KDV, kurumlar vergisi, SGK ve maaş karşılıklarını otomatik ayırın; dönem sonunu sürprizsiz karşılayın.',
    primaryPain: 'Maaş-SGK-vergi yükü ayrışmaz, nakit rezervi yanlış kurulur.',
    valueAxis: ['R', 'P'],
    resultSignal: 'Yük: Ayrışmış',
    sceneKey: 'bordro-yuku',
    prohibited: ['huni', 'takvim'],
    fingerprint: {
      layout: 'sag-agirlikli-ui',
      heroObject: 'bordro-yuk-agirligi',
      uiModule: 'grafik',
      perspective: 'karisik',
      accent: 'yellow',
      resultSignal: 'Yük: Ayrışmış',
    },
    ui: {
      type: 'grafik',
      heading: 'Bordro Yükü',
      metrics: [
        { label: 'Maaş', value: '₺286K' },
        { label: 'SGK', value: '₺91K' },
        { label: 'Vergi', value: '₺62K', tone: 'uyari' },
        { label: 'Personel', value: '24' },
      ],
      rows: [
        { left: 'Maaş', right: '₺286K', tone: 'olumlu' },
        { left: 'SGK', right: '₺91K', tone: 'notr' },
        { left: 'Vergi', right: '₺62K', tone: 'uyari' },
      ],
    },
  },
  {
    id: 'visual-banka-kredi-taksit',
    slug: 'banka-kredi-ve-taksit-takip-sistemi',
    title: 'Banka Kredi ve Taksit Takip Sistemi',
    story: 'Tüm banka hesaplarını, kredi taksitlerini ve ödeme tarihlerini tek sayfada izleyin; yaklaşan taksitleri önceden görün.',
    primaryPain: 'Taksit ve kalan borç takibi dağınık, gecikme yük olur.',
    valueAxis: ['K', 'Z'],
    resultSignal: 'Taksit: İzleniyor',
    sceneKey: 'kredi-taksit',
    prohibited: ['huni', 'staff-list'],
    fingerprint: {
      layout: 'sol-metin-sag-sahne',
      heroObject: 'taksit-plani-tablosu',
      uiModule: 'liste',
      perspective: 'on-ui',
      accent: 'blue',
      resultSignal: 'Taksit: İzleniyor',
    },
    ui: {
      type: 'liste',
      heading: 'Ödeme Planı',
      metrics: [
        { label: 'Borç', value: '₺2,24M' },
        { label: 'Kalan', value: '₺1,42M' },
        { label: 'Taksit', value: '₺95K' },
        { label: 'Geciken', value: '₺20K', tone: 'uyari' },
      ],
      rows: [
        { left: 'Taksit 5', right: '₺95K', tone: 'olumlu' },
        { left: 'Taksit 6', right: '₺95K', tone: 'olumlu' },
        { left: 'Geciken', right: '₺20K', tone: 'uyari' },
      ],
    },
  },
];

const manifestBySlug = new Map(urunHikayeManifestleri.map((m) => [m.slug, m]));

export function urunGorselManifestiBul(slug: string): UrunGorselManifest | null {
  return manifestBySlug.get(slug) ?? null;
}

// Kart için katalog ögesi üretir; manifesti olmayan ürün yedek görsele düşer.
// Fiyat ve kategori her zaman canlı veriden gelir; uydurma değildir.
export function urunHikayeKatalogOgesi(template: TemplateViewModel): UrunHikayeKatalogOgesi {
  const manifest = urunGorselManifestiBul(template.slug);
  const gorsel = manifest ?? yedekGorselOlustur(template.slug, template.name);
  return {
    id: gorsel.id,
    slug: template.slug,
    title: template.name,
    href: template.url,
    priceText: `₺${template.priceTL.toLocaleString('tr-TR')}`,
    badge: 'Excel Şablonu',
    categorySlug: template.categorySlug,
    categoryLabel: template.categoryName,
    kapak: template.kapak,
    visual: gorsel,
  };
}
