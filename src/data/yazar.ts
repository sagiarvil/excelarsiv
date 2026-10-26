/** Excel Arşiv E-E-A-T yazar entity — tek kaynak (Person schema + UI). */

export const YAZAR = {
  ad: 'Barış Bağırlar',
  unvan: 'Finansal Karar Sistemleri',
  rolKisa: 'Tasarlayan ve denetleyen',
  rolUzun:
    'Her Excel şablonunu tasarlayan ve denetleyen; finansal karar sistemleri mimarı.',
  ozet:
    'Bankacılık kökenli finansal karar sistemleri uzmanı. Excel Arşiv şablonlarını saha gerçekliği, denetlenebilir formül mimarisi ve yöneticiye sunulabilir karar çıktısı standardında tasarlar ve denetler.',
  neden:
    'İşletmelerin kritik kararları dağınık tablolara ve tahminlere bırakılmamalı. Bu şablonları, bankacılık ve saha deneyimimle ölçülebilir, denetlenebilir ve yöneticiye sunulabilir karar sistemleri olarak kuruyorum.',
  foto: '/images/baris-bagirlar.jpg',
  fotoKucuk: '/images/baris-bagirlar-96.jpg',
  fotoGenislik: 512,
  fotoYukseklik: 512,
  eposta: 'barisbagirlar@gmail.com',
  profilYolu: '/hakkinda',
  sameAs: [
    'https://www.wikidata.org/wiki/Q11589432',
    'https://www.wikidata.org/wiki/Q11190',
    'https://www.linkedin.com/in/barisbagirlar/',
    'https://www.tarimkon.org/danisma-kurulu/',
    'https://drfin.com.tr',
    'https://excelarsiv.com',
    'https://inversioncore.com',
    'https://belginkuyumculuk.com',
    'https://saatchi.com.tr',
    'https://dilekceyazdir.com.tr',
    'https://htmlandhtml.com',
    'https://degerlet.com',
  ],
  kimlikler: [
    'Bankacı',
    'Tarımkon — Uluslararası Tarım ve Gıda Konfederasyonu Danışma Kurulu Üyesi',
  ],
  platformlar: [
    { ad: 'DrFin.com.tr', url: 'https://drfin.com.tr' },
    { ad: 'ExcelArsiv.com', url: 'https://excelarsiv.com' },
    { ad: 'InversionCore.com', url: 'https://inversioncore.com' },
    { ad: 'BelginKuyumculuk.com', url: 'https://belginkuyumculuk.com' },
    { ad: 'Saatchi.com.tr', url: 'https://saatchi.com.tr' },
    { ad: 'DilekceYazdir.com.tr', url: 'https://dilekceyazdir.com.tr' },
    { ad: 'HTMLandHTML.com', url: 'https://htmlandhtml.com' },
    { ad: 'Degerlet.com', url: 'https://degerlet.com' },
  ],
} as const;

export function yazarPersonId(site: URL | string): string {
  const origin = typeof site === 'string' ? site : site.toString();
  return new URL('/hakkinda#person', origin).href;
}

export function yazarPersonLd(site: URL | string) {
  const origin = typeof site === 'string' ? new URL(site) : site;
  const id = yazarPersonId(origin);
  const imageUrl = new URL(YAZAR.foto, origin).href;
  return {
    '@type': 'Person',
    '@id': id,
    name: YAZAR.ad,
    url: new URL(YAZAR.profilYolu, origin).href,
    image: {
      '@type': 'ImageObject',
      url: imageUrl,
      width: YAZAR.fotoGenislik,
      height: YAZAR.fotoYukseklik,
      caption: `${YAZAR.ad} — ${YAZAR.unvan}`,
    },
    jobTitle: YAZAR.unvan,
    description: YAZAR.ozet,
    email: YAZAR.eposta,
    worksFor: { '@id': new URL('/#organization', origin).href },
    sameAs: [...YAZAR.sameAs],
    knowsAbout: [
      'Finansal karar sistemleri',
      'Excel karar destek modelleri',
      'Nakit akışı ve işletme finansı',
      'Denetlenebilir formül mimarisi',
    ],
  };
}
