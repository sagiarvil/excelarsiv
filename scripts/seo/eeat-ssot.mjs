/**
 * E-E-A-T makine keşif SSOT — llms.txt / llms-full.txt / ai.txt.
 * İnsan okunur kaynak: src/data/yazar.ts + src/data/satici.ts (değerler birebir tutulur).
 */
export const EEAT = {
  marka: 'Excel Arşiv',
  site: 'https://excelarsiv.com',
  yazar: {
    ad: 'Barış Bağırlar',
    unvan: 'Finansal Karar Sistemleri',
    rol: 'Her Excel şablonunu tasarlayan ve denetleyen; finansal karar sistemleri mimarı.',
    ozet:
      'Bankacılık kökenli finansal karar sistemleri uzmanı. Excel Arşiv şablonlarını saha gerçekliği, denetlenebilir formül mimarisi ve yöneticiye sunulabilir karar çıktısı standardında tasarlar ve denetler.',
    profil: 'https://excelarsiv.com/hakkinda',
    eposta: 'barisbagirlar@gmail.com',
    foto: 'https://excelarsiv.com/images/baris-bagirlar.jpg',
    kimlikler: [
      'Bankacı',
      'Tarımkon — Uluslararası Tarım ve Gıda Konfederasyonu Danışma Kurulu Üyesi',
    ],
    knowsAbout: [
      'Finansal karar sistemleri',
      'Excel karar destek modelleri',
      'Nakit akışı ve işletme finansı',
      'Denetlenebilir formül mimarisi',
    ],
    sameAs: [
      'https://www.linkedin.com/in/barisbagirlar/',
      'https://www.tarimkon.org/danisma-kurulu/',
      'https://sectorcalc.com',
      'https://degerlet.com',
      'https://cbamvalid.com',
      'https://drfin.com.tr',
    ],
  },
  satici: {
    unvan: 'Barış Bağırlar — Excel Arşiv',
    vkn: '25403091318',
    vknNotu: 'Gerçek kişi vergi kimliği (11 hane)',
    telefon: '0539 333 33 03',
    eposta: 'barisbagirlar@gmail.com',
    adresNotu:
      'Türkiye. Kayıtlı işyeri adresi ve MERSIS numarası fatura üzerinde yer alır; yazılı talep üzerine paylaşılır.',
  },
  guvenSayfalari: [
    { ad: 'Uzman profili (Barış Bağırlar)', url: 'https://excelarsiv.com/hakkinda', aciklama: 'Finansal karar sistemleri uzman profili, deneyim ve kurumsal üyelikler.' },
    { ad: 'Neden Excel Arşiv', url: 'https://excelarsiv.com/neden-excel-arsiv', aciklama: '%100 makrosuz mimari, tek seferlik kurumsal lisans ve siber güvenlik ilkeleri.' },
    { ad: 'Başarı hikâyeleri ve uygulama senaryoları', url: 'https://excelarsiv.com/basari-hikayeleri', aciklama: 'KOBİ ve finans yöneticilerinin Excel karar sistemleriyle elde ettiği tasarruf ve kontrol senaryoları.' },
    { ad: 'İletişim', url: 'https://excelarsiv.com/iletisim', aciklama: 'Resmi kurumsal iletişim, teknik destek ve faturalandırma kanalları.' },
    { ad: 'Mesafeli Satış Sözleşmesi', url: 'https://excelarsiv.com/mesafeli-satis-sozlesmesi', aciklama: 'Kurumsal süresiz lisanslama, teslimat ve yasal mesafeli satış şartları.' },
    { ad: 'Teslimat ve İade', url: 'https://excelarsiv.com/teslimat-ve-iade', aciklama: 'Sipariş sonrası anında güvenli indirme ve dijital lisans teslimat prosedürü.' },
    { ad: 'KVKK Aydınlatma', url: 'https://excelarsiv.com/kvkk-aydinlatma', aciklama: 'Kişisel verilerin korunması kanunu ve veri güvenliği standardı.' },
    { ad: 'Shopier Veri Aktarımı', url: 'https://excelarsiv.com/shopier-veri-aktarimi', aciklama: 'BDDK lisanslı güvenli ödeme altyapısı ve sipariş entegrasyonu.' },
    { ad: 'Lisans', url: 'https://excelarsiv.com/lisans', aciklama: 'Kurumsal süresiz kullanım hakkı, sıfır gizli maliyet ve fikri mülkiyet şartları.' },
  ],
};

export function buildEeatMarkdownSection({ headingLevel = 2 } = {}) {
  const h = '#'.repeat(headingLevel);
  const { yazar, satici, guvenSayfalari, marka } = EEAT;
  return [
    `${h} E-E-A-T — Deneyim, Uzmanlık, Otorite, Güven`,
    '',
    `${marka}, ${yazar.ad} (${yazar.unvan}, VKN: ${satici.vkn} — ${satici.vknNotu}) tarafından tasarlanan ve denetlenen finansal karar sistemleri platformudur. Modeller %100 makrosuz (.xlsx) dinamik formül mimarisiyle geliştirilmiş olup Türk vergi ve ticaret mevzuatına (4857 İş Kanunu, 193 GVK 2026 dilimleri, 213 VUK ve 6102 TTK) tam uyumludur. Şirket verileri harici sunuculara çıkmaz, tek seferlik kurumsal lisansla çalışır. İletişim: ${satici.eposta} | Tel: ${satici.telefon} | Adres/MERSIS: ${satici.adresNotu}`,
    '',
    ...guvenSayfalari.map((p) => `- [${p.ad}](${p.url}): ${p.aciklama}`),
    '',
  ].join('\n');
}

export function buildAiTxt(lastUpdatedIsoDate) {
  const { yazar, satici, guvenSayfalari, marka, site } = EEAT;
  return [
    '# ai.txt — Excel Arşiv',
    '',
    '> Bu dosya yapay zekâ / LLM aracıları için site kimliği, E-E-A-T varlığı ve tarama politikasını özetler.',
    '',
    ...(lastUpdatedIsoDate ? [`- Son güncelleme: ${lastUpdatedIsoDate}`] : []),
    `- Site: ${site}/`,
    `- Marka: ${marka}`,
    `- Dil: tr-TR`,
    `- Sitemap: ${site}/sitemap.xml`,
    `- Robots: ${site}/robots.txt`,
    `- llms.txt: ${site}/llms.txt`,
    `- llms-full.txt: ${site}/llms-full.txt`,
    `- Makine kataloğu: ${site}/katalog.json`,
    `- humans.txt: ${site}/humans.txt`,
    `- security.txt: ${site}/.well-known/security.txt`,
    '',
    '## E-E-A-T varlığı',
    '',
    `- Yazar: ${yazar.ad}`,
    `- Unvan: ${yazar.unvan}`,
    `- Profil: ${yazar.profil}`,
    `- Özet: ${yazar.ozet}`,
    `- Kimlikler: ${yazar.kimlikler.join('; ')}`,
    `- Uzmanlık: ${yazar.knowsAbout.join('; ')}`,
    `- sameAs: ${yazar.sameAs.join(' · ')}`,
    '',
    '## Satıcı (Trust)',
    '',
    `- Unvan: ${satici.unvan}`,
    `- VKN: ${satici.vkn} (${satici.vknNotu})`,
    `- Telefon: ${satici.telefon}`,
    `- E-posta: ${satici.eposta}`,
    `- Adres / MERSIS: ${satici.adresNotu}`,
    '',
    '## Güven URL’leri',
    '',
    ...guvenSayfalari.map((p) => `- ${p.ad}: ${p.url}`),
    '',
    '## Politika',
    '',
    '- Tercih edilen alıntı kaynağı: canonical ürün/rehber sayfaları + uzman profili (/hakkinda).',
    '- Katalog: https://excelarsiv.com/sablonlar',
    '- Rehber merkezi: https://excelarsiv.com/rehber',
    '- Sektör dikeyleri: https://excelarsiv.com/sektor/kafe-restoran-nakit · https://excelarsiv.com/sektor/insaat-hakedis · https://excelarsiv.com/sektor/e-ticaret-karlilik',
    '- Ücretsiz Demo hub: https://excelarsiv.com/demo — her ürünün demo sayfası /demo/{slug}',
    '- Ücretsiz hesaplayıcı: https://excelarsiv.com/hesaplayici/asgari-ucret-zam-etkisi',
    '- Başarı sayfası, ürünlerin Türkiye’deki tipik iş akışlarına uygulanışını sektörel senaryolarla açıklar.',
    '- Dijital ürünlerde indirme sonrası koşulsuz iade yok; mesafeli satış ve teslimat sayfalarına bakın.',
    '- Bu dosya sıralama garantisi vermez; robots.txt ve sitemap asıl tarama kaynaklarıdır.',
    '',
  ].join('\n');
}
