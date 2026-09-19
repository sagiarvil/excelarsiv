export interface ProductDecisionData {
  problem: string;
  painPrice: string;
  whoIsItFor: string;
  inputs: string[];
  systemComputes: string;
  decisionAction: string;
  wrongDataEffect: string;
  whyPayPrice: string;
  concreteDecisionExample: string;
}

// 7 Temel Kategori ve Amiral Ürünler İçin Zenginleştirilmiş Karar Anatomisi
export const specificDecisionData: Record<string, Partial<ProductDecisionData>> = {
  '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi': {
    problem: 'Kasada bugün para olması, üç hafta sonra çek ve vergi ödemelerini yapabileceğiniz anlamına gelmez.',
    painPrice: 'Nakit açığını 2 hafta önceden görememek; karşılıksız çek cezası, banka limitlerinin dondurulması ve fahiş faizli acil spot kredi maliyeti demektir.',
    whoIsItFor: 'Haftalık çek, vergi, maaş ve tedarikçi ödemesi olan şirket sahipleri, CFO\'lar ve finans yöneticileri.',
    inputs: ['Haftalık nakit girişleri (tahsilat, POS, nakit)', 'Haftalık sabit giderler (kira, maaş, SGK, vergi)', 'Tedarikçi çekleri ve kredi taksit vadeleri'],
    systemComputes: '13 haftalık dinamik likidite dengesini, kritik nakit açığı haftalarını ve gerekli acil finansman ihtiyacını hesaplar.',
    decisionAction: 'Hangi haftada nakit açığı oluşacağını 14 gün önceden görüp tahsilatları öne çekme veya tedarikçi vadesini yeniden sıralama kararı verirsiniz.',
    wrongDataEffect: 'Eksik veya hatalı veri durumunda sistem formülü bozulmaz; hücre bazlı doğrulama alarmları ve negatif bakiye uyarıları devreye girer.',
    whyPayPrice: 'Tek bir karşılıksız çek veya acil spot kredinin komisyonu en az 40.000 TL iken, bu karar sistemi 999 TL tek ödemeyle ömür boyu nakit kontrolü sağlar.',
    concreteDecisionExample: '4. haftada 186.000 TL kümülatif açık oluşuyor → 2 hafta önceden vadeli tahsilatlar öne çekilmeli, 3. hafta tedarikçi ödemesi 5 gün ötelenmeli.'
  },
  'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi': {
    problem: 'Hangi müşterinin nakdinizi kilitlediğini ve batık riski taşıdığını bilmeden mal vermeye devam etmek.',
    painPrice: 'Tahsilatı geciken alacakların enflasyon karşısında erimesi, tahsil edilemeyen faturalar ve şüpheli alacak karşılıkları.',
    whoIsItFor: 'Açık hesap ve vadeli çalışan toptancılar, üreticiler ve dağıtım firmaları.',
    inputs: ['Fatura tarihleri ve tutarları', 'Vade günleri ve yapılan tahsilatlar', 'Müşteri bazlı kredi limitleri'],
    systemComputes: 'Dinamik cari yaşlandırma dilimlerini (0-30, 31-60, 61-90+ gün), ortalama tahsilat süresini (DSO) ve müşteri risk skorunu hesaplar.',
    decisionAction: 'Risk skoru düşen ve vadesi 45 günü aşan müşterilere yeni mal sevkiyatını durdurma ve yasal takip başlatma kararı verirsiniz.',
    wrongDataEffect: 'Mükerrer fatura veya eksik tahsilat girildiğinde bakiye kontrol hücresi kırmızıya döner, sistem tutarsızlığı anında işaret eder.',
    whyPayPrice: '100.000 TL\'lik tek bir batık müşteri tüm yıllık kârınızı silerken, 999 TL\'lik sistem her gün hangi müşteriden para isteneceğini gösterir.',
    concreteDecisionExample: 'ABC Ticaret 74 gündür ödeme yapmıyor ve limitini %120 aştı → Açık hesap sevkiyat derhal bloke edilmeli, teminat mektubu istenmeli.'
  },
  'banka-kredi-ve-taksit-takip-sistemi': {
    problem: 'Farklı bankalardaki kredi, rotatif, BCH ve leasing taksitlerinin hangi aylarda şirketi sıkıştıracağını görememek.',
    painPrice: 'Taksit günü geldiğinde hesapta para olmaması nedeniyle gecikme faizi ödemek ve Merkez Bankası KKB kredi notunun bozulması.',
    whoIsItFor: '3\'ten fazla bankayla çalışan, kredi ve finansman kullanan tüm KOBİ ve şirket yöneticileri.',
    inputs: ['Kredi tutarı, türü (spot, rotatif, taksitli, leasing)', 'Faiz oranı, ilk taksit tarihi ve vade sayısı', 'Banka bazlı limit ve teminat bilgileri'],
    systemComputes: 'Tüm bankaların aylık anapara ve faiz ödeme takvimini, toplam finansman yükünü ve kalan borç bakiyesini tek tabloda toplar.',
    decisionAction: 'Hangi ay finansman baskısının zirve yapacağını görerek erken kredi kapama, rotatif faiz yenileme veya borç yapılandırma kararı alırsınız.',
    wrongDataEffect: 'Faiz veya taksit sayısı yanlış girilirse sağlama toplamı banka ödeme planıyla eşleşmez ve uyarı verir.',
    whyPayPrice: 'Gereksiz yere rotatifte bırakılan kredinin 1 aylık faiz farkı binlerce lira tutarken, 799 TL ile tüm banka portföyünüzü kontrol edersiniz.',
    concreteDecisionExample: 'Ekim ayında 3 ayrı bankanın taksiti aynı haftaya denk geliyor (₺420.000) → Rotatif kredi limiti devreye alınmalı veya 1 kredi ötelenmeli.'
  },
  'aylik-patron-finans-paneli': {
    problem: 'Ciro artarken kasada neden para kalmadığını, hangi giderin şirketi kemirdiğini tek bakışta görememek.',
    painPrice: 'Muhasebecinin getirdiği teknik mizanları anlamaya çalışırken geçen haftalar ve yanlış stratejik büyüme kararları.',
    whoIsItFor: 'Şirket ortakları, patronlar, genel müdürler ve yönetim kurulu üyeleri.',
    inputs: ['Aylık net satışlar ve maliyetler', 'Operasyonel giderler ve personel yükü', 'Banka borçları, kasa ve alacak bakiyeleri'],
    systemComputes: 'EBITDA (FAVÖK), brüt kâr marjı, net nakit pozisyonu, başabaş noktası ve aylık büyüme hızını tek ekranda özetler.',
    decisionAction: 'Hangi operasyonel kalemin bütçeyi aştığını görerek masraf kısma, yeni yatırım yapma veya kâr dağıtma kararını verirsiniz.',
    wrongDataEffect: 'Gelir tablosu ile bilanço bakiyesi tutmadığında yönetici paneli otomatik kırmızı alarm verir.',
    whyPayPrice: 'Bir finans direktörü (CFO) istihdam etmek ayda 100.000 TL maliyet yaratırken, 1.499 TL tek ödemeyle CFO yönetim kokpitine sahip olursunuz.',
    concreteDecisionExample: 'Ciro %18 arttı ancak brüt marj %32\'den %24\'e geriledi → Satış fiyatları enflasyonun altında kalmış, acil %12 zam yapılmalı.'
  },
  'sube-karlilik-ve-nakit-hesaplayici': {
    problem: 'Hangi şubenin gerçekten kâr ettiğini, hangi şubenin diğerlerinin kazandığı parayı yuttuğunu ayırt edememek.',
    painPrice: 'Zarar eden şubeyi aylarca açık tutmanın yarattığı kira, personel, stok ve genel merkez yükü.',
    whoIsItFor: 'Birden fazla mağazası, restoranı veya satış noktası olan zincir işletme sahipleri.',
    inputs: ['Şube bazlı ciro ve satılan malın maliyeti', 'Şube direkt giderleri (kira, personel, fatura)', 'Genel merkez dağıtım anahtarları'],
    systemComputes: 'Her şubenin net kârlılığını, katkı payını, metrekare başına verimliliğini ve başabaş cirosunu hesaplar.',
    decisionAction: 'Kâr etmeyen şubeyi kapatma, kirayı yeniden müzakere etme veya personel sayısını revize etme kararını netleştirirsiniz.',
    wrongDataEffect: 'Merkez gider dağıtımında eksik veri kalırsa şube kârlılıkları kilitlenmez, uyarı sütunu devreye girer.',
    whyPayPrice: 'Zarar eden bir şubenin aylık net kaybı en az 80.000 TL iken, 799 TL\'lik sistemle şube kapatma/büyütme kararını objektif alırsınız.',
    concreteDecisionExample: 'Kadıköy şubesi ciroda 2. sırada ancak yüksek kira ve personel nedeniyle aylık -₺34.000 zarar yazıyor → Kira indirimi istenmeli veya kapatılmalı.'
  },
  'stok-satis-ve-nakit-baglanma-sistemi': {
    problem: 'Depoda ne kadar nakdin kilitlendiğini ve hangi ürünlerin ölü stoğa dönüştüğünü fark edememek.',
    painPrice: 'Rafta aylarca bekleyen malların finansman maliyeti, depolama gideri ve modası/SKT\'si geçen ürünlerin hurdaya çıkması.',
    whoIsItFor: 'Perakendeciler, e-ticaret satıcıları, toptancılar ve imalatçılar.',
    inputs: ['Stok giriş-çıkış hareketleri ve birim maliyetler', 'Mevcut depo sayımları ve tedarik süreleri', 'Aylık ortalama satış hızları'],
    systemComputes: 'Stok devir hızını, gün cinsinden stokta kalma süresini (DIO) ve ölü stoğa kilitlenen net nakit tutarını hesaplar.',
    decisionAction: 'Hangi üründen acilen sipariş verilmesi gerektiğini, hangi ürün için kampanya yapılıp nakde çevrileceğini belirlersiniz.',
    wrongDataEffect: 'Negatif stok veya tutarsız maliyet girişinde sistem satırı sarı renkle işaretler.',
    whyPayPrice: 'Depoda kilitlenen 300.000 TL\'lik ölü stok şirketin nakit akışını felç ederken, 799 TL ile nakdinizi serbest bırakırsınız.',
    concreteDecisionExample: 'X modelinde 94 günlük stok var ve ₺140.000 nakit bağlı → Acil %15 indirimle tasfiye edilmeli, sipariş durdurulmalı.'
  }
};

export function getProductDecisionAnatomy(slug: string, summary: string, priceTL: number, inputs: string[], outputs: string[]): ProductDecisionData {
  const specific = specificDecisionData[slug];
  if (specific && specific.problem && specific.painPrice && specific.concreteDecisionExample) {
    return {
      problem: specific.problem,
      painPrice: specific.painPrice,
      whoIsItFor: specific.whoIsItFor ?? 'İlgili alanda karar almak zorunda olan işletme sahipleri, finans yöneticileri ve uzmanlar.',
      inputs: specific.inputs ?? inputs.slice(0, 3),
      systemComputes: specific.systemComputes ?? outputs.slice(0, 2).join(' ve ') + ' hesaplar.',
      decisionAction: specific.decisionAction ?? 'Eksik veya hatalı kararları önleyerek zamanında doğru operasyonel adımı atmanızı sağlar.',
      wrongDataEffect: specific.wrongDataEffect ?? 'Formüller kilitli değildir; mantıksal kontrol hücreleri hatalı veri girişinde anında görsel uyarı verir.',
      whyPayPrice: specific.whyPayPrice ?? `Danışmanlık veya yazılım maliyetlerinin yanında şirketinize özel teklifle anında kullanıma hazır karar altyapısı sağlar.`,
      concreteDecisionExample: specific.concreteDecisionExample
    };
  }

  // Genel Fallback Mantığı
  return {
    problem: `${summary} Dağınık tablolarda yapılan hatalar şirkette nakit ve kâr kaybına yol açar.`,
    painPrice: 'Plansız yürütülen süreçler nedeniyle oluşan zaman kaybı, hatalı hesaplamalar ve geciken yönetim kararları.',
    whoIsItFor: 'Sürecini Excel üzerinden hatasız, hızlı ve kurumsal standartta yönetmek isteyen işletmeler.',
    inputs: inputs.length > 0 ? inputs.slice(0, 3) : ['İlgili dönem verileri', 'Birim tutarlar ve miktarlar', 'Tarih ve vade bilgileri'],
    systemComputes: outputs.length > 0 ? outputs.slice(0, 2).join(' ve ') : 'Gerekli finansal ve operasyonel rasyoları otomatik hesaplar.',
    decisionAction: 'Tahminlere göre değil, doğrulanmış matematiksel formüllere göre aksiyon alma imkanı sunar.',
    wrongDataEffect: 'Hücre doğrulama ve mantıksal denetim formülleri sayesinde hatalı girişlerde erken uyarı sinyali verir.',
    whyPayPrice: `Yüz binlerce liralık hantal yazılımlar yerine şirketinize özel teklifle ömür boyu sınırsız kullanım elde edersiniz.`,
    concreteDecisionExample: `Veriler sisteme girildiği anda ${outputs[0] || 'kritik karar çıktısı'} ekranda belirir ve derhal aksiyona dönüştürülür.`
  };
}
