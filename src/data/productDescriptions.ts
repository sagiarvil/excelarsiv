import { specificDecisionData } from './productDecisionAnatomy';

export interface ProductDescription {
  hookQuestion: string;
  subTitle: string;
  salesPitch: string[];
  keyBenefits: { title: string; desc: string }[];
  sheetDescriptions: { name: string; title: string; description: string; kind: 'input' | 'calculation' | 'output' }[];
  howToUseSteps: { step: number; title: string; desc: string }[];
  whyExcelAdvantage: string;
}

// 7 Ana Kategori İçin Psikolojik Satış Kurgusu ve Argüman Şablonları
const categoryHooks: Record<string, { hookPrefix: string; painFocus: string; solutionFocus: string; valueProposition: string }> = {
  'finansal-analiz': {
    hookPrefix: 'Şirketinizin finansal sağlığını, ciro artarken kârlılığın neden düştüğünü tek bir ekranda net olarak görmekte zorlanıyor musunuz?',
    painFocus: 'Dağınık muhasebe çıktıları, karmaşık mizanlar ve geciken raporlar arasında patron ve yöneticiler şirketin gerçek durumunu zamanında göremez. Nakdin nereye gittiği ve hangi operasyonun kârı kemirdiği belirsizleşir.',
    solutionFocus: 'Bu profesyonel Excel sistemi; tüm finansal göstergeleri, kâr marjlarını, nakit akışını ve risk sinyallerini otomatik hesaplayarak doğrudan karar aldıran bir yönetim kokpiti sunar.',
    valueProposition: 'Aylık binlerce liralık pahalı ERP ve finans yazılımlarına abone olmak yerine, tek seferlik ödemeyle ömür boyu şifresiz ve açık formüllü bir finansal analiz sistemine sahip olun.'
  },
  'nakit-akisi': {
    hookPrefix: 'Gelecek haftalarda hangi gün veya haftada nakit açığı yaşayacağınızı önceden kestiremiyor musunuz?',
    painFocus: 'Kasada bugün para olması, 2 hafta sonra çek, vergi, maaş veya tedarikçi ödemelerinin sorunsuz yapılacağı anlamına gelmez. Nakit açığını son anda fark etmek fahiş faizli acil spot kredilere ve karşılıksız çek cezalarına yol açar.',
    solutionFocus: 'Bu nakit akışı sistemi; tüm tahsilat, ödeme ve vadeleri dinamik bir takvimde birleştirerek nakit darboğazlarını haftalar öncesinden haber verir.',
    valueProposition: 'Ödemelerinizi sürpriz olmaktan çıkarın; tek bir dosya ile haftalık likidite dengenizi güvenceye alın ve maliyetli finansman tuzaklarından kurtulun.'
  },
  'muhasebe-ve-vergi': {
    hookPrefix: 'Vergi, SGK ve mevzuat hesaplamalarında tek bir formül veya eksik belge hatası yüzünden ceza alma riski mi taşıyorsunuz?',
    painFocus: 'Sürekli güncellenen mevzuat hükümleri, karmaşık matrah hesapları ve tevkifat/iade listeleri manuel yapıldığında hem günlerce zaman alır hem de vergi incelemelerinde ciddi tarhiyat riski doğurur.',
    solutionFocus: 'Mevzuat kurallarını ve çapraz kontrolleri otomatik işleten bu sistem; hesaplamaları sıfır hata ile tamamlar, vergi avantajlarını ve hak kayıplarını anında görünür kılar.',
    valueProposition: 'YMM ve mali müşavir standartlarında kurgulanmış kontrol motorlarıyla şirketinizi vergi risklerine karşı zırhlandırın.'
  },
  'butce-ve-planlama': {
    hookPrefix: 'Yıl başında koyduğunuz hedeflerin ne kadar gerçekleştiğini ve bütçe sapmalarının nereden kaynaklandığını takip etmekte zorlanıyor musunuz?',
    painFocus: 'Planlanan ile gerçekleşen arasındaki farklar düzenli takip edilmediğinde maliyet aşımları fark edilmez ve kârlılık hedefleri kağıt üzerinde kalır.',
    solutionFocus: 'Dönemsel bütçe, gerçekleşen harcamalar ve sapma analizlerini tek bir dinamik modelde toplayarak bütçe disiplinini anlık olarak sağlar.',
    valueProposition: 'Tahminlere değil, somut verilere dayalı bir bütçe yönetimiyle şirketinizi güvenle büyütün.'
  },
  'stok-ve-uretim': {
    hookPrefix: 'Depodaki stoklarınızda ne kadar nakit bağlandığını, kritik seviyeye inen ürünleri veya hammadde kayıplarını anlık göremiyor musunuz?',
    painFocus: 'Fazla stok gereksiz sermaye bağlarken, eksik stok müşteri kaybına ve üretim duruşuna neden olur. Fire, kayıp-kaçak ve zam yansıtmaları hesaplanamadığında kâr sessizce erir.',
    solutionFocus: 'Giriş-çıkış hareketlerini, maliyet katmanlarını, re-order noktalarını ve fire oranlarını tek ekranda kontrol altına alır.',
    valueProposition: 'Stok ve üretim operasyonunuzu sıfır karmaşıklıkla yönetin; depoda yatan parayı nakde çevirin.'
  },
  'satis-ve-fiyatlama': {
    hookPrefix: 'Satış yaparken komisyon, kargo, hammadde zamları ve operasyonel giderlerden sonra cebinize gerçekte ne kadar net kâr kaldığını biliyor musunuz?',
    painFocus: 'Brüt ciroya aldanıp gizli giderleri hesaba katmamak, kâr ettiğinizi sanırken sermayeyi tüketmenize yol açar.',
    solutionFocus: 'Tüm komisyon, kesinti, iskonto ve birim maliyetleri hesaplayarak hangi ürünün karlı, hangisinin zararına satıldığını net olarak ortaya koyar.',
    valueProposition: 'Doğru fiyatlama ile marjınızı koruyun; zararına satış yapma riskini tamamen ortadan kaldırın.'
  },
  'personel-ve-bordro': {
    hookPrefix: 'İşçilik maliyetleri, kıdem-ihbar yükü, kaçırılan SGK teşvikleri ve dava risklerini yönetmekte güçlük mü çekiyorsunuz?',
    painFocus: 'Personel maliyeti sadece net maaştan ibaret değildir. SGK primleri, kıdem karşılıkları, mesai riskleri ve teşvik hataları işletmeye gizli ve ağır bir yük bindirir.',
    solutionFocus: 'Tüm bordro, teşvik, kıdem ve mesai hesaplamalarını güncel yasal parametrelerle otomatikleştirerek iş gücü maliyetinizi optimize eder.',
    valueProposition: 'Yasal risklerden korunun, çalışan maliyetinizi şeffaf bir şekilde yönetin ve hak ettiğiniz tüm SGK teşviklerinden eksiksiz yararlanın.'
  }
};

// Sayfa İsimlerini Anlaşılır Başlıklara Çevirici
function formatSheetName(name: string): string {
  return name
    .replace(/_/g, ' ')
    .replace(/\b([a-zğüşıöç])/gi, (char) => char.toLocaleUpperCase('tr-TR'));
}

// Sayfa Amacını Ürüne Özgü Net Açıklamaya Dönüştürücü
function enhanceSheetDescription(
  sheet: { name: string; purpose: string; kind: 'input' | 'calculation' | 'output' },
  productName: string
): string {
  const n = sheet.name.toUpperCase();
  if (sheet.purpose && sheet.purpose.toLowerCase().includes(productName.toLowerCase().slice(0, 12))) {
    return sheet.purpose;
  }
  if (n.includes('KAPAK')) {
    return `${productName} çalışma kitabının sürüm kimliği ve ana sekme gezinti indeksidir (${sheet.purpose}).`;
  }
  if (n.includes('HIZLI_BASLANGIC') || n.includes('KILAVUZ')) {
    return `${productName} için adım adım veri girişi ve formül kullanım yönergesi (${sheet.purpose}).`;
  }
  if (n.includes('PANO') || n.includes('DASHBOARD')) {
    return `${productName} sonuç özetleri, KPI kartları ve karar göstergeleri ekranı (${sheet.purpose}).`;
  }
  if (n.includes('GIRDI') || n.includes('VERI') || sheet.kind === 'input') {
    return `${productName} veri giriş katmanı: ${sheet.purpose}. Açık hücrelere kendi kayıtlarınızı yazabilirsiniz.`;
  }
  if (n.includes('KONTROL') || n.includes('ALARM')) {
    return `${productName} tutarlılık denetimi: ${sheet.purpose}. Eksik veya dengesiz kayıtları anında işaretler.`;
  }
  if (n.includes('KARAR')) {
    return `${productName} karar motoru: ${sheet.purpose}. Eşikleri ve önerilen aksiyonu öne çıkarır.`;
  }
  if (n.includes('RAPOR')) {
    return `${productName} çıktı ekranı: ${sheet.purpose}. A4 PDF veya yönetici sunumu olarak dışa aktarılır.`;
  }
  return `${productName} (${formatSheetName(sheet.name)}): ${sheet.purpose}.`;
}

export function getProductFullDescription(template: {
  id: string;
  name: string;
  summary: string;
  category: string;
  priceTL: number;
  sheetMap: { name: string; purpose: string; kind: 'input' | 'calculation' | 'output' }[];
  inputs: string[];
  outputs: string[];
  suitableFor: string[];
  notSuitableFor: string[];
}): ProductDescription {
  const categoryConfig = categoryHooks[template.category] ?? categoryHooks['finansal-analiz'];
  const decisionData = specificDecisionData[template.id] ?? {};

  // Özel Kanca Sorusu
  const hookQuestion = decisionData.problem
    ? `${decisionData.problem} ${template.name} ile bu sorunu kökten çözün.`
    : `${categoryConfig.hookPrefix} ${template.name} tam olarak bunun için tasarlandı.`;

  // Alt Başlık (Someka tarzı)
  const subTitle = `Kullanıma hazır profesyonel Excel sistemi. ${template.summary}`;

  // Satış Paragrafları (Ürüne özel girdi, çıktı ve sayfa mimarisiyle dinamik)
  const inputList = template.inputs.slice(0, 3).join(', ');
  const outputList = template.outputs.slice(0, 3).join(', ');
  const rolesList = template.suitableFor.slice(0, 3).join(', ');
  const salesPitch = [
    `${template.name}, özellikle ${rolesList || 'yöneticiler ve finans ekipleri'} için ${inputList} verilerini işleyerek ${outputList} çıktılarını üreten profesyonel bir karar modelidir.`,
    decisionData.painPrice
      ? `${decisionData.painPrice} ${template.name} modeli ${template.sheetMap.length} çalışma sayfasıyla hesaplama kontrollerini tek yapıda toplar.`
      : `${template.name} içinde ${inputList} bilgilerini girmenizle birlikte ${outputList} göstergeleri ${template.sheetMap.length} sayfalık formül zincirinde otomatik güncellenir.`,
    `${template.name} çalışma kitabı makrosuz (.xlsx) yapısıyla harici sunuculara veri göndermeden şirketinize özel ${outputList} analizi sunar.`
  ];

  // Temel Faydalar
  const keyBenefits = [
    {
      title: `${template.name}: Otomatik Hesaplama Disiplini`,
      desc: `${inputList} verilerinden ${outputList} sonuçlarını hatasız üreten ${template.sheetMap.length} sayfalı açık formül mimarisi sunar.`
    },
    {
      title: `${rolesList || 'Yönetim'} İçin Anlık Pano`,
      desc: `${template.name} yönetim ekranında kritik sapmaları, eşik kontrollerini ve ${outputList} özetlerini tek bakışta izleyin.`
    },
    {
      title: 'Erken Uyarı ve Kontrol Katmanı',
      desc: decisionData.decisionAction
        ? `${decisionData.decisionAction}`
        : `${template.name} içindeki tutarlılık kontrolleri, ${inputList} girişlerindeki eksik veya çelişkili kayıtları anında işaretler.`
    },
    {
      title: 'Kurulumsuz Yerel .xlsx Mülkiyeti',
      desc: `Aylık abonelik gerektirmeden Windows ve macOS Excel üzerinde çalışır; ${template.name} kurumsal kullanım için açık formüllüdür.`
    }
  ];

  // Sayfa Açıklamaları
  const sheetDescriptions = template.sheetMap.map((sheet, index) => ({
    name: sheet.name,
    title: `${index + 1}- ${formatSheetName(sheet.name)}`,
    description: enhanceSheetDescription(sheet, template.name),
    kind: sheet.kind
  }));

  // 3 Adımda Kullanım Kılavuzu
  const howToUseSteps = [
    {
      step: 1,
      title: `${template.name} Girdilerini Tanımlayın`,
      desc: `${template.inputs.slice(0, 3).join(', ')} alanlarına şirket verilerinizi girin; örnek kayıtları tek hamlede temizleyin.`
    },
    {
      step: 2,
      title: 'Formül ve Tutarlılık Motorunu Çalıştırın',
      desc: `${template.sheetMap.slice(0, 4).map((s) => s.name).join(', ')} sekmeleri arasındaki dinamik bağlantılar girilen verileri otomatik işler.`
    },
    {
      step: 3,
      title: `${template.outputs[0] || 'Yönetim'} Çıktısını Raporlayın`,
      desc: `${template.outputs.slice(0, 3).join(', ')} sonuçlarını yönetici panosundan okuyun veya A4 PDF rapor olarak paylaşın.`
    }
  ];

  return {
    hookQuestion,
    subTitle,
    salesPitch,
    keyBenefits,
    sheetDescriptions,
    howToUseSteps,
    whyExcelAdvantage: categoryConfig.valueProposition
  };
}
