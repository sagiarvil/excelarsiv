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

// Sayfa Amacını Kullanıcı Dostu Someka Tarzı Açıklamaya Dönüştürücü
function enhanceSheetDescription(sheet: { name: string; purpose: string; kind: 'input' | 'calculation' | 'output' }): string {
  const n = sheet.name.toUpperCase();
  if (n.includes('KAPAK')) {
    return 'Dosyanın genel kimliğini, sürüm bilgilerini ve tek tıkla ilgili sayfalara gitmenizi sağlayan gezinti butonlarını içerir.';
  }
  if (n.includes('HIZLI_BASLANGIC') || n.includes('KILAVUZ')) {
    return 'Şablonu ilk kez açan kullanıcılar için adım adım kullanım talimatlarını ve formül mantığını özetleyen rehber sayfadır.';
  }
  if (n.includes('PANO') || n.includes('DASHBOARD')) {
    return 'Tüm verilerin özetlendiği, grafikler ve kilit performans göstergeleriyle (KPI) desteklenmiş ana yönetim ekranıdır. Kararlarınızı tek bakışta bu sayfadan verirsiniz.';
  }
  if (n.includes('GIRDI') || n.includes('VERI') || sheet.kind === 'input') {
    return `${sheet.purpose}. Bu alanda yalnızca beyaz renkli giriş hücrelerini doldurmanız yeterlidir; formüller verileri otomatik olarak işler.`;
  }
  if (n.includes('HESAP') || n.includes('MOTOR') || sheet.kind === 'calculation') {
    return `${sheet.purpose}. Arka planda matematiksel formüller, oranlar ve koşullu mantıklar hatasız olarak çalışır; elle müdahaleye gerek kalmaz.`;
  }
  if (n.includes('KONTROL') || n.includes('ALARM')) {
    return 'Hatalı, eksik veya tutarsız veri girişlerini anında yakalayan; bakiye ve mantık uyumsuzluklarında kırmızı uyarı veren denetim sayfasıdır.';
  }
  if (n.includes('KARAR')) {
    return 'Ham rakamları somut iş kararlarına dönüştüren karar motorudur. Hangi aksiyonun ne zaman alınması gerektiğini net önerilerle sunar.';
  }
  if (n.includes('RAPOR')) {
    return 'Yönetime, bankalara veya ortaklara sunulmaya hazır; yazdırılabilir ve PDF olarak dışa aktarılabilir profesyonel dönem raporudur.';
  }
  return `${sheet.purpose}. İş sürecinizi hızlandırmak ve hatasız sonuç üretmek için özel olarak tasarlanmıştır.`;
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

  // Satış Paragrafları (İkna kabiliyeti yüksek, psikolojisi güçlü)
  const salesPitch = [
    `${template.name}, işletmenizin en kritik operasyonel ve finansal süreçlerini tek bir dosyada kusursuz bir şekilde yönetmeniz için geliştirilmiş kapsamlı bir Excel çözümüdür. Dağınık notlar, birbirini tutmayan tablolar ve son anda fark edilen hatalar yüzünden yaşanan stres ve zaman kaybını tamamen ortadan kaldırır.`,
    decisionData.painPrice
      ? `Bu alanda yapılan hataların işletmeye faturası çok ağır olabilir: ${decisionData.painPrice} ${template.name}, bu riskleri sıfıra indirmek ve her adımda doğru kararı vermenizi sağlamak için arka planda çalışan gelişmiş formül ağıyla donatılmıştır.`
      : `${categoryConfig.painFocus} ${template.name} sayesinde verilerinizi sadece ilgili giriş alanlarına yazmanız yeterlidir; sistem kalan tüm hesaplamaları, kontrolleri ve raporları sizin yerinize saniyeler içinde tamamlar.`,
    `Başka hiçbir pahalı yazılıma, karmaşık ERP modülüne veya aylık abonelik ücretine gerek kalmadan; kendi bilgisayarınızda %100 gizlilikle, internet bağımlılığı olmadan hemen çalışmaya başlayabilirsiniz. Tamamen şifresiz, açık formüllü ve şirketinize göre özelleştirilebilir bir yapı sunar.`
  ];

  // Temel Faydalar
  const keyBenefits = [
    {
      title: 'Zaman Tasarrufu ve Sıfır Formül Hatası',
      desc: 'Günlerce süren manuel hesaplama, tablo eşleştirme ve hata arama süreçlerini sonlandırır. Formüller önceden test edilmiş ve doğrulanmıştır.'
    },
    {
      title: 'Anlık Yönetici Raporları ve Görsel Dashboard',
      desc: 'Rakamları boğucu tablolardan çıkarıp tek bakışta anlaşılan grafikler ve karar göstergelerine dönüştürür. Ortaklarınıza veya bankanıza gururla sunabilirsiniz.'
    },
    {
      title: 'Erken Uyarı ve Risk Tespiti',
      desc: decisionData.decisionAction
        ? `${decisionData.decisionAction}`
        : 'Beklenmeyen nakit açıkları, geciken ödemeler ve kârlılık sapmalarını henüz oluşmadan yakalar ve aksiyon almanızı sağlar.'
    },
    {
      title: 'Şeffaf Teklif Modeli, Ömür Boyu Sahiplik',
      desc: 'Aylık veya yıllık yinelenen gizli lisans ücreti yoktur. Şirketinize özel tek seferlik teklifle dosya tamamen firmanıza devredilir, süre sınırı olmadan kullanabilirsiniz.'
    },
    {
      title: 'Windows ve Mac ile %100 Uyumlu',
      desc: 'Microsoft Excel 2016 ve üzerindeki tüm masaüstü sürümlerde sorunsuz çalışır. Dışarıya veri göndermez; tüm şirket bilgileriniz kendi bilgisayarınızda kalır.'
    }
  ];

  // Sayfa Açıklamaları (Someka Görsel 1 & 2'deki gibi 1- Dashboard, 2- Ekip Verileri vb.)
  const sheetDescriptions = template.sheetMap.map((sheet, index) => ({
    name: sheet.name,
    title: `${index + 1}- ${formatSheetName(sheet.name)}`,
    description: enhanceSheetDescription(sheet),
    kind: sheet.kind
  }));

  // 3 Adımda Kullanım Kılavuzu
  const howToUseSteps = [
    {
      step: 1,
      title: 'Verilerinizi Girin',
      desc: `Dosyayı açın ve açık renkli girdi hücrelerine temel işletme verilerinizi (${template.inputs.slice(0, 3).join(', ')}) yazın. Hazır örnek verileri silip kendi rakamlarınızı eklemeniz birkaç dakika sürer.`
    },
    {
      step: 2,
      title: 'Otomatik Hesaplama ve Kontrolleri İzleyin',
      desc: 'Siz verileri girdikçe arkadaki formül motoru tüm oranları, toplamları ve dengeleri otomatik hesaplar. Dahili kontrol mekanizması hatalı girişleri anında uyarır.'
    },
    {
      step: 3,
      title: 'Karar Çıktılarını ve Raporunuzu Alın',
      desc: `Yönetim panelinden (${template.outputs.slice(0, 3).join(', ')}) gibi kritik çıktıları görün. Dilerseniz tek tuşla yazdırabilir veya PDF olarak paylaşabilirsiniz.`
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
