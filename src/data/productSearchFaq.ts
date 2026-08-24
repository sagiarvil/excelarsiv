type TemplateLike = {
  name: string;
  category: string;
  inputs: string[];
  outputs: string[];
  minExcelVersion: string;
  macCompatible: boolean;
};

type FaqItem = { question: string; answer: string };

const categoryIntent: Record<string, { query: string; audience: string; value: string }> = {
  'finansal-analiz': { query: 'finansal analiz Excel', audience: 'finans yöneticileri, mali müşavirler ve işletme sahipleri', value: 'finansal veriyi karar göstergelerine dönüştürmek' },
  'nakit-akisi': { query: 'nakit akışı Excel', audience: 'finans yöneticileri, muhasebe ekipleri ve işletme sahipleri', value: 'nakit açığını, ödeme baskısını ve kritik dönemleri önceden görmek' },
  'muhasebe-ve-vergi': { query: 'muhasebe Excel şablonu', audience: 'muhasebeciler, mali müşavirler ve finans ekipleri', value: 'kayıt, kontrol ve mevzuata bağlı hesaplamaları standartlaştırmak' },
  'butce-ve-planlama': { query: 'bütçe Excel şablonu', audience: 'bütçe, finans ve yönetim ekipleri', value: 'planlanan ve gerçekleşen sonuçları karşılaştırıp sapmayı erken görmek' },
  'stok-ve-uretim': { query: 'stok takip Excel', audience: 'stok, üretim, satın alma ve finans ekipleri', value: 'stok miktarı, maliyet ve sermaye bağlanmasını birlikte izlemek' },
  'satis-ve-fiyatlama': { query: 'satış kârlılık Excel', audience: 'satış, finans ve işletme yöneticileri', value: 'fiyat, marj ve gerçek kârlılığı aynı karar ekranında görmek' },
  'personel-ve-bordro': { query: 'personel maliyeti Excel', audience: 'muhasebe, insan kaynakları ve finans ekipleri', value: 'personel maliyetini ve değişim etkisini kontrollü hesaplamak' },
};

export function getProductSearchFaq(t: TemplateLike): FaqItem[] {
  const intent = categoryIntent[t.category] ?? categoryIntent['finansal-analiz'];
  const inputText = t.inputs.slice(0, 3).join(', ');
  const outputText = t.outputs.slice(0, 3).join(', ');

  return [
    {
      question: `${t.name} ne işe yarar?`,
      answer: `${t.name}, ${intent.value} için hazırlanmış bir Excel sistemidir. Temel girdiler ${inputText}; temel çıktılar ise ${outputText} alanlarında toplanır.`,
    },
    {
      question: `${t.name} muhasebeciler ve finansçılar için uygun mu?`,
      answer: `Evet. Bu sistem özellikle ${intent.audience} için tasarlanmıştır. Tekrarlanan manuel hesaplamaları azaltır, aynı veriden tutarlı çıktı üretir ve kontrol edilmesi gereken alanları görünür hale getirir.`,
    },
    {
      question: `${intent.query} arayan biri neden ${t.name} kullanmalı?`,
      answer: `Basit bir kayıt tablosundan farklı olarak ${t.name}, yalnız veri saklamayı değil karar üretmeyi hedefler. ${outputText} gibi sonuçları aynı iş akışında sunarak kullanıcıyı ayrı dosyalar ve manuel kontrol adımları arasında dolaştırmaz.`,
    },
    {
      question: `${t.name} için hangi Excel sürümü gerekir?`,
      answer: `Minimum gereksinim ${t.minExcelVersion} sürümüdür. ${t.macCompatible ? 'Windows ve Mac Excel ile kullanılabilir.' : 'Windows Excel kullanımı esas alınmıştır.'} Ürünün teknik künyesindeki sürüm ve uyumluluk bilgileri satın alma öncesinde kontrol edilmelidir.`,
    },
  ];
}
