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
      question: `${t.name} hangi girdilerle hangi karar çıktısını üretir?`,
      answer: `${t.name}, ${intent.value} amacıyla ${inputText} verilerini işleyerek ${outputText} raporlarını oluşturur.`,
    },
    {
      question: `${intent.query} ihtiyacında ${t.name} kimlere uygundur?`,
      answer: `${t.name}; ${intent.audience} için ${t.minExcelVersion} (${t.macCompatible ? 'Windows ve Mac' : 'Windows'}) üzerinde makrosuz ${outputText} takibi sağlar.`,
    },
  ];
}
