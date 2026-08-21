export interface KararPageSpec {
  slug: string;
  title: string;
  description: string;
  primaryQuery: string;
  primaryProductSlug?: string;
  alternativeProductSlugs: string[];
  intro: string;
  situations: Array<{ situation: string; productSlug: string; reason: string }>;
  notFor: string[];
  faq: Array<{ q: string; a: string }>;
}

type BaseSpec = Omit<KararPageSpec, 'notFor' | 'faq'> & { notFor?: string[]; faq?: KararPageSpec['faq'] };

function makePage(spec: BaseSpec): KararPageSpec {
  const primary = spec.primaryProductSlug ? 'önerilen ürünün' : 'seçim rehberinin';
  return {
    ...spec,
    notFor: spec.notFor ?? [
      'ERP veya bulut tabanlı çok kullanıcılı operasyon yazılımı arıyorsanız.',
      'Uzman görüşü gerektiren hukuki veya vergisel sonucu yalnız bir Excel dosyasına bırakmak istiyorsanız.',
      'Ürün sayfasındaki kapsam ve uyumluluk bilgisini incelemeden satın alma kararı vermek istiyorsanız.',
    ],
    faq: spec.faq ?? [
      { q: 'Bu karar rehberi ne işe yarar?', a: `İşletme ihtiyacını ürün kapsamıyla eşleştirir ve ${primary} hangi durumda daha uygun olabileceğini açıklar. Kesin kapsam için ürün detay sayfası esas alınır.` },
      { q: 'Fiyat bilgisi güncel mi?', a: 'Karar sayfasındaki fiyatlar ürün veri kaynağından build sırasında alınır; karar metnine elle yazılmaz.' },
      { q: 'Satın almadan önce inceleyebilir miyim?', a: 'Uygun ürünlerde demo ve gerçek ekran örnekleri ürün detay sayfasında sunulur. Satın almadan önce kapsamı ve çıktıları kontrol edin.' },
      { q: 'Bu sayfa uzman görüşünün yerine geçer mi?', a: 'Hayır. Sayfa ürün seçim rehberidir. Vergi, hukuk, ihale veya mevzuat sonucu doğuran konularda güncel düzenleme ve uzman değerlendirmesi ayrıca gerekir.' },
      { q: 'Ürün sayfası ile karar sayfası farklıysa hangisi esas?', a: 'Ürün kapsamı, fiyat, uyumluluk ve satın alma bilgisi bakımından ürün detay sayfası esas kaynaktır.' },
    ],
  };
}

export const kararPages: KararPageSpec[] = [
  makePage({
    slug: 'hangi-excel-sistemini-almaliyim',
    title: 'Hangi Excel sistemini almalıyım?',
    description: 'Nakit, cari, stok, vergi, personel ve kârlılık ihtiyacınıza göre hangi Excel sisteminden başlamanız gerektiğini karşılaştırın.',
    primaryQuery: 'hangi Excel sistemini almalıyım',
    alternativeProductSlugs: ['13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi','akilli-kasa-defteri-ve-nakit-kontrol-sistemi','aylik-patron-finans-paneli'],
    intro: 'Tek bir Excel sistemi herkes için doğru değildir. Önce çözmek istediğiniz işletme problemini belirleyin; sonra o probleme ait girdi, kontrol ve karar çıktısını sağlayan ürünü seçin.',
    situations: [
      { situation: 'Önümüzdeki haftalarda nakit açığını görmek', productSlug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi', reason: 'Haftalık giriş, çıkış ve kümülatif bakiyeyi birlikte gösterir.' },
      { situation: 'Günlük kasa hareketlerini kontrol etmek', productSlug: 'akilli-kasa-defteri-ve-nakit-kontrol-sistemi', reason: 'Günlük hareket ve sayım farkına odaklanır.' },
      { situation: 'Aylık yönetim özeti hazırlamak', productSlug: 'aylik-patron-finans-paneli', reason: 'Temel finans göstergelerini tek yönetim görünümünde toplar.' },
    ],
  }),
  makePage({
    slug: 'kobi-nakit-akisi-excel', title: "KOBİ için nakit akışı Excel'i: hangisini seçmelisiniz?", description: 'KOBİ nakit akışını haftalık planlamak için 13 haftalık model ile günlük kasa kontrolü arasındaki farkı görün.', primaryQuery: "KOBİ için nakit akışı Excel'i", primaryProductSlug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi', alternativeProductSlugs: ['akilli-kasa-defteri-ve-nakit-kontrol-sistemi','kobi-finans-yonetim-paketi'], intro: 'Gelecek haftalardaki tahsilat ve ödemeleri birlikte planlamak istiyorsanız haftalık nakit akışı modeli gerekir. Yalnız günlük kasayı izlemek istiyorsanız daha basit kasa sistemi yeterlidir.', situations: [
      { situation: '13 haftalık tahsilat ve ödeme planı', productSlug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi', reason: 'Haftalık projeksiyon ve kritik bakiye görünümü sağlar.' },
      { situation: 'Günlük kasa kontrolü', productSlug: 'akilli-kasa-defteri-ve-nakit-kontrol-sistemi', reason: 'Fiili kasa hareketlerine odaklanır.' },
      { situation: 'Daha geniş finans görünümü', productSlug: 'kobi-finans-yonetim-paketi', reason: 'Birden çok finans sürecini aynı çalışma yapısında toplar.' },
    ] }),
  makePage({
    slug: 'kasa-defteri-excel', title: 'Günlük kasa defteri Excel: hangi sistemi seçmelisiniz?', description: 'Günlük kasa hareketi, kapanış bakiyesi ve sayım farkını izlemek için uygun Excel sistemini karşılaştırın.', primaryQuery: 'günlük kasa defteri Excel', primaryProductSlug: 'akilli-kasa-defteri-ve-nakit-kontrol-sistemi', alternativeProductSlugs: ['13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi'], intro: 'Günlük gelir ve giderlerin fiili kasa bakiyesiyle uyuşup uyuşmadığını görmek için kasa odaklı bir sistem gerekir. Gelecek haftaları tahmin etmek ayrı bir problemdir.', situations: [
      { situation: 'Günlük açılış-kapanış ve sayım farkı', productSlug: 'akilli-kasa-defteri-ve-nakit-kontrol-sistemi', reason: 'Kasa hareketi ve fark kontrolüne odaklanır.' },
      { situation: 'Gelecek ödeme baskısını görmek', productSlug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi', reason: 'Günlük kayıt yerine haftalık projeksiyon sunar.' },
    ] }),
  makePage({
    slug: 'mali-musavir-cari-takip-excel', title: 'Cari hesap ve tahsilat takibi için hangi Excel sistemi?', description: 'Cari bakiye, vade ve müşteri tahsilat riskini izlemek için uygun Excel sistemini karşılaştırın.', primaryQuery: 'cari hesap takip Excel', primaryProductSlug: 'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi', alternativeProductSlugs: ['cari-ba-bs-toplu-mutabakat','cek-senet-ve-vade-risk-sistemi'], intro: 'Cari takipte amaç yalnız bakiye görmek değil; vade, gecikme ve tahsilat önceliğini birlikte değerlendirmektir. Mutabakat ve çek-senet takibi farklı alt problemlerdir.', situations: [
      { situation: 'Müşteri bazlı bakiye ve gecikme', productSlug: 'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi', reason: 'Tahsilat ve müşteri riskine odaklanır.' },
      { situation: 'Cari ile Ba-Bs mutabakatı', productSlug: 'cari-ba-bs-toplu-mutabakat', reason: 'Mutabakat farklarını ayırır.' },
      { situation: 'Çek-senet vade riski', productSlug: 'cek-senet-ve-vade-risk-sistemi', reason: 'Vadeli kıymetli evrak takibine odaklanır.' },
    ] }),
  makePage({
    slug: 'pos-komisyon-kontrol-excel', title: 'POS komisyon kontrolü için hangi Excel sistemi?', description: 'POS brüt satış, komisyon, iade ve net tahsilatı karşılaştırmak için uygun Excel sistemini görün.', primaryQuery: 'POS komisyon kontrol Excel', primaryProductSlug: 'pos-komisyon-ve-net-tahsilat-kontrol-sistemi', alternativeProductSlugs: ['gunluk-gelir-gider-ve-gercek-karlilik-sistemi'], intro: 'POS cirosu bankaya geçen net tutarla aynı değildir. Komisyon, iade ve diğer kesintiler ayrı izlendiğinde kanalın gerçek tahsilat maliyeti görünür hale gelir.', situations: [
      { situation: 'POS kanal bazlı net tahsilat', productSlug: 'pos-komisyon-ve-net-tahsilat-kontrol-sistemi', reason: 'Brüt satıştan net tahsilata geçişi gösterir.' },
      { situation: 'Günlük genel kârlılık', productSlug: 'gunluk-gelir-gider-ve-gercek-karlilik-sistemi', reason: 'POS dışındaki gelir-gider kalemlerini de kapsar.' },
    ] }),
  makePage({
    slug: 'trendyol-pazaryeri-net-kar-excel', title: 'Pazaryeri net kârı için hangi Excel sistemi?', description: 'Komisyon, kargo, reklam, iade ve hakediş sonrası pazaryeri net kârını izlemek için uygun Excel sistemini karşılaştırın.', primaryQuery: 'Trendyol komisyon sonrası net kâr Excel', primaryProductSlug: 'trendyol-komisyon-sonrasi-net-kar', alternativeProductSlugs: ['pazaryeri-net-kar-ve-eksik-hakedis-yakalayici'], intro: 'Pazaryerinde satış fiyatı tek başına kârlılığı göstermez. Komisyon, kargo, reklam, kampanya ve iadeler gerçek marjı etkiler; hakediş kontrolü ise beklenen ödeme ile gerçekleşeni karşılaştırır.', situations: [
      { situation: 'Trendyol satışında net kâr', productSlug: 'trendyol-komisyon-sonrasi-net-kar', reason: 'Kanalın temel kesintileri sonrası karar çıktısı verir.' },
      { situation: 'Beklenen ve gerçek hakedişi karşılaştırmak', productSlug: 'pazaryeri-net-kar-ve-eksik-hakedis-yakalayici', reason: 'Eksik hakediş farkına odaklanır.' },
    ] }),
  makePage({
    slug: 'kdv-iade-dosyasi-excel', title: 'KDV iade listesi hazırlamak için hangi Excel sistemi?', description: 'KDV iade listeleri, azami alacak ve tevkifat mahsup süreçleri için uygun Excel sistemlerini karşılaştırın.', primaryQuery: 'KDV iade listesi hazırlama Excel', primaryProductSlug: 'kdv-iade-listesi-robotu-gib7', alternativeProductSlugs: ['kdv-iadesi-azami-alacak-hesabi-dosya-hazirlayici','kdv-tevkifat-mahsup-iade-listesi'], intro: 'KDV iade sürecinde liste hazırlama, azami iade alacağını hesaplama ve tevkifat/mahsup ayrımı farklı işlerdir. Doğru ürün, ihtiyaç duyduğunuz çıktıya göre seçilmelidir.', situations: [
      { situation: 'GİB iade listelerini hazırlamak', productSlug: 'kdv-iade-listesi-robotu-gib7', reason: 'Liste üretimine odaklanır.' },
      { situation: 'Azami iade alacağını hesaplamak', productSlug: 'kdv-iadesi-azami-alacak-hesabi-dosya-hazirlayici', reason: 'İade tutarı ve dosya hazırlığına odaklanır.' },
      { situation: 'Tevkifat mahsup/iade ayrımı', productSlug: 'kdv-tevkifat-mahsup-iade-listesi', reason: 'Tevkifat ve mahsup kararını destekler.' },
    ] }),
  makePage({
    slug: 'amortisman-yeniden-degerleme-excel', title: 'Amortisman ve yeniden değerleme için hangi Excel sistemi?', description: 'Amortisman, yeniden değerleme ve sabit kıymet satış zamanlaması için uygun Excel sistemlerini karşılaştırın.', primaryQuery: 'amortisman hesaplama Excel', primaryProductSlug: 'amortisman-2026-yeniden-degerleme', alternativeProductSlugs: ['amortisman-ve-sabit-kiymet-satis-zamanlama-stratejisti','yeniden-degerleme-yapmali-miyim-vergi-tasarruf-analizi'], intro: 'Amortisman hesabı, yeniden değerleme kararı ve sabit kıymetin hangi dönemde satılacağı aynı problem değildir. Kullanacağınız sistem ihtiyaç duyduğunuz karar çıktısına göre seçilmelidir.', situations: [
      { situation: 'Amortisman ve yeniden değerleme hesabı', productSlug: 'amortisman-2026-yeniden-degerleme', reason: 'Temel amortisman ve yeniden değerleme etkisini birlikte ele alır.' },
      { situation: 'Sabit kıymet satış zamanlaması', productSlug: 'amortisman-ve-sabit-kiymet-satis-zamanlama-stratejisti', reason: 'Dönemler arası satış etkisini karşılaştırır.' },
      { situation: 'Yeniden değerleme faydasını ölçmek', productSlug: 'yeniden-degerleme-yapmali-miyim-vergi-tasarruf-analizi', reason: 'Vergi tasarrufu ve maliyet etkisine odaklanır.' },
    ] }),
  makePage({
    slug: 'kidem-ihbar-maliyeti-excel', title: 'Kıdem ve ihbar maliyeti için hangi Excel sistemi?', description: 'Personel çıkış maliyeti ve fazla mesai riskini ayrı değerlendirerek uygun Excel sistemini seçin.', primaryQuery: 'kıdem ihbar hesaplama Excel', primaryProductSlug: 'kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici', alternativeProductSlugs: ['fazla-mesai-ve-isci-dava-riski-tespit-dosyasi'], intro: 'Personel çıkışında kıdem ve ihbar yükü ile fazla mesai alacağı aynı hesap değildir. Çıkış maliyetini planlamak için tazminat odaklı sistem; çalışma süresi uyuşmazlıkları için farklı sistem gerekir.', situations: [
      { situation: 'Personel çıkış maliyetini planlamak', productSlug: 'kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici', reason: 'Kıdem ve ihbar yüküne odaklanır.' },
      { situation: 'Fazla mesai kaynaklı riskleri görmek', productSlug: 'fazla-mesai-ve-isci-dava-riski-tespit-dosyasi', reason: 'Çalışma süresi ve olası alacak riskine odaklanır.' },
    ] }),
  makePage({
    slug: 'sgk-tesvik-optimizasyon-excel', title: 'SGK teşvik analizi için hangi Excel sistemi?', description: 'Kaçırılan SGK teşviklerini ve teşvikli bordro seçeneklerini değerlendirmek için uygun Excel sistemini karşılaştırın.', primaryQuery: 'SGK teşvik hesaplama Excel', primaryProductSlug: 'kacirilan-sgk-tesvikleri-ve-gercek-iscilik-maliyeti-analizi', alternativeProductSlugs: ['tesvikli-bordro-optimizasyon','tesvikli-bordro-avantajli-tesvik'], intro: 'SGK teşvik analizinde kaçırılan avantajı görmek ve çalışan bazında hangi teşvik setinin daha avantajlı olduğunu değerlendirmek farklı sorulardır. Ürün seçimi bu ayrıma göre yapılmalıdır.', situations: [
      { situation: 'Kaçırılan teşvik ve gerçek işçilik maliyeti', productSlug: 'kacirilan-sgk-tesvikleri-ve-gercek-iscilik-maliyeti-analizi', reason: 'Teşvik etkisini toplam işçilik maliyetiyle birlikte gösterir.' },
      { situation: 'Çoklu teşvik senaryosu', productSlug: 'tesvikli-bordro-optimizasyon', reason: 'Farklı teşvik kombinasyonlarını karşılaştırmaya odaklanır.' },
      { situation: 'Çalışan için avantajlı teşvik seçimi', productSlug: 'tesvikli-bordro-avantajli-tesvik', reason: 'Kişi bazlı seçim problemine odaklanır.' },
    ] }),
  makePage({
    slug: 'restoran-kafe-maliyet-excel', title: 'Restoran ve kafe maliyeti için hangi Excel sistemi?', description: 'Reçete maliyeti, fire ve mutfak kayıp-kaçak analizi için uygun Excel sistemlerini karşılaştırın.', primaryQuery: 'restoran reçete maliyet Excel', primaryProductSlug: 'restoran-recete-maliyet-fire', alternativeProductSlugs: ['mutfak-kayip-kacak-hesaplayici'], intro: 'Restoran maliyetinde reçete bazlı teorik maliyet ile fiili tüketim farkı iki ayrı kontroldür. Menü fiyatlaması için reçete/fire hesabı; mutfak sapmasını görmek için kayıp-kaçak analizi gerekir.', situations: [
      { situation: 'Porsiyon ve reçete maliyetini hesaplamak', productSlug: 'restoran-recete-maliyet-fire', reason: 'Reçete ve fire maliyetine odaklanır.' },
      { situation: 'Teorik tüketim ile fiili tüketimi karşılaştırmak', productSlug: 'mutfak-kayip-kacak-hesaplayici', reason: 'Kayıp ve kaçak farkını görünür kılar.' },
    ] }),
  makePage({
    slug: 'insaat-hakedis-excel', title: 'İnşaat hakediş takibi için hangi Excel sistemi?', description: 'Şantiye maliyeti, hakediş fiyat farkı ve taşeron mutabakatı için uygun Excel sistemlerini karşılaştırın.', primaryQuery: 'inşaat hakediş takip Excel', primaryProductSlug: 'insaat-hakedis-santiye-maliyet', alternativeProductSlugs: ['hakedis-fiyat-farki-hak-kaybi-cetveli','taseron-hakedis-kesinti-mutabakati'], intro: 'İnşaatta hakediş takibi, fiyat farkı hesabı ve taşeron mutabakatı aynı süreç içinde görünse de farklı kontrol noktalarıdır. Ana ihtiyacınıza göre ilgili sistemi seçmek gerekir.', situations: [
      { situation: 'Hakediş ve şantiye maliyetini birlikte izlemek', productSlug: 'insaat-hakedis-santiye-maliyet', reason: 'Hakediş ve maliyet sapmasına odaklanır.' },
      { situation: 'Fiyat farkı ve hak kaybını hesaplamak', productSlug: 'hakedis-fiyat-farki-hak-kaybi-cetveli', reason: 'Endeks/fiyat farkı hesabına odaklanır.' },
      { situation: 'Taşeron hakediş-kesinti mutabakatı', productSlug: 'taseron-hakedis-kesinti-mutabakati', reason: 'Alt yüklenici mutabakatına odaklanır.' },
    ] }),
  makePage({
    slug: 'ihale-teklif-sinir-deger-excel', title: 'İhale teklif ve sınır değer hesabı için hangi Excel sistemi?', description: 'Sınır değer, teklif aralığı ve aşırı düşük teklif savunması için uygun Excel sistemlerini karşılaştırın.', primaryQuery: 'ihaleye kaç TL teklif vermeliyim', primaryProductSlug: 'ihaleye-kac-tl-teklif-vermeliyim', alternativeProductSlugs: ['asiri-dusuk-teklif-savunma-robotu'], intro: 'Teklif fiyatını belirlemek ile aşırı düşük teklif sorgusuna açıklama hazırlamak farklı aşamalardır. İlk aşamada sınır değer ve teklif aralığı; ikinci aşamada açıklama ve savunma yapısı önem kazanır.', situations: [
      { situation: 'Teklif aralığını ve sınır değeri hesaplamak', productSlug: 'ihaleye-kac-tl-teklif-vermeliyim', reason: 'Teklif öncesi karar problemine odaklanır.' },
      { situation: 'Aşırı düşük teklif açıklamasını hazırlamak', productSlug: 'asiri-dusuk-teklif-savunma-robotu', reason: 'Sorgu sonrası savunma yapısına odaklanır.' },
    ], faq: [
      { q: 'Sınır değer hesabı teklif kararını tek başına belirler mi?', a: 'Hayır. Maliyet, rekabet, şartname ve riskler ayrıca değerlendirilmelidir.' },
      { q: 'Aşırı düşük teklif savunması aynı ürün mü?', a: 'Hayır. Savunma, teklif sonrası farklı bir süreçtir ve ayrı ürünle ele alınır.' },
      { q: 'Resmi ihale sonucunu önceden belirler mi?', a: 'Hayır. Araç karar desteği sağlar; idarenin değerlendirmesi ve hukuki sonuç ayrı süreçlerdir.' },
      { q: 'Rakip teklifleri girmek gerekir mi?', a: 'Ürünün hesap yapısında istenen veri alanlarını ürün detay sayfasından kontrol edin.' },
      { q: 'Ürün fiyatı nasıl belirleniyor?', a: 'Karar sayfasındaki ürün fiyatı mevcut ürün veri kaynağından alınır.' },
    ] }),
  makePage({
    slug: 'stok-devir-nakit-baglanma-excel', title: 'Stok devir ve stokta bağlı nakit için hangi Excel sistemi?', description: 'Stok devir hızı, ortalama maliyet ve stokta bağlı nakdi görmek için uygun Excel sistemini karşılaştırın.', primaryQuery: 'stok devir nakit bağlanma Excel', primaryProductSlug: 'stok-satis-ve-nakit-baglanma-sistemi', alternativeProductSlugs: ['ithalat-depo-teslim-rafa-gelen-net-birim-maliyet'], intro: 'Stok miktarını bilmek tek başına yeterli değildir. Devir hızı ve stokta bağlı nakit işletmenin finansman yükünü gösterir. İthal ürünlerde rafa gelen birim maliyet ayrı bir hesaplama problemidir.', situations: [
      { situation: 'Stok devir hızı ve bağlı nakit', productSlug: 'stok-satis-ve-nakit-baglanma-sistemi', reason: 'Stok hareketi ile nakit bağlanmasını birlikte gösterir.' },
      { situation: 'İthal ürünün depo teslim birim maliyeti', productSlug: 'ithalat-depo-teslim-rafa-gelen-net-birim-maliyet', reason: 'İthalat maliyet bileşenlerine odaklanır.' },
    ] }),
  makePage({
    slug: 'sube-karlilik-analizi-excel', title: 'Şube kârlılık analizi için hangi Excel sistemi?', description: 'Şube bazlı gelir, gider, nakit ve kârlılık kararını değerlendirmek için uygun Excel sistemini karşılaştırın.', primaryQuery: 'şube kârlılık analizi Excel', primaryProductSlug: 'sube-karlilik-ve-nakit-hesaplayici', alternativeProductSlugs: ['aylik-patron-finans-paneli','proje-ve-is-bazinda-gercek-karlilik-sistemi'], intro: 'Şubenin cirosu yüksek olsa bile gider, ortak maliyet ve nakit etkisi nedeniyle gerçek kârlılık zayıf olabilir. Şube kararı için şube bazlı sonuç; genel yönetim ve proje kârlılığı için farklı araçlar gerekir.', situations: [
      { situation: 'Şube bazında kâr ve nakit', productSlug: 'sube-karlilik-ve-nakit-hesaplayici', reason: 'Şube kararına odaklanır.' },
      { situation: 'Şirket genel yönetim özeti', productSlug: 'aylik-patron-finans-paneli', reason: 'Şube dışındaki genel finans göstergelerini toplar.' },
      { situation: 'Proje veya iş bazlı kârlılık', productSlug: 'proje-ve-is-bazinda-gercek-karlilik-sistemi', reason: 'Şube yerine iş/proje boyutuna odaklanır.' },
    ] }),
  makePage({
    slug: 'ttk-376-sermaye-kaybi-excel', title: 'TTK 376 sermaye kaybı için hangi Excel sistemi?', description: 'Öz kaynak erimesi, sermaye kaybı ve borca batıklık göstergelerini değerlendirmek için uygun Excel sistemini görün.', primaryQuery: 'TTK 376 sermaye kaybı hesaplama Excel', primaryProductSlug: 'sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli', alternativeProductSlugs: ['konkordato-nakit-akis-on-projesi'], intro: 'TTK 376 kapsamında sermaye kaybı ve öz kaynak durumu, nakit sıkışıklığı veya konkordato projeksiyonuyla aynı değildir. Önce bilanço göstergelerinin hangi eşiğe geldiğini görmek gerekir.', situations: [
      { situation: 'Öz kaynak ve sermaye kaybı eşiğini izlemek', productSlug: 'sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli', reason: 'TTK 376 göstergelerine odaklanır.' },
      { situation: 'Konkordato için nakit akış ön projesi', productSlug: 'konkordato-nakit-akis-on-projesi', reason: 'Farklı bir finansal ve hukuki süreç için projeksiyon sunar.' },
    ] }),
  makePage({
    slug: 'doviz-acik-pozisyon-kur-riski-excel', title: 'Döviz açık pozisyonu ve kur riski için hangi Excel sistemi?', description: 'Döviz varlık ve borçlarını netleştirip kur şoku etkisini görmek için uygun Excel sistemini karşılaştırın.', primaryQuery: 'döviz açık pozisyon Excel', primaryProductSlug: 'doviz-acik-pozisyonu-ve-kur-riski-stres-testi', alternativeProductSlugs: ['kkeg-ve-finansman-gider-kisitlamasi-vergi-savunma-seti'], intro: 'Kur riski için önce döviz cinsinden varlık ve yükümlülüklerin net pozisyonu çıkarılmalıdır. Finansman gider kısıtlaması ise ayrı bir vergi problemidir.', situations: [
      { situation: 'Net döviz pozisyonu ve kur şoku', productSlug: 'doviz-acik-pozisyonu-ve-kur-riski-stres-testi', reason: 'Döviz açık pozisyonuna odaklanır.' },
      { situation: 'Finansman gider kısıtlaması', productSlug: 'kkeg-ve-finansman-gider-kisitlamasi-vergi-savunma-seti', reason: 'Vergisel finansman gideri problemine odaklanır.' },
    ] }),
  makePage({
    slug: 'logo-erp-cari-yaslandirma-excel', title: 'Logo/ERP cari yaşlandırma için hangi Excel sistemi?', description: 'Logo veya ERP cari hareket çıktısını yaşlandırma ve tahsilat kararına dönüştürmek için uygun Excel sistemini değerlendirin.', primaryQuery: 'Logo cari yaşlandırma raporu Excel', primaryProductSlug: 'logo-sql-cari-yaslandirma-tahsilat-karar-motoru', alternativeProductSlugs: ['cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi'], intro: 'ERP çıktısı tek başına tahsilat önceliğini göstermez. Cari hareketlerin yaşlandırılması, açık kalemlerin sınıflandırılması ve aksiyon sırasına dönüştürülmesi ayrı bir analiz katmanıdır. Logo adı burada uyumluluk bağlamında kullanılır; marka sahibiyle resmi ilişki ima edilmez.', situations: [
      { situation: 'Logo/SQL cari hareketinden yaşlandırma üretmek', productSlug: 'logo-sql-cari-yaslandirma-tahsilat-karar-motoru', reason: 'ERP çıktısından tahsilat odaklı yaşlandırma üretir.' },
      { situation: 'Genel cari ve müşteri risk takibi', productSlug: 'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi', reason: 'Belirli ERP çıktısına bağlı olmadan cari risk yönetimine odaklanır.' },
    ], notFor: ['Logo ERP eklentisi veya resmi çözüm ortağı ürünü arıyorsanız.', 'Mac uyumluluğu zorunluysa ürün sayfasındaki platform bilgisini kontrol etmeden satın almayın.', 'ERP veritabanına doğrudan yazma yapan entegrasyon arıyorsanız.'] }),
];

export const kararPageBySlug = new Map(kararPages.map((page) => [page.slug, page]));
