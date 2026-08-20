#!/usr/bin/env python3
"""
expand_karar_content.py — Adds high-quality, authentic domain paragraphs to each karar page
ensuring all 18 pages have strictly 780-920 words in <main>.
"""

import json
import re

ADDITIONAL_PARAS = {
    "kobi-nakit-akisi-excel": [
        "Finansal dayanıklılık analizi, şirketin beklenmeyen kriz anlarında kaç gün operasyonlarını sürdürebileceğini gösteren en kritik göstergedir. KOBİ'ler genellikle banka kredilerine veya anlık faktoring işlemlerine aşırı bağımlı hale geldiğinde likidite kırılganlığı artar. 13 haftalık nakit akışı modeli, nakit tamponu rezervi kavramını operasyonel süreçlerin merkezine yerleştirir. Her hafta için hedeflenen minimum emniyet nakit seviyesi belirlenir ve gerçekleşen bakiyelerin bu eşiğin altına inmesi durumunda yönetim ekranında otomatik ikaz mekanizmaları devreye girer.",
        "Müşteri segmentasyonu ve ödeme alışkanlıkları analizi, nakit girişlerinin güvenilirliğini artırmak için modele dahil edilmiştir. Büyük kurumsal müşterilerin uzun vadeli ödeme vadeleri ile perakende müşterilerin anlık nakit tahsilatları ayrı ağırlık katsayılarıyla projekte edilir. Geçmiş dönem tahsilat sapmaları modele işlenerek geleceğe dönük nakit giriş tahminlerinin doğruluk oranı en üst düzeye çıkarılır. Bu yaklaşım finans ekibinin nakit tahminlerinde aşırı iyimser olmasını engeller.",
        "Hammadde satın alma ve tedarikçi vade optimizasyonu sayfası, şirket nakdinin en yoğun çıktığı dönemleri tespit ederek toplu alım iskontoları ile vade uzatma alternatifleri arasında matematiksel karşılaştırma yapar. Peşin ödeme iskontosu almanın şirketin nakit maliyetine göre avantaj sağlayıp sağlamadığı anında analiz edilir."
    ],
    "kasa-defteri-excel": [
        "Kasa kontrol süreçlerinde kurumsal güvenliğin sağlanması, çift onay mekanizması ve yetkilendirme limitlerinin uygulanmasıyla mümkündür. Günlük nakit hareketlerinde harcama yapan personel, harcamayı onaylayan birim amiri ve tediye işlemini gerçekleştiren kasa sorumlusu bilgileri her hareket satırında kayıt altına alınır. Bu kayıt disiplini işletme içinde belgesiz veya onaysız harcama yapılmasını tamamen sonlandırır.",
        "Dönemsel harcama eğilimleri ve departman bütçe kontrolleri, kasa defteri özet tablosunda dinamik grafiklerle raporlanır. Şirket genel müdürü veya finans müdürü hangi masraf merkezinin ay içinde nakit kullanımını artırdığını, hangi günlerde kasa çıkışlarının yoğunlaştığını tek bir analitik ekranda inceler. Bütçeyi aşan masraf kalemleri için erken uyarı verilir.",
        "Kasa sayım tutanakları ve bankaya nakit yatırma takvimi sayfası, kasada tutulan fiziki para miktarının sigorta limitlerini aşmasını engeller. Gün sonunda biriken nakdin belirlenen üst sınırı aşması durumunda banka hesabına aktarım talimatı otomatik olarak oluşturulur. Fiziksel güvenlik ve finansal disiplin eş zamanlı sağlanır."
    ],
    "mali-musavir-cari-takip-excel": [
        "Müşteri tahsilat performansının objektif kriterlerle ölçülmesi, şirketin nakit döngüsünü kısaltan en etkili yöntemdir. Cari takip sistemi, müşterilerin fatura ödeme disiplinini DSO (Günlük Satışların Tahsilat Süresi) metriğiyle hesaplar. Sektör ortalamasının üzerinde vade kullanan veya ödemelerini sürekli erteleyen cari hesaplar için risk katsayısı otomatik olarak yükseltilir ve yeni sipariş onaylarında finans onayı zorunlu tutulur.",
        "Hukuki takip ve şüpheli alacak yönetim modülü, vadesi 90 günü aşan problemli alacaklar için avukat bildirim listeleri ve icra takip dosyası özetleri üretir. VUK 323 sayılı kanun uyarınca dava ve icra safhasına gelen alacaklar için ayrılması gereken karşılık tutarı ve bu karşılığın kurumlar vergisi matrahına etkisi hesaplanır.",
        "Toplu cari bakiye mutabakatı ve e-posta bildirim sayfası, ay sonlarında tüm müşterilere standart formatta bakiye mutabakat mektupları gönderilmesini kolaylaştırır. Müşteriden gelen itirazlar ve fatura uyuşmazlıkları mutabakat tablosunda kayıt altına alınarak karşılıklı hesap farkları hızla çözülür."
    ],
    "pos-komisyon-kontrol-excel": [
        "POS operasyonlarında maliyetlerin asgari seviyeye indirilmesi, ciro hacmine göre dinamik POS yönlendirme stratejisi kurmayı gerektirir. Farklı bankaların kart programları (Bonus, World, Maximum, CardFinans, Axess) için sunduğu kampanya komisyon oranları ve puan maliyetleri sisteme tanımlanır. Kasa personeline hangi kart için hangi banka POS cihazının kullanılması gerektiği en düşük maliyetli seçenek olarak önerilir.",
        "Blokeli alacakların iskonto edilmesi ve erken nakde çevrilmesi süreçlerinde bankaların uyguladığı faiz oranları ile piyasa mevduat oranları karşılaştırılır. Şirketin likidite ihtiyacı olmadığında paranın blokede beklemesinin mi yoksa erken çözülerek repoda veya işletme sermayesinde değerlendirilmesinin mi daha kârlı olduğu analiz edilir.",
        "Yıllık POS ciro ve verimlilik karnesi, bankalarla yapılacak yıllık üye işyeri sözleşme yenilemelerinde masaya somut verilerle oturulmasını sağlar. Bankanın şirketten kestiği toplam komisyon ve hizmet bedeli yekünü gösterilerek komisyon indirimi ve aidat muafiyeti talep edilir."
    ],
    "trendyol-pazaryeri-net-kar-excel": [
        "Pazaryeri operasyonlarında net kârlılığı korumanın yolu, ürün listeleme aşamasında tüm gizli maliyet bileşenlerini doğru hesaplamaktan geçer. Ürünün paketleme malzemesi, koli bandı, fatura çıktısı, depolama alanı maliyeti ve platformun kestiği işlem bedelleri birim maliyete dahil edilir. Bu titiz maliyet yapısı sayesinde satıcı hiçbir zaman gerçek maliyetinin altında fiyatlandırma yapmaz.",
        "Kampanya ve flaş indirim kârlılık simülatörü, pazaryerinin sunduğu 'yüzde on indirim yap, komisyonu yüzde iki düşürelim' tekliflerinin şirkete gerçek kâr getirip getirmeyeceğini önceden test eder. Çok satan kampanyalarda artan hacmin kargo barem maliyetini nasıl değiştirdiği ve toplam katkı payını nasıl etkilediği açıkça modellenir.",
        "Kayıp kargo, hasarlı ürün ve müşteri iadesi takip cetveli, kargo firmalarının ve platformun satıcıya yansıttığı haksız kesintileri kayıt altına alır. Tazmin talebi oluşturulması gereken hasarlı kargolar listelenerek platformdan hak edilen tazminatların tahsilatı hızlandırılır."
    ],
    "kdv-iade-dosyasi-excel": [
        "KDV iade taleplerinde vergi inceleme riskini en aza indirmek için yüklenim dağıtım anahtarlarının matematiksel ve mevzuatsal tutarlılığı eksiksiz kurulmalıdır. Üretim işletmelerinde doğrudan ilk madde ve malzeme giderleri ile genel üretim ve amortisman payları ihracat teslimlerine formüllerle paylaştırılır. Vergi müfettişlerinin talep edeceği tüm ara hesap tabloları şeffaf biçimde sunulur.",
        "GİB İnternet Vergi Dairesi sistemine yüklenecek Excel dosyalarındaki mükerrer fatura, geçersiz vergi kimlik numarası, kapalı mükellef ve hatalı KDV oranı uyarıları dosya oluşturulurken anında yakalanır. Hatalı kayıtlar liste dışına alınarak vergi dairesi otomasyon sisteminin dosyayı ilk yüklemede kabul etmesi sağlanır.",
        "Yeminli Mali Müşavir (YMM) KDV iadesi tasdik raporu çalışma kağıtları formatında hazırlanan özet tablolar, denetim sürecini hızlandırır. Mahsuben KDV iade talepleri ile nakden KDV iade talepleri arasındaki teminat mektubu ve YMM raporu gereksinimleri mevzuat kurallarına göre listelenir."
    ],
    "amortisman-yeniden-degerleme-excel": [
        "Enflasyonist dönemlerde şirketlerin aktifinde yer alan binalar, arsalar, fabrika tesisleri ve makineler tarihi maliyet bedelleriyle bilançoda kaldığında özkaynaklar erir ve fiktif kâr üzerinden yüksek kurumlar vergisi ödenir. VUK Mükerrer 298/Ç maddesi kapsamındaki sürekli yeniden değerleme mekanizması, duran varlıkları güncel Yİ-ÜFE oranlarıyla değerleyerek şirketin bilanço gücünü gerçek seviyesine taşır.",
        "Yeniden değerleme sonucu pasifte oluşan değer artış fonu ve amortisman farklarının vergi matrahına etkisi detaylı tablolarda modellenir. Şirketin her yıl elde edeceği ilave amortisman gideri sayesinde ödenecek kurumlar vergisinde yasal ve kalıcı tasarruf sağlanır.",
        "Binek otomobiller için Gelir Vergisi Kanunu 40. maddesindeki amortisman gider kısıtlaması tavanları, KDV ve ÖTV toplam maliyet sınırları sisteme tanımlanmıştır. Binek araç alımlarında gider yazılabilecek azami tutarlar ve kanunen kabul edilmeyen gider (KKEG) ayrımı hatasız hesaplanır."
    ],
    "kidem-ihbar-maliyeti-excel": [
        "İş hukuku ve insan kaynakları yönetiminde işten ayrılış süreçlerinin maliyetlendirilmesi, kıdem tazminatı tavanı ve giydirilmiş ücret hesaplamalarının güncel mevzuata tam uyumlu yapılmasını gerektirir. Yemek yardımı, servis imkanı, yakacak desteği, düzenli primler ve bayram harçlıkları giydirilmiş brüt ücrete yasal oranlarda dahil edilerek olası işçilik alacağı davalarının önüne geçilir.",
        "Yıllık izin karşılığı hesaplama sayfası, personelin hak ettiği ancak kullanmadığı izin günlerini son brüt ücret üzerinden hesaplayarak yasal kesintilerini (SGK primi, gelir vergisi, damga vergisi) ayrıntılı bordro icmalinde gösterir. İhbar önelinin kullandırılması ile peşin ödenmesi durumları nakit çıkışı açısından kıyaslanır.",
        "Dönem sonu kıdem tazminatı karşılığı cetveli, TMS 19 ve vergi mevzuatı standartlarına uygun olarak şirketin tüm aktif personeli için gelecekte oluşacak toplam kıdem yükümlülüğünü aktüeryal varsayımlarla hesaplar. Şirketin finansal tablolarında şeffaf karşılık ayrılmasına olanak tanır."
    ],
    "sgk-tesvik-optimizasyon-excel": [
        "İstihdam teşviklerinin işletme bütçesine sağladığı katkı, personel işe alım süreçlerinde teşvik kriterlerinin önceden sorgulanmasıyla maksimize edilir. İŞKUR kayıt durumu, mezuniyet belgesi, son 6 aydaki sigortalılık geçmişi ve cinsiyet kriterleri analiz edilerek adayın hangi teşvik kanunu kapsamında istihdam edilebileceği mülakat aşamasında netleştirilir.",
        "Teşvikli bordro karşılaştırma sayfası, teşvik uygulanmayan standart bordro maliyeti ile teşvik uygulanan optimize bordro maliyetini personel bazında yan yana listeler. Şirketin her ay elde ettiği net prim tasarrufu toplam bordro gideri içinde yüzde olarak raporlanır.",
        "SGK prim borcu sorgulama ve yapılandırma takip takvimi, teşviklerden yararlanmanın ön şartı olan düzenli prim ödeme ve borçsuzluk durumunu kontrol altında tutar. Teşvik hakkının kaybedilmemesi için son ödeme tarihleri finans ajandasına entegre edilir."
    ],
    "restoran-kafe-maliyet-excel": [
        "Restoran ve kafe işletmeciliğinde kârlılığın temel direği, mutfaktaki reçete gramajlarının ve hammadde firelerinin kuruşu kuruşuna kontrol edilmesidir. Etin pişme firesi, sebzenin ayıklama kaybı, sosların porsiyonlama sapmaları reçete maliyet kartlarında matematiksel olarak modellenir. Menüdeki her yemeğin gerçek porsiyon maliyeti ve brüt kâr marjı şeffaf biçimde hesaplanır.",
        "Menü mühendisliği (Menu Engineering) matrisi, satış adetleri ile tabak kârlılıklarını 4 ana kategoride (Yıldızlar, İş Atları, Bulmacalar, Köpekler) sınıflandırır. Çok satan ama düşük kâr marjlı ürünlerin porsiyon veya fiyat ayarlamaları yapılarak restoranın genel kâr marjı yukarı çekilir.",
        "Haftalık ve aylık fiili hammadde sayım sayfası, mutfak deposundaki fiziki stoklar ile satış adetlerine göre tüketilmesi gereken teorik stokları karşılaştırır. Porsiyon aşımı, mutfak zayiatı veya personel kaynaklı kaçaklar anında tespit edilerek maliyet kontrolü sağlanır."
    ],
    "insaat-hakedis-excel": [
        "İnşaat ve taahhüt projelerinde kârlılığın korunması, şantiye harcamalarının ve taşeron hakedişlerinin kuruşu kuruşuna denetlenmesine bağlıdır. Yapılmayan imalatların hakedişe yazılması veya avans kesintilerinin unutulması projeleri zarara sürükler. Yeşil defter ve metraj cetvelleri doğrudan hakediş icmaline bağlanarak imalat pursantajları şeffafça hesaplanır.",
        "Resmi kamu ihaleleri için TÜİK girdi endeksleri üzerinden hesaplanan fiyat farkı cetveli, hakediş dosyasına otomatik eklenerek idareden hak kaybı yaşanmadan ödeme alınmasını sağlar. Yıllara sari inşaat stopajı, nakit teminat ve nefaset kesintileri mevzuata tam uyumlu olarak düşülür.",
        "Taşeron mutabakat sayfaları, her taşeronun kümülatif hak edişini, ödenen avansları ve kalan kesin teminat bakiyelerini şeffaf şekilde tutarak iş teslimindeki anlaşmazlıkları önler. Şantiye bütçesinin gerçekleşme oranları düzenli raporlanır.",
        "Şantiye malzeme alımları ve makine yakıt giderleri proje maliyet kodlarına bağlanarak bütçelenen metraj maliyetleri ile gerçekleşen faturalar arasındaki sapmalar düzenli raporlanır. Proje nakit akışı ve hak ediş tahsilat takvimi eşleştirilerek malzeme tedarikçilerine yapılacak ödemeler finansman krizine yol açmadan planlanır."
    ],
    "ihale-teklif-sinir-deger-excel": [
        "Kamu ihalelerinde teklif hazırlama süreci matematiksel bir optimizasyon gerektirir. Çok düşük teklif vermek aşırı düşük sorgulamasına takılarak elenme riskini doğururken, çok yüksek teklif vermek ihaleyi kaybettirir. İhaleye Kaç TL Teklif Vermeliyim sistemi, Kamu İhale Kurumu sınır değer formüllerini ve yaklaşık maliyet katsayılarını simüle eder.",
        "Aşırı düşük sorgulama sınırının hemen üstünde kalarak ihaleyi kazanma ihtimalini en üst düzeye çıkaran optimize teklif tutarı belirlenir. İhale kârlılığı ve nakit akışı güvenceye alınır. İhale teminat mektubu komisyonları, sözleşme damga vergisi ve KİK payı gibi zorunlu ihale maliyetleri teklif kârlılık cetveline doğrudan yansıtılır.",
        "Geçmiş ihalelerdeki rakip tenzilat ortalamaları modele girilerek istatistiki kazanma ihtimali hesaplanır. Firmanın nakit akışına en uygun teklif tutarı tespit edilir. İhale komisyonu tarafından açıklanan yaklaşık maliyet ile sınır değer arasındaki hassasiyet aralığı modellenerek teklif dosyasının risk derecesi puanlanır.",
        "Birim fiyat analizleri ve analiz girdi cetvelleri sayfası, aşırı düşük savunması istenmesi durumunda KİK mevzuatına uygun analiz formatında savunma dosyası hazırlanmasına olanak tanır."
    ],
    "stok-devir-nakit-baglanma-excel": [
        "Depoda atıl bekleyen her stok kalemi, işletmenin nakit kasasından çekilmiş ve raflara kilitlenmiş sermayedir. Stok devir hızının düşmesi depolama maliyetlerini artırırken şirketi likidite krizine sokar. Stok ve Nakit Bağlanma Sistemi, ürün bazında satış hızını ve stokta kalma gün süresini hesaplar.",
        "ABC analizi ile cironun yüzde seksenini oluşturan kritik ürünleri öne çıkarırken ölü stokları listeler. Optimum sipariş miktarı ve emniyet stoku seviyeleri formüle edilerek satın alma bütçesi doğru ürünlere tahsis edilir. Depoda bağlanan nakit finansman maliyeti aylık bazda raporlanır.",
        "Tedarikçi teslim süreleri ve sipariş karşılama performansları izlenerek yok satma riskleri ve fazla stok maliyetleri dengelenir. Stok finansman maliyeti aylık ticari kredi faiz oranları üzerinden simüle edilir. Depoda fazla ürün tutmanın şirkete maliyeti açıkça görülür.",
        "Ürün bazında brüt kâr marjı ile stok devir hızının çarpımından elde edilen GMROI (Brüt Kâr Yatırım Getirisi) metriği, hangi ürün grubunun şirket sermayesini en verimli şekilde katladığını ortaya koyar."
    ],
    "sube-karlilik-analizi-excel": [
        "Çok şubeli işletmelerde ciro yüksekliği yanıltıcı olabilir. Yüksek ciro yapan bir şube, yüksek kira ve personel gideri nedeniyle şirketin diğer kârlı mağazalarının ürettiği nakdi tüketiyor olabilir. Şube Kârlılık Hesaplayıcı, her şubenin brüt gelirini, doğrudan işletme giderlerini ve merkezden payına düşen ortak genel yönetim masraflarını ayrıştırır.",
        "Her lokasyonun başabaş noktası net biçimde hesaplanır. Metrekare başına satış verimliliği ve personel başına ciro performansı şubeler arasında karşılaştırılır. Düşük marjlı şubelerin gider yapısı detaylı olarak analiz edilir.",
        "Şube kapatma simülasyonu, zarar eden bir lokasyonun kapatılması durumunda kurtarılacak nakit tutarı ile merkezde kalacak sabit giderleri karşılaştırarak stratejik karar desteği sunar. Hangi şubenin kâra katkı sağladığı, hangisinin ise şirkete zarar verdiği net olarak ortaya konur.",
        "Bölgesel kârlılık haritası, hangi coğrafi bölgede yeni şube açmanın şirket genel marjını yükselteceğini yatırım geri dönüş süresi (ROI) hesaplarıyla modeller. Mağaza kârlılığı güvenceye alınır."
    ],
    "ttk-376-sermaye-kaybi-excel": [
        "Türk Ticaret Kanunu 376. maddesi, şirket yönetim kurullarına sermaye kaybı ve borca batıklık durumunda ağır yasal sorumluluklar yükler. Sermaye ve yedek akçelerin karşılıksız kalması halinde acil genel kurul çağrısı zorunludur. TTK 376 Cetveli, şirketin bilanço kalemlerini yasal tebliğ standartlarıyla analiz eder.",
        "Yabancı para borçlardan doğan kur farkı zararlarının hesaplama dışı bırakılması opsiyonunu mevzuata uygun şekilde uygular. Sermaye koruma oranını kuruşu kuruşuna belirler ve şirketin borca batıklıktan çıkması için gereken asgari sermaye artırımı veya sermaye tamamlama fonu tutarını hesaplar.",
        "Ortaklar kuruluna sunulacak resmi durum tespit raporu ve iyileştirme tedbirleri tablosu, yasal denetimlerde şirket yöneticilerini hukuki güvence altına alır. Ara dönem bilançoları üzerinden yapılan projeksiyonlar, yıl sonu kapanışında şirketin hangi hukuki statüde yer alacağını önceden gösterir.",
        "Bağımsız denetim standartlarına uygun borca batıklık ara bilançosu tablosu, aktiflerin olası tasfiye değerleri üzerinden şirketin borçlarını karşılama gücünü matematiksel olarak belgeler."
    ],
    "doviz-acik-pozisyon-kur-riski-excel": [
        "Döviz cinsinden borcu veya hammadde ithalatı olan şirketler için kur dalgalanmaları en büyük bilanço riskidir. Döviz gelirleri ile döviz borçları arasındaki vade uyuşmazlığı kur şoklarında faaliyet kârını tamamen silebilir. Döviz Açık Pozisyonu ve Kur Riski modeli, şirketin dövizli nakit, alacak, ticari borç ve banka kredisi varlıklarını tek bir tabloda toplar.",
        "Net yabancı para pozisyonunu ve kur duyarlılığını hesaplar. Dolar ve Euro için yüzde on, yirmi beş ve elli oranındaki kur artış senaryolarında şirketin maruz kalacağı kambiyo zararını ve özkaynak erimesini stres testiyle simüle eder.",
        "İhracat gelirlerinin döviz borçlarını karşılama oranı ve forward türev araçlarının bilanço koruma etkisi matematiksel modellerle test edilir. Para birimi bazında net pozisyon ayrı ayrı izlenerek çapraz kur parite riskleri de analiz edilir.",
        "Türev finansal araçlar (Forward, Opsiyon) kullanımının bilanço kur riskini ne oranda sınırlandıracağı simülasyon sayfasında maliyet ve fayda ekseninde modellenir. Şirketin kur şoklarına dayanıklılığı artırılır."
    ],
    "logo-erp-cari-yaslandirma-excel": [
        "Logo ve kurumsal ERP muhasebe programları yoğun hareket kaydı tutabilir ancak şirket yöneticilerine doğrudan karar aldıracak dinamik yaşlandırma özetlerini her zaman pratik olarak sunamaz. Logo Uyumlu Cari Yaşlandırma Motoru, muhasebe yazılımından alınan ham mizan ve muavin hareket dökümlerini işler.",
        "Borç ve alacak kapatmalarını FIFO mantığıyla yaşlandırarak dinamik tahsilat paneline dönüştürür. Müşteri bazlı ortalama tahsilat vadesi ve vade sapmaları analiz edilerek hangi müşterilerin sözleşme vadesini aştığı net biçimde listelenir.",
        "Kritik vade eşiğini aşan müşteriler için otomatik tahsilat aksiyon planı ve hukuk takip özeti oluşturulur. Finans biriminin raporlama süresi dakikalara indirilir.",
        "Satış temsilcilerinin prim hesaplamalarında tahsilat vadelerini kriter alan performans matrisi, satış ekibinin vadesi geçen alacakları toplamasını teşvik eder."
    ]
}


# Update scripts/build_all_karar_pages.py
with open('veri/urunler.json') as f:
    products = {p['slug']: p for p in json.load(f)}
