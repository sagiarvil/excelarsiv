export interface ProductSeoEntry {
  title: string;
  description: string;
  primaryQuery: string;
  guideSlug?: string;
  guideLinkLabel?: string;
}

export const productSeo: Record<string, ProductSeoEntry> = {
  '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi': { title:'13 Haftalık Nakit Akışı Excel | Ödeme Planlama', description:'13 haftalık nakit akış tablosunda tahsilat ve ödemeleri haftalık planlayın; nakit açığını, ödeme baskısını ve kasa dengesini önceden görün.', primaryQuery:'13 haftalık nakit akışı excel', guideSlug:'13-haftalik-nakit-akisi-excel', guideLinkLabel:'13 haftalık nakit akışı Excel rehberini okuyun' },
  'akilli-kasa-defteri-ve-nakit-kontrol-sistemi': { title:'Kasa Defteri Excel | Gelir Gider ve Nakit Takibi', description:'Kasa defteri Excel sistemiyle günlük nakit giriş çıkışını, açılış-kapanış bakiyesini ve gelir gider hareketlerini düzenli takip edin.', primaryQuery:'kasa defteri excel', guideSlug:'kasa-defteri-excel', guideLinkLabel:'Kasa defteri Excel rehberini okuyun' },
  'aylik-patron-finans-paneli': { title:'Finans Paneli Satın Al | Aylık Patron Excel Paneli', description:'Aylık finans panelini satın almadan önce gerçek Excel ekranlarını ve demoyu inceleyin. Ciro, kâr, nakit ve alacak göstergelerini tek ekranda görün.', primaryQuery:'finans paneli satın al' },
  'banka-kredi-ve-taksit-takip-sistemi': { title:'Kredi Takip Excel | Banka Kredisi ve Taksit Takibi', description:'Banka kredilerini, taksit tarihlerini, kalan borcu ve ödeme yükünü tek Excel sisteminde izleyerek kredi ödeme takviminizi yönetin.', primaryQuery:'kredi takip excel' },
  'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi': { title:'Cari Hesap Takip Excel | Tahsilat ve Risk', description:'Cari hesap takip Excel sistemiyle tahsilatları, geciken alacakları ve müşteri riskini izleyin; kimin ne kadar borçlu olduğunu net görün.', primaryQuery:'cari hesap takip excel', guideSlug:'cari-hesap-takip-excel', guideLinkLabel:'Cari hesap takip Excel rehberini okuyun' },
  'cek-senet-ve-vade-risk-sistemi': { title:'Çek Senet Takip Excel | Vade ve Risk Kontrolü', description:'Çek ve senetleri vade, tutar, taraf ve tahsilat durumuyla tek Excel sisteminde izleyin; yaklaşan ödeme ve tahsilat yoğunluğunu görün.', primaryQuery:'çek senet takip excel' },
  'gunluk-gelir-gider-ve-gercek-karlilik-sistemi': { title:'Gelir Gider Tablosu Excel | Gerçek Kârlılık', description:'Günlük gelir gider tablosunu Excel üzerinde yönetin; hareketlerin gerçek kârlılığa etkisini dönem ve kategori bazında takip edin.', primaryQuery:'gelir gider tablosu excel' },
  'kobi-finans-yonetim-paketi': { title:'KOBİ Finans Takip Excel | Nakit, Borç ve Kârlılık', description:'KOBİ nakit akışı, cari hesap, borç, stok ve kârlılık takibini tek pakette birleştiren Excel finans yönetim sistemini inceleyin.', primaryQuery:'kobi finans takip excel' },
  'pos-komisyon-ve-net-tahsilat-kontrol-sistemi': { title:'POS Komisyon Hesaplama Excel | Net Tahsilat', description:'POS satışlarını, banka komisyonlarını ve hesaba geçecek net tahsilatı Excel üzerinde karşılaştırın; kesinti ve vade farkını görün.', primaryQuery:'pos komisyon hesaplama excel' },
  'proje-ve-is-bazinda-gercek-karlilik-sistemi': { title:'Proje Maliyet Takip Excel | Gerçek Kârlılık', description:'Proje gelirini, doğrudan giderleri ve iş bazında gerçek kârlılığı Excel üzerinde takip ederek hangi işin ne kadar kazandırdığını görün.', primaryQuery:'proje maliyet takip excel' },
  'stok-satis-ve-nakit-baglanma-sistemi': { title:'Stok Takip Excel Şablonu | Satış ve Nakit', description:'Stok takip Excel şablonuyla giriş çıkışı, satışları, ürün bazlı kârlılığı ve stokta bağlı nakdi tek sistem üzerinden izleyin.', primaryQuery:'stok takip excel şablonu', guideSlug:'stok-takip-excel', guideLinkLabel:'Stok takip Excel rehberini okuyun' },
  'vergi-sgk-ve-maas-karsilik-ayirma-sistemi': { title:'Vergi SGK Takip Excel | Maaş ve Ödeme Planı', description:'Vergi, SGK ve maaş yükümlülükleri için dönemsel karşılık ayırın; yaklaşan ödeme ihtiyacını tek Excel planında görün ve takip edin.', primaryQuery:'vergi sgk takip excel' },
  'asiri-dusuk-teklif-savunma-robotu': { title:'Aşırı Düşük Teklif Açıklaması Excel | İhale Analizi', description:'Aşırı düşük teklif açıklaması için maliyet kalemlerini, teklif dayanaklarını ve savunma hesaplarını düzenli bir Excel çalışma yapısında hazırlayın.', primaryQuery:'aşırı düşük teklif açıklaması excel' },
  'ihaleye-kac-tl-teklif-vermeliyim': { title:'İhale Teklif Hesaplama Excel | Sınır Değer', description:'İhale teklif tutarını maliyet, marj ve sınır değer yaklaşımıyla değerlendirmek için karar destekli Excel hesaplama sistemini kullanın.', primaryQuery:'ihale teklif hesaplama excel' },
  'hakedis-fiyat-farki-hak-kaybi-cetveli': { title:'Hakediş Fiyat Farkı Hesaplama Excel | Hak Kaybı', description:'Hakediş dönemlerini ve fiyat farkını Excel üzerinde hesaplayıp karşılaştırın; eksik hesaplanan tutarları ve olası hak kaybını görün.', primaryQuery:'hakediş fiyat farkı hesaplama excel' },
  'yillara-sari-insaat-stopaj-nakit-akis-planlayici': { title:'Yıllara Sari İnşaat Stopaj Excel | Nakit Akışı', description:'Yıllara sari inşaat işlerinde stopaj etkisini ve dönemsel nakit akışını Excel üzerinde birlikte planlayın ve ödeme baskısını görün.', primaryQuery:'yıllara sari inşaat stopaj excel' },
  'taseron-hakedis-kesinti-mutabakati': { title:'Taşeron Hakediş Takip Excel | Kesinti Mutabakatı', description:'Taşeron ve alt yüklenici hakedişlerini, kesintileri ve mutabakat farklarını tek Excel çalışma yapısında takip edip kontrol edin.', primaryQuery:'taşeron hakediş takip excel' },
  'kacirilan-sgk-tesvikleri-ve-gercek-iscilik-maliyeti-analizi': { title:'SGK Teşvik Hesaplama Excel | İşçilik Maliyeti', description:'SGK teşvik hesaplama ve gerçek işçilik maliyetini Excel üzerinde karşılaştırın; kullanılmayan teşvikleri ve maliyet farklarını görün.', primaryQuery:'sgk teşvik hesaplama excel' },
  'kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici': { title:'Kıdem İhbar Hesaplama Excel | Personel Maliyeti', description:'Kıdem, ihbar ve personel çıkarma maliyetini çalışan bazında Excel üzerinde hesaplayın ve toplam yükümlülüğü görün.', primaryQuery:'kıdem ihbar hesaplama excel' },
  'fazla-mesai-ve-isci-dava-riski-tespit-dosyasi': { title:'Fazla Mesai Hesaplama Excel | İşçi Dava Riski', description:'Fazla mesai kayıtlarını ve olası işçi alacak riskini Excel üzerinde analiz ederek dönemsel dava riskini ve mali yükü görün.', primaryQuery:'fazla mesai hesaplama excel' },
  'asgari-ucret-zam-etkisi-fiyat-ayarlama-cetveli': { title:'Asgari Ücret Zam Etkisi Excel | Fiyat Ayarlama', description:'Asgari ücret artışının işçilik maliyeti ve satış fiyatı üzerindeki etkisini Excel üzerinde hesaplayarak fiyat ayarlama ihtiyacını görün.', primaryQuery:'asgari ücret zam etkisi excel' },
  'ithalat-depo-teslim-rafa-gelen-net-birim-maliyet': { title:'İthalat Maliyet Hesaplama Excel | Net Birim Maliyet', description:'Ürün bedeli, kur, navlun, vergi ve diğer ithalat giderlerini Excel üzerinde birleştirerek depoya gelen net birim maliyeti hesaplayın.', primaryQuery:'ithalat maliyet hesaplama excel' },
  'amortisman-ve-sabit-kiymet-satis-zamanlama-stratejisti': { title:'Amortisman Hesaplama Excel | Satış Zamanlama', description:'Normal ve azalan bakiyeli amortismanı, kıst hesabını ve çeyrek satış senaryolarını Excel üzerinde karşılaştırarak satış dönemini değerlendirin.', primaryQuery:'amortisman satış zamanlama excel' },
  'doviz-acik-pozisyonu-ve-kur-riski-stres-testi': { title:'Döviz Açık Pozisyon Excel | Kur Riski Stres Testi', description:'Dövizli varlık ve borçların net açık pozisyonunu hesaplayın; kur şoku senaryolarında kâr/zarar etkisini Excel üzerinde görün.', primaryQuery:'döviz açık pozisyon excel' },
  'kdv-iadesi-azami-alacak-hesabi-dosya-hazirlayici': { title:'KDV İadesi Hesaplama Excel | Azami Alacak', description:'Devreden KDV ve iade sınırıyla azami KDV iade alacağını Excel üzerinde hesaplayın; belge takibini ve dosya hazırlığını düzenleyin.', primaryQuery:'kdv iadesi azami alacak excel' },
  'kkeg-ve-finansman-gider-kisitlamasi-vergi-savunma-seti': { title:'KKEG Hesaplama Excel | Finansman Gider Kısıtlaması', description:'Yabancı kaynak/öz kaynak farkına göre kısıtlamaya tabi finansman giderini ve KKEG tutarını Excel üzerinde hesaplayın.', primaryQuery:'kkeg finansman gider kısıtlaması excel' },
  'mutfak-kayip-kacak-hesaplayici': { title:'Mutfak Kayıp/Kaçak Excel | Menü Kârlılık Analizi', description:'Reçete bazlı teorik tüketimi fiili stokla karşılaştırın; mutfak kayıp oranını ve tutarını Excel üzerinde tespit edin.', primaryQuery:'mutfak kayıp kaçak excel' },
  'nakliye-maliyeti-hesaplayici': { title:'Sefer Başına Nakliye Maliyeti Excel | Fiyatlandırma', description:'Araç filosu ve sefer kayıtlarından yakıt, bakım, sürücü ve amortisman maliyetini Excel üzerinde km ve ton-km bazında hesaplayın.', primaryQuery:'nakliye maliyeti excel' },
  'ortaklar-cari-ve-kasa-adat-faiz-faturasi-hesaplayici': { title:'Ortaklar Cari Adat Faiz Excel | Fatura Hesabı', description:'Ortak cari bakiyeleri ve kasa adatları üzerinden dönem faizini Excel üzerinde hesaplayın; fatura tutarını üretin.', primaryQuery:'ortaklar cari adat faiz excel' },
  'pazaryeri-net-kar-ve-eksik-hakedis-yakalayici': { title:'Pazaryeri Net Kâr Excel | Eksik Hakediş', description:'Komisyon, kargo, reklam ve iade sonrası net kârı hesaplayın; beklenen ile gerçek hakediş farkını Excel üzerinde yakalayın.', primaryQuery:'pazaryeri net kâr excel' },
  'sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli': { title:'TTK 376 Sermaye Kaybı Excel | Öz Kaynak Cetveli', description:'Bilanço bileşenlerinden öz kaynağı ve sermaye kaybını hesaplayın; TTK 376 eşiklerini Excel üzerinde izleyin.', primaryQuery:'ttk 376 sermaye kaybı excel' },
  'sube-karlilik-ve-nakit-hesaplayici': { title:'Bu Şubeyi Kapatmalı Mıyım Excel | Şube Kârlılık', description:'Şube bazında gelir, gider ve nakit akışını toplayın; kâr marjını ve nakit pozisyonunu Excel üzerinde karşılaştırın.', primaryQuery:'şube kârlılık excel' },
  'uretim-recetesi-ve-zam-yansitma-hesaplayici': { title:'Üretim Reçetesi Excel | Zam Yansıtma', description:'Reçetedeki hammadde zamlarının ürün maliyetine etkisini hesaplayın; hedef marjı koruyan satış fiyatını Excel üzerinde görün.', primaryQuery:'üretim reçetesi zam yansıtma excel' },
  'vergi-sgk-borcunu-tecil-etmeli-miyim-kredi-mi-tecil-mi': { title:'Vergi SGK Tecil Excel | Kredi Karşılaştırması', description:'Tecil, banka kredisi ve peşin ödeme seçeneklerinin maliyetini bugünkü değerle Excel üzerinde karşılaştırın.', primaryQuery:'vergi sgk tecil kredi excel' },
  'yeniden-degerleme-yapmali-miyim-vergi-tasarruf-analizi': { title:'Yeniden Değerleme Excel | Vergi Tasarruf Analizi', description:'Yeniden değerleme katsayısı, fon vergisi ve ek amortisman tasarrufunu Excel üzerinde karşılaştırarak net faydayı görün.', primaryQuery:'yeniden değerleme vergi tasarrufu excel' },
  'amortisman-2026-yeniden-degerleme': { title:'Amortisman 2026 Excel | Yeniden Değerleme', description:'Amortisman tutarını 298/Ç ve Geç. 32 yeniden değerleme kararıyla birlikte Excel üzerinde görün; hangi yöntemin vergi etkisini netleştirin.', primaryQuery:'amortisman 2026 yeniden değerleme excel' },
  'cari-ba-bs-toplu-mutabakat': { title:'Ba Bs Mutabakat Excel | Cari Fark Kovası', description:'Cari bakiyeleri Form Ba/Bs tutarlarıyla eşleştirin; fark kovalarını ve mutabakat dışı satırları Excel üzerinde listeleyin.', primaryQuery:'ba bs mutabakat excel' },
  'defter-beyan-e-arsiv-aktarim': { title:'Defter Beyan Excel | e-Arşiv Aktarım', description:'e-Arşiv satırlarını Defter Beyan kolon haritasına dönüştürün; aktarım hatalarını Excel üzerinde üretip düzeltin.', primaryQuery:'defter beyan e-arşiv aktarım excel' },
  'e-fatura-satir-defteri-pdf-kaniti': { title:'e-Fatura Satır Defteri Excel | PDF Kanıtı', description:'e-Fatura satırlarını muhasebe kayıtlarıyla eşleştirin; tutar farkını ve eksik belgeyi Excel üzerinde işaretleyin.', primaryQuery:'e-fatura satır defteri excel' },
  'e-fatura-toplu-donusturucu': { title:'e-Fatura Toplu Dönüştürücü Excel | Defter ve PDF', description:'Toplu e-Fatura export satırlarından Excel defteri ve PDF kanıt üretin; makro veya XML API olmadan dönüştürün.', primaryQuery:'e-fatura toplu dönüştürücü excel' },
  'insaat-hakedis-santiye-maliyet': { title:'İnşaat Hakediş Excel | Şantiye Maliyet', description:'Hakediş kuyruğunu ve şantiye maliyet sapmasını tek Excel ekranında izleyin; geciken hak edişi ve bütçe farkını görün.', primaryQuery:'inşaat hakediş şantiye maliyet excel' },
  'kdv-iade-listesi-robotu-gib7': { title:'KDV İade Listesi Excel | GİB 7 Liste', description:'GİB’in istediği indirilecek KDV, yüklenilen KDV, GÇB ve tamamlayıcı listeleri Excel üzerinden üretin.', primaryQuery:'kdv iade listesi gib excel' },
  'kdv-tevkifat-mahsup-iade-listesi': { title:'KDV Tevkifat Mahsup Excel | İade Ayrımı', description:'Tevkifat mahsup listesi ile iade satırlarını ayırın; MAHSUP, İADE veya BEKLE kararını Excel üzerinde üretin.', primaryQuery:'kdv tevkifat mahsup iade excel' },
  'kira-avans-takip-dekont': { title:'Kira Avans Takip Excel | Dekont', description:'Kira ve avans mahsup bakiyesini Excel üzerinde izleyin; yazdırılabilir dekont için kalan tutarı netleştirin.', primaryQuery:'kira avans takip excel' },
  'konkordato-nakit-akis-on-projesi': { title:'Konkordato Nakit Akış Excel | Ön Proje', description:'Konkordato için üç yıllık bilanço ve nakit akış ön projesini Excel üzerinde mahkeme sunumuna hazırlayın.', primaryQuery:'konkordato nakit akış excel' },
  'restoran-recete-maliyet-fire': { title:'Restoran Reçete Maliyet Excel | Fire', description:'Fire dahil porsiyon maliyetini ve bileşim yüzde 100 kontrolünü Excel üzerinde hesaplayın.', primaryQuery:'restoran reçete maliyet fire excel' },
  'tesvikli-bordro-avantajli-tesvik': { title:'Teşvikli Bordro Excel | En Avantajlı Kod', description:'Çalışan bazında en avantajlı SGK teşvik kodunu seçin; kaçırılan tasarrufu Excel üzerinde görün.', primaryQuery:'teşvikli bordro excel' },
  'tesvikli-bordro-optimizasyon': { title:'Teşvikli Bordro Optimizasyon Excel | Senaryo', description:'Kişi ve teşvik kodu matrisinde çok senaryolu bordro setini Excel üzerinde optimize edin.', primaryQuery:'teşvikli bordro optimizasyon excel' },
  'trendyol-komisyon-sonrasi-net-kar': { title:'Trendyol Net Kâr Excel | Komisyon Sonrası', description:'Trendyol komisyon, TY Plus, flash, reklam ve kargo sonrası net kârı Excel üzerinde hesaplayın; SAT, ZAM veya ÇEKİL kararını görün.', primaryQuery:'trendyol net kâr excel' },
  'ymm-tasdik-kontrol-robotu': { title:'YMM Tasdik Kontrol Excel | Kanıt Paketi', description:'YMM tasdik kontrol listesini, eksik belge kuyruğunu ve kanıt paketini Excel üzerinde yönetin; TASDİK, EKSİK veya DURDUR kararını üretin.', primaryQuery:'ymm tasdik kontrol excel' },
  'logo-sql-cari-yaslandirma-tahsilat-karar-motoru': { title:'Logo SQL Cari Yaşlandırma Excel | Tahsilat Karar Motoru', description:'Logo/SQL cari hareket çıktısını FIFO açık kalem ve yaşlandırma ile tahsilat kararına dönüştüren Excel karar motorunu inceleyin; VUK 323 ön-elemesi ve 13 haftalık tahsilat planı içerir.', primaryQuery:'logo sql cari yaşlandırma excel' },
  'gayrimenkul-yatirim-fizibilite-ve-karlilik': { title:'Gayrimenkul Yatırım Fizibilite Excel | Cap Rate ve Getiri', description:'Ticari ve konut gayrimenkul yatırımlarında kira projeksiyonu, Cap Rate, net işletme geliri (NOI) ve IRR getirisini Excel ile hesaplayın.', primaryQuery:'gayrimenkul yatırım fizibilite excel' },
  'ges-yatirim-fizibilite-ve-proje-finansmani': { title:'GES Yatırım Fizibilite Excel | Proje Finansmanı ve DSCR', description:'Güneş enerjisi santrali (GES) yatırımlarında elektrik üretim satışı, degradasyon, DSCR ve amortisman süresini Excel ile hesaplayın.', primaryQuery:'ges yatırım fizibilite excel' },
  'otel-fizibilite-ve-karlilik': { title:'Otel Fizibilite Excel | Doluluk ve RevPAR Kârlılık', description:'Otel yatırımlarında sezonluk doluluk, RevPAR, ADR, yiyecek içecek gelirleri ve başabaş doluluk eşiğini dinamik Excel ile modelleyin.', primaryQuery:'otel fizibilite excel' },
  'uretim-kapasite-yatirim-ve-karlilik': { title:'Üretim Kapasite Yatırım Excel | OEE ve Birim Maliyet', description:'Yeni makine ve hat yatırımlarında OEE kapasitesi, çevrim süresi, marjinal birim maliyet ve başabaş üretim adedini Excel ile analiz edin.', primaryQuery:'üretim kapasite yatırım excel' },
  'arsa-ve-gayrimenkul-proje-gelistirme-fizibilite': { title:'Arsa Proje Geliştirme Excel | Hasılat Paylaşımı', description:'Arsa geliştirme projelerinde KAKS/TAKS emsal hesabı, kat karşılığı ve hasılat paylaşımı senaryolarını kurumsal Excel ile planlayın.', primaryQuery:'arsa proje geliştirme excel' },
  'sirket-satin-alma-ve-ortaklik-devir-analizi': { title:'Şirket Satın Alma Excel | Değerleme ve M&A Analizi', description:'Şirket satın alma ve hisse devirlerinde normalize EBITDA, Net Borç köprüsü, DCF ve adil hisse değerini gelişmiş Excel ile hesaplayın.', primaryQuery:'şirket satın alma excel' },
  'saas-finansal-planlama-runway-ve-yatirim': { title:'SaaS Finansal Planlama Excel | Runway ve MRR Büyüme', description:'Abonelik modelli yazılım şirketleri için MRR, Churn, CAC/LTV, aylık net nakit yakımı ve runway süresini Excel tablosunda takip edin.', primaryQuery:'saas finansal planlama excel' },
  'gayrimenkul-tut-sat-karar': { title:'Gayrimenkul Tut Sat Karar Excel | Alternatif Getiri', description:'Gayrimenkulü kirada tutmak ile satıp mevduat veya fon yatırımına yönelme kararını 5 yıllık net bugünkü değerle Excel üzerinde kıyaslayın.', primaryQuery:'gayrimenkul tut sat karar excel' },
  'ortakli-gayrimenkul-yatirimi-kar-dagitim': { title:'Ortaklı Gayrimenkul Kâr Dağıtım Excel | Waterfall Modeli', description:'Ortaklı gayrimenkul projelerinde tercihli getiri (Preferred Return) ve kademeli Hurdle waterfall kâr paylaşımını Excel ile dağıtın.', primaryQuery:'ortaklı gayrimenkul kâr dağıtım excel' },
  'borcla-sirket-satin-alma-ve-yatirim-getirisi': { title:'LBO Şirket Satın Alma Excel | Borçla Finansman ve Getiri', description:'Kaldıraçlı satın alma (LBO) analizinde serbest nakit akışıyla borç servisi, kıdemli kredi amortismanı ve kaldıraçlı IRR getirisini Excel ile çözün.', primaryQuery:'borçla şirket satın alma excel' },
  'startup-yatirim-alma-ve-nakit-runway': { title:'Startup Yatırım ve Runway Excel | Cap Table Sulanma', description:'Erken aşama girişimlerde yatırım öncesi/sonrası değerleme, Cap Table sulanma hesabı ve aylık nakit pisti (runway) süresini Excel ile izleyin.', primaryQuery:'startup yatırım runway excel' },
  'emsal-sirket-carpanlariyla-degerleme': { title:'Emsal Şirket Çarpanları Excel | Piyasa Değerlemesi', description:'Sektörel emsal şirketlerin EV/EBITDA, P/E ve F/K çarpanlarını normalize ederek şirketin adil piyasa değerini Excel modeliyle belirleyin.', primaryQuery:'emsal şirket çarpanları excel' },
  'wacc-hesaplama-ve-sermaye-maliyeti': { title:'WACC Hesaplama Excel | Ağırlıklı Sermaye Maliyeti', description:'CAPM özkaynak maliyeti, vergi sonrası borçlanma maliyeti ve sermaye yapısıyla şirkete özgü WACC iskonto oranını Excel üzerinde hesaplayın.', primaryQuery:'wacc hesaplama excel' },
  'profesyonel-hizmet-sirketi-karlilik-ve-kapasite': { title:'Hizmet Şirketi Kârlılık Excel | Kapasite ve Adam Saat', description:'Danışmanlık ve ajanslarda personel faturalandırılabilir kapasite oranı (Utilization) ve müşteri net kâr marjını Excel ile denetleyin.', primaryQuery:'hizmet şirketi kârlılık excel' },
  'perakende-magaza-acilis-fizibilite': { title:'Mağaza Açılış Fizibilite Excel | Perakende Başabaş', description:'Cadde ve AVM mağaza açılışlarında fit-out capex, yaya trafiği, ciro kirası ve başabaş ciro eşiğini kurumsal Excel sistemiyle planlayın.', primaryQuery:'mağaza açılış fizibilite excel' },
  'filo-yatirim-ve-arac-yenileme-fizibilite': { title:'Filo Araç Yenileme Excel | Satın Al vs Kirala TCO', description:'Şirket araçlarında satın alma ile uzun dönem operasyonel kiralama (Buy vs Lease) maliyetlerini ve TCO farkını Excel ile karşılaştırın.', primaryQuery:'filo araç yenileme excel' },
  'franchise-yatirim-fizibilite': { title:'Franchise Yatırım Fizibilite Excel | Royalty ve Kâr', description:'Franchise yatırımlarında isim hakkı bedeli, aylık royalty kesintisi, ürün tedarik marjı ve yatırım geri dönüş süresini Excel ile hesaplayın.', primaryQuery:'franchise yatırım fizibilite excel' },
  'klinik-saglik-merkezi-fizibilite-ve-karlilik': { title:'Klinik Fizibilite Excel | Sağlık Merkezi Kârlılık', description:'Tıp merkezi ve kliniklerde cihaz leasingi, hekim hakediş primleri, hasta kapasitesi ve branş kârlılığını dinamik Excel modeliyle yönetin.', primaryQuery:'klinik fizibilite excel' },
  'restoran-kafe-yatirim-fizibilite-ve-karlilik': { title:'Restoran Kafe Yatırım Excel | Food Cost ve Başabaş', description:'Restoran ve kafe açılışlarında masa devir hızı, adisyon büyüklüğü, food cost reçete maliyeti ve başabaş ciro noktasını Excel ile hesaplayın.', primaryQuery:'restoran kafe yatırım excel' },
};

const qaApprovedCatalogSlugs = [
  "asgari-kurumlar-vergisi-simulasyon-motoru",
  "banka-kredi-covenant-erken-uyari",
  "dava-masraf-harc-hesaplama",
  "depo-doluluk-lokasyon",
  "dernek-kooperatif-mali-yonetim",
  "e-belge-zorunluluk-radari",
  "e-ticaret-gercek-karlilik-fiyatlama",
  "elektrik-faturasi-dogrulama",
  "emlak-pipeline-komisyon",
  "fason-uretim-takip-sistemi",
  "fesih-maliyeti-simulatoru",
  "filo-arac-maliyet-komutasi",
  "ges-uretim-performans",
  "ges-yatirim-fizibilite",
  "gida-lot-maliyet-izlenebilirlik",
  "hizmet-ihracati-100-indirim-transfer-takip",
  "hukuk-burosu-dosya-vekalet",
  "ihale-fiyat-farki-eskalasyon-pro",
  "ihracat-siparis-karlilik-kur",
  "insaat-hakedis-yonetim-sistemi",
  "isg-risk-degerlendirme-pro-6331",
  "ithalat-landed-cost-motoru",
  "kargo-desi-maliyet-optimizasyonu",
  "kat-karsiligi-hasilat-paylasimi-simulator",
  "kdv-iadesi-tutar-surec-simulasyonu",
  "kik-asiri-dusuk-savunma-sinir-deger",
  "kira-portfoyu-getiri-komutasi",
  "konkordato-ttk376-kriz-paketi",
  "kvkk-veri-envanteri-verbis-uyum",
  "makine-bakim-kalibrasyon-durus-maliyeti",
  "mini-mrp-bom-malzeme-kapasite",
  "mizan-anomali-denetim-oncesi-kontrol",
  "oee-durus-kok-neden-komutasi",
  "on-uc-haftalik-nakit-odeme-onceligi",
  "ortulu-sermaye-emsal-faiz-tf-paketi",
  "pazaryeri-hakedis-mutabakat-motoru",
  "proje-finansmani-dscr-llcr-paketi",
  "puantaj-vardiya-fazla-mesai-kanit-sistemi",
  "recete-maliyeti-menu-muhendisligi",
  "sarj-istasyonu-yatirim",
  "sekiz-d-duzeltici-faaliyet-dosya",
  "sera-kurulum-fizibilite",
  "sevkiyat-fiyatlama-navlun",
  "sgk-prim-tesvikleri-optimizasyon-motoru",
  "site-apartman-yonetim-sistemi",
  "spc-proses-yetenek-analizi",
  "stok-optimizasyon-abc-olu-stok-nakit",
  "sut-surusu-yonetim",
  "tahsilat-riski-vade-komuta-paneli",
  "tarimsal-destek-uygunluk",
  "uretim-kari-125-kv-optimizasyonu",
  "yem-rasyonu-maliyet",
  "yeniden-degerleme-komuta-merkezi",
  "yurt-disi-yapilanma-vergi-simulatoru"
] as const;

function seoLabelFromSlug(slug: string): string {
  const small = new Set(['ve','ile','100','125','6331']);
  return slug.split('-').map((part, index) => {
    if (small.has(part)) return part === 've' || part === 'ile' ? part : part;
    const first = part.charAt(0).toLocaleUpperCase('tr-TR');
    return first + part.slice(1);
  }).join(' ');
}

for (const slug of qaApprovedCatalogSlugs) {
  if (productSeo[slug]) continue;
  const label = seoLabelFromSlug(slug);
  productSeo[slug] = {
    title: `${label} Excel | ExcelArşiv`,
    description: `${label} için kalite kapılarından geçmiş Excel karar, kontrol ve raporlama sistemini inceleyin; gerçek ürün ekranları ve doğrulama kanıtlarıyla değerlendirin.`,
    primaryQuery: `${label.toLocaleLowerCase('tr-TR')} excel`,
  };
}

export function getProductSeoByPath(pathname: string): ProductSeoEntry | null {
  const match = pathname.match(/^\/sablon\/([^/]+)\/?$/);
  return match?.[1] ? productSeo[match[1]] ?? null : null;
}
