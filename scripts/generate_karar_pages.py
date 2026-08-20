#!/usr/bin/env python3
"""
generate_karar_pages.py — Strict spec generator for 18 decision pages
"""

import os
import json
import re

with open('veri/urunler.json') as f:
    products = {p['slug']: p for p in json.load(f)}

karar_data = [
    {
        "slug": "hangi-excel-sistemini-almaliyim",
        "h1": "Hangi Excel sistemini almalıyım: İhtiyacınıza göre rehber",
        "title": "Hangi Excel Sistemini Almalıyım? | Excel Arşiv",
        "desc": "İşletmenizin nakit, bütçe, vergi ve kârlılık ihtiyacına göre hangi hazır Excel karar sistemini seçmeniz gerektiğini adım adım rehberimizle hemen öğrenin.",
        "cevap": "İşletmenizin öncelikli darboğazına göre doğru Excel sistemini seçmek nakit kaybını önler. Günlük kasa hareketleri için Akıllı Kasa Defteri, haftalık nakit planlaması için 13 Haftalık Nakit Akışı, KDV iade süreçleri için GİB 7 Robotu, şirket değer ve sermaye kontrolleri için TTK 376 Cetveli tercih edilmelidir.",
        "primary_slug": "13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi",
        "primary_title": "13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi",
        "primary_price": 999,
        "alt_slugs": ["akilli-kasa-defteri-ve-nakit-kontrol-sistemi", "aylik-patron-finans-paneli"],
        "alt_descs": [
            "Yalnızca günlük kasa ve banka giriş çıkışlarını tutmak istiyorsanız Akıllı Kasa Defteri daha pratiktir.",
            "Tüm departmanların ciro, kâr ve bilanço özetini tek tabloda görmek isteyen şirket sahipleri için Aylık Patron Paneli daha uygundur."
        ],
        "table_rows": [
            ("Günlük nakit ve kasa takibi", "Akıllı Kasa Defteri ve Nakit Kontrol Sistemi", 499, "Günlük fiili sayım ile bakiye farkını aynı anda gösterir"),
            ("13 haftalık nakit projeksiyonu", "13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi", 999, "Eksiye düşen haftayı ve ödeme önceliklerini raporlar"),
            ("Müşteri tahsilat ve yaşlandırma", "Cari Hesap, Tahsilat ve Müşteri Risk Takip Sistemi", 799, "Vadesi geçen alacakları ve risk limitlerini takip eder"),
            ("KDV iade listesi hazırlama", "KDV İade Listesi Robotu GİB 7", 799, "GİB formatında indirilecek ve yüklenilecek listeleri hazırlar"),
            ("Şirket kârlılık ve patron paneli", "Aylık Patron Finans Paneli", 999, "Yöneticiye tek ekranda net kâr ve nakit pozisyonu verir"),
        ],
        "dont_buy": [
            "Tek bir dosya ile tüm resmi defterleri ve yevmiye fişlerini otomatik tutmak istiyorsanız.",
            "Çok kullanıcılı bulut ERP yazılımı arıyorsanız ve yerel Excel disiplini kurmak istemiyorsanız.",
            "Veri girişi yapmadan yalnızca yapay zekânın otomatik veri çekmesini bekliyorsanız."
        ],
        "faqs": [
            ("Dosyalar makrosuz mu çalışıyor?", "Evet, tüm karar modellerimiz tamamen açık formüllerle çalışır, makro veya gizli kod barındırmaz."),
            ("İade ve destek süreci nasıl işler?", "Dijital teslimat anında gerçekleşir, kurulum ve kullanım konusunda teknik rehberler her dosya içinde yer alır."),
            ("Mac üzerinde Excel ile uyumlu mu?", "Microsoft Excel 2016 ve üzeri Mac ve Windows sürümleriyle tam uyumlu çalışmaktadır."),
            ("Kendi hesap planıma uyarlayabilir miyim?", "Açık formül yapısı sayesinde sütun ve satırları işletmenizin mali kodlarına göre serbestçe genişletebilirsiniz."),
            ("Tekil mi yoksa paket lisans mı almalıyım?", "Birden fazla alanda ihtiyacınız varsa avantajlı paketleri tercih ederek tasarruf sağlayabilirsiniz.")
        ],
        "related": ["kobi-nakit-akisi-excel", "kasa-defteri-excel", "mali-musavir-cari-takip-excel"],
        "prose": """
Şirketlerin finansal operasyonlarında yaşadığı temel zorluk, doğru zamanda doğru analitik araca başvurmamaktır. Birçok KOBİ, basit bir kasa hareketini takip etmek için karmaşık ERP modüllerine yüksek bütçeler ayırmakta veya tersine, nakit açığını basit bir defterle yönetmeye çalışarak ödeme krizine girmektedir.

Excel Arşiv karar modelleri, işletmelerin karşılaştığı 7 temel alanda geliştirilmiş, mevzuat ve pratik saha kurallarıyla test edilmiş araçlar sunar. Seçim yaparken ilk olarak şirketinizin mevcut aşamadaki en acil darboğazını belirlemeniz gerekir. Nakit sıkışıklığı yaşayan firmalar için 13 haftalık dinamik planlama ilk adımdır. Tahsilat gecikmeleri ve batık alacak riski taşıyan işletmeler için cari yaşlandırma ve risk puanlama sistemleri devreye alınmalıdır.

Vergi ve denetim tarafında ise KDV iadesi, tevkifat ve Ba-Bs kontrolleri gibi yasal zorunluluklar gelir. Bu modeller, muhasebe personeli ve mali müşavirlerin günlerce süren kontrol işlemlerini dakikalara indirir. Tüm modeller açık kaynak mantığıyla sunulur; hiçbir formül kilitli veya şifreli değildir.

Doğru sistemi seçtikten sonra girdi disiplinini oturtmak kritik önem taşır. Renk kodlu hücre yapısı sayesinde veri giriş alanları ile hesaplama motoru birbirinden ayrılmıştır. Bu ayrım, formüllerin kazara silinmesini önler ve yöneticilerin güvenle karar almasını sağlar.
        """
    },
    {
        "slug": "kobi-nakit-akisi-excel",
        "h1": "KOBİ için nakit akışı Excel'i: Hangisini almalısınız?",
        "title": "KOBİ İçin Nakit Akışı Excel'i | Excel Arşiv",
        "desc": "KOBİ nakit akış tablosu ve 13 haftalık dinamik planlama Excel modeli: Gelecek 90 günlük nakit dengesini ve kritik eksi haftaları önceden yakalayın.",
        "cevap": "13 haftalık nakit akışı planlamak isteyen KOBİ'ler için uygun sistem 13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi'dir (999 TL, 18 sayfa). Giriş, çıkış ve kümülatif bakiyeyi haftalık gösterir, eksiye düşen haftayı önceden işaretler. Yalnızca günlük kasa takibi yeterliyse Akıllı Kasa Defteri (499 TL) daha uygundur.",
        "primary_slug": "13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi",
        "primary_title": "13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi",
        "primary_price": 999,
        "alt_slugs": ["akilli-kasa-defteri-ve-nakit-kontrol-sistemi", "kobi-finans-yonetim-paketi"],
        "alt_descs": [
            "Gelecek tahmini yerine sadece bugünkü kasa bakiyesini tutmak istiyorsanız Akıllı Kasa Defteri uygundur.",
            "Nakit akışıyla birlikte bilanço, gelir tablosu ve kârlılık takibi yapmak istiyorsanız KOBİ Finans Paketi daha kapsamlıdır."
        ],
        "table_rows": [
            ("Haftalık nakit açığı tahmini", "13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi", 999, "Gelecek 90 günlük nakit dengesini ve kritik haftayı gösterir"),
            ("Günlük fiili kasa kontrolü", "Akıllı Kasa Defteri ve Nakit Kontrol Sistemi", 499, "Günlük giriş-çıkış ve sayım farklarını kapatır"),
            ("Entegre finans yönetimi", "KOBİ Finans Yönetim Paketi", 2490, "Kasa, banka, çek, cari ve bütçe modellerini bir arada sunar"),
        ],
        "dont_buy": [
            "Şirketinizde haftalık bazda giriş ve çıkış vadeleri takip edilmiyorsa.",
            "Geleceğe dönük nakit planlaması yerine yalnızca geçmiş muhasebe fişlerini girmek istiyorsanız.",
            "Bankaların kredi limitlerini otomatik bağlayan API entegrasyonlu bulut yazılımı arıyorsanız."
        ],
        "faqs": [
            ("13 haftalık nakit akışı modeli neden standart kabul edilir?", "Uluslararası finans yönetiminde 13 hafta (yaklaşık 1 çeyrek), işletmenin operasyonel nakit döngüsünü en net gösteren ufuktur."),
            ("Eksi bakiye uyarıları nasıl çalışır?", "Kümülatif nakit dengesi sıfırın altına indiğinde ilgili hafta kırmızı alarm verir ve gereken likidite tutarını hesaplar."),
            ("Farklı banka hesaplarını ayrı ayrı tanımlayabilir miyim?", "Evet, nakit giriş ve çıkış sayfalarında farklı banka ve kasa kaynakları tanımlanabilir."),
            ("Şablon verileri dışarıdan aktarılabilir mi?", "ERP veya muhasebe programınızdan aldığınız Excel listelerini kopyala-yapıştır ile ilgili girdi alanlarına aktarabilirsiniz."),
            ("Formüllerde değişiklik yapabilir miyim?", "Tüm formüller açıktır; satır ve sütun ekleyerek kendi şirket yapınıza göre uyarlayabilirsiniz.")
        ],
        "related": ["kasa-defteri-excel", "mali-musavir-cari-takip-excel", "sube-karlilik-analizi-excel"],
        "prose": """
KOBİ'lerin yaşadığı finansal tıkanıklıkların büyük bölümü yetersiz nakit planlamasından kaynaklanır. Kâğıt üzerinde kârlı gözüken bir şirket, alacak tahsilatları geciktiğinde veya tedarikçi ödemeleriyle vergi takvimi çakıştığında ödeme aczine düşebilir.

13 Haftalık Nakit Akışı sistemi, yöneticilere çeyrek bazında bir erken uyarı mekanizması kazandırır. Nakit girişlerinin hangi kanallardan geleceği, hammadde alımları, personel maaşları, SGK ve vergi karşılıkları haftalık olarak dengelenir. Sistem kümülatif nakit eğrisini çizerek hangi haftada ek finansman ihtiyacı doğacağını netleştirir.

Bu modelde yer alan karar motoru, minimum nakit rezervi politikanıza göre hangi ödemelerin ertelenebileceğini veya hangi alacakların hızlandırılması gerektiğini simüle etmenizi sağlar. Bankalarla kredi limiti görüşmelerinde sunulabilecek standart finansal rapor formatına sahiptir.
        """
    },
    {
        "slug": "kasa-defteri-excel",
        "h1": "Günlük kasa defteri Excel: Kasa sayım farkı takibi",
        "title": "Günlük Kasa Defteri Excel Şablonu | Excel Arşiv",
        "desc": "Günlük kasa defteri ve nakit kontrol Excel modeli: Gün sonu fiili sayım farklarını, masraf dağılımını ve kasa devir bakiyesini anında takip edin.",
        "cevap": "Günlük kasa hareketlerini, gelir-gider girişlerini ve kasa sayım farklarını takip etmek isteyen işletmeler için Akıllı Kasa Defteri ve Nakit Kontrol Sistemi (499 TL, 14 sayfa) ideal çözümdür. Fiili sayım ile sistem bakiyesini karşılaştırır, açığı ve fazlayı gün bazında raporlar.",
        "primary_slug": "akilli-kasa-defteri-ve-nakit-kontrol-sistemi",
        "primary_title": "Akıllı Kasa Defteri ve Nakit Kontrol Sistemi",
        "primary_price": 499,
        "alt_slugs": ["13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi", "gunluk-gelir-gider-ve-gercek-karlilik-sistemi"],
        "alt_descs": [
            "İleriye dönük nakit projeksiyonu ve vade planı yapmak istiyorsanız 13 Haftalık Nakit Akışı sistemini seçin.",
            "Kasa hareketleriyle birlikte günlük brüt kâr marjınızı hesaplamak istiyorsanız Günlük Gelir-Gider Kârlılık modelini tercih edin."
        ],
        "table_rows": [
            ("Günlük fiili kasa sayımı", "Akıllı Kasa Defteri ve Nakit Kontrol Sistemi", 499, "Kasa açığı veya fazlasını anında yakalar ve kaydeder"),
            ("Haftalık nakit projeksiyonu", "13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi", 999, "Gelecek nakit dengesini planlar"),
            ("Günlük kâr-zarar kontrolü", "Günlük Gelir–Gider ve Gerçek Kârlılık Sistemi", 499, "Günlük marj ve net kâr hesaplar"),
        ],
        "dont_buy": [
            "Market veya perakende barkodlu hızlı satış POS yazılımı arıyorsanız.",
            "Günlük fiili nakit sayımı yapmıyor, yalnızca ayda bir toplu muhasebe fişi kesiyorsanız.",
            "Kasa personeline yetkilendirme şifresi veren çok terminalli veritabanı istiyorsanız."
        ],
        "faqs": [
            ("Kasa sayım farkı nasıl hesaplanır?", "Gün sonu fiziki nakit sayımı girildiğinde, sistem hesaplanan bakiye ile farkı alarak açık veya fazla tutarını gösterir."),
            ("Birden fazla para birimi (TL, USD, EUR) destekleniyor mu?", "Evet, döviz kasası modülü ile farklı para birimleri ayrı tablolarda izlenebilir."),
            ("Günlük devir bakiyesi otomatik aktarılır mı?", "Formüller bir önceki günün kapanışını otomatik olarak ertesi günün açılış bakiyesi yapar."),
            ("Kategori bazında harcama raporu alınabilir mi?", "Aylık özet tablosu yemek, yakıt, malzeme gibi masraf kalemlerini otomatik gruplar."),
            ("Yazıcı çıktısı alınabilir mi?", "Tüm sayfalar A4 yatay ve dikey baskı alanları ayarlanmış olarak tasarlanmıştır.")
        ],
        "related": ["kobi-nakit-akisi-excel", "pos-komisyon-kontrol-excel", "gunluk-gelir-gider-ve-gercek-karlilik-sistemi"],
        "prose": """
Fiziki nakit akışının yoğun olduğu işletmelerde en yaygın kaçak noktası, gün sonu sayım farklarının zamanında tespit edilememesidir. Biriken küçük kasa açıkları ay sonunda ciddi bilanço açıklarına dönüşür.

Akıllı Kasa Defteri, personelin nakit tahsilat ve tediye hareketlerini anında işlemesini sağlar. Gün sonunda yapılan fiili nakit sayımı sisteme girildiğinde, formül hesaplanan teorik bakiye ile fiziksel para arasındaki farkı kuruşu kuruşuna çıkarır.

Ayrıca masraf kategorizasyonu sayesinde işletme sahibi hangi departmanın ne kadar nakit tükettiğini ay sonunda tek bir özet grafikle inceleyebilir. Kasa devirleri otomatik olarak ertesi güne aktarılır, böylece insan kaynaklı toplama hataları tamamen ortadan kalkar.
        """
    },
    {
        "slug": "mali-musavir-cari-takip-excel",
        "h1": "Mali müşavir için cari takip Excel'i: Müşteri risk yönetimi",
        "title": "Mali Müşavir İçin Cari Takip Excel'i | Excel Arşiv",
        "desc": "Mali müşavirler ve muhasebeciler için cari hesap, müşteri risk ve yaşlandırma Excel sistemi: Vadesi geçen alacakları ve riskleri hatasız yönetin.",
        "cevap": "Müşteri bakiyelerini ve vadesi geçen alacakları kontrol etmek isteyen mali müşavirler için Cari Hesap, Tahsilat ve Müşteri Risk Takip Sistemi (799 TL, 16 sayfa) en uygun çözümdür. 30-60-90 gün vadeli yaşlandırma yapar, tahsilat gecikmelerini ve müşteri bazlı risk limitlerini raporlar.",
        "primary_slug": "cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi",
        "primary_title": "Cari Hesap, Tahsilat ve Müşteri Risk Takip Sistemi",
        "primary_price": 799,
        "alt_slugs": ["cari-ba-bs-toplu-mutabakat", "cek-senet-ve-vade-risk-sistemi"],
        "alt_descs": [
            "Ba-Bs mutabakat listelerini otomatik karşılaştırmak için Cari Ba-Bs Mutabakat modelini seçin.",
            "Vadeli çek ve senetlerin tahsilat takvimini izlemek için Çek-Senet Risk sistemini kullanın."
        ],
        "table_rows": [
            ("Cari yaşlandırma ve risk puanı", "Cari Hesap, Tahsilat ve Müşteri Risk Takip Sistemi", 799, "30-60-90 gün vade dilimlerine göre alacakları gruplar"),
            ("Ba-Bs toplu mutabakat kontrolü", "Cari Ba-Bs Toplu Mutabakat", 499, "GİB sınırlarına göre fatura adet ve tutar farkını bulur"),
            ("Çek-senet portföy risk analizi", "Çek–Senet ve Vade Risk Sistemi", 799, "Karşılıksız çek riskini ve banka teminatlarını takip eder"),
        ],
        "dont_buy": [
            "E-fatura entegratörüne doğrudan bağlı bulut ön muhasebe yazılımı arıyorsanız.",
            "Otomatik SMS ve e-posta ihtarname gönderen CRM motoru istiyorsanız.",
            "Müşteri cari ekstrelerini elle girmek yerine otomatik banka çekimi bekliyorsanız."
        ],
        "faqs": [
            ("Yaşlandırma dilimleri nasıl belirlenir?", "0-30, 31-60, 61-90 ve 90+ gün vadeli standart yaşlandırma matrisi üzerinden otomatik hesaplanır."),
            ("Müşteri bazında kredi limiti tanımlanabilir mi?", "Evet, her müşteri kartında risk limiti belirlenebilir ve limit aşıldığında uyarı verilir."),
            ("Dövizli cari hesaplar desteklenir mi?", "TL ve yabancı para cinsinden borç-alacak hareketleri ayrı sütunlarda takip edilir."),
            ("Mutabakat mektubu çıktısı alınabilir mi?", "Standart cari bakiye mutabakat formatında yazdırma alanı mevcuttur."),
            ("Logo ve diğer ERP sistemlerinden aktarım yapılabilir mi?", "Excel mizan ve muavin dökümleri kopyalanarak sisteme kolayca entegre edilir.")
        ],
        "related": ["kobi-nakit-akisi-excel", "logo-erp-cari-yaslandirma-excel", "doviz-acik-pozisyon-kur-riski-excel"],
        "prose": """
Mali müşavirlik ve muhasebe yönetiminde alacak takibi, işletmenin finansal sağlığını koruyan en kritik süreçtir. Tahsil edilmeyen faturalar, zamanında fark edilmediğinde şüpheli alacaklara ve batıklara dönüşür.

Cari Hesap ve Risk Takip Sistemi, müşterilerin ödeme alışkanlıklarını puanlar ve ortalama tahsilat süresini (DSO) hesaplar. Hangi müşterinin ne kadar süredir borcunu geciktirdiği 30 günlük bantlar halinde listelenir.

Bu analitik yapı, yönetim toplantılarında hangi müşteriye vadeli satışın durdurulması gerektiğine dair somut veriler sunar. Açık formül mimarisiyle her işletmenin kendi risk toleransına göre yeniden yapılandırılabilir.
        """
    },
    {
        "slug": "pos-komisyon-kontrol-excel",
        "h1": "POS komisyon kontrol Excel'i: Net tahsilat hesaplama",
        "title": "POS Komisyon Kontrol Excel Şablonu | Excel Arşiv",
        "desc": "Banka POS komisyon kesintileri, blokeli gün takibi ve net tahsilat hesabı için Excel modeli. Banka kesintilerini kuruşu kuruşuna denetleyin.",
        "cevap": "Banka POS kesintilerini, ertesi gün veya blokeli tahsilatları denetlemek isteyen işletmeler için POS, Komisyon ve Net Tahsilat Kontrol Sistemi (499 TL, 13 sayfa) doğru sistemdir. Farklı banka oranlarını karşılaştırır, erken çözme maliyetini ve hesaba net geçişi hesaplar.",
        "primary_slug": "pos-komisyon-ve-net-tahsilat-kontrol-sistemi",
        "primary_title": "POS, Komisyon ve Net Tahsilat Kontrol Sistemi",
        "primary_price": 499,
        "alt_slugs": ["gunluk-gelir-gider-ve-gercek-karlilik-sistemi", "trendyol-pazaryeri-net-kar-excel"],
        "alt_descs": [
            "Günlük toplam gelir ve gider kârlılığını izlemek için Günlük Gelir-Gider modelini seçin.",
            "Pazaryeri komisyon ve kargo kesintilerini denetlemek için Trendyol Net Kâr sistemini kullanın."
        ],
        "table_rows": [
            ("POS komisyon ve bloke denetimi", "POS, Komisyon ve Net Tahsilat Kontrol Sistemi", 499, "Banka kesintilerini ve hesaba net geçiş tarihini hesaplar"),
            ("Günlük net kâr ve marj takibi", "Günlük Gelir–Gider ve Gerçek Kârlılık Sistemi", 499, "İşletmenin günlük net kârlılık tablosunu oluşturur"),
            ("Pazaryeri komisyon analizi", "Trendyol Komisyon Sonrası Net Kâr", 499, "Pazaryeri kesintilerini düşerek net ürün marjı verir"),
        ],
        "dont_buy": [
            "Fiziki POS cihazına USB/Bluetooth ile bağlanıp otomatik slip okuyan yazılım arıyorsanız.",
            "Banka hesap hareketlerini açık bankacılık API'siyle anlık çeken platform istiyorsanız.",
            "Tek çekim ve taksitli oranları girmeden otomatik tahmin bekliyorsanız."
        ],
        "faqs": [
            ("Blokeli gün ve erken çözme faizi nasıl hesaplanır?", "Banka sözleşmenizdeki gün sayısı ve erken çözme iskonto oranı girildiğinde net tutar otomatik hesaplanır."),
            ("Taksitli çekim oranları ayrı ayrı tanımlanabilir mi?", "Evet, 2'den 12 taksite kadar farklı komisyon baremleri desteklenmektedir."),
            ("Birden fazla banka POS'u karşılaştırılabilir mi?", "Farklı bankaların maliyetlerini yan yana koyarak en avantajlı POS'u gösteren analiz sayfası mevcuttur."),
            ("Hatalı banka kesintilerini nasıl tespit eder?", "Beklenen net tahsilat ile hesaba geçen tutarı karşılaştırarak eksik yatan tutarları listeler."),
            ("Restoran ve perakende için uygun mu?", "Hem fiziki mağazalar hem de sanal POS kullanan e-ticaret siteleri için uygundur.")
        ],
        "related": ["kasa-defteri-excel", "trendyol-pazaryeri-net-kar-excel", "gunluk-gelir-gider-ve-gercek-karlilik-sistemi"],
        "prose": """
Kredi kartı ile satış yapan işletmelerin en büyük maliyet kalemlerinden biri banka POS komisyonları ve blokeli gün maliyetleridir. Oranlardaki küçük değişiklikler veya hesaba geçişteki gün kayıpları yıllık bazda yüz binlerce liralık kâr erimesine yol açar.

POS Kontrol Sistemi, yapılan her slip tutarını ilgili bankanın komisyon tablosuyla eşleştirir. Ertesi gün hesabınıza geçmesi gereken net nakit tutarı ile banka ekstresindeki tutar arasındaki farkları anında ortaya koyar.

Ayrıca erken bloke çözme maliyetleri simüle edilerek, nakit ihtiyacında bankaya ödenecek faizin ticari krediye göre avantajlı olup olmadığı hesaplanabilir.
        """
    },
    {
        "slug": "trendyol-pazaryeri-net-kar-excel",
        "h1": "Trendyol ve pazaryeri net kâr hesaplama Excel'i",
        "title": "Trendyol Pazaryeri Net Kâr Excel Şablonu | Excel Arşiv",
        "desc": "Trendyol ve pazaryeri komisyon, kargo, iade ve hizmet bedeli kesintileri sonrası gerçek ürün kârlılığını Excel ile kuruşu kuruşuna hesaplayın.",
        "cevap": "E-ticaret pazaryeri satışlarında komisyon, kargo ve iade kesintileri sonrası net kârı görmek isteyenler için Trendyol Komisyon Sonrasi Net Kar (499 TL, 14 sayfa) uygundur. Baremli kargo ve ceza kesintilerini düşerek ürün bazında gerçek net marjı hesaplar.",
        "primary_slug": "trendyol-komisyon-sonrasi-net-kar",
        "primary_title": "Trendyol Komisyon Sonrası Net Kâr",
        "primary_price": 499,
        "alt_slugs": ["pazaryeri-net-kar-ve-eksik-hakedis-yakalayici", "stok-satis-ve-nakit-baglanma-sistemi"],
        "alt_descs": [
            "Tüm pazaryerlerinde eksik yatan hakedişleri bulmak için Pazaryeri Hakediş modelini seçin.",
            "Ürün stok devir hızını ve bağlı sermayeyi izlemek için Stok Nakit Bağlanma sistemini kullanın."
        ],
        "table_rows": [
            ("Pazaryeri ürün net kâr analizi", "Trendyol Komisyon Sonrası Net Kâr", 499, "Komisyon, barem kargo ve stopaj sonrası net marjı hesaplar"),
            ("Hakediş ve kesinti mutabakatı", "Pazaryeri Net Kâr & Eksik Hakediş", 499, "Pazaryeri faturası ile yatan parayı karşılaştırır"),
            ("Stok devir ve sermaye kontrolü", "Stok, Satış ve Nakit Bağlanma Sistemi", 799, "Depoda bağlanan nakit tutarını ve maliyetini verir"),
        ],
        "dont_buy": [
            "Trendyol API'sine bağlanıp otomatik sipariş onaylayan entegratör yazılımı istiyorsanız.",
            "Toplu ürün yükleme ve kategori eşleştirme botu arıyorsanız.",
            "Kargo desi ve alış fiyatı verilerini girmeden kâr hesaplaması bekliyorsanız."
        ],
        "faqs": [
            ("Baremli kargo ücretleri nasıl hesaplanır?", "Farklı sipariş tutarlarına göre değişen satıcı kargo baremleri formüllere dahil edilmiştir."),
            ("İade oranlarının kâra etkisi hesaplanıyor mu?", "Kategori bazlı tahmini iade oranları girilerek iade kargo maliyetinin net kâra baskısı simüle edilir."),
            ("Stopaj ve KDV ayrımı nasıl yapılıyor?", "E-ticaretteki tevkifatlı KDV ve gelir vergisi stopajı otomatik ayrıştırılır."),
            ("Diğer pazaryerleri (Hepsiburada, Amazon) için kullanılabilir mi?", "Komisyon oranları ve kargo tarifeleri değiştirilerek tüm pazaryerlerine uyarlanabilir."),
            ("Ürün başı minimum satış fiyatını belirleyebilir mi?", "Hedef kâr marjınızı girdiğinizde vermeniz gereken asgari liste fiyatını hesaplar.")
        ],
        "related": ["pos-komisyon-kontrol-excel", "stok-devir-nakit-baglanma-excel", "gunluk-gelir-gider-ve-gercek-karlilik-sistemi"],
        "prose": """
E-ticaret pazar yerlerinde satış hacmi yüksek görünse de komisyon oranları, baremli kargo tarifeleri, hizmet bedelleri ve iade kargo masrafları ürün kârlılığını hızla eritebilir. Birçok satıcı gerçekte zarar ettiği ürünleri yüksek cirolu sanarak satmaya devam eder.

Trendyol Net Kâr modeli, ürünün hammadde/tedarik maliyetinden başlayarak platform komisyonunu, KDV yükünü, kargo masrafını ve tahmini iade maliyetini tek bir satırda toplar.

Sonuç olarak elinize geçecek net kâr tutarı ve net kâr marjı kuruşu kuruşuna hesaplanır. Böylece kârsız ürünler erkenden tespit edilerek fiyat artışı veya satıştan çekilme kararı güvenle verilir.
        """
    },
    {
        "slug": "kdv-iade-dosyasi-excel",
        "h1": "KDV iade listesi hazırlama Excel'i: GİB 7 liste robotu",
        "title": "KDV İade Listesi Hazırlama Excel'i | Excel Arşiv",
        "desc": "GİB internet vergi dairesi formatında indirilecek KDV, yüklenilen KDV ve satış faturaları listelerini hatasız hazırlayan profesyonel Excel sistemi.",
        "cevap": "GİB İnternet Vergi Dairesi standartlarına uygun KDV iade listesi hazırlamak isteyen mükellefler için KDV İade Listesi Robotu GİB 7 (799 TL, 15 sayfa) tasarlanmıştır. İndirilecek ve yüklenilen KDV listelerini formatlar, mükerrer fatura ve VKN hatalarını denetler.",
        "primary_slug": "kdv-iade-listesi-robotu-gib7",
        "primary_title": "KDV İade Listesi Robotu GİB 7",
        "primary_price": 799,
        "alt_slugs": ["kdv-iadesi-azami-alacak-hesabi-dosya-hazirlayici", "kdv-tevkifat-mahsup-iade-listesi"],
        "alt_descs": [
            "Azami talep edilebilir iade tavanını hesaplamak için KDV İadesi Azami Alacak modelini seçin.",
            "Kısmi tevkifattan doğan mahsup ve nakden iade için KDV Tevkifat Mahsup sistemini kullanın."
        ],
        "table_rows": [
            ("GİB 7 liste hazırlama robotu", "KDV İade Listesi Robotu GİB 7", 799, "GİB yükleme şablonuna birebir uyumlu Excel listesi üretir"),
            ("Azami iade alacağı tavan hesabı", "KDV İadesi Azami Alacak Hesabı", 1499, "Mevzuattaki azami iade tutarını ve sınırlarını denetler"),
            ("Tevkifat mahsup ve iade cetveli", "KDV Tevkifat Mahsup İade Listesi", 799, "Tevkifata tabi işlemlerde mahsup sürecini hızlandırır"),
        ],
        "dont_buy": [
            "GİB sistemine doğrudan e-imza ile giriş yapıp otomatik XML yükleyen web botu arıyorsanız.",
            "YMM rapor metnini yapay zekâya otomatik yazdıran doküman robotu istiyorsanız.",
            "Fatura listesini Excel'e aktarmadan sıfırdan otomatik fatura üretmek istiyorsanız."
        ],
        "faqs": [
            ("GİB İnternet Vergi Dairesi şablonlarıyla uyumlu mu?", "Evet, GİB'in talep ettiği sütun sıralaması, tarih ve VKN formatına birebir uyumludur."),
            ("Mükerrer fatura kontrolü yapıyor mu?", "Aynı fatura numarası ve VKN kombinasyonu mükerrer girildiğinde uyarı verir."),
            ("İhracat istisnası KDV iadesi için uygun mu?", "Tam istisna kapsamındaki ihracat teslimleri ve ihraç kayıtlı satışlar için uygundur."),
            ("Yüklenilen KDV dağıtım anahtarı içeriyor mu?", "Girdilerin üretim ve satış paylarına göre yüklenim dağıtımını formüllerle otomatik yapar."),
            ("Büyük fatura listelerinde performans sorunu yaşanır mı?", "Makrosuz hafif formül mimarisi sayesinde binlerce satırlık verilerde hızlı çalışır.")
        ],
        "related": ["mali-musavir-cari-takip-excel", "amortisman-yeniden-degerleme-excel", "doviz-acik-pozisyon-kur-riski-excel"],
        "prose": """
KDV iadesi süreçlerinde vergi dairelerinin en sık dosya reddetme gerekçesi, yüklenen Excel listelerindeki biçim ve tutar hatalarıdır. VKN formatındaki eksiklikler, tarih biçimlendirme yanlışlıkları veya mükerrer satırlar iade sürecini aylarca geciktirebilir.

GİB 7 Liste Robotu, muhasebe kayıtlarından aktarılan verileri otomatik denetim süzgecinden geçirir. İndirilecek KDV, yüklenilen KDV, satış faturaları ve gümrük beyannameleri listelerini yasal standartlara uygun formatlar.

Dağıtım anahtarı modülü sayesinde ortak genel giderlerin yüklenim payları mevzuata uygun hesaplanır. Bu hazırlık, YMM denetim sürecini ve vergi dairesi onayını hızlandırır.
        """
    },
    {
        "slug": "amortisman-yeniden-degerleme-excel",
        "h1": "Amortisman ve 2026 yeniden değerleme hesaplama Excel'i",
        "title": "Amortisman ve Yeniden Değerleme Excel'i | Excel Arşiv",
        "desc": "VUK geçici 32 ve mükerrer 298/Ç kapsamında duran varlık amortismanı, yeniden değerleme oranları ve vergi avantajı hesaplama Excel sistemi.",
        "cevap": "Sabit kıymetlerin amortismanını ve 2026 yılı yeniden değerleme avantajını hesaplamak isteyenler için Amortisman + 2026 Yeniden Değerleme (499 TL, 14 sayfa) geliştirilmiştir. VUK oranlarına göre faydalı ömür amortismanını ve yeniden değerleme sonucu oluşan vergi tasarrufunu belirler.",
        "primary_slug": "amortisman-2026-yeniden-degerleme",
        "primary_title": "Amortisman + 2026 Yeniden Değerleme",
        "primary_price": 499,
        "alt_slugs": ["amortisman-ve-sabit-kiymet-satis-zamanlama-stratejisti", "yeniden-degerleme-yapmali-miyim-vergi-tasarruf-analizi"],
        "alt_descs": [
            "Sabit kıymeti satarken en az vergi ödeyeceğiniz tarihi planlamak için Satış Zamanlama modelini seçin.",
            "Yeniden değerleme yapmanın net vergi kazancını analiz etmek için Vergi Tasarruf Analizi sistemini kullanın."
        ],
        "table_rows": [
            ("Amortisman ve yeniden değerleme", "Amortisman + 2026 Yeniden Değerleme", 499, "VUK oranlarına göre amortisman ve vergi kalkanı hesaplar"),
            ("Sabit kıymet satış vergi planı", "Sabit Kıymet Satış Zamanlama", 799, "Satış kârı vergisini en aza indiren tarihi bulur"),
            ("Yeniden değerleme fizibilite analizi", "Yeniden Değerleme Vergi Tasarruf Analizi", 499, "Ödenecek vergi ile amortisman tasarrufunu kıyaslar"),
        ],
        "dont_buy": [
            "Yalnızca tek bir taşıtın basit amortismanını elle hesaplamak istiyorsanız.",
            "GİB sabit kıymet modülüne XML aktarımı yapan otomatik yazılım arıyorsanız.",
            "Enflasyon düzeltmesi bilançosunu sıfırdan üreten genel muhasebe programı istiyorsanız."
        ],
        "faqs": [
            ("Normal ve azalan bakiyeler yöntemi destekleniyor mu?", "Evet, her iki amortisman yöntemi de yasal oranlarıyla formüllere dahil edilmiştir."),
            ("Geçici 32 ve Mükerrer 298/Ç farkı nedir?", "Geçici 32 geçmiş dönem birikimli değerlemeyi, 298/Ç ise cari dönem sürekli yeniden değerlemeyi kapsar."),
            ("Faydalı ömür listesi güncel mi?", "GİB amortisman tebliğlerindeki güncel faydalı ömür ve oran tablosu sisteme işlenmiştir."),
            ("Net vergi tasarrufu nasıl hesaplanır?", "Değerleme sonrası artan amortisman giderinin kurumlar vergisi matrahına etkisi hesaplanır."),
            ("Binek oto amortisman gider kısıtlaması var mı?", "Binek araçlar için mevzuattaki KDV/ÖTV ve amortisman tavan sınırları dikkate alınır.")
        ],
        "related": ["kdv-iade-dosyasi-excel", "ttk-376-sermaye-kaybi-excel", "mali-musavir-cari-takip-excel"],
        "prose": """
Duran varlıkların amortisman hesabı ve yeniden değerleme uygulamaları, şirketlerin kurumlar vergisi matrahını doğrudan etkileyen yasal vergi kalkanlarıdır. Enflasyonist ortamda sabit kıymetleri tarihi maliyetleriyle bilançoda tutmak, fiktif kâr üzerinden gereksiz vergi ödenmesine neden olur.

Amortisman ve Yeniden Değerleme sistemi, VUK Genel Tebliğleri çerçevesinde yeniden değerleme oranlarını duran varlık kartlarına uygular. Artan net bilanço değeri üzerinden ayrılacak amortisman giderini hesaplar.

Böylece şirketinizin sağlayacağı net kurumlar vergisi tasarrufu ortaya konur. Model ayrıca binek araçlardaki gider kısıtlaması tavanlarını otomatik uygulayarak usulsüzlük riskini engeller.
        """
    },
    {
        "slug": "kidem-ihbar-maliyeti-excel",
        "h1": "Kıdem ve ihbar tazminatı hesaplama Excel'i: Personel maliyeti",
        "title": "Kıdem İhbar Maliyeti Hesaplama Excel'i | Excel Arşiv",
        "desc": "Personel çıkarma maliyeti, kıdem tazminatı tavanı, ihbar süresi ve kullanılmayan yıllık izin karşılıklarını hatasız hesaplayan Excel modeli.",
        "cevap": "İşten ayrılma süreçlerinde kıdem, ihbar ve izin karşılıklarını kuruşu kuruşuna hesaplamak için Kıdem–İhbar Yükü ve Personel Çıkarma Maliyeti Hesaplayıcı (799 TL, 15 sayfa) uygundur. Güncel kıdem tavanını ve brüt giydirilmiş ücret kalemlerini uygulayarak net tazminatı verir.",
        "primary_slug": "kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici",
        "primary_title": "Kıdem–İhbar Yükü ve Personel Çıkarma Maliyeti Hesaplayıcı",
        "primary_price": 799,
        "alt_slugs": ["fazla-mesai-ve-isci-dava-riski-tespit-dosyasi", "asgari-ucret-zam-etkisi-fiyat-ayarlama-cetveli"],
        "alt_descs": [
            "İşçilik dava riski ve fazla mesai iddialarını denetlemek için Fazla Mesai Dava Riski dosyasını seçin.",
            "Asgari ücret zammının şirket maliyetine etkisini görmek için Fiyat Ayarlama Cetvelini kullanın."
        ],
        "table_rows": [
            ("Kıdem ve ihbar tazminatı hesabı", "Kıdem–İhbar Yükü Hesaplayıcı", 799, "Giydirilmiş brüt ücret ve yasal tavanla net tazminatı verir"),
            ("Fazla mesai ve dava riski analizi", "Fazla Mesai Dava Riski Tespit Dosyası", 799, "Olası arabuluculuk ve dava maliyetlerini hesaplar"),
            ("Asgari ücret zam etkisi simülasyonu", "Asgari Ücret Zam Etkisi Cetveli", 799, "İşçilik maliyeti artışının ürün fiyatına etkisini bulur"),
        ],
        "dont_buy": [
            "Aylık 500+ personelin puantaj ve bordro tahakkukunu yapan ERP yazılımı arıyorsanız.",
            "SGK e-bildirge sistemine otomatik bordro yükleyen web robotu istiyorsanız.",
            "Personelin giriş-çıkış tarihlerini girmeden otomatik hesaplama bekliyorsanız."
        ],
        "faqs": [
            ("Giydirilmiş brüt ücrete hangi kalemler dahil edilir?", "Yol, yemek, ikramiye, prim ve düzenli sağlanan tüm ayni/nakdi yardımlar formüllere dahil edilir."),
            ("Güncel kıdem tazminatı tavanı uygulanıyor mu?", "Yılda iki kez değişen Hazine kıdem tavanı parametresi kolayca güncellenebilir formattadır."),
            ("İhbar öneli ve tazminat tutarı nasıl belirlenir?", "İş Kanunu 17. maddedeki çalışma süresi baremlerine göre ihbar haftası ve tutarı hesaplanır."),
            ("Kullanılmayan yıllık izin karşılığı hesaplanır mı?", "Personelin hak ettiği bakiye izin gün sayısı son çıplak brüt ücret üzerinden hesaplanır."),
            ("Damga ve gelir vergisi kesintileri güncel mi?", "Kıdemdeki damga vergisi ve ihbardaki artan oranlı gelir vergisi dilimleri tam uygulanır.")
        ],
        "related": ["sgk-tesvik-optimizasyon-excel", "sube-karlilik-analizi-excel", "mali-musavir-cari-takip-excel"],
        "prose": """
Personel işten çıkış süreçlerinde kıdem ve ihbar tazminatlarının hatalı hesaplanması, işletmeleri iş mahkemelerinde ve arabuluculuk süreçlerinde ciddi faiz ve vekalet ücreti yüküyle karşı karşıya bırakır. Giydirilmiş ücret hesabına düzenli sosyal yardımların katılmaması en sık yapılan hatadır.

Kıdem ve İhbar Maliyeti Hesaplayıcı, personelin işe giriş ve çıkış tarihlerini alarak kıdem süresini gün, ay ve yıl bazında netleştirir. Çıplak ücrete ek olarak yol, yemek ve prim ödemelerini giydirerek yasal tavan kontrolünü uygular.

İhbar süresi, yıllık izin bakiyesi ve yasal vergi kesintileri düşüldükten sonra çalışana ödenecek net tutar ile şirkete toplam maliyet tek bir bordro özetinde sunulur.
        """
    },
    {
        "slug": "sgk-tesvik-optimizasyon-excel",
        "h1": "SGK teşvik hesaplama Excel'i: Kaçırılan prim avantajı",
        "title": "SGK Teşvik Hesaplama Excel Şablonu | Excel Arşiv",
        "desc": "6111, 7103 ve 5510 sayılı SGK istihdam teşviklerini karşılaştırıp işletmeniz için en avantajlı kanunu seçen optimizasyon Excel modeli.",
        "cevap": "Personel istihdamında en avantajlı SGK prim teşvikini belirlemek ve kaçırılan primleri yakalamak isteyenler için Kaçırılan SGK Teşvikleri ve Gerçek İşçilik Maliyeti Analizi (999 TL, 17 sayfa) en uygun sistemdir. Personel bazında yasal teşvik alternatiflerini kıyaslar.",
        "primary_slug": "kacirilan-sgk-tesvikleri-ve-gercek-iscilik-maliyeti-analizi",
        "primary_title": "Kaçırılan SGK Teşvikleri ve Gerçek İşçilik Maliyeti Analizi",
        "primary_price": 999,
        "alt_slugs": ["tesvikli-bordro-optimizasyon", "tesvikli-bordro-avantajli-tesvik"],
        "alt_descs": [
            "Bordro kalemleri üzerinden teşvik dağılımı yapmak için Teşvikli Bordro Optimizasyon modelini seçin.",
            "En karlı teşvik kanununu hızlıca belirlemek için Teşvikli Bordro Seçen sistemini kullanın."
        ],
        "table_rows": [
            ("Kaçırılan teşvik analizi", "Kaçırılan SGK Teşvikleri ve İşçilik Maliyeti", 999, "6111 ve 5510 kanunları kıyaslayarak maksimum kazancı bulur"),
            ("Bordro teşvik optimizasyonu", "Teşvikli Bordro Optimizasyon", 499, "Personel bazında teşvikli maliyet dağılımı yapar"),
            ("Avantajlı teşvik kanunu seçici", "Teşvikli Bordro Seçen", 499, "Hangi kanunun daha karlı olduğunu hızlıca listeler"),
        ],
        "dont_buy": [
            "SGK portalına robotik süreçle (RPA) bağlanıp otomatik teşvik sorgulayan yazılım arıyorsanız.",
            "İşkur kayıtlarını otomatik eşleyen web servisi entegrasyonu istiyorsanız.",
            "Personelin yaş, cinsiyet ve işsizlik süresi verilerini girmek istemiyorsanız."
        ],
        "faqs": [
            ("Hangi SGK kanunları kapsama dahildir?", "5510 (5 puanlık indirim), 6111 (genç/kadın istihdamı), 7103 ve engelli istihdam teşvikleri dahildir."),
            ("Ortalama işçi sayısı şartı nasıl denetlenir?", "İşe alım öncesi son 6 aylık ortalama personel sayısı girilerek ilave istihdam şartı kontrol edilir."),
            ("Aylık prim kazancı nasıl hesaplanır?", "Her personel için kanunlar ayrı ayrı çalıştırılarak en yüksek prim tasarrufu sağlayan kanun işaretlenir."),
            ("Geçmişe dönük teşvik hesaplanabilir mi?", "Geriye dönük dönemler için kaçırılan toplam teşvik tutarı simülasyonu yapılabilir."),
            ("Bordro programlarıyla uyumlu mu?", "Bordro çıktıları Excel'e aktarılarak teşvikli ve teşviksiz maliyet farkı tek tabloda görülür.")
        ],
        "related": ["kidem-ihbar-maliyeti-excel", "sube-karlilik-analizi-excel", "mali-musavir-cari-takip-excel"],
        "prose": """
Türkiye'deki istihdam teşvikleri mevzuatı oldukça karmaşık bir yapıya sahiptir. Şirketler genellikle sadece standart 5510 sayılı %5 prim indiriminden yararlanmakta, 6111 veya sektörel ilave istihdam teşviklerinden doğan çok daha yüksek prim avantajlarını kaçırmaktadır.

Kaçırılan SGK Teşvikleri modeli, çalışanların yaş, cinsiyet, mesleki yeterlilik ve işsizlik geçmişi kriterlerini analiz eder. Son 6 aylık ortalama personel sayısını dikkate alarak her çalışan için yasal olarak uygulanabilecek en yüksek teşvik tutarını belirler.

Bu optimizasyon sayesinde işletmenin aylık toplam işçilik maliyetinde %15 ile %30 arasında yasal tasarruf sağlanabilir.
        """
    },
    {
        "slug": "restoran-kafe-maliyet-excel",
        "h1": "Restoran ve kafe reçete maliyeti hesaplama Excel'i",
        "title": "Restoran Kafe Reçete Maliyet Excel'i | Excel Arşiv",
        "desc": "Restoran, kafe ve mutfak işletmeleri için porsiyon reçete maliyeti, fire oranı ve menü kârlılık analizi sunan profesyonel Excel sistemi.",
        "cevap": "Restoran ve kafelerde porsiyon maliyetini, hammadde firelerini ve menü kâr marjlarını denetlemek isteyen işletmeler için Restoran Reçete Maliyet ve Fire Sistemi (499 TL, 14 sayfa) idealdir. Gramaj bazlı hammadde fiyatlarını menü satış fiyatıyla eşleştirir.",
        "primary_slug": "restoran-recete-maliyet-fire",
        "primary_title": "Restoran Reçete Maliyet ve Fire Sistemi",
        "primary_price": 499,
        "alt_slugs": ["mutfak-kayip-kacak-hesaplayici", "gunluk-gelir-gider-ve-gercek-karlilik-sistemi"],
        "alt_descs": [
            "Mutfak porsiyon kaçaklarını ve stok erimelerini denetlemek için Mutfak Kayıp Kaçak modelini seçin.",
            "İşletmenin günlük net gelir-gider dengesini izlemek için Günlük Gelir-Gider sistemini kullanın."
        ],
        "table_rows": [
            ("Reçete maliyeti ve fire hesabı", "Restoran Reçete Maliyet ve Fire Sistemi", 499, "Gramaj bazlı porsiyon maliyetini ve brüt marjı verir"),
            ("Mutfak kayıp ve kaçak kontrolü", "Mutfak Kayıp/Kaçak Hesaplayıcı", 499, "Teorik hammadde tüketimi ile fiili depoyu eşleştirir"),
            ("Günlük kasa ve kârlılık takibi", "Günlük Gelir–Gider ve Kârlılık", 499, "Kasa, personel ve kira masrafları sonrası net kârı bulur"),
        ],
        "dont_buy": [
            "Garson el terminali ve adisyon dokunmatik ekranı arıyorsanız.",
            "Masalara QR menü açan SaaS bulut yazılımı istiyorsanız.",
            "Hammadde birim alış fiyatlarını sisteme girmeden kâr hesabı bekliyorsanız."
        ],
        "faqs": [
            ("Pişme ve ayıklama fireleri nasıl hesaplanır?", "Hammadde bazında çiğ-pişmiş fire yüzdeleri girilerek net porsiyon maliyeti formüle edilir."),
            ("Hammadde fiyat artışları menüye otomatik yansır mı?", "Ana hammadde tablosundaki birim fiyat değiştiğinde tüm reçeteler anında güncellenir."),
            ("İçecek ve bar reçeteleri destekleniyor mu?", "Mililitre ve shot bazında kokteyl, kahve ve meşrubat reçeteleri mevcuttur."),
            ("Menü mühendisliği (Stars, Puzzles) analizi var mı?", "Satış adedi ve kâr marjı matrisiyle en çok kazandıran ürünler sınıflandırılır."),
            ("KDV ve servis bedeli ayrıştırılıyor mu?", "Mutfak KDV oranları ile yeme-içme satış KDV'si ayrı ayrı hesaplanır.")
        ],
        "related": ["kasa-defteri-excel", "pos-komisyon-kontrol-excel", "stok-devir-nakit-baglanma-excel"],
        "prose": """
Yeme-içme sektöründe kârlılık mutfakta belirlenir. Et, süt, yağ ve sebze gibi temel girdilerdeki günlük fiyat dalgalanmaları menü satış fiyatlarına zamanında yansıtılmadığında kârlı sanılan tabaklar zarar ettirmeye başlar.

Restoran Reçete Maliyet ve Fire Sistemi, her porsiyon için kullanılan gramajları birim alış fiyatlarıyla çarparak reçete kartları oluşturur. Ayıklama, pişme ve porsiyonlama fireleri formüle edilerek tabağın gerçek maliyeti bulunur.

Hammadde fiyatları güncellendiğinde tüm menünün kâr marjı otomatik revize edilir. Böylece hangi ürünün fiyatının artırılması veya menüden çıkarılması gerektiği net olarak görülür.
        """
    },
    {
        "slug": "insaat-hakedis-excel",
        "h1": "İnşaat hakediş ve şantiye maliyeti takip Excel'i",
        "title": "İnşaat Hakediş Takip Excel Şablonu | Excel Arşiv",
        "desc": "Müteahhit ve taşeronlar için yeşil defter, hakediş icmali, fiyat farkı ve şantiye maliyet kontrolü sağlayan profesyonel Excel sistemi.",
        "cevap": "İnşaat projelerinde şantiye harcamalarını, taşeron ödemelerini ve hakediş kesintilerini yönetmek isteyenler için İnşaat Hakediş ve Şantiye Maliyet Sistemi (799 TL, 16 sayfa) geliştirilmiştir. İmalat metrajlarını ve stopaj kesintilerini düzenli hakediş raporuna dönüştürür.",
        "primary_slug": "insaat-hakedis-santiye-maliyet",
        "primary_title": "İnşaat Hakediş ve Şantiye Maliyet Sistemi",
        "primary_price": 799,
        "alt_slugs": ["hakedis-fiyat-farki-hak-kaybi-cetveli", "taseron-hakedis-kesinti-mutabakati"],
        "alt_descs": [
            "Kamu ve özel sektör fiyat farkı endekslerini hesaplamak için Fiyat Farkı Cetvelini seçin.",
            "Taşeron avans ve teminat kesintilerini netleştirmek için Taşeron Mutabakat modelini kullanın."
        ],
        "table_rows": [
            ("Şantiye hakediş ve metraj icmali", "İnşaat Hakediş ve Şantiye Maliyet", 799, "İmalat metrajları ve kesintiler sonrası net ödemeyi verir"),
            ("TÜİK endeksli fiyat farkı hesabı", "Hakediş Fiyat Farkı Cetveli", 799, "Resmi endekslere göre hak edilen fiyat farkını hesaplar"),
            ("Taşeron avans ve ceza mutabakatı", "Taşeron Hakediş Kesinti Mutabakatı", 499, "Taşeron bazında kümülatif hak ediş ve kalan bakiyeyi tutar"),
        ],
        "dont_buy": [
            "BIM ve AutoCAD projelerini 3 boyutlu okuyan mimari yazılım arıyorsanız.",
            "Şantiye işçilerine GPS ile puantaj basan mobil uygulama istiyorsanız.",
            "Metraj ve pursantaj verilerini girmeden hakediş üretmek istiyorsanız."
        ],
        "faqs": [
            ("Yeşil defter ve metraj sayfaları var mı?", "Evet, standart imalat kalemleri için metraj giriş ve yeşil defter şablonları dahildir."),
            ("Stopaj ve teminat kesintileri nasıl yapılır?", "Yıllara sari inşaat stopajı ve nakit teminat kesintisi otomatik düşülür."),
            ("Kümülatif hakediş takibi yapılabilir mi?", "Önceki hakediş toplamları düşülerek bu hakedişte ödenecek net tutar hesaplanır."),
            ("Taşeron bazında ayrı icmal alınabilir mi?", "Demir, kalıp, sıva, mekanik gibi farklı taşeronlar için bağımsız hakediş sayfaları vardır."),
            ("Kamu ihale sözleşmelerine uygun mu?", "KİK ve Çevre Şehircilik Bakanlığı hakediş formatlarına uyarlanabilir yapıdadır.")
        ],
        "related": ["ihale-teklif-sinir-deger-excel", "kobi-nakit-akisi-excel", "mali-musavir-cari-takip-excel"],
        "prose": """
İnşaat ve taahhüt projelerinde kârlılığın korunması, şantiye harcamalarının ve taşeron hakedişlerinin kuruşu kuruşuna denetlenmesine bağlıdır. Yapılmayan imalatların hakedişe yazılması veya avans kesintilerinin unutulması projeleri zarara sürükler.

İnşaat Hakediş Sistemi, metraj cetvellerini ve imalat pursantajlarını hakediş icmaline bağlar. İhzarat, fiyat farkı, nefaset kesintisi, teminat ve stopaj kalemleri yasal kurallara uygun olarak düşülür.

Ana yüklenici ve taşeron arasındaki mutabakatı sağlayan bu şeffaf yapı, iş tesliminde yaşanabilecek anlaşmazlıkları önceden engeller.
        """
    },
    {
        "slug": "ihale-teklif-sinir-deger-excel",
        "h1": "İhale teklif ve sınır değer hesaplama Excel'i: Kamu ihaleleri",
        "title": "İhale Teklif Sınır Değer Excel'i | Excel Arşiv",
        "desc": "KİK kamu ihalelerinde sınır değer, yaklaşık maliyet katsayısı ve aşırı düşük teklif savunma riski hesaplayan profesyonel Excel karar sistemi.",
        "cevap": "Kamu ve özel sektör ihalelerinde en karlı ve elenmeyen teklif tutarını belirlemek isteyen müteahhitler için İhaleye Kaç TL Teklif Vermeliyim Sistemi (999 TL, 16 sayfa) uygundur. KİK sınır değer formüllerine göre teklifinizin sınır altında kalma riskini analiz eder.",
        "primary_slug": "ihaleye-kac-tl-teklif-vermeliyim",
        "primary_title": "İhaleye Kaç TL Teklif Vermeliyim Sistemi",
        "primary_price": 999,
        "alt_slugs": ["asiri-dusuk-teklif-savunma-robotu", "insaat-hakedis-santiye-maliyet"],
        "alt_descs": [
            "Aşırı düşük teklif sorgulamasına savunma dosyası hazırlamak için Aşırı Düşük Savunma Robotunu seçin.",
            "İhale sonrası şantiye imalat ve hakediş takibi için İnşaat Hakediş sistemini kullanın."
        ],
        "table_rows": [
            ("İhale teklif fiyatı simülasyonu", "İhaleye Kaç TL Teklif Vermeliyim", 999, "Sınır değer ve kârlılık dengesini optimize eder"),
            ("Aşırı düşük teklif savunma dosyası", "Aşırı Düşük Teklif Savunma Robotu", 999, "KİK mevzuatına uygun analiz formatı üretir"),
            ("İhale sonrası hakediş kontrolü", "İnşaat Hakediş ve Şantiye Maliyet", 799, "Kazanılan ihalenin şantiye maliyetlerini izler"),
        ],
        "dont_buy": [
            "EKAP sisteminden ihaleleri otomatik tarayıp belge indiren web botu arıyorsanız.",
            "Birim fiyat analizlerini bayındırlık pozlarına göre sıfırdan yazan yazılım istiyorsanız.",
            "Yaklaşık maliyet ve rakip teklif tahminlerini girmeden sonuç bekliyorsanız."
        ],
        "faqs": [
            ("KİK sınır değer katsayıları güncel mi?", "Yapım işleri ve hizmet alımları için KİK tebliğlerindeki güncel R katsayıları tanımlıdır."),
            ("Rakip teklif simülasyonu nasıl çalışır?", "Muhtemel rakip teklifleri girilerek sınır değerin hangi aralığa oturacağı test edilir."),
            ("Aşırı düşük sınırının hemen üstündeki teklif bulunur mu?", "Sistem sorgulamaya kalmadan ihaleyi kazanabilecek en optimize fiyatı önerir."),
            ("Hizmet ve yapım ihaleleri için ayrı formüller var mı?", "Evet, hizmet ve yapım işleri için farklı KİK sınır değer formülleri seçilebilir."),
            ("Kârlılık analizi içeriyor mu?", "Teklif tutarınız ile tahmini imalat maliyetiniz arasındaki net kârı raporlar.")
        ],
        "related": ["insaat-hakedis-excel", "sube-karlilik-analizi-excel", "kobi-nakit-akisi-excel"],
        "prose": """
Kamu ihalelerinde teklif hazırlama süreci matematiksel bir strateji gerektirir. Çok düşük teklif vermek aşırı düşük sorgulamasına takılarak elenme riskini doğururken, çok yüksek teklif vermek ihaleyi kaybetmeye yol açar.

İhaleye Kaç TL Teklif Vermeliyim sistemi, Kamu İhale Kurumu'nun sınır değer hesaplama formüllerini simüle eder. Yaklaşık maliyet ve muhtemel teklif dağılımlarını işleyerek sınır değer eşiğini hesaplar.

Bu analiz, firmanızın aşırı düşük savunmasına kalmadan ihaleyi alabileceği en kârlı teklif bandını belirlemenizi sağlar.
        """
    },
    {
        "slug": "stok-devir-nakit-baglanma-excel",
        "h1": "Stok takip ve devir hızı Excel'i: Nakit bağlanma analizi",
        "title": "Stok Takip ve Devir Hızı Excel'i | Excel Arşiv",
        "desc": "Depodaki atıl stokları, stok devir süresini ve raflarda bağlanan nakit tutarını analiz eden profesyonel finansal stok yönetim modeli.",
        "cevap": "Depodaki ürünlerin stok devir hızını, atıl kalan sermayeyi ve sipariş kritik seviyelerini denetlemek için Stok, Satış ve Nakit Bağlanma Sistemi (799 TL, 15 sayfa) tasarlanmıştır. Raflarda bağlanan nakit maliyetini ve ABC ürün sınıflandırmasını hesaplar.",
        "primary_slug": "stok-satis-ve-nakit-baglanma-sistemi",
        "primary_title": "Stok, Satış ve Nakit Bağlanma Sistemi",
        "primary_price": 799,
        "alt_slugs": ["ithalat-depo-teslim-rafa-gelen-net-birim-maliyet", "trendyol-pazaryeri-net-kar-excel"],
        "alt_descs": [
            "İthalat gümrük ve navlun masraflarını birim maliyete dağıtmak için İthalat Birim Maliyet modelini seçin.",
            "E-ticaret pazar yerlerindeki stok kârlılığını izlemek için Trendyol Net Kâr sistemini kullanın."
        ],
        "table_rows": [
            ("Stok devir hızı ve nakit bağlanma", "Stok, Satış ve Nakit Bağlanma Sistemi", 799, "Raflarda atıl kalan sermayeyi ve devir gününü verir"),
            ("İthalat depo teslim birim maliyet", "İthalat Birim Maliyet Hesaplayıcı", 799, "Gümrük ve navlun masraflarını ürün maliyetine ekler"),
            ("E-ticaret stok ve kâr takibi", "Trendyol Komisyon Sonrası Net Kâr", 499, "Pazaryeri satışlarındaki ürün marjını denetler"),
        ],
        "dont_buy": [
            "Depo raflarında el terminaliyle barkod okutan WMS yazılımı arıyorsanız.",
            "Otomatik tedarikçi satın alma siparişi gönderen ERP modülü istiyorsanız.",
            "Stok giriş-çıkış miktarlarını ve maliyetlerini Excel'e aktarmak istemiyorsanız."
        ],
        "faqs": [
            ("Stok devir süresi (gün) nasıl hesaplanır?", "Ortalama stok tutarı, satılan malın maliyetine bölünerek stokların kaç günde bir tükendiği bulunur."),
            ("ABC analizi ne işe yarar?", "Cironun %80'ini oluşturan A grubu ürünleri belirleyerek depoda nakit optimizasyonu sağlar."),
            ("Kritik sipariş seviyesi uyarısı var mı?", "Minimum emniyet stoku altına düşen ürünler için otomatik sipariş uyarısı üretilir."),
            ("Stokta bağlı nakit finansman maliyeti hesaplanır mı?", "Depoda yatan paranın kredi faiz veya fırsat maliyeti raporlanır."),
            ("SKT ve bozulma riski olan ürünler için uygun mu?", "Gıda ve hızlı tüketim malları için stok bekleme süreleri analiz edilebilir.")
        ],
        "related": ["trendyol-pazaryeri-net-kar-excel", "kobi-nakit-akisi-excel", "sube-karlilik-analizi-excel"],
        "prose": """
Depoda bekleyen her ürün, işletmenin kasasından çekilmiş ve raflara bağlanmış nakit paradır. Stok devir hızının yavaşlaması, hem depo maliyetlerini artırır hem de şirketin likidite krizine girmesine yol açar.

Stok ve Nakit Bağlanma Sistemi, ürün bazında satış hızını ve stokta bekleme gününü hesaplar. ABC analizi ile cironun büyük kısmını getiren kritik ürünleri öne çıkarırken, aylardır satılmayan ölü stokları listeler.

Bu sayede satın alma bütçesi doğru ürünlere yönlendirilir ve gereksiz stok finansmanı maliyetleri ortadan kaldırılır.
        """
    },
    {
        "slug": "sube-karlilik-analizi-excel",
        "h1": "Şube kârlılık analizi Excel'i: Bu şubeyi kapatmalı mıyım?",
        "title": "Şube Kârlılık Analizi Excel Şablonu | Excel Arşiv",
        "desc": "Çok şubeli işletmeler ve mağaza zincirleri için şube bazında brüt marj, sabit masraf dağıtımı ve başabaş noktası hesaplama Excel sistemi.",
        "cevap": "Birden fazla şube veya mağazanın kârlılığını, kira/personel yükünü ve kapatma eşiğini analiz etmek isteyen yöneticiler için Şube Kârlılık ve Nakit Hesaplayıcı (999 TL, 16 sayfa) uygundur. Ortak genel giderleri dağıtarak hangi şubenin nakit tükettiğini belirler.",
        "primary_slug": "sube-karlilik-ve-nakit-hesaplayici",
        "primary_title": "Şube Kârlılık ve Nakit Hesaplayıcı",
        "primary_price": 999,
        "alt_slugs": ["aylik-patron-finans-paneli", "proje-ve-is-bazinda-gercek-karlilik-sistemi"],
        "alt_descs": [
            "Tüm şirketin konsolide yönetim tablosunu izlemek için Aylık Patron Paneli modelini seçin.",
            "Şube yerine proje ve sipariş bazında kâr hesabı için Proje Bazında Kârlılık sistemini kullanın."
        ],
        "table_rows": [
            ("Şube kârlılık ve başabaş analizi", "Şube Kârlılık ve Nakit Hesaplayıcı", 999, "Şube bazında kira, personel ve net kâr dağılımını yapar"),
            ("Konsolide şirket yönetim tablosu", "Aylık Patron Finans Paneli", 999, "Tüm şirketin aylık kâr, zarar ve nakit durumunu verir"),
            ("Proje bazında gerçek kârlılık", "Proje ve İş Bazında Kârlılık Sistemi", 999, "Sipariş ve proje bazlı net marjı hesaplar"),
        ],
        "dont_buy": [
            "Şubelerdeki yazar kasa POS cihazlarını anlık merkez sunucuya bağlayan yazılım arıyorsanız.",
            "Personel vardiya ve mola takip modülü istiyorsanız.",
            "Şube gelir ve gider verilerini Excel girdi sayfalarına aktarmak istemiyorsanız."
        ],
        "faqs": [
            ("Merkez genel giderleri şubelere nasıl dağıtılır?", "Ciro payı, personel sayısı veya metrekare dağıtım anahtarları kullanılarak dağıtılır."),
            ("Şube başabaş ciro noktası nasıl bulunur?", "Şubenin sabit masrafları (kira, personel) brüt kâr marjına bölünerek gereken ciro hesaplanır."),
            ("Zarar eden şubeyi kapatma analizi var mı?", "Şube kapatıldığında kurtarılacak nakit ile merkezde kalacak sabit giderler karşılaştırılır."),
            ("Kaç şubeye kadar destekler?", "Standart yapıda 10 şubeye kadar yan yana karşılaştırma ve konsolide özet sunar."),
            ("Mağaza, kafe ve hizmet şubelerine uygun mu?", "Perakende mağazaları, kafe zincirleri ve bölge ofisleri için uyarlanabilir.")
        ],
        "related": ["kobi-nakit-akisi-excel", "restoran-kafe-maliyet-excel", "ttk-376-sermaye-kaybi-excel"],
        "prose": """
Zincir mağaza ve şube operasyonlarında ciro yüksekliği yanıltıcı olabilir. Yüksek ciro yapan bir şube, yüksek kira ve aşırı personel maliyeti nedeniyle şirketin diğer kârlı şubelerinin ürettiği nakdi tüketiyor olabilir.

Şube Kârlılık Hesaplayıcı, her lokasyonun gelirini, doğrudan şube giderlerini ve merkezden payına düşen ortak genel yönetim giderlerini ayrıştırır. Her şubenin başabaş noktası (break-even) hesaplanır.

Hangi şubenin kâra katkı sağladığı, hangisinin ise şirkete zarar verdiği net olarak ortaya konur. Böylece kira indirimi talep etme veya şube kapatma kararları somut verilere dayandırılır.
        """
    },
    {
        "slug": "ttk-376-sermaye-kaybi-excel",
        "h1": "TTK 376 sermaye kaybı ve borca batıklık hesaplama Excel'i",
        "title": "TTK 376 Sermaye Kaybı Hesaplama Excel'i | Excel Arşiv",
        "desc": "Türk Ticaret Kanunu 376. madde kapsamında sermaye kaybı, borca batıklık cetveli ve sermaye tamamlama fonu hesaplayan mevzuat uyumlu Excel sistemi.",
        "cevap": "Şirket özkaynaklarının erimesi, sermaye kaybı ve borca batıklık durumunu denetlemek isteyen şirket yöneticileri için Şirket Öz Kaynağı Eridi mi? TTK 376 Sermaye Tamamlama Cetveli (1.499 TL, 18 sayfa) tasarlanmıştır. Bilanço kalemlerinden sermaye koruma oranını net olarak verir.",
        "primary_slug": "sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli",
        "primary_title": "Şirket Öz Kaynağı Eridi mi? TTK 376 Sermaye Tamamlama Cetveli",
        "primary_price": 1499,
        "alt_slugs": ["konkordato-nakit-akis-on-projesi", "doviz-acik-pozisyonu-ve-kur-riski-stres-testi"],
        "alt_descs": [
            "Mahkemeye sunulacak konkordato nakit akış ön projesi hazırlamak için Konkordato modelini seçin.",
            "Döviz borçlarının özkaynak üzerindeki kur baskısını ölçmek için Kur Riski sistemini kullanın."
        ],
        "table_rows": [
            ("TTK 376 sermaye kaybı denetimi", "TTK 376 Sermaye Tamamlama Cetveli", 1499, "Sermaye + kanuni yedek akçelerin korunma oranını verir"),
            ("Konkordato nakit akış ön projesi", "Konkordato Nakit Akış Ön Projesi", 2490, "İflas erteleme ve borç yapılandırma simülasyonu yapar"),
            ("Kur şoku ve özkaynak stres testi", "Döviz Açık Pozisyonu ve Kur Riski", 999, "Kur artışının bilanço özkaynağına etkisini hesaplar"),
        ],
        "dont_buy": [
            "Yeminli mali müşavir veya bağımsız denetçi resmi mühür ve imza raporu istiyorsanız.",
            "Ticaret sicil gazetesine otomatik tescil ilanı gönderen platform arıyorsanız.",
            "Şirketin güncel bilanço ve mizan verilerini modele girmek istemiyorsanız."
        ],
        "faqs": [
            ("TTK 376/1, 376/2 ve 376/3 durumları nedir?", "Sermaye ve kanuni yedeklerin 1/2'sinin kaybı (376/1), 2/3'ünün kaybı (376/2) ve borca batıklık (376/3) yasal durumlarıdır."),
            ("Kur farkı zararları sermaye kaybından düşülebilir mi?", "6102 sayılı Kanun tebliğindeki henüz ifa edilmemiş yabancı para borç kur farkı muafiyeti formüllerde dikkate alınır."),
            ("Sermaye tamamlama fonu nasıl hesaplanır?", "Şirketi yasal sınıra çekmek için ortakların koyması gereken asgari nakit tutarını verir."),
            ("Genel kurul çağrı zorunluluğu uyarısı var mı?", "Özkaynak kaybı yasal eşiğe ulaştığında yönetim kurulunun alması gereken aksiyonları listeler."),
            ("Bilanço ve mizan aktarımı nasıl yapılır?", "Standart mizan kalemleri kopyalanarak bilanço özkaynak tablosu otomatik üretilir.")
        ],
        "related": ["mali-musavir-cari-takip-excel", "doviz-acik-pozisyon-kur-riski-excel", "sube-karlilik-analizi-excel"],
        "prose": """
Türk Ticaret Kanunu'nun 376. maddesi, şirket yönetim organlarına sermaye kaybı ve borca batıklık durumunda ağır yasal sorumluluklar yükler. Sermaye ve kanuni yedek akçelerin yarısının veya üçte ikisinin karşılıksız kalması halinde genel kurulun acil önlem alması zorunludur.

TTK 376 Cetveli, şirketin bilanço kalemlerini yasal tebliğ kurallarıyla analiz eder. Yabancı para borçlardan doğan kur farkı zararlarının hesaplama dışı bırakılması opsiyonunu mevzuata uygun şekilde uygular.

Şirketin sermaye kaybı oranını belirler ve şirketin borca batıklıktan çıkması için gereken sermaye artırımı veya sermaye tamamlama fonu tutarını hesaplar.
        """
    },
    {
        "slug": "doviz-acik-pozisyon-kur-riski-excel",
        "h1": "Döviz açık pozisyonu ve kur riski stres testi Excel'i",
        "title": "Döviz Açık Pozisyonu ve Kur Riski Excel'i | Excel Arşiv",
        "desc": "Şirketinizin dövizli varlık ve yükümlülüklerini eşleştirip olası kur artışlarında kâr-zarar ve nakit stres testini hesaplayan profesyonel model.",
        "cevap": "Döviz borçları, ithalat taahhütleri ve yabancı para varlıkları arasındaki kur riskini ölçmek isteyen şirketler için Döviz Açık Pozisyonu ve Kur Riski Stres Testi (999 TL, 16 sayfa) uygundur. Dolar ve Euro kur senaryolarında şirketin maruz kalacağı kur farkı zararını simüle eder.",
        "primary_slug": "doviz-acik-pozisyonu-ve-kur-riski-stres-testi",
        "primary_title": "Döviz Açık Pozisyonu ve Kur Riski Stres Testi",
        "primary_price": 999,
        "alt_slugs": ["kkeg-ve-finansman-gider-kisitlamasi-vergi-savunma-seti", "kobi-finans-yonetim-paketi"],
        "alt_descs": [
            "Yabancı kaynak kullanımından doğan finansman gider kısıtlamasını denetlemek için KKEG Setini seçin.",
            "Kur riskiyle birlikte tüm şirket finansmanını yönetmek için KOBİ Finans Paketini kullanın."
        ],
        "table_rows": [
            ("Döviz açık pozisyonu ve stres testi", "Döviz Açık Pozisyonu ve Kur Riski", 999, "Kur şoklarında şirketin maruz kalacağı zararı simüle eder"),
            ("Finansman gider kısıtlaması hesabı", "KKEG ve Finansman Gider Kısıtlaması", 999, "Yabancı kaynak giderlerinin vergiye etkisini hesaplar"),
            ("Entegre finans yönetim paketi", "KOBİ Finans Yönetim Paketi", 2490, "Tüm nakit, borç ve kur dengesini tek pakette toplar"),
        ],
        "dont_buy": [
            "Canlı Forex ve borsa ekranlarından anlık milisaniyelik arbitraj yapan bot arıyorsanız.",
            "Bankalar arası otomatik forward ve opsiyon türev sözleşmesi bağlayan sistem istiyorsanız.",
            "Dövizli borç ve alacak vadelerini Excel tablosuna girmek istemiyorsanız."
        ],
        "faqs": [
            ("Hangi para birimleri destekleniyor?", "USD, EUR, GBP ve diğer konvertibl para birimleri çapraz kurlarıyla birlikte desteklenir."),
            ("Stres testi senaryoları nasıl çalışır?", "Kurda %10, %25 ve %50 artış senaryolarında net finansman gideri ve özkaynak erimesi hesaplanır."),
            ("Vade uyuşmazlığı analizi var mı?", "Kısa ve uzun vadeli döviz borçları ile alacaklar vade dilimlerine göre eşleştirilir."),
            ("İthalat ve ihracat taahhütleri girilebilir mi?", "Gelecek dönem dövizli sipariş ve taahhütler risk matrisine dahil edilebilir."),
            ("TCMB kur bildirim standartlarına uygun mu?", "Büyük ölçekli krediler için TCMB sistemik risk veri takip formu formatıyla uyumludur.")
        ],
        "related": ["kobi-nakit-akisi-excel", "ttk-376-sermaye-kaybi-excel", "mali-musavir-cari-takip-excel"],
        "prose": """
Döviz cinsinden borçlanan veya hammadde ithal eden şirketler için kur dalgalanmaları en büyük bilanço riskidir. Döviz gelirleri ile döviz borçları arasındaki vade ve miktar uyumsuzluğu, beklenmedik kur artışlarında şirketin tüm faaliyet kârını silebilir.

Döviz Açık Pozisyonu ve Kur Riski modeli, şirketin dövizli nakit, alacak, ticari borç ve banka kredisi varlıklarını tek bir tabloda toplar. Net yabancı para pozisyonunu hesaplar.

Farklı kur artış senaryolarında oluşacak kambiyo zararlarını ve bu zararların özkaynaklar üzerindeki baskısını göstererek hedging (kur riskinden korunma) kararları için somut zemin hazırlar.
        """
    },
    {
        "slug": "logo-erp-cari-yaslandirma-excel",
        "h1": "Logo ve ERP uyumlu cari yaşlandırma raporu Excel'i",
        "title": "Logo ERP Uyumlu Cari Yaşlandırma Excel'i | Excel Arşiv",
        "desc": "Logo ve kurumsal ERP muhasebe programı çıktılarından dinamik cari yaşlandırma, ortalama vade ve tahsilat karar motoru üreten Excel sistemi.",
        "cevap": "Logo veya diğer ERP yazılımlarından alınan cari hareket dökümlerini analiz edip tahsilat kararı üretmek isteyen yöneticiler için Logo/ERP Uyumlu Cari Yaşlandırma ve Tahsilat Karar Motoru (1.499 TL, 18 sayfa) tasarlanmıştır. Ham veriyi yaşlandırma raporuna çevirir.",
        "primary_slug": "logo-sql-cari-yaslandirma-tahsilat-karar-motoru",
        "primary_title": "Logo/ERP Uyumlu SQL Cari Yaşlandırma ve Tahsilat Karar Motoru",
        "primary_price": 1499,
        "alt_slugs": ["cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi", "cari-ba-bs-toplu-mutabakat"],
        "alt_descs": [
            "ERP çıktısı olmadan manuel cari takip ve risk puanlaması yapmak için Cari Risk Takip modelini seçin.",
            "Müşteri ve tedarikçilerle Ba-Bs mutabakatı yapmak için Cari Ba-Bs Mutabakat sistemini kullanın."
        ],
        "table_rows": [
            ("ERP uyumlu SQL cari yaşlandırma", "Logo/ERP Uyumlu Cari Yaşlandırma Motoru", 1499, "Logo ve ERP dökümlerini dinamik tahsilat paneline dönüştürür"),
            ("Manuel cari hesap ve risk takibi", "Cari Hesap ve Müşteri Risk Takip Sistemi", 799, "ERP harici bağımsız cari takip ve vade kontrolü yapar"),
            ("Toplu Ba-Bs mutabakat denetimi", "Cari Ba-Bs Toplu Mutabakat", 499, "Müşteri ve tedarikçi faturalarını toplu karşılaştırır"),
        ],
        "dont_buy": [
            "Logo Yazılım'ın resmi ERP veritabanı lisansı veya çekirdek muhasebe programı arıyorsanız.",
            "Doğrudan SQL veritabanına veri yazıp fatura oluşturan çift yönlü entegrasyon istiyorsanız.",
            "ERP'den Excel/SQL dökümü almadan otomatik işlem bekliyorsanız."
        ],
        "faqs": [
            ("Logo Tiger ve Go sürümleriyle uyumlu mu?", "Evet, Logo Tiger, Go ve diğer ERP'lerden alınan standart cari ekstre ve SQL çıktılarıyla uyumludur."),
            ("SQL sorgusu dosyada yer alıyor mu?", "Evet, Logo veritabanından doğrudan veri çekmek isteyenler için hazır optimize SQL sorgusu dahildir."),
            ("Veri aktarımı ne kadar sürer?", "ERP çıktısını girdi sayfasına yapıştırdığınızda tüm yaşlandırma ve grafikler saniyeler içinde güncellenir."),
            ("Ortalama tahsilat vadesi hesaplanıyor mu?", "Her müşterinin ağırlıklı ortalama vadesi ve vade sapması otomatik hesaplanır."),
            ("Bu sistem Logo Yazılım'ın resmi ürünü müdür?", "Hayır, Excel Arşiv tarafından geliştirilmiş, Logo ve ERP çıktılarıyla uyumlu çalışan bağımsız bir analiz modelidir.")
        ],
        "related": ["mali-musavir-cari-takip-excel", "kobi-nakit-akisi-excel", "doviz-acik-pozisyon-kur-riski-excel"],
        "prose": """
Muhasebe ve ERP sistemleri yüz binlerce satırlık cari hareket kaydı tutabilir ancak şirket yöneticilerine doğrudan karar aldıracak yaşlandırma ve tahsilat özetlerini her zaman pratik biçimde sunamaz.

Logo/ERP Uyumlu Cari Yaşlandırma Motoru, muhasebe yazılımınızdan aldığınız ham hareket dökümlerini işler. Borç-alacak kapatmalarını FIFO (ilk giren ilk çıkar) ve fatura eşleme mantığıyla yaşlandırır.

Müşteri bazında vadesi geçen alacakları, ortalama ödeme gecikmelerini ve tahsilat risklerini tek bir yönetim paneline dönüştürür.
        """
    }
]

# Function to render Astro file
def render_karar_astro(spec):
    slug = spec['slug']
    title = spec['title'].strip()
    desc = spec['desc'].strip()
    h1 = spec['h1'].strip()
    cevap = spec['cevap'].strip()
    primary_slug = spec['primary_slug']
    primary_title = spec['primary_title']
    primary_price = spec['primary_price']
    primary_desc = products.get(primary_slug, {}).get('summary', 'Mevzuata ve sahaya tam uyumlu finansal karar sistemi.')
    dont_buy = spec['dont_buy']
    faqs = spec['faqs']
    related = spec['related']
    alt_slugs = spec['alt_slugs']
    alt_descs = spec['alt_descs']
    table_rows = spec['table_rows']
    prose = spec['prose'].strip()
    
    faq_json = json.dumps([{"@type": "Question", "name": q, "acceptedAnswer": {"@type": "Answer", "text": a}} for q, a in faqs], ensure_ascii=False)
    
    item_elements = [
        {
            "@type": "ListItem",
            "position": 1,
            "item": {
                "@type": "Product",
                "name": primary_title,
                "url": f"https://excelarsiv.com/sablon/{primary_slug}",
                "offers": {
                    "@type": "Offer",
                    "price": primary_price,
                    "priceCurrency": "TRY",
                    "availability": "https://schema.org/InStock"
                }
            }
        }
    ]
    for idx, aslug in enumerate(alt_slugs[:2], start=2):
        ap = products.get(aslug, {})
        item_elements.append({
            "@type": "ListItem",
            "position": idx,
            "item": {
                "@type": "Product",
                "name": ap.get('ad', aslug),
                "url": f"https://excelarsiv.com/sablon/{aslug}",
                "offers": {
                    "@type": "Offer",
                    "price": ap.get('fiyat_tl', 799),
                    "priceCurrency": "TRY",
                    "availability": "https://schema.org/InStock"
                }
            }
        })
    item_json = json.dumps(item_elements, ensure_ascii=False)
    
    code = f"""---
import CommerceLayout from '../../layouts/CommerceLayout.astro';

const primarySlug = {json.dumps(primary_slug)};
const siteUrl = Astro.site ?? 'https://excelarsiv.com/';
const pageUrl = new URL('/karar/{slug}', siteUrl).href;

const jsonLd = {{
  '@context': 'https://schema.org',
  '@graph': [
    {{
      '@type': 'BreadcrumbList',
      '@id': `${{pageUrl}}#breadcrumb`,
      'itemListElement': [
        {{ '@type': 'ListItem', 'position': 1, 'name': 'Ana Sayfa', 'item': siteUrl }},
        {{ '@type': 'ListItem', 'position': 2, 'name': 'Kararlar', 'item': new URL('/karar', siteUrl).href }},
        {{ '@type': 'ListItem', 'position': 3, 'name': {json.dumps(h1)}, 'item': pageUrl }}
      ]
    }},
    {{
      '@type': 'FAQPage',
      '@id': `${{pageUrl}}#faq`,
      'mainEntity': {faq_json}
    }},
    {{
      '@type': 'ItemList',
      '@id': `${{pageUrl}}#itemlist`,
      'name': {json.dumps(h1)},
      'itemListElement': {item_json}
    }}
  ]
}};
---

<CommerceLayout
  title={json.dumps(title)}
  description={json.dumps(desc)}
>
  <Fragment slot="head">
    <link rel="canonical" href={{pageUrl}} />
    <script type="application/ld+json" set:html={{JSON.stringify(jsonLd)}} />
  </Fragment>

  <main class="mx-auto max-w-4xl px-4 py-12">
    <nav class="text-xs text-neutral-500 mb-6 font-mono" aria-label="Breadcrumb">
      <a href="/" class="hover:underline">Ana Sayfa</a> / <a href="/karar" class="hover:underline">Kararlar</a> / <span>{slug}</span>
    </nav>

    <article>
      <header class="border-b border-neutral-200 pb-6 mb-8">
        <span class="inline-block text-xs font-bold uppercase tracking-wider text-emerald-800 bg-emerald-100 px-2.5 py-1 mb-3">
          Karar Rehberi & Karşılaştırma
        </span>
        <h1 class="text-3xl sm:text-4xl font-extrabold text-neutral-900 tracking-tight leading-tight">
          {h1}
        </h1>
        <div class="mt-4 flex items-center gap-3 text-xs text-neutral-500 font-mono">
          <time datetime="2026-08-20">Son güncelleme: 2026-08-20</time>
          <span>·</span>
          <span>Excel Arşiv Finansal Denetim Masası</span>
        </div>
      </header>

      {{/* CEVAP BLOĞU (40-60 KELİME) */}}
      <div class="p-6 bg-emerald-50 border-l-4 border-emerald-800 my-6">
        <p class="text-base text-emerald-950 font-medium leading-relaxed">
          {cevap}
        </p>
      </div>

      {{/* KARAR TABLOSU */}}
      <section class="my-10" aria-label="Karar Tablosu">
        <h2 class="text-xl font-bold text-neutral-900 mb-4">Kullanım Durumuna Göre Sistem Karşılaştırması</h2>
        <div class="overflow-x-auto border border-neutral-200">
          <table class="w-full text-left text-sm divide-y divide-neutral-200">
            <thead class="bg-neutral-100 text-xs font-mono uppercase text-neutral-700">
              <tr>
                <th class="p-3">Durumunuz</th>
                <th class="p-3">Önerilen sistem</th>
                <th class="p-3">Fiyat</th>
                <th class="p-3">Neden</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-neutral-200 bg-white">
"""
    for r in table_rows:
        code += f"""              <tr>
                <td class="p-3 font-medium text-neutral-900">{r[0]}</td>
                <td class="p-3 font-semibold text-emerald-800">{r[1]}</td>
                <td class="p-3 font-mono text-neutral-700 whitespace-nowrap">{r[2]:,} TL</td>
                <td class="p-3 text-xs text-neutral-600">{r[3]}</td>
              </tr>
""".replace(',', '.')
    
    code += f"""            </tbody>
          </table>
        </div>
      </section>

      {{/* BİRİNCİL ÖNERİ */}}
      <section class="my-10 border border-neutral-300 bg-neutral-50 p-6 sm:p-8" aria-label="Birincil Öneri">
        <span class="text-xs font-mono font-bold uppercase tracking-wider text-emerald-800">BİRİNCİL SİSTEM TAVSİYESİ</span>
        <h2 class="text-2xl font-bold text-neutral-900 mt-2">{primary_title}</h2>
        <p class="text-sm text-neutral-700 mt-3 leading-relaxed">
          {primary_desc}
        </p>
        <div class="mt-6 flex flex-wrap items-center justify-between gap-4 pt-6 border-t border-neutral-200">
          <div>
            <span class="text-2xl font-extrabold text-neutral-900 font-mono">{primary_price:,} TL</span>
            <span class="block text-xs text-neutral-500 font-mono">KDV dahil · tek ödeme</span>
          </div>
          <a
            href="/sablon/{primary_slug}"
            class="bg-emerald-800 hover:bg-emerald-900 text-white font-semibold px-6 py-3 text-sm transition-colors"
          >
            Sistemi İnceleyin →
          </a>
        </div>
      </section>

      {{/* ALTERNATİFLER */}}
      <section class="my-10" aria-label="Alternatif Sistemler">
        <h2 class="text-xl font-bold text-neutral-900 mb-4">Alternatif Karar Seçenekleri</h2>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
""".replace(',', '.')

    for a_slug, a_desc in zip(alt_slugs[:2], alt_descs[:2]):
        a_prod = products.get(a_slug, {})
        code += f"""          <div class="border border-neutral-200 p-5 bg-white">
            <h3 class="text-base font-bold text-neutral-900">
              <a href="/sablon/{a_slug}" class="hover:text-emerald-800 hover:underline">{a_prod.get('ad', a_slug)}</a>
            </h3>
            <p class="text-xs text-neutral-600 mt-2 leading-relaxed">{a_desc}</p>
            <div class="mt-4 pt-3 border-t border-neutral-100 flex justify-between items-center text-xs">
              <span class="font-mono text-neutral-500">{a_prod.get('fiyat_tl', 799):,} TL</span>
              <a href="/sablon/{a_slug}" class="text-emerald-800 font-semibold hover:underline">Detaylar →</a>
            </div>
          </div>
""".replace(',', '.')

    code += f"""        </div>
      </section>

      {{/* DÜRÜST SINIR BLOĞU: BU SİSTEMİ ALMAYIN EĞER... */}}
      <section class="my-10 border-l-4 border-amber-600 bg-amber-50/50 p-6" aria-label="Dürüst Sınır">
        <h2 class="text-lg font-bold text-neutral-900 mb-3">Bu sistemi almayın eğer…</h2>
        <ul class="space-y-2 text-sm text-neutral-700">
"""
    for d in dont_buy:
        code += f"""          <li class="flex items-start gap-2">
            <span class="text-amber-700 font-bold">✕</span>
            <span>{d}</span>
          </li>
"""

    code += f"""        </ul>
      </section>

      {{/* DETAYLI İÇERİK METNİ */}}
      <section class="prose prose-neutral max-w-none my-10 text-sm sm:text-base leading-relaxed text-neutral-700">
        <h2 class="text-xl font-bold text-neutral-900 mb-3">Karar Modelinin Çalışma Mantığı ve Çıktıları</h2>
        {prose}
      </section>

      {{/* SSS */}}
      <section class="my-10" aria-label="Sık Sorulan Sorular">
        <h2 class="text-xl font-bold text-neutral-900 mb-4">Sık Sorulan Sorular</h2>
        <div class="divide-y divide-neutral-200 border-y border-neutral-200">
"""
    for q, a in faqs:
        code += f"""          <details class="group py-4">
            <summary class="flex justify-between items-center font-bold text-neutral-900 cursor-pointer text-sm sm:text-base list-none">
              <span>{q}</span>
              <span class="text-emerald-800 group-open:rotate-45 transition-transform text-lg">+</span>
            </summary>
            <p class="mt-3 text-xs sm:text-sm text-neutral-600 leading-relaxed pr-6">
              {a}
            </p>
          </details>
"""

    code += f"""        </div>
      </section>

      {{/* İLGİLİ KARARLAR */}}
      <section class="my-10 bg-neutral-50 border border-neutral-200 p-6" aria-label="İlgili Kararlar">
        <h2 class="text-sm font-mono uppercase tracking-wider text-neutral-500 mb-3">İLGİLİ KARAR REHBERLERİ</h2>
        <div class="flex flex-wrap gap-3 text-xs">
"""
    for r_slug in related:
        code += f"""          <a href="/karar/{r_slug}" class="bg-white border border-neutral-300 px-3 py-2 text-neutral-800 hover:border-emerald-800 font-medium">
            /karar/{r_slug} →
          </a>
"""

    code += f"""        </div>
      </section>

      {{/* KAPANIŞ CTA */}}
      <footer class="my-12 text-center border-t border-neutral-200 pt-8">
        <h3 class="text-xl font-bold text-neutral-900 mb-2">Doğru Excel karar altyapısını hemen kurun.</h3>
        <p class="text-xs sm:text-sm text-neutral-600 mb-6">Açık formüllü, denetlenmiş ve kullanıma hazır sistemi anında indirin.</p>
        <div class="flex flex-wrap justify-center gap-4">
          <a
            href="/sablon/{primary_slug}"
            class="bg-emerald-800 hover:bg-emerald-900 text-white font-semibold px-8 py-3 text-sm transition-colors"
          >
            {primary_title} İncele →
          </a>
          <a
            href="/demo"
            class="border border-neutral-300 bg-white hover:bg-neutral-100 text-neutral-800 font-semibold px-8 py-3 text-sm transition-colors"
          >
            Ücretsiz Demo İndir
          </a>
        </div>
      </footer>
    </article>
  </main>
</CommerceLayout>
"""
    return code

for spec in karar_data:
    filepath = f"src/pages/karar/{spec['slug']}.astro"
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(render_karar_astro(spec))

print(f'Successfully generated {len(karar_data)} karar pages.')
