#!/usr/bin/env python3
"""
build_all_karar_pages.py — Complete, self-contained generator for all 18 decision pages.
"""

import os
import json
import re

with open('veri/urunler.json') as f:
    products = {p['slug']: p for p in json.load(f)}

def kelimeler(metin):
    return re.findall(r'\w+', metin.lower(), flags=re.UNICODE)

karar_pages = [
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
        "paragraphs": [
            "Şirketlerin finansal operasyonlarında yaşadığı temel zorluk, doğru zamanda doğru analitik araca başvurmamaktır. Birçok işletme basit bir kasa hareketini takip etmek için karmaşık ERP modüllerine yüksek bütçeler ayırmakta veya tersine, nakit açığını basit bir not defteriyle yönetmeye çalışarak ödeme krizine girmektedir. Excel Arşiv karar modelleri, işletmelerin karşılaştığı temel operasyonel alanlarda geliştirilmiş, mevzuat ve pratik saha kurallarıyla test edilmiş analitik araçlar sunar.",
            "Seçim yaparken ilk olarak şirketinizin mevcut aşamadaki en acil darboğazını belirlemeniz gerekir. Nakit sıkışıklığı yaşayan firmalar için 13 haftalık dinamik planlama ilk adımdır. Tahsilat gecikmeleri ve alacak riski taşıyan işletmeler için cari yaşlandırma ve risk puanlama sistemleri devreye alınmalıdır. Vergi ve denetim tarafında ise KDV iadesi, tevkifat ve Ba-Bs kontrolleri gibi yasal zorunluluklar gelir. Bu modeller, muhasebe personeli ve mali müşavirlerin günlerce süren kontrol işlemlerini dakikalara indirir.",
            "İşletmenizin operasyonel hacmine göre tekil ürün veya paket seçimi yapmak maliyet avantajı sağlar. Örneğin tek bir şantiyenin hakedişini tutmak için inşaat hakediş dosyası yeterliyken, genel şirket finansmanını yönetmek isteyen KOBİ'ler için Kasa, Banka, Çek ve Bütçe modellerini içeren KOBİ Finans Paketi daha doğru bir yatırımdır. Her model birbirini tamamlayacak şekilde standart hücre ve formül disipliniyle inşa edilmiştir.",
            "Finansal karar süreçlerinde güvenilirlik, modellerin şeffaf formül mimarisinden kaynaklanır. Karmaşık makrolar veya gizli yazılımlar yerine tamamen açık Excel formülleri kullanılır. Bu durum denetim firmaları ve mali müşavirlerin tüm hesaplama adımlarını adım adım doğrulamasına imkan tanır. Hata payı en aza indirilir.",
            "Tüm modeller açık kaynak mantığıyla sunulur; hiçbir formül kilitli veya şifreli değildir. Doğru sistemi seçtikten sonra girdi disiplinini oturtmak kritik önem taşır. Renk kodlu hücre yapısı sayesinde veri giriş alanları ile hesaplama motoru birbirinden ayrılmıştır. Bu ayrım, formüllerin kazara bozulmasını önler ve yöneticilerin güvenle karar almasını sağlar. Şirketinizin büyüklüğüne göre tekil modellerden paket çözümlere geçiş yapabilirsiniz."
        ]
    },
    {
        "slug": "kobi-nakit-akisi-excel",
        "h1": "KOBİ için nakit akışı Excel'i: Hangisini almalısınız?",
        "title": "KOBİ İçin Nakit Akışı Excel Şablonu | Excel Arşiv",
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
        "paragraphs": [
            "KOBİ'lerin yaşadığı finansal tıkanıklıkların büyük bölümü yetersiz nakit planlamasından kaynaklanır. Kâğıt üzerinde kârlı gözüken bir şirket, alacak tahsilatları geciktiğinde veya tedarikçi ödemeleriyle vergi takvimi çakıştığında ödeme aczine düşebilir. Nakit akışı yönetimi, kâr-zarar tablosundan tamamen farklı bir likidite mantığına dayanır.",
            "13 Haftalık Nakit Akışı sistemi, yöneticilere çeyrek bazında bir erken uyarı mekanizması kazandırır. Nakit girişlerinin hangi kanallardan geleceği, hammadde alımları, personel maaşları, SGK ve vergi karşılıkları haftalık olarak dengelenir. Sistem kümülatif nakit eğrisini çizerek hangi haftada ek finansman ihtiyacı doğacağını netleştirir. Bu sayede panik borçlanmaların önüne geçilir.",
            "Nakit açığı yönetiminde en kritik faktör, vadesi gelen borçların şirketin gerçek tahsilat gücüyle uyumlandırılmasıdır. Çok sayıda işletme kâğıt üzerinde kâr elde ederken, nakit giriş ve çıkış vadelerindeki asimetri sebebiyle ani likidite krizine girer. 13 haftalık nakit akışı planlaması, her haftanın başlangıç bakiyesini, beklenen müşteri tahsilatlarını, zorunlu tedarikçi ödemelerini, personel maaşlarını ve vergi yükümlülüklerini tek bir dinamik projeksiyonda birleştirir.",
            "Senaryo analizi yeteneği sayesinde satışların yüzde yirmi düşmesi veya tahsilatların iki hafta gecikmesi gibi risk senaryolarında kasanın kaçıncı haftada eksiye düşeceği önceden simüle edilir. Bu öngörü, şirket yönetimine kredi limitlerini açtırma, vadeli mal alımlarını yeniden müzakere etme veya gereksiz nakit çıkışlarını dondurma imkanı tanır.",
            "Model içerisindeki ödeme önceliklendirme matrisi, yasal yükümlülükler ile ticari borçları risk derecelerine göre sınıflandırır. Gecikmesi durumunda ticari faiz veya itibar kaybı doğuracak ödemeler ilk sıraya alınırken ertelenebilir genel yönetim harcamaları net olarak ayrıştırılır.",
            "Finans yöneticisi her hafta gerçekleşen fiili rakamları sisteme girdiğinde model gelecek 12 haftayı otomatik olarak günceller. Bütçelenen nakit akışı ile gerçekleşen nakit akışı arasındaki sapmalar haftalık raporla şirket ortaklarına sunulur."
        ]
    },
    {
        "slug": "kasa-defteri-excel",
        "h1": "Günlük kasa defteri Excel: Kasa sayım farkı takibi",
        "title": "Günlük Kasa Defteri Excel Şablonu | Excel Arşiv",
        "desc": "Günlük kasa defteri ve nakit kontrol Excel modeli: Gün sonu fiili sayım farklarını, masraf dağılımını ve kasa devir bakiyesini anında takip edin.",
        "cevap": "Günlük kasa hareketlerini, gelir-gider girişlerini ve kasa sayım farklarını takip etmek isteyen işletmeler için Akıllı Kasa Defteri ve Nakit Kontrol Sistemi (499 TL, 14 sayfa) ideal çözümdür. Fiili sayım ile sistem bakiyesini karşılaştırır, açığı ve fazlayı gün bazında kuruşu kuruşuna raporlar.",
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
        "paragraphs": [
            "Fiziki nakit sirkülasyonunun yüksek olduğu perakende ve hizmet işletmelerinde kasa açıkları genellikle gün sonu sayım disiplini kurulmadığı için birikir. Küçük sayım farkları zamanında tespit edilmediğinde ay sonunda açıklanamayan bilanço açıklarına dönüşür. Akıllı Kasa Defteri, her nakit giriş ve çıkış hareketini anında makbuz veya fiş numarasıyla eşleştirir.",
            "Gün sonunda kasada bulunan fiziki para adetleri ve kupürleri girildiğinde sistem teorik bakiye ile fiili nakit arasındaki farkı kuruşu kuruşuna çıkarır. Açık veya fazla durumu anında renk kodlu uyarı verir ve kasa sorumlusunun açıklama girmesini zorunlu kılar.",
            "Departman ve masraf merkezi bazında harcama kodlaması sayesinde şirket yönetimi yemek, yakıt, kırtasiye ve personel avansı gibi kalemlerin nakit tüketim paylarını aylık özet tablosunda inceler. Kontrolsüz harcamalar anında frenlenir.",
            "Kasa devir bakiyeleri bir sonraki güne formüllerle otomatik olarak aktarılır. Manuel toplama veya taşıma hataları tamamen engellenir. Gün sonu imzalı mutabakat çıktısı alınarak kasa güvenliği kurumsal seviyeye taşınır.",
            "Sistem, kasa sorumlusunun yetkisi dahilindeki harcamaları limit kontrolünden geçirir. Belirlenen tavan tutarı aşan nakit çıkışlarında onay uyarısı verir. Bu mekanizma, şirket içinde kontrolsüz avans dağıtımını ve belgesiz harcamaları disipline eder.",
            "Farklı kasa lokasyonları veya şube kasaları tek bir merkez özet tablosuna bağlanabilir. Gün sonunda her kasanın devir bakiyesi konsolide nakit durumuna aktarılır. Bu sayede genel müdür toplam fiziksel nakit varlığını tek ekranda izler."
        ]
    },
    {
        "slug": "mali-musavir-cari-takip-excel",
        "h1": "Mali müşavir için cari takip Excel'i: Müşteri risk yönetimi",
        "title": "Mali Müşavir Cari Takip Excel Şablonu | Excel Arşiv",
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
        "paragraphs": [
            "Mali müşavirlerin ve kurumsal muhasebe departmanlarının karşılaştığı en büyük risk, vadesi geçen müşteri alacaklarının yaşlandırma yapılmadan takip edilmesidir. Alacakların vadesi uzadıkça tahsil edilme olasılığı hızla düşer ve şüpheli ticari alacak karşılığı ayırma zorunluluğu doğar. Cari Yaşlandırma Sistemi, tüm müşteri hesaplarını 30, 60, 90 ve 90 üzeri gün dilimlerine göre anında gruplar.",
            "Müşteri bazlı kredi limiti ve risk skoru mekanizması, borcunu düzenli ödemeyen müşterilere yeni vadeli satış yapılmasını engeller. Finans birimi ve satış ekibi arasındaki koordinasyon somut kurallara bağlanır.",
            "Dövizli cari hesaplar için kur değerleme modülü, yabancı para alacak ve borçların değerleme tarihindeki kur farkı kâr veya zararını otomatik hesaplar. Dönem sonu mizan mutabakatları dakikalar içinde tamamlanır.",
            "Otomatik mutabakat mektubu ve hesap ekstresi şablonları, müşterilere düzenli bakiye teyidi gönderilmesini sağlar. Hatalı fatura kayıtları veya eksik ödemeler erken aşamada çözüme kavuşturulur.",
            "VUK 323 şüpheli ticari alacak karşılığı ayırma şartları modelde kontrol edilebilir. İcra veya dava aşamasına gelen alacaklar için ayrılacak karşılık tutarı ve vergi matrahına etkisi hesaplanır.",
            "Toplu müşteri risk skoru dağılımı grafiği, şirketin toplam alacak portföyünün ne kadarının güvenli, ne kadarının ise kritik risk bandında olduğunu tek bakışta gösterir. Finansman yöneticisi faktoring veya teminat kararlarını bu tabloya göre alır."
        ]
    },
    {
        "slug": "pos-komisyon-kontrol-excel",
        "h1": "POS komisyon kontrol Excel'i: Net tahsilat hesaplama",
        "title": "POS Komisyon Kontrol Excel Şablonu | Excel Arşiv",
        "desc": "Banka POS komisyon kesintileri, blokeli gün takibi ve net tahsilat hesabı için Excel modeli. Banka kesintilerini kuruşu kuruşuna denetleyin.",
        "cevap": "Banka POS kesintilerini, ertesi gün veya blokeli tahsilatları denetlemek isteyen işletmeler için POS, Komisyon ve Net Tahsilat Kontrol Sistemi (499 TL, 13 sayfa) doğru sistemdir. Farklı banka oranlarını karşılaştırır, erken bloke çözme maliyetini ve hesaba net geçişi kuruşu kuruşuna eksiksiz hesaplar.",
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
        "paragraphs": [
            "Kredi kartı ve POS cihazları üzerinden tahsilat yapan işletmelerin en sık yaşadığı finansal kayıp, banka komisyon oranlarının ve bloke sürelerinin doğru denetlenmemesidir. Anlaşma şartlarındaki oran değişiklikleri veya hesaba geçiş tarihlerindeki kaymalar yıllık bazda kârlılığı eritir.",
            "POS Kontrol Sistemi, bankadan gelen günlük ve aylık slipleri sözleşmedeki komisyon baremleriyle eşleştirir. Ertesi gün hesaba geçmesi gereken net tutar ile banka ekstresindeki fiili bakiye kuruşu kuruşuna karşılaştırılır. Hatalı kesintiler hemen listelenir.",
            "Farklı bankaların sunduğu taksitli çekim komisyonları ve bloke gün şartları yan yana kıyaslanır. Müşteriye uygulanacak taksit farkı ve vade maliyeti en kârlı şekilde belirlenir.",
            "Erken bloke çözme faiz maliyeti simülasyonu, işletmenin acil nakit ihtiyacında POS blokesini çözmenin ticari kredi kullanmaktan daha avantajlı olup olmadığını matematiksel olarak hesaplar.",
            "Farklı bankaların POS teklifleri arasında simülasyon yaparak yıllık ciro hacminizde hangi bankanın daha az komisyon keseceği hesaplanır. Bu rapor banka pazarlıklarında işletmeye güçlü bir koz sunar.",
            "Banka üye işyeri sözleşmelerindeki aidat, yazılım bedeli ve operasyonel kesintiler de ek maliyet sekmesinde toplanarak fiili efektif komisyon oranı hesaplanır. Böylece gizli maliyetler net biçimde görünür hale gelir."
        ]
    },
    {
        "slug": "trendyol-pazaryeri-net-kar-excel",
        "h1": "Trendyol ve pazaryeri net kâr hesaplama Excel'i",
        "title": "Trendyol Pazaryeri Net Kâr Excel | Excel Arşiv",
        "desc": "Trendyol ve pazaryeri komisyon, kargo, iade ve hizmet bedeli kesintileri sonrası gerçek ürün kârlılığını Excel ile kuruşu kuruşuna hesaplayın.",
        "cevap": "E-ticaret pazaryeri satışlarında komisyon, kargo, ceza ve iade kesintileri sonrası net kârı görmek isteyen satıcılar için Trendyol Komisyon Sonrası Net Kâr (499 TL, 14 sayfa) uygundur. Baremli kargo ve ceza kesintilerini düşerek ürün bazında gerçek net kâr marjını kuruşu kuruşuna hesaplar.",
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
        "paragraphs": [
            "Pazaryerlerinde satış yapan firmaların en büyük yanılgısı, ciro büyüklüğünü kârlılık ile karıştırmaktır. Platform komisyonu, baremli kargo maliyeti, hizmet bedelleri, KDV tevkifatı ve iade kargo masrafları bir araya geldiğinde yüksek hacimli satışlar dahi net zarara yol açabilir.",
            "Trendyol Net Kâr modeli, her ürün için alış fiyatını, paketleme masrafını, kategori komisyonunu ve desi bazlı kargo ücretini tek bir formülde birleştirir. İade oranı tahmini düşüldükten sonra satıcının cebine kalacak net kâr tutarı ve net kâr marjı kuruşu kuruşuna hesaplanır.",
            "Pazaryeri reklam harcamaları ve kupon maliyetleri ürün birim kârına paylaştırılarak hangi kampanyanın gerçekten net kâr ürettiği tespit edilir. Zararına satış yaptıran kampanyalar anında durdurulur.",
            "Pazaryeri hakediş mutabakat tablosu, platformun kestiği faturalar ile satıcı hesabına yatan nakit transferlerini karşılaştırır. Eksik ödemeler ve haksız cezalar resmi itiraz formatında listelenir.",
            "Farklı satış hacimlerinde baremli kargo fiyat geçişleri simüle edilerek kargo barem sınırına yakın ürünlerin fiyatlandırmasında optimum satış fiyatı belirlenir.",
            "Sonuç olarak elinize geçecek net kâr tutarı ve net kâr marjı kuruşu kuruşuna hesaplanır. Böylece kârsız ürünler erkenden tespit edilerek fiyat artışı veya satıştan çekilme kararı güvenle verilir. Kampanya dönemlerinde indirim marjları sağlıklı planlanır."
        ]
    },
    {
        "slug": "kdv-iade-dosyasi-excel",
        "h1": "KDV iade listesi hazırlama Excel'i: GİB 7 liste robotu",
        "title": "KDV İade Listesi Hazırlama Excel | Excel Arşiv",
        "desc": "GİB internet vergi dairesi formatında indirilecek KDV, yüklenilen KDV ve satış faturaları listelerini hatasız hazırlayan profesyonel Excel sistemi.",
        "cevap": "GİB İnternet Vergi Dairesi standartlarına uygun KDV iade listesi hazırlamak isteyen mükellefler için KDV İade Listesi Robotu GİB 7 (799 TL, 15 sayfa) tasarlanmıştır. İndirilecek ve yüklenilen KDV listelerini formatlar, mükerrer fatura ve VKN hatalarını denetler ve listeleri hazırlar.",
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
            ("Büyük fatura listelerinde performans sorunu yaşanır mı?", "Makrosuz hafif formül mimarisi sayesinde uzun listelerde hızlı çalışır.")
        ],
        "related": ["mali-musavir-cari-takip-excel", "amortisman-yeniden-degerleme-excel", "doviz-acik-pozisyon-kur-riski-excel"],
        "paragraphs": [
            "KDV iadesi dosyalarının vergi daireleri tarafından reddedilmesinin veya incelemeye sevk edilmesinin başlıca sebebi, yüklenen Excel listelerindeki biçimsel ve matematiksel tutarsızlıklardır. VKN hataları, mükerrer faturalar ve intaç tarihi uyuşmazlıkları süreci aylarca kilitler.",
            "GİB 7 Liste Robotu, muhasebe kayıtlarından aktarılan indirilecek KDV, yüklenilen KDV, satış faturaları ve gümrük beyannameleri listelerini İnternet Vergi Dairesi yükleme şablonuna birebir uyumlu hale getirir. Tüm alanlar otomatik doğrulanır.",
            "Yüklenilen KDV dağıtım anahtarı modülü, genel imalat ve yönetim giderlerinin ihracat veya indirimli oranlı teslimlere düşen payını yasal mevzuata tam uyumlu formüllerle dağıtır.",
            "Karşıt inceleme ve mükellef risk denetim sayfaları, listedeki tedarikçilerin vergi uyum durumlarını ön kontrolden geçirerek YMM raporunun ve nakden iade talebinin hızla onaylanmasını sağlar.",
            "Gümrük çıkış beyannamesi (GÇB) kapanış tarihleri ile fatura intaç tarihleri arasındaki uyum otomatik test edilir. İade talep dönemine ait olmayan teslimler listeden ayıklanır.",
            "KDV iade kontrol raporu sayfası, vergi dairesi kontrol memurlarının en sık sorguladığı oran ve matrah tutarlılıklarını önceden denetleyerek dosyanın ilk seferde eksiksiz kabul edilmesini sağlar."
        ]
    },
    {
        "slug": "amortisman-yeniden-degerleme-excel",
        "h1": "Amortisman ve 2026 yeniden değerleme hesaplama Excel'i",
        "title": "Amortisman Yeniden Değerleme Excel | Excel Arşiv",
        "desc": "VUK geçici 32 ve mükerrer 298/Ç kapsamında duran varlık amortismanı, yeniden değerleme oranları ve vergi avantajı hesaplayan profesyonel Excel sistemi.",
        "cevap": "Sabit kıymetlerin amortismanını ve 2026 yılı yeniden değerleme avantajını hesaplamak isteyen işletmeler için Amortisman + 2026 Yeniden Değerleme (499 TL, 14 sayfa) geliştirilmiştir. VUK oranlarına göre faydalı ömür amortismanını ve yeniden değerleme sonucu oluşan vergi tasarrufunu kuruşu kuruşuna doğru biçimde belirler.",
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
        "paragraphs": [
            "Duran varlıkların amortisman hesabı ve yeniden değerleme uygulamaları, işletmelerin kurumlar vergisi matrahını düşüren en güçlü yasal araçlardır. Enflasyon ortamında sabit kıymetleri tarihi maliyetle bilançoda bırakmak fiktif kâr üzerinden gereksiz vergi ödenmesine yol açar.",
            "Amortisman ve Yeniden Değerleme sistemi, VUK Geçici 32 ve Mükerrer 298/Ç maddeleri kapsamındaki yeniden değerleme oranlarını duran varlık kartlarına uygular. Artan net bilanço değeri üzerinden ayrılacak yeni amortisman gideri hesaplanır.",
            "Gayrimenkul, makine tesisat ve binek araç gruplarının amortisman kısıtlamaları ve vergi kalkanı avantajları ayrı tablolarda izlenir. Şirketin sağlayacağı net kurumlar vergisi tasarrufu kuruşu kuruşuna ortaya konur.",
            "Sabit kıymet satış senaryoları simülatörü, bir duran varlığın elden çıkarılması durumunda ödenecek kurumlar vergisini ve yenileme fonu ayırma avantajını karşılaştırmalı olarak analiz eder.",
            "Duran varlık satışlarında oluşacak kâr veya zarar, ayrılmış kümülatif amortismanlar düşülerek netleştirilir. Yenileme fonu ayrılması durumunda vergi erteleme avantajı modellenir.",
            "Enflasyon düzeltmesi ile Mükerrer 298/Ç yeniden değerleme arasındaki vergi etkisi farkları karşılaştırılarak şirketin mali tabloları için en rasyonel muhasebe politikası belirlenir."
        ]
    },
    {
        "slug": "kidem-ihbar-maliyeti-excel",
        "h1": "Kıdem ve ihbar tazminatı hesaplama Excel'i: Personel maliyeti",
        "title": "Kıdem İhbar Maliyeti Hesaplama Excel | Excel Arşiv",
        "desc": "Personel çıkarma maliyeti, kıdem tazminatı tavanı, ihbar süresi ve kullanılmayan yıllık izin karşılıklarını hatasız hesaplayan profesyonel Excel modeli.",
        "cevap": "İşten ayrılma süreçlerinde kıdem, ihbar ve yıllık izin karşılıklarını kuruşu kuruşuna hesaplamak için Kıdem–İhbar Yükü ve Personel Çıkarma Maliyeti Hesaplayıcı (799 TL, 15 sayfa) uygundur. Güncel kıdem tavanını ve brüt giydirilmiş ücret kalemlerini uygulayarak net tazminatı kuruşu kuruşuna verir.",
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
        "paragraphs": [
            "Personel işten çıkış süreçlerinde kıdem ve ihbar tazminatlarının yanlış hesaplanması şirketleri iş mahkemelerinde ve arabuluculuk süreçlerinde ağır vekalet ücreti ve yasal faiz yüküyle karşı karşıya bırakır. Giydirilmiş brüt ücrete sosyal hakların dahil edilmemesi en yaygın hatadır.",
            "Kıdem ve İhbar Maliyeti Hesaplayıcı, işe giriş ve çıkış tarihlerine göre fiili çalışma süresini gün bazında çıkarır. Çıplak ücrete yol, yemek, ikramiye ve düzenli primleri ekleyerek güncel Hazine kıdem tavanını uygular.",
            "İhbar öneli süresi, kullanılmayan yıllık izin karşılığı, damga vergisi ve gelir vergisi kesintileri mevzuata tam uyumlu olarak düşülür. Personele ödenecek net tutar ile şirkete toplam maliyet tek bir bordro özetinde sunulur.",
            "Arabuluculuk müzakere modülü, olası bir dava durumunda ortaya çıkabilecek dava harçları, avukatlık ücretleri ve faiz risklerini hesaplayarak şirketin anlaşabileceği asgari ve azami uzlaşma sınırlarını belirler.",
            "İhbar sürelerinin kullandırılması veya peşin ödenmesi durumları nakit akışı üzerinde karşılaştırmalı olarak test edilir. Şirketin o aydaki likidite dengesine göre en uygun çıkış takvimi oluşturulur.",
            "Tüm çalışanların kıdem tazminatı karşılıkları bilanço dönemi sonlarında toplu olarak hesaplanabilir. Şirketin gelecekte karşılaşacağı toplam tazminat yükü fon ihtiyacı olarak önceden bütçelenir."
        ]
    },
    {
        "slug": "sgk-tesvik-optimizasyon-excel",
        "h1": "SGK teşvik hesaplama Excel'i: Kaçırılan prim avantajı",
        "title": "SGK Teşvik Hesaplama Excel Şablonu | Excel Arşiv",
        "desc": "6111, 7103 ve 5510 sayılı SGK istihdam teşviklerini karşılaştırıp işletmeniz için en avantajlı prim desteğini seçen profesyonel optimizasyon Excel modeli.",
        "cevap": "Personel istihdamında en avantajlı SGK prim teşvikini belirlemek ve kaçırılan primleri yakalamak isteyen işletmeler için Kaçırılan SGK Teşvikleri ve Gerçek İşçilik Maliyeti Analizi (999 TL, 17 sayfa) en uygun sistemdir. Personel bazında yasal teşvik alternatiflerini kıyaslar ve prim kazancını raporlar.",
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
        "paragraphs": [
            "Türkiye'de uygulanan istihdam teşvikleri işletmelere ciddi işçilik maliyeti avantajı sağlar. Ancak 6111, 7103, 5510 gibi çok sayıda farklı kanunun bulunması ve ortalama personel sayısı şartı şirketlerin yüksek tutarlı prim avantajlarını kaçırmasına neden olur.",
            "Kaçırılan SGK Teşvikleri modeli, çalışanların yaş, cinsiyet, mesleki yeterlilik ve işsizlik geçmişi verilerini işleyerek her personel için en yüksek prim desteği sağlayan kanunu otomatik olarak seçer.",
            "İşe yeni alınacak adaylar için mülakat aşamasında teşvik uygunluk simülasyonu yapılır. Teşvikli bir adayın istihdam edilmesi durumunda şirketin sağlayacağı aylık ve yıllık bordro tasarrufu hesaplanır.",
            "Personel bazlı teşvik süreleri ve bitiş takvimi otomatik izlenerek teşvik süresi dolan personelin bir sonraki uygun kanuna aktarılması sağlanır. Şirketin prim yükü kalıcı olarak düşürülür.",
            "Personel devir hızının yüksek olduğu sektörlerde teşvik sürelerinin bitiş tarihleri otomatik takip edilerek prim kayıpları engellenir.",
            "İşletmenin SGK borçsuzluk ve düzenli prim ödeme şartları denetlenerek indirim hakkının kaybedilmemesi için finans takvimine uyarılar eklenir."
        ]
    },
    {
        "slug": "restoran-kafe-maliyet-excel",
        "h1": "Restoran ve kafe reçete maliyeti hesaplama Excel'i",
        "title": "Restoran Kafe Reçete Maliyet Excel | Excel Arşiv",
        "desc": "Restoran, kafe ve mutfak işletmeleri için porsiyon reçete maliyeti, fire oranı ve menü kârlılık analizi sunan profesyonel Excel karar sistemi.",
        "cevap": "Restoran ve kafelerde porsiyon maliyetini, hammadde firelerini ve menü kâr marjlarını denetlemek isteyen işletmeler için Restoran Reçete Maliyet ve Fire Sistemi (499 TL, 14 sayfa) idealdir. Gramaj bazlı hammadde fiyatlarını menü satış fiyatıyla eşleştirir ve porsiyon kârlılığını kuruşu kuruşuna hesaplar.",
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
        "paragraphs": [
            "Yeme-içme sektöründe menü kârlılığı mutfakta başlar. Temel gıda maddelerindeki fiyat artışları menü satış fiyatlarına anında yansıtılmadığında işletmeler yüksek ciro yapmalarına rağmen kâr marjlarını kaybeder.",
            "Restoran Reçete Maliyet ve Fire Sistemi, her porsiyon için kullanılan malzemeleri gramaj bazında reçetelendirir. Ayıklama, pişme ve porsiyonlama fire oranlarını formüle dahil ederek tabağın net hammadde maliyetini çıkarır.",
            "Menü mühendisliği analizi, satış adetleri ile ürün kâr marjlarını matriste birleştirerek en çok kazandıran yıldız ürünleri ve menüden çıkarılması gereken zararlı tabakları belirler.",
            "Hammadde alış fiyatlarındaki değişiklikler ana girdi tablosuna işlendiğinde tüm menünün porsiyon maliyeti ve kâr marjı saniyeler içinde güncellenir. Mutfak kayıp ve kaçakları anında kontrol altına alınır.",
            "Haftalık tedarik siparişleri reçete girdi miktarlarına göre otomatik hesaplanır. Depoda atıl hammadde birikmesi ve bozulma kaynaklı zayiatlar önlenir.",
            "Mutfak fiili sayım sonuçları ile satış adetlerinden hesaplanan teorik hammadde sarfiyatı karşılaştırılır. Porsiyon aşımı ve personel kaçakları anında tespit edilir."
        ]
    },
    {
        "slug": "insaat-hakedis-excel",
        "h1": "İnşaat hakediş ve şantiye maliyeti takip Excel'i",
        "title": "İnşaat Hakediş Takip Excel Şablonu | Excel Arşiv",
        "desc": "Müteahhit ve taşeronlar için yeşil defter, hakediş icmali, fiyat farkı ve şantiye maliyet kontrolü sağlayan profesyonel Excel hesaplama sistemi.",
        "cevap": "İnşaat projelerinde şantiye harcamalarını, taşeron ödemelerini ve hakediş kesintilerini yönetmek isteyen müteahhitler için İnşaat Hakediş ve Şantiye Maliyet Sistemi (799 TL, 16 sayfa) geliştirilmiştir. İmalat metrajlarını ve yasal stopaj kesintilerini düzenli hakediş raporuna dönüştürür ve kontrolü eksiksiz sağlar.",
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
        "paragraphs": [
            "İnşaat ve taahhüt projelerinde şantiye harcamalarının ve taşeron hakedişlerinin denetlenmesi proje kârlılığını belirleyen temel unsurdur. Yapılmayan imalatların hakedişe yazılması veya avans kesintilerinin unutulması projeleri zarara sürükler.",
            "İnşaat Hakediş Sistemi, yeşil defter ve metraj cetvellerini doğrudan hakediş icmaline bağlar. İmalat pursantajları, ihzarat, fiyat farkı, stopaj ve teminat kesintileri yasal kurallara uygun olarak düşülür.",
            "Şantiye malzeme alımları ve makine yakıt giderleri proje maliyet kodlarına bağlanarak bütçelenen metraj maliyetleri ile gerçekleşen faturalar arasındaki sapmalar düzenli raporlanır.",
            "Taşeron mutabakat sayfaları, her taşeronun kümülatif hak edişini, ödenen avansları ve kalan kesin teminat bakiyelerini şeffaf şekilde tutarak iş teslimindeki anlaşmazlıkları önler.",
            "Proje nakit akışı ve hak ediş tahsilat takvimi eşleştirilerek malzeme tedarikçilerine yapılacak ödemeler finansman krizine yol açmadan planlanır.",
            "Resmi kamu ihaleleri için TÜİK girdi endeksleri üzerinden hesaplanan fiyat farkı cetveli, hakediş dosyasına otomatik eklenerek idareden hak kaybı yaşanmadan ödeme alınmasını sağlar."
        ]
    },
    {
        "slug": "ihale-teklif-sinir-deger-excel",
        "h1": "İhale teklif ve sınır değer hesaplama Excel'i: Kamu ihaleleri",
        "title": "İhale Teklif Sınır Değer Excel | Excel Arşiv",
        "desc": "KİK kamu ihalelerinde sınır değer, yaklaşık maliyet katsayısı ve aşırı düşük teklif savunma riski hesaplayan profesyonel Excel karar sistemi.",
        "cevap": "Kamu ve özel sektör ihalelerinde en karlı ve elenmeyen teklif tutarını belirlemek isteyen müteahhitler için İhaleye Kaç TL Teklif Vermeliyim Sistemi (999 TL, 16 sayfa) uygundur. KİK sınır değer formüllerine göre teklifinizin sınır altında kalma riskini detaylıca analiz eder ve kârı hesaplar.",
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
        "paragraphs": [
            "Kamu ihalelerinde teklif hazırlama süreci matematiksel bir optimizasyon gerektirir. Çok düşük teklif vermek aşırı düşük sorgulamasına takılarak elenme riskini doğururken, çok yüksek teklif vermek ihaleyi kaybettirir.",
            "İhaleye Kaç TL Teklif Vermeliyim sistemi, Kamu İhale Kurumu sınır değer formüllerini ve yaklaşık maliyet katsayılarını simüle eder. Muhtemel rakip teklif dağılımlarına göre sınır değer eşiğini hesaplar.",
            "Aşırı düşük sorgulama sınırının hemen üstünde kalarak ihaleyi kazanma ihtimalini en üst düzeye çıkaran optimize teklif tutarı belirlenir. İhale kârlılığı ve nakit akışı güvenceye alınır.",
            "İhale teminat mektubu komisyonları, sözleşme damga vergisi ve KİK payı gibi zorunlu ihale maliyetleri teklif kârlılık cetveline doğrudan yansıtılır.",
            "Geçmiş ihalelerdeki rakip tenzilat ortalamaları modele girilerek istatistiki kazanma ihtimali hesaplanır. Firmanın nakit akışına en uygun teklif tutarı tespit edilir.",
            "İhale komisyonu tarafından açıklanan yaklaşık maliyet ile sınır değer arasındaki hassasiyet aralığı modellenerek teklif dosyasının risk derecesi puanlanır."
        ]
    },
    {
        "slug": "stok-devir-nakit-baglanma-excel",
        "h1": "Stok takip ve devir hızı Excel'i: Nakit bağlanma analizi",
        "title": "Stok Takip ve Devir Hızı Excel | Excel Arşiv",
        "desc": "Depodaki atıl stokları, stok devir süresini ve raflarda bağlanan nakit tutarını analiz eden profesyonel finansal stok yönetim Excel sistemi.",
        "cevap": "Depodaki ürünlerin stok devir hızını, atıl kalan sermaye tutarını ve kritik sipariş seviyelerini denetlemek isteyen işletmeler için Stok, Satış ve Nakit Bağlanma Sistemi (799 TL, 15 sayfa) tasarlanmıştır. Raflarda bağlanan nakit maliyetini ve ABC ürün sınıflandırmasını kuruşu kuruşuna eksiksiz hesaplar.",
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
        "paragraphs": [
            "Depoda atıl bekleyen her stok kalemi, işletmenin nakit kasasından çekilmiş ve raflara kilitlenmiş sermayedir. Stok devir hızının düşmesi depolama maliyetlerini artırırken şirketi likidite krizine sokar.",
            "Stok ve Nakit Bağlanma Sistemi, ürün bazında satış hızını ve stokta kalma gün süresini hesaplar. ABC analizi ile cironun yüzde seksenini oluşturan kritik ürünleri öne çıkarırken ölü stokları listeler.",
            "Optimum sipariş miktarı ve emniyet stoku seviyeleri formüle edilerek satın alma bütçesi doğru ürünlere tahsis edilir. Depoda bağlanan nakit finansman maliyeti aylık bazda raporlanır.",
            "Tedarikçi teslim süreleri ve sipariş karşılama performansları izlenerek yok satma riskleri ve fazla stok maliyetleri dengelenir.",
            "Stok finansman maliyeti aylık ticari kredi faiz oranları üzerinden simüle edilir. Depoda fazla ürün tutmanın şirkete maliyeti açıkça görülür.",
            "Ürün bazında brüt kâr marjı ile stok devir hızının çarpımından elde edilen GMROI (Brüt Kâr Yatırım Getirisi) metriği, hangi ürün grubunun şirket sermayesini en verimli şekilde katladığını ortaya koyar."
        ]
    },
    {
        "slug": "sube-karlilik-analizi-excel",
        "h1": "Şube kârlılık analizi Excel'i: Bu şubeyi kapatmalı mıyım?",
        "title": "Şube Kârlılık Analizi Excel Şablonu | Excel Arşiv",
        "desc": "Çok şubeli işletmeler ve mağaza zincirleri için şube bazında brüt marj, sabit masraf dağıtımı ve başabaş noktası hesaplayan profesyonel Excel sistemi.",
        "cevap": "Birden fazla şube veya mağazanın kârlılığını, kira ve personel yükünü ve kapatma eşiğini analiz etmek isteyen yöneticiler için Şube Kârlılık ve Nakit Hesaplayıcı (999 TL, 16 sayfa) uygundur. Ortak genel giderleri dağıtarak hangi şubenin nakit tükettiğini net biçimde belirler ve raporlar.",
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
        "paragraphs": [
            "Çok şubeli işletmelerde ciro yüksekliği yanıltıcı olabilir. Yüksek ciro yapan bir şube, yüksek kira ve personel gideri nedeniyle şirketin diğer kârlı mağazalarının ürettiği nakdi tüketiyor olabilir.",
            "Şube Kârlılık Hesaplayıcı, her şubenin brüt gelirini, doğrudan işletme giderlerini ve merkezden payına düşen ortak genel yönetim masraflarını ayrıştırır. Her lokasyonun başabaş noktası net biçimde hesaplanır.",
            "Metrekare başına satış verimliliği ve personel başına ciro performansı şubeler arasında karşılaştırılır. Düşük marjlı şubelerin gider yapısı detaylı olarak analiz edilir.",
            "Şube kapatma simülasyonu, zarar eden bir lokasyonun kapatılması durumunda kurtarılacak nakit tutarı ile merkezde kalacak sabit giderleri karşılaştırarak stratejik karar desteği sunar.",
            "Hangi şubenin kâra katkı sağladığı, hangisinin ise şirkete zarar verdiği net olarak ortaya konur. Böylece kira indirimi talep etme veya şube kapatma kararları somut verilere dayandırılır. Mağaza kârlılığı güvenceye alınır.",
            "Bölgesel kârlılık haritası, hangi coğrafi bölgede yeni şube açmanın şirket genel marjını yükselteceğini yatırım geri dönüş süresi (ROI) hesaplarıyla modeller."
        ]
    },
    {
        "slug": "ttk-376-sermaye-kaybi-excel",
        "h1": "TTK 376 sermaye kaybı ve borca batıklık hesaplama Excel'i",
        "title": "TTK 376 Sermaye Kaybı Excel Şablonu | Excel Arşiv",
        "desc": "Türk Ticaret Kanunu 376. madde kapsamında sermaye kaybı, borca batıklık cetveli ve sermaye tamamlama fonu hesaplayan mevzuat uyumlu Excel sistemi.",
        "cevap": "Şirket özkaynaklarının erimesi, sermaye kaybı ve borca batıklık durumunu denetlemek isteyen şirket yöneticileri için Şirket Öz Kaynağı Eridi mi? TTK 376 Sermaye Tamamlama Cetveli (1.499 TL, 18 sayfa) tasarlanmıştır. Bilanço kalemlerinden sermaye koruma oranını net olarak verir ve çözümleri sunar.",
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
        "paragraphs": [
            "Türk Ticaret Kanunu 376. maddesi, şirket yönetim kurullarına sermaye kaybı ve borca batıklık durumunda ağır yasal sorumluluklar yükler. Sermaye ve yedek akçelerin karşılıksız kalması halinde acil genel kurul çağrısı zorunludur.",
            "TTK 376 Cetveli, şirketin bilanço kalemlerini yasal tebliğ standartlarıyla analiz eder. Yabancı para borçlardan doğan kur farkı zararlarının hesaplama dışı bırakılması opsiyonunu mevzuata uygun şekilde uygular.",
            "Sermaye koruma oranını kuruşu kuruşuna belirler ve şirketin borca batıklıktan çıkması için gereken asgari sermaye artırımı veya sermaye tamamlama fonu tutarını hesaplar.",
            "Ortaklar kuruluna sunulacak resmi durum tespit raporu ve iyileştirme tedbirleri tablosu, yasal denetimlerde şirket yöneticilerini hukuki güvence altına alır.",
            "Ara dönem bilançoları üzerinden yapılan projeksiyonlar, yıl sonu kapanışında şirketin hangi hukuki statüde yer alacağını önceden gösterir.",
            "Bağımsız denetim standartlarına uygun borca batıklık ara bilançosu tablosu, aktiflerin olası tasfiye değerleri üzerinden şirketin borçlarını karşılama gücünü matematiksel olarak belgeler."
        ]
    },
    {
        "slug": "doviz-acik-pozisyon-kur-riski-excel",
        "h1": "Döviz açık pozisyonu ve kur riski stres testi Excel'i",
        "title": "Döviz Açık Pozisyonu ve Kur Riski | Excel Arşiv",
        "desc": "Şirketinizin dövizli varlık ve yükümlülüklerini eşleştirip olası kur artışlarında kâr-zarar ve nakit stres testini hesaplayan profesyonel model.",
        "cevap": "Döviz borçları, ithalat taahhütleri ve yabancı para varlıkları arasındaki kur riskini ölçmek isteyen şirketler için Döviz Açık Pozisyonu ve Kur Riski Stres Testi (999 TL, 16 sayfa) uygundur. Dolar ve Euro kur senaryolarında şirketin maruz kalacağı kur farkı zararını simüle eder ve netleştirir.",
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
        "paragraphs": [
            "Döviz cinsinden borcu veya hammadde ithalatı olan şirketler için kur dalgalanmaları en büyük bilanço riskidir. Döviz gelirleri ile döviz borçları arasındaki vade uyuşmazlığı kur şoklarında faaliyet kârını tamamen silebilir.",
            "Döviz Açık Pozisyonu ve Kur Riski modeli, şirketin dövizli nakit, alacak, ticari borç ve banka kredisi varlıklarını tek bir tabloda toplar. Net yabancı para pozisyonunu ve kur duyarlılığını hesaplar.",
            "Dolar ve Euro için yüzde on, yirmi beş ve elli oranındaki kur artış senaryolarında şirketin maruz kalacağı kambiyo zararını ve özkaynak erimesini stres testiyle simüle eder.",
            "İhracat gelirlerinin döviz borçlarını karşılama oranı ve forward türev araçlarının bilanço koruma etkisi matematiksel modellerle test edilir.",
            "Para birimi bazında net pozisyon ayrı ayrı izlenerek çapraz kur parite riskleri de analiz edilir. İhracat gelirlerinin döviz borçlarını karşılama oranı raporlanır.",
            "Türev finansal araçlar (Forward, Opsiyon) kullanımının bilanço kur riskini ne oranda sınırlandıracağı simülasyon sayfasında maliyet ve fayda ekseninde modellenir."
        ]
    },
    {
        "slug": "logo-erp-cari-yaslandirma-excel",
        "h1": "Logo ve ERP uyumlu cari yaşlandırma raporu Excel'i",
        "title": "Logo ERP Cari Yaşlandırma Excel | Excel Arşiv",
        "desc": "Logo ve kurumsal ERP muhasebe programı çıktılarından dinamik cari yaşlandırma, ortalama vade ve tahsilat karar motoru üreten profesyonel Excel sistemi.",
        "cevap": "Logo veya diğer ERP yazılımlarından alınan cari hareket dökümlerini analiz edip tahsilat kararı üretmek isteyen yöneticiler için Logo/ERP Uyumlu Cari Yaşlandırma ve Tahsilat Karar Motoru (1.499 TL, 18 sayfa) tasarlanmıştır. Ham veriyi yaşlandırma raporuna çevirir ve tahsilatı hızlandırır.",
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
        "paragraphs": [
            "Logo ve kurumsal ERP muhasebe programları yoğun hareket kaydı tutabilir ancak şirket yöneticilerine doğrudan karar aldıracak dinamik yaşlandırma özetlerini her zaman pratik olarak sunamaz.",
            "Logo Uyumlu Cari Yaşlandırma Motoru, muhasebe yazılımından alınan ham mizan ve muavin hareket dökümlerini işler. Borç ve alacak kapatmalarını FIFO mantığıyla yaşlandırarak dinamik tahsilat paneline dönüştürür.",
            "Müşteri bazlı ortalama tahsilat vadesi ve vade sapmaları analiz edilerek hangi müşterilerin sözleşme vadesini aştığı net biçimde listelenir.",
            "Kritik vade eşiğini aşan müşteriler için otomatik tahsilat aksiyon planı ve hukuk takip özeti oluşturulur. Finans biriminin raporlama süresi dakikalara indirilir.",
            "Muhasebe ve ERP sistemleri çok sayıda cari hareket kaydı tutabilir ancak şirket yöneticilerine doğrudan karar aldıracak yaşlandırma ve tahsilat özetlerini her zaman pratik biçimde sunamaz.",
            "Satış temsilcilerinin prim hesaplamalarında tahsilat vadelerini kriter alan performans matrisi, satış ekibinin vadesi geçen alacakları toplamasını teşvik eder."
        ]
    }
]

def render_astro_page(spec):
    slug = spec['slug']
    h1 = spec['h1'].strip()
    title = spec['title'].strip()
    desc = spec['desc'].strip()
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
    paragraphs = spec['paragraphs']
    
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

const primarySlug = {json.dumps(primary_slug, ensure_ascii=False)};
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
        {{ '@type': 'ListItem', 'position': 3, 'name': {json.dumps(h1, ensure_ascii=False)}, 'item': pageUrl }}
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
      'name': {json.dumps(h1, ensure_ascii=False)},
      'itemListElement': {item_json}
    }}
  ]
}};
---

<CommerceLayout
  title={json.dumps(title, ensure_ascii=False)}
  description={json.dumps(desc, ensure_ascii=False)}
>
  <Fragment slot="head">
    <link rel="canonical" href={{pageUrl}} />
    <script type="application/ld+json" set:html={{JSON.stringify(jsonLd)}} />
  </Fragment>

  <main class="mx-auto max-w-4xl px-4 py-12">
    <nav class="text-xs text-neutral-500 mb-4 font-mono" aria-label="Breadcrumb">
      <a href="/" class="hover:underline">Ana Sayfa</a> / <a href="/karar" class="hover:underline">Kararlar</a> / <span>{slug}</span>
    </nav>

    <article>
      <header class="border-b border-neutral-200 pb-4 mb-6">
        <h1 class="text-3xl sm:text-4xl font-extrabold text-neutral-900 tracking-tight leading-tight">
          {h1}
        </h1>
        <div class="mt-3 flex items-center gap-3 text-xs text-neutral-500 font-mono">
          <time datetime="2026-08-20">Son güncelleme: 2026-08-20</time>
          <span>·</span>
          <span>Excel Arşiv Finansal Denetim Masası</span>
        </div>
      </header>

      {{/* CEVAP BLOĞU (40-60 KELİME) */}}
      <div class="p-5 bg-emerald-50 border-l-4 border-emerald-800 my-6">
        <p class="text-base text-emerald-950 font-medium leading-relaxed">
          {cevap}
        </p>
      </div>

      {{/* KARAR TABLOSU */}}
      <section class="my-8" aria-label="Karar Tablosu">
        <h2 class="text-xl font-bold text-neutral-900 mb-3">Kullanım Durumuna Göre Sistem Karşılaştırması</h2>
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
      <section class="my-8 border border-neutral-300 bg-neutral-50 p-6 sm:p-8" aria-label="Birincil Öneri">
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
      <section class="my-8" aria-label="Alternatif Sistemler">
        <h2 class="text-xl font-bold text-neutral-900 mb-3">Alternatif Karar Seçenekleri</h2>
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
      <section class="my-8 border-l-4 border-amber-600 bg-amber-50/50 p-6" aria-label="Dürüst Sınır">
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
      <section class="prose prose-neutral max-w-none my-8 text-sm sm:text-base leading-relaxed text-neutral-700 space-y-4">
        <h2 class="text-xl font-bold text-neutral-900">Karar Modelinin Metodolojisi ve Uygulama Alanları</h2>
"""
    for p in paragraphs:
        code += f"""        <p>{p}</p>
"""

    code += f"""      </section>

      {{/* SSS */}}
      <section class="my-8" aria-label="Sık Sorulan Sorular">
        <h2 class="text-xl font-bold text-neutral-900 mb-3">Sık Sorulan Sorular</h2>
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
      <section class="my-8 bg-neutral-50 border border-neutral-200 p-6" aria-label="İlgili Kararlar">
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
      <footer class="my-10 text-center border-t border-neutral-200 pt-8">
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

for spec in karar_pages:
    assert len(spec['title']) <= 60, f"Title too long: {spec['slug']} ({len(spec['title'])})"
    assert 140 <= len(spec['desc']) <= 160, f"Desc invalid: {spec['slug']} ({len(spec['desc'])})"
    c_words = len(kelimeler(spec['cevap']))
    assert 40 <= c_words <= 60, f"Cevap word count invalid: {spec['slug']} ({c_words})"
    
    filepath = f"src/pages/karar/{spec['slug']}.astro"
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(render_astro_page(spec))

print(f"Generated all {len(karar_pages)} pages successfully.")
