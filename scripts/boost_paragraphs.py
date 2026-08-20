#!/usr/bin/env python3
"""
boost_paragraphs.py — Adds 2 rich domain paragraphs to each karar page, bringing all to 780-920 words.
"""

EXTRA_BOOST = {
    "hangi-excel-sistemini-almaliyim": [
        "Finansal model seçimi yaparken şirketin işlem sıklığı ile raporlama periyodu arasındaki denge iyi kurgulanmalıdır. Günlük perakende satışı olan işletmeler için anlık kasa ve POS modelleri öncelik taşırken, proje bazlı çalışan taahhüt firmaları için hakediş ve sınır değer modelleri ön plana çıkar. Excel Arşiv sistemleri her sektörün kendine özgü operasyonel ritmine uyum sağlayacak esneklikte tasarlanmıştır.",
        "Ayrıca paket sistem tercihi yapan işletmeler, kasa, nakit akışı, bütçe ve bilanço modelleri arasında standart veri akışı sağlayarak kurum içi finansal entegrasyonu en düşük maliyetle tamamlar. Açık formül mimarisi sayesinde tüm modeller birbirine formülle bağlanabilir."
    ],
    "kobi-nakit-akisi-excel": [
        "Nakit akışı modellerinde senaryo duyarlılığı, işletmenin piyasa dalgalanmalarına karşı sigortasıdır. Faiz oranlarındaki artışlar veya müşteri vadelerindeki uzamalar nakit projeksiyonuna anında yansıtılarak şirketin borçlanma ihtiyacı haftalık takvimde önceden görülür. Bu disiplin KOBİ'lerin bankalar karşısında güçlü kalmasını sağlar."
    ],
    "kasa-defteri-excel": [
        "Kasada tutulan nakit tutarının günlük olarak denetlenmesi, personelin sorumluluk bilincini artırır ve şirket içi iç kontrol sisteminin temel taşını oluşturur. Günlük kasa raporları Excel üzerinden PDF olarak arşivlenerek geriye dönük denetimlerde mali müşavire eksiksiz döküm sunulur."
    ],
    "mali-musavir-cari-takip-excel": [
        "Müşteri risk analizi raporları, satış ekibi ile muhasebe departmanı arasındaki bilgi kopukluğunu giderir. Hangi müşterinin risk sınırına ulaştığı ve tahsilat yapılmadan yeni sipariş verilmemesi gerektiği tek ekranda netleşir. Bu şeffaflık şirket alacaklarının kalitesini yükseltir."
    ],
    "pos-komisyon-kontrol-excel": [
        "Banka POS komisyon kesintilerinin faturadaki hizmet bedelleriyle karşılaştırılması, yıl sonunda yüksek tutarlı haksız kesintilerin iadesini sağlar. Sanal ve fiziki POS cihazlarının kart bazlı maliyet analizleri, e-ticaret sitelerinin ödeme adımı dönüşüm oranını ve kârlılığını optimize eder.",
        "POS bloke gün sayısının nakit akışına etkisi, 13 haftalık nakit projeksiyonuna otomatik veri aktarımı ile izlenebilir. Hangi gün ne kadar net POS tahsilatının banka hesabına geçeceği finans yöneticisinin kontrolünde olur."
    ],
    "trendyol-pazaryeri-net-kar-excel": [
        "E-ticarette ürün birim kârlılığı kadar sepet ortalaması ve kargo barem sınırları da kârlılığı doğrudan etkiler. 200 TL altı ve üstü siparişlerde uygulanan satıcı kargo baremleri, ürün satış fiyatının doğru barem eşiğinde tutulmasını zorunlu kılar. Sistem bu optimizasyonu otomatik hesaplar.",
        "Pazaryeri komisyon iadeleri ve müşteri iptalleri sonrası geri dönen komisyon tutarları da mutabakat tablosunda denetlenerek platformun satıcıya eksik iade yapması tamamen önlenir."
    ],
    "kdv-iade-dosyasi-excel": [
        "KDV iade listelerindeki fatura sıra ve seri numarası formatlarının GİB İnternet Vergi Dairesi sistemine tam uyumu, liste onay aşamasında oluşabilecek teknik yükleme hatalarını ortadan kaldırır. İndirilecek ve yüklenilen listeler dakikalar içinde hazır hale getirilir.",
        "İade tutarının mahsuben elektrik, SGK ve vergi borçlarına aktarılması süreçlerinde vergi dairesinin talep edeceği mahsup talep dilekçesi ekleri mevzuat formatında otomatik doldurulur."
    ],
    "amortisman-yeniden-degerleme-excel": [
        "Amortisman cetvellerinde faydalı ömür ve amortisman oranlarının VUK tebliğlerine tam uyumu, vergi denetimlerinde usulsüzlük cezası riskini sıfıra indirir. Binek otomobillerdeki gider kısıtlaması tavanları her yıl güncellenen yasal tutarlara göre otomatik revize edilir.",
        "Duran varlıkların enflasyon düzeltmesi sonrasındaki yeni değerleri üzerinden Mükerrer 298/Ç değerlemesi yapılması, şirketin kurumlar vergisi kalkanını kalıcı olarak güçlendirir."
    ],
    "kidem-ihbar-maliyeti-excel": [
        "Kıdem tazminatı tavanı yılda iki kez Hazine ve Maliye Bakanlığı tarafından güncellendiğinde sistemdeki tavan hücresi değiştirilerek tüm hesaplamalar anında yeni tavana uyarlanır. İhbar öneli süreleri İş Kanunu'ndaki kıdem baremlerine göre hatasız uygulanır.",
        "İşten ayrılan personele imzalatılacak ibraname ve tazminat hesap pusulası sayfası, yasal geçerliliği olan standart formatta çıktı almaya hazır olarak tasarlanmıştır."
    ],
    "sgk-tesvik-optimizasyon-excel": [
        "İstihdam teşviklerinin işletme bütçesine sağladığı kazanç, personel bazında sağlanan prim tasarrufunun yıllık toplamda yüksek tutarlara ulaşmasıyla şirket kârlılığına doğrudan katkı sağlar. Teşvikli işe alım stratejisi insan kaynakları planlamasının merkezine oturur.",
        "Geçmişe dönük teşvik sorgulamaları ile şirketin son dönemde kaçırdığı prim avantajları simüle edilerek SGK'dan geriye dönük mahsup talebinde bulunma imkanı araştırılır."
    ],
    "restoran-kafe-maliyet-excel": [
        "Reçete porsiyon maliyetlerinde gramaj standardı kurmak, şef ve mutfak personelinin keyfi porsiyonlama yapmasını engeller. Her tabağın net maliyeti ve hedeflenen kâr marjı menü fiyat listesine temel oluşturur.",
        "İçecek, tatlı ve ana yemek gruplarının kârlılık dağılımı özet grafikte analiz edilerek garsonların yüksek marjlı ürünleri önermesi için satış prim sistemi kurgulanabilir."
    ],
    "insaat-hakedis-excel": [
        "Şantiye maliyet kontrolünde yeşil defter ve ataşman kayıtlarının hakediş raporlarına günü gününe bağlanması, proje bitiminde taşeronlarla yaşanacak hukuki ihtilafları kökünden çözer. İmalat metrajları şeffafça onaylanır.",
        "Proje bazında gerçekleşen malzeme, işçilik ve taşeron harcamaları bütçelenen birim fiyatlarla kıyaslanarak hangi imalat kaleminde maliyet aşımı yaşandığı anında tespit edilir."
    ],
    "ihale-teklif-sinir-deger-excel": [
        "Kamu ihalelerinde yaklaşık maliyet tahmininin doğru modellenmesi, sınır değer katsayısı (R) üzerinden hesaplanan sınır değerin hangi aralığa oturacağını yüksek isabetle tahmin etmeyi sağlar. Müteahhit teklif zarfını bu veriye göre hazırlar.",
        "Aşırı düşük teklif sorgulaması gelmesi durumunda analiz formatında malzeme, işçilik ve makine analizleri KİK tebliğ kurallarına göre hızlıca savunma dosyasına dönüştürülür."
    ],
    "stok-devir-nakit-baglanma-excel": [
        "Stok devir süresi uzayan ürünler için erken indirim ve tasfiye kampanyaları düzenlemek, depoda kilitli kalan nakdin kurtarılarak yüksek devirli kârlı ürünlere yatırılmasını sağlar. Şirketin likiditesi rahatlar.",
        "Tedarikçi bazında sipariş teslim süreleri ve gecikme oranları takip edilerek emniyet stoku seviyeleri optimize edilir. Depo kiralama ve finansman maliyetleri asgari düzeye indirilir."
    ],
    "sube-karlilik-analizi-excel": [
        "Şube bazında kira, personel, elektrik ve yerel pazarlama giderleri doğrudan şube kâr-zarar tablosuna yazılırken merkez genel yönetim giderleri belirlenen ciro payına göre adilce paylaştırılır. Her lokasyonun net katkısı görülür.",
        "Zarar eden şubelerin kapatılması veya kira revizyonu yapılması süreçlerinde somut finansal analiz tabloları mülk sahipleri ve yönetim kurulu ile yapılan müzakerelerde temel dayanak olur."
    ],
    "ttk-376-sermaye-kaybi-excel": [
        "Sermaye kaybı ve borca batıklık hesaplamalarında tebliğde tanınan kur farkı zararlarının hesaplama dışı bırakılması imkanı şirketleri gereksiz sermaye artırımı yapmaktan veya tasfiye riskinden korur.",
        "Ortaklar kuruluna sunulacak TTK 376 durum raporu, şirketin yasal durumunu ve sermaye tamamlama fonu ihtiyacını net bir şekilde ortaya koyarak yöneticilerin hukuki sorumluluğunu güvenceye alır."
    ],
    "doviz-acik-pozisyon-kur-riski-excel": [
        "Döviz açık pozisyonunun bilanço üzerindeki kur baskısı, farklı kur senaryolarında şirketin kâr-zarar tablosuna ve özkaynaklarına etkileriyle ayrıntılı olarak raporlanır. Finansal risk yönetimi profesyonel boyuta taşınır.",
        "Forward ve vadeli döviz alım sözleşmelerinin maliyeti ile olası kur artışındaki zarar karşılaştırılarak hedging yapmanın rasyonel olup olmadığı matematiksel olarak ortaya konur."
    ],
    "logo-erp-cari-yaslandirma-excel": [
        "Logo ve diğer kurumsal ERP sistemlerinden alınan muavin dökümleri saniyeler içinde dinamik cari yaşlandırma paneline dönüştürülür. Müşteri bazlı ortalama tahsilat vadesi ve geciken alacaklar listelenir.",
        "Tahsilat ekibine verilecek günlük ve haftalık arama listeleri risk skorlarına göre önceliklendirilerek şirketin nakit tahsilat hızı ve operasyonel verimliliği maksimize edilir."
    ]
}

# Run expander
import scripts.build_all_karar_pages as b

for page in b.karar_pages:
    slug = page['slug']
    if slug in EXTRA_BOOST:
        page['paragraphs'].extend(EXTRA_BOOST[slug])

for page in b.karar_pages:
    code = b.render_astro_page(page)
    slug = page['slug']
    with open(f'src/pages/karar/{slug}.astro', 'w', encoding='utf-8') as f:
        f.write(code)

print('Successfully boosted all 18 pages!')
