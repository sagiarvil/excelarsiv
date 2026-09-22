'use strict';

/**
 * Proof Demo sözleşmeleri.
 * Bu dosya yalnızca demo değerini göstermek için kullanılan basitleştirilmiş hesapları içerir.
 * Premium MOTOR / AYARLAR / eşik seti / analitik formüller burada bulunmaz.
 */

const SPECS = Object.freeze({
  "13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi": {
    "karar": "Hangi haftada nakit açığı oluşacağını ve ne kadar önlem gerektiğini gösterir.",
    "girisBasliklari": [
      "Hafta",
      "Açılış (₺)",
      "Gelen (₺)",
      "Giden (₺)",
      "Kapanış (₺)"
    ],
    "ornek": [
      [
        "Hafta 1",
        25000,
        42000,
        38500,
        "=B6+C6-D6"
      ],
      [
        "Hafta 2",
        "=E6",
        38000,
        51000,
        "=B7+C7-D7"
      ],
      [
        "Hafta 3",
        "=E7",
        45000,
        36000,
        "=B8+C8-D8"
      ],
      [
        "Hafta 4",
        "=E8",
        30000,
        44000,
        "=B9+C9-D9"
      ],
      [
        "Hafta 5",
        "=E9",
        52000,
        39000,
        "=B10+C10-D10"
      ],
      [
        "Hafta 6",
        "=E10",
        34000,
        47000,
        "=B11+C11-D11"
      ]
    ],
    "metrikler": [
      [
        "En düşük kapanış",
        "=MIN(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Negatif hafta sayısı",
        "=COUNTIF(DEMO_GIRIS!E6:E25,\"<0\")",
        "sayi"
      ],
      [
        "Dönem net nakit",
        "=SUM(DEMO_GIRIS!C6:C25)-SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Gerekli asgari önlem",
        "=MAX(0,-B6)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B7>0,\"DURDUR\",IF(B6<10000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Negatif haftanın ödemelerini önceki haftalara yay.",
      "Tahsilatı öne çekebileceğin müşterileri ayır.",
      "Tam sürümde 13 hafta, senaryo, duyarlılık ve aksiyon motoru birlikte çalışır."
    ]
  },
  "akilli-kasa-defteri-ve-nakit-kontrol-sistemi": {
    "karar": "Kasadaki kullanılabilir para, açık riski ve ödeme baskısını gösterir.",
    "girisBasliklari": [
      "Tarih",
      "Açıklama",
      "Gelir (₺)",
      "Gider (₺)",
      "Yakın ödeme (₺)"
    ],
    "ornek": [
      [
        "01.08.2026",
        "Günlük satış",
        18500,
        0,
        0
      ],
      [
        "02.08.2026",
        "Tedarikçi ödemesi",
        0,
        7200,
        0
      ],
      [
        "03.08.2026",
        "Müşteri tahsilatı",
        9600,
        0,
        0
      ],
      [
        "04.08.2026",
        "Kira",
        0,
        5200,
        0
      ],
      [
        "05.08.2026",
        "Yaklaşan vergi",
        0,
        0,
        12000
      ]
    ],
    "metrikler": [
      [
        "Toplam gelir",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Toplam gider",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Yakın ödeme",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Kullanılabilir nakit",
        "=B6-B7-B8",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9<0,\"DURDUR\",IF(B9<5000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Negatif kullanılabilir nakitte yeni ödeme taahhüdü verme.",
      "Yakın ödemeleri tahsilat tarihleriyle eşleştir.",
      "Tam sürümde kasa farkı, anomali ve ödeme önerisi birlikte çalışır."
    ]
  },
  "amortisman-2026-yeniden-degerleme": {
    "karar": "Amortisman tutarı ile 298/Ç ve Geç.32 yeniden değerleme kararını birlikte gösterir.",
    "girisBasliklari": [
      "Kıymet",
      "Maliyet (₺)",
      "Amortisman (₺)",
      "YD artış (₺)",
      "Net etki (₺)"
    ],
    "ornek": [
      [
        "Makine A",
        500000,
        100000,
        80000,
        180000
      ],
      [
        "Bina B",
        2000000,
        40000,
        250000,
        290000
      ],
      [
        "Taşıt C",
        800000,
        160000,
        0,
        160000
      ],
      [
        "Demirbaş",
        120000,
        24000,
        15000,
        39000
      ]
    ],
    "metrikler": [
      [
        "Maliyet toplam",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Amortisman",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "YD artış",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Net etki",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B8>200000,\"YAP\",IF(B8>0,\"BEKLE\",\"YAPMA\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "YD artışı yüksek kıymetlerde 298/Ç vs Geç.32 karşılaştırın.",
      "Amortisman cetvelini yıl sonu kapanışına kilitleyin.",
      "Tam sürümde kural yılı ve vaka doğrulaması açılır."
    ]
  },
  "amortisman-ve-sabit-kiymet-satis-zamanlama-stratejisti": {
    "karar": "Normal/azalan bakiyeli amortisman ve çeyrek bazlı satış senaryosunun net nakit etkisini gösterir.",
    "girisBasliklari": [
      "Kıymet",
      "Maliyet (₺)",
      "Ömür (yıl)",
      "Satış bedeli (₺)",
      "Net nakit (₺)"
    ],
    "ornek": [
      [
        "Makine 1",
        1200000,
        8,
        680000,
        512000
      ],
      [
        "Makine 2",
        640000,
        6,
        290000,
        205000
      ],
      [
        "Taşıt 1",
        480000,
        5,
        185000,
        132000
      ],
      [
        "Bina 1",
        3200000,
        25,
        2450000,
        1900000
      ]
    ],
    "metrikler": [
      [
        "Toplam maliyet",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam satış bedeli",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Toplam net nakit",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Kâr oranı",
        "=IFERROR(B7/B6,0)",
        "oran"
      ],
      [
        "Demo karar",
        "=IF(B8<=0,\"DURDUR\",IF(B9<0.25,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "En yüksek net nakitli çeyreği satış dönemi olarak değerlendir.",
      "Vergi etkisini güncel oranla doğrula.",
      "Tam sürümde kıst amortisman, NBD ve dönem karşılaştırması açılır."
    ]
  },
  "asgari-ucret-zam-etkisi-fiyat-ayarlama-cetveli": {
    "karar": "Asgari ücret zammının işçilik maliyetine ve fiyat ayarlamasına etkisini gösterir.",
    "girisBasliklari": [
      "Ürün",
      "Satış fiyatı (₺)",
      "İşçilik payı (₺)",
      "Zam etkisi %",
      "Yeni fiyat (₺)"
    ],
    "ornek": [
      [
        "Ürün A",
        890,
        310,
        0.1,
        "=B6*(1+D6)"
      ],
      [
        "Ürün B",
        1490,
        520,
        0.1,
        "=B7*(1+D7)"
      ],
      [
        "Ürün C",
        640,
        220,
        0.1,
        "=B8*(1+D8)"
      ],
      [
        "Ürün D",
        1990,
        640,
        0.1,
        "=B9*(1+D9)"
      ]
    ],
    "metrikler": [
      [
        "Toplam işçilik payı",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Zam sonrası işçilik payı",
        "=B6*(1+AVERAGE(DEMO_GIRIS!D6:D25))",
        "para"
      ],
      [
        "Zam kaynaklı maliyet artışı",
        "=B7-B6",
        "para"
      ],
      [
        "Ortalama yeni fiyat",
        "=AVERAGE(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B8/B6>0.15,\"DURDUR\",IF(B8/B6>0.05,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Zam etkisini ürün bazlı fiyata yansıt.",
      "İşçilik yoğun ürünlerde fiyat artışını önceliklendir.",
      "Tam sürümde çalışan bazlı maliyet, SGK yükü ve senaryo motoru açılır."
    ]
  },
  "asiri-dusuk-teklif-savunma-robotu": {
    "karar": "Aşırı düşük teklif açıklamasında sınırın altında kalan tutarı ve açıklama yeterliliğini gösterir.",
    "girisBasliklari": [
      "Kalem",
      "Maliyet (₺)",
      "Belge durumu",
      "Açıklama dahil mi",
      "Not"
    ],
    "ornek": [
      [
        "Malzeme",
        185000,
        "Belgeli",
        "Evet",
        ""
      ],
      [
        "İşçilik",
        96000,
        "Belgeli",
        "Evet",
        ""
      ],
      [
        "Makine",
        42000,
        "Belgesiz",
        "Hayır",
        ""
      ],
      [
        "Nakliye",
        18000,
        "Belgeli",
        "Evet",
        ""
      ],
      [
        "Genel gider",
        30000,
        "Belgesiz",
        "Hayır",
        ""
      ]
    ],
    "metrikler": [
      [
        "Toplam açıklama maliyeti",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Açıklama kapsamındaki tutar",
        "=SUMPRODUCT(--(DEMO_GIRIS!D6:D25=\"Evet\"),DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Belgesiz kalem tutarı",
        "=SUMPRODUCT(--(DEMO_GIRIS!C6:C25=\"Belgesiz\"),DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Açıklama kapsama oranı",
        "=IFERROR(B7/B6,0)",
        "oran"
      ],
      [
        "Demo karar",
        "=IF(B8>B6*0.25,\"DURDUR\",IF(B9<0.75,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Belgesiz ve kapsam dışı kalemleri açıklamaya ekle.",
      "Açıklama tutarını sınır değerin altında kalan bölümle karşılaştır.",
      "Tam sürümde sınır hesapları, endeks farkı ve ihale kararı motoru birlikte çalışır."
    ]
  },
  "aylik-patron-finans-paneli": {
    "karar": "İşletmenin finansal durumunu tek ekranda yorumlar.",
    "girisBasliklari": [
      "Gösterge",
      "Tutar (₺)",
      "Not",
      "Dönem",
      "Kaynak"
    ],
    "ornek": [
      [
        "Nakit",
        185000,
        "",
        "Ağustos",
        "Kasa+Banka"
      ],
      [
        "Alacak",
        420000,
        "",
        "Ağustos",
        "Cari"
      ],
      [
        "Stok",
        310000,
        "",
        "Ağustos",
        "Stok"
      ],
      [
        "Kısa borç",
        560000,
        "",
        "Ağustos",
        "Borç"
      ],
      [
        "Aylık gelir",
        740000,
        "",
        "Ağustos",
        "Satış"
      ],
      [
        "Aylık gider",
        675000,
        "",
        "Ağustos",
        "Muhasebe"
      ]
    ],
    "metrikler": [
      [
        "Toplam likit varlık",
        "=SUM(DEMO_GIRIS!B6:B8)",
        "para"
      ],
      [
        "Kısa borç",
        "=DEMO_GIRIS!B9",
        "para"
      ],
      [
        "Aylık net",
        "=DEMO_GIRIS!B10-DEMO_GIRIS!B11",
        "para"
      ],
      [
        "Likidite tamponu",
        "=B6-B7",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9<0,\"DURDUR\",IF(B8<0,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Likidite tamponunu kısa borçla birlikte değerlendir.",
      "Aylık net ile nakit değişimini aynı şey kabul etme.",
      "Tam sürümde yönetici KPI, trend, senaryo ve aksiyonlar tek sayfada birleşir."
    ]
  },
  "banka-kredi-ve-taksit-takip-sistemi": {
    "karar": "Borç yükünü, yaklaşan taksit baskısını ve refinansman ihtiyacını gösterir.",
    "girisBasliklari": [
      "Banka / kredi",
      "Kalan borç (₺)",
      "Aylık taksit (₺)",
      "Vade gün",
      "Aylık serbest nakit (₺)"
    ],
    "ornek": [
      [
        "Banka A - İşletme",
        420000,
        32000,
        5,
        95000
      ],
      [
        "Banka B - Spot",
        260000,
        24000,
        12,
        95000
      ],
      [
        "Banka C - Taşıt",
        145000,
        16000,
        20,
        95000
      ],
      [
        "Finansman - Ekipman",
        98000,
        12000,
        8,
        95000
      ],
      [
        "Banka D - Kobi Destek",
        180000,
        11000,
        25,
        95000
      ]
    ],
    "metrikler": [
      [
        "Toplam kalan borç",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Aylık toplam taksit",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "15 gün içindeki taksit sayısı",
        "=COUNTIF(DEMO_GIRIS!D6:D25,\"<=15\")",
        "sayi"
      ],
      [
        "Taksit / serbest nakit",
        "=IFERROR(B7/MAX(DEMO_GIRIS!E6:E25),0)",
        "yuzde"
      ],
      [
        "Demo karar",
        "=IF(B9>1,\"DURDUR\",IF(B9>0.6,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Taksit yükü serbest nakdi aşıyorsa yeni borçlanmayı durdur.",
      "15 gün içindeki taksitleri tahsilat planıyla eşleştir.",
      "Tam sürümde banka/kredi yoğunlaşması, takvim ve refinansman sinyali detaylanır."
    ]
  },
  "cari-ba-bs-toplu-mutabakat": {
    "karar": "Cari ile Form Ba/Bs tutarlarını eşleştirir; fark kovalarını listeler.",
    "girisBasliklari": [
      "VKN",
      "Cari (₺)",
      "Ba/Bs (₺)",
      "Fark (₺)",
      "Kova"
    ],
    "ornek": [
      [
        "1234567890",
        125000,
        125000,
        0,
        "Eşleşti"
      ],
      [
        "2345678901",
        88000,
        86000,
        2000,
        "Tutar farkı"
      ],
      [
        "3456789012",
        45000,
        0,
        45000,
        "Eşleşmedi"
      ],
      [
        "4567890123",
        67000,
        67050,
        -50,
        "Tolerans"
      ]
    ],
    "metrikler": [
      [
        "Cari toplam",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Ba/Bs toplam",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Fark mutlak",
        "=SUMPRODUCT(ABS(DEMO_GIRIS!D6:D25))",
        "para"
      ],
      [
        "Kayıt",
        "=COUNTA(DEMO_GIRIS!A6:A25)",
        "adet"
      ],
      [
        "Demo karar",
        "=IF(B8>20000,\"DURDUR\",IF(B8>0,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Eşleşmeyen VKN satırlarını cariye işleyin.",
      "Tutar farkını fatura ile kapatın.",
      "Tam sürümde dört kova ve kök neden açılır."
    ]
  },
  "cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi": {
    "karar": "Kime satış durdurulmalı ve kimden tahsilat hızlandırılmalı sorusunu yanıtlar.",
    "girisBasliklari": [
      "Müşteri",
      "Bakiye (₺)",
      "Gecikme (gün)",
      "Teminat (₺)",
      "Planlanan tahsilat (₺)"
    ],
    "ornek": [
      [
        "Alfa Yapı",
        85000,
        35,
        25000,
        40000
      ],
      [
        "Beta Tekstil",
        32000,
        7,
        0,
        12000
      ],
      [
        "Gama Gıda",
        120000,
        18,
        45000,
        30000
      ],
      [
        "Delta İnşaat",
        18000,
        0,
        0,
        18000
      ],
      [
        "Epsilon Ltd.",
        54000,
        28,
        15000,
        20000
      ]
    ],
    "metrikler": [
      [
        "Toplam açık bakiye",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "30+ gün geciken müşteri",
        "=COUNTIF(DEMO_GIRIS!C6:C25,\">30\")",
        "sayi"
      ],
      [
        "Teminat açığı",
        "=MAX(0,SUM(DEMO_GIRIS!B6:B25)-SUM(DEMO_GIRIS!D6:D25)-SUM(DEMO_GIRIS!E6:E25))",
        "para"
      ],
      [
        "En yüksek gecikme",
        "=MAX(DEMO_GIRIS!C6:C25)",
        "sayi"
      ],
      [
        "Demo karar",
        "=IF(B7>=2,\"DURDUR\",IF(B8>50000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "30+ gün geciken müşterilerde yeni vadeli satışı incele.",
      "Teminat ve planlanan tahsilatı bakiye ile karşılaştır.",
      "Tam sürümde müşteri bazlı risk sıralaması ve dinamik tahsilat aksiyonu açılır."
    ]
  },
  "cek-senet-ve-vade-risk-sistemi": {
    "karar": "Ödeme yoğunlaşmasının hangi tarihte oluştuğunu ve karşılıksız kalma riskini gösterir.",
    "girisBasliklari": [
      "Vade",
      "Belge",
      "Tutar (₺)",
      "Hazır karşılık (₺)",
      "Durum"
    ],
    "ornek": [
      [
        "12.08.2026",
        "Çek-001",
        48000,
        40000,
        "Bekliyor"
      ],
      [
        "14.08.2026",
        "Çek-002",
        72000,
        60000,
        "Bekliyor"
      ],
      [
        "14.08.2026",
        "Senet-003",
        55000,
        45000,
        "Bekliyor"
      ],
      [
        "28.08.2026",
        "Çek-004",
        32000,
        32000,
        "Hazır"
      ],
      [
        "02.09.2026",
        "Senet-005",
        64000,
        55000,
        "Bekliyor"
      ]
    ],
    "metrikler": [
      [
        "Toplam vade yükü",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Toplam hazır karşılık",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Karşılık açığı",
        "=MAX(0,B6-B7)",
        "para"
      ],
      [
        "En büyük tek belge",
        "=MAX(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B8>75000,\"DURDUR\",IF(B8>25000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Aynı haftaya yığılan vadeleri yeniden planla.",
      "Karşılığı ayrılmamış büyük belgeleri önceliklendir.",
      "Tam sürümde tarih yoğunluğu, senaryo ve kök-neden risk analizi açılır."
    ]
  },
  "defter-beyan-e-arsiv-aktarim": {
    "karar": "e-Arşiv satırlarını Defter Beyan kolon haritasına dönüştürür ve hata listesini üretir.",
    "girisBasliklari": [
      "Fatura No",
      "Tutar (₺)",
      "Map durumu",
      "Hata",
      "Aktarım"
    ],
    "ornek": [
      [
        "EA-1001",
        11800,
        "Tam",
        "",
        "Hazır"
      ],
      [
        "EA-1002",
        23600,
        "Eksik alan",
        "VKN boş",
        "Durdur"
      ],
      [
        "EA-1003",
        5900,
        "Tam",
        "",
        "Hazır"
      ],
      [
        "EA-1004",
        17700,
        "Format",
        "Tarih hatalı",
        "İncele"
      ]
    ],
    "metrikler": [
      [
        "Satır sayısı",
        "=COUNTA(DEMO_GIRIS!A6:A25)",
        "adet"
      ],
      [
        "Tutar toplam",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Hatalı satır",
        "=COUNTIF(DEMO_GIRIS!E6:E25,\"Durdur\")+COUNTIF(DEMO_GIRIS!E6:E25,\"İncele\")",
        "adet"
      ],
      [
        "Hazır satır",
        "=COUNTIF(DEMO_GIRIS!E6:E25,\"Hazır\")",
        "adet"
      ],
      [
        "Demo karar",
        "=IF(B8>2,\"DURDUR\",IF(B8>0,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Hatalı satırları kolon haritasına göre düzeltin.",
      "Hazır satırları DBS aktarımına alın.",
      "Tam sürümde API yok; yapıştır-çalıştır modelidir."
    ]
  },
  "doviz-acik-pozisyonu-ve-kur-riski-stres-testi": {
    "karar": "Dövizli varlık ve borçların net açık pozisyonunu ve kur şoku etkisini gösterir.",
    "girisBasliklari": [
      "Kalem",
      "Varlık (₺)",
      "Borç (₺)",
      "Döviz cinsi",
      "Net pozisyon (₺)"
    ],
    "ornek": [
      [
        "Mevduat",
        850000,
        0,
        "USD",
        850000
      ],
      [
        "İhracat alacağı",
        1200000,
        0,
        "EUR",
        1200000
      ],
      [
        "Kredi",
        0,
        2400000,
        "USD",
        -2400000
      ],
      [
        "Tedarikçi borcu",
        0,
        480000,
        "EUR",
        -480000
      ]
    ],
    "metrikler": [
      [
        "Toplam varlık",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam borç",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Net pozisyon",
        "=B6-B7",
        "para"
      ],
      [
        "Kur şoku etkisi (%15)",
        "=B8*0.15",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9>500000,\"DURDUR\",IF(B9>200000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Açık pozisyonu dengeleyici döviz işlemi değerlendir.",
      "Kur şoku senaryosunu AYARLAR oranıyla doğrula.",
      "Tam sürümde cins bazlı pozisyon, stres senaryoları ve karar motoru açılır."
    ]
  },
  "e-fatura-satir-defteri-pdf-kaniti": {
    "karar": "e-Fatura satırları ile muhasebe kayıtlarını eşleştirir; tutar farkı ve eksik belgeyi işaretler.",
    "girisBasliklari": [
      "UUID/No",
      "Tarih",
      "e-Fatura (₺)",
      "Muhasebe (₺)",
      "Fark (₺)"
    ],
    "ornek": [
      [
        "UUID-001",
        "01.07.2026",
        11800,
        11800,
        0
      ],
      [
        "UUID-002",
        "05.07.2026",
        23600,
        23000,
        600
      ],
      [
        "UUID-003",
        "12.07.2026",
        5900,
        5900,
        0
      ],
      [
        "UUID-004",
        "20.07.2026",
        17700,
        0,
        17700
      ]
    ],
    "metrikler": [
      [
        "e-Fatura toplam",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Muhasebe toplam",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Fark toplam",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Kayıt sayısı",
        "=COUNTA(DEMO_GIRIS!A6:A25)",
        "adet"
      ],
      [
        "Demo karar",
        "=IF(B8>10000,\"DURDUR\",IF(B8>0,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Eşleşmeyen UUID satırlarını muhasebeye işleyin.",
      "Tutar farkını belgeyle kapatın.",
      "Tam sürümde PDF kanıt raporu baskıya hazırdır."
    ]
  },
  "fazla-mesai-ve-isci-dava-riski-tespit-dosyasi": {
    "karar": "Fazla mesai karşılığını ve işçi alacağı riskini gösterir.",
    "girisBasliklari": [
      "Çalışan",
      "Brüt ücret (₺)",
      "Aylık fazla mesai (saat)",
      "Ödenen (₺)",
      "Eksik ödeme (₺)"
    ],
    "ornek": [
      [
        "Çalışan 1",
        42000,
        45,
        18000,
        3000
      ],
      [
        "Çalışan 2",
        38000,
        28,
        12000,
        0
      ],
      [
        "Çalışan 3",
        51000,
        62,
        15000,
        7000
      ],
      [
        "Çalışan 4",
        36000,
        15,
        8000,
        0
      ]
    ],
    "metrikler": [
      [
        "Toplam fazla mesai karşılığı",
        "=SUM(DEMO_GIRIS!C6:C25)* (AVERAGE(DEMO_GIRIS!B6:B25)/225)*1.5",
        "para"
      ],
      [
        "Toplam ödenen",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Toplam eksik ödeme",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Eksik ödeme oranı",
        "=IFERROR(B8/MAX(B6,1),0)",
        "oran"
      ],
      [
        "Demo karar",
        "=IF(B9>0.35,\"DURDUR\",IF(B8>8000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Eksik ödeme olan çalışanları öncelikli dava riski olarak işaretle.",
      "Aylık fazla mesai kayıtlarını sözleşmeyle doğrula.",
      "Tam sürümde saat sınırı, anomali ve dava riski motoru birlikte çalışır."
    ]
  },
  "gunluk-gelir-gider-ve-gercek-karlilik-sistemi": {
    "karar": "Ciro ile gerçek nakit kazancı arasındaki farkı gösterir.",
    "girisBasliklari": [
      "Gün",
      "Ciro (₺)",
      "Direkt gider (₺)",
      "Sabit/nakit gider (₺)",
      "Tahsil edilen (₺)"
    ],
    "ornek": [
      [
        "Pazartesi",
        42000,
        21000,
        8500,
        36000
      ],
      [
        "Salı",
        38000,
        20500,
        8500,
        25000
      ],
      [
        "Çarşamba",
        51000,
        26000,
        8500,
        47000
      ],
      [
        "Perşembe",
        33000,
        19000,
        8500,
        28000
      ],
      [
        "Cuma",
        62000,
        31500,
        8500,
        50000
      ]
    ],
    "metrikler": [
      [
        "Toplam ciro",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Gerçek nakit gider",
        "=SUM(DEMO_GIRIS!C6:C25)+SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Hesaplanan kâr",
        "=B6-B7",
        "para"
      ],
      [
        "Tahsilat-kâr farkı",
        "=SUM(DEMO_GIRIS!E6:E25)-B7",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B8<0,\"DURDUR\",IF(B9<B8*0.5,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Ciro yüksek olsa da nakit gideri aşmıyorsa büyümeyi sorgula.",
      "Tahsilat ile hesaplanan kâr arasındaki farkı izle.",
      "Tam sürümde gerçek kârlılık, nakit bağlanması ve senaryo analizi ayrıştırılır."
    ]
  },
  "hakedis-fiyat-farki-hak-kaybi-cetveli": {
    "karar": "Hakediş fiyat farkını ve endeks kaynaklı hak kaybını gösterir.",
    "girisBasliklari": [
      "Dönem",
      "Hakediş (₺)",
      "Sözleşme endeks",
      "Güncel endeks",
      "Katsayı"
    ],
    "ornek": [
      [
        "Haziran",
        240000,
        220.4,
        231.8,
        0.8
      ],
      [
        "Temmuz",
        265000,
        220.4,
        238.1,
        0.8
      ],
      [
        "Ağustos",
        250000,
        220.4,
        245.6,
        0.8
      ],
      [
        "Eylül",
        280000,
        220.4,
        252.2,
        0.8
      ]
    ],
    "metrikler": [
      [
        "Toplam hakediş",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Ortalama endeks artışı",
        "=IFERROR(AVERAGE(DEMO_GIRIS!D6:D25)/AVERAGE(DEMO_GIRIS!C6:C25)-1,0)",
        "oran"
      ],
      [
        "Hesaplanan fiyat farkı",
        "=SUM(DEMO_GIRIS!B6:B25)*(B7)*AVERAGE(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Endeks farkı oranı",
        "=IFERROR(B8/MAX(B6,1),0)",
        "oran"
      ],
      [
        "Demo karar",
        "=IF(B9>0.15,\"DURDUR\",IF(B8<10000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Endeks artışını sözleşme hükmüyle eşleştir.",
      "Hesaplanan fiyat farkını hakedişe yansıt.",
      "Tam sürümde dönem bazlı fark, hak kaybı ve rapor motoru açılır."
    ]
  },
  "ihaleye-kac-tl-teklif-vermeliyim": {
    "karar": "Rakip teklifler üzerinden sınır değeri ve önerilen teklif aralığını gösterir.",
    "girisBasliklari": [
      "Teklif veren",
      "Teklif (₺)",
      "Yaklaşık maliyet (₺)",
      "Aralık durumu",
      "Not"
    ],
    "ornek": [
      [
        "Firma A",
        1180000,
        1200000,
        "Üst bölge",
        ""
      ],
      [
        "Firma B",
        1320000,
        1200000,
        "Üst bölge",
        ""
      ],
      [
        "Firma C",
        1040000,
        1200000,
        "Alt bölge",
        ""
      ],
      [
        "Firma D",
        960000,
        1200000,
        "Aşırı düşük",
        ""
      ],
      [
        "Firma E",
        1250000,
        1200000,
        "Üst bölge",
        ""
      ]
    ],
    "metrikler": [
      [
        "Ortalama rakip teklif",
        "=AVERAGE(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Sınır değer (tahmini)",
        "=IFERROR(B6*0.92,0)",
        "para"
      ],
      [
        "Önerilen alt sınır",
        "=IFERROR(B7*0.95,0)",
        "para"
      ],
      [
        "Aşırı düşük sayısı",
        "=COUNTIF(DEMO_GIRIS!D6:D25,\"Aşırı düşük\")",
        "sayi"
      ],
      [
        "Demo karar",
        "=IF(B9>=1,\"İNCELE\",IF(B8>1200000,\"DURDUR\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Sınır değerin altındaki teklifler aşırı düşük sorgusu riski taşır.",
      "Önerilen aralığı kâr hedefinle karşılaştır.",
      "Tam sürümde önceki ihale trendi, maliyet motoru ve senaryo analizi açılır."
    ]
  },
  "insaat-hakedis-santiye-maliyet": {
    "karar": "Hakediş kuyruğu ve şantiye maliyet sapmasını tek ekranda gösterir.",
    "girisBasliklari": [
      "Hakediş",
      "Tutar (₺)",
      "Maliyet (₺)",
      "Gecikme (gün)",
      "Durum"
    ],
    "ornek": [
      [
        "H-01",
        850000,
        720000,
        0,
        "Onay"
      ],
      [
        "H-02",
        420000,
        480000,
        12,
        "Bekliyor"
      ],
      [
        "H-03",
        610000,
        590000,
        3,
        "Bekliyor"
      ],
      [
        "H-04",
        300000,
        250000,
        25,
        "Geciken"
      ]
    ],
    "metrikler": [
      [
        "Hakediş toplam",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Maliyet toplam",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Sapma",
        "=B6-B7",
        "para"
      ],
      [
        "Ort. gecikme",
        "=AVERAGE(DEMO_GIRIS!D6:D25)",
        "adet"
      ],
      [
        "Demo karar",
        "=IF(B9>15,\"KRİTİK\",IF(OR(B8<0,B9>5),\"DİKKAT\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Maliyet aşan şantiye kalemini inceleyin.",
      "Geciken hakedişleri önceliklendirin.",
      "Tam sürümde durum makinesi ve yaşlandırma açılır."
    ]
  },
  "ithalat-depo-teslim-rafa-gelen-net-birim-maliyet": {
    "karar": "İthal ürünün FOB değerinden depo teslim net birim maliyetine geçişini gösterir.",
    "girisBasliklari": [
      "Ürün",
      "FOB (₺)",
      "CIF (₺)",
      "Vergi toplam (₺)",
      "Net birim maliyet (₺)"
    ],
    "ornek": [
      [
        "Ürün A",
        2400,
        2650,
        1180,
        3960
      ],
      [
        "Ürün B",
        850,
        940,
        420,
        1490
      ],
      [
        "Ürün C",
        5200,
        5750,
        2560,
        8790
      ],
      [
        "Ürün D",
        1360,
        1490,
        660,
        2410
      ]
    ],
    "metrikler": [
      [
        "Toplam FOB",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam vergi",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Vergi / CIF oranı",
        "=IFERROR(B7/SUM(DEMO_GIRIS!C6:C25),0)",
        "oran"
      ],
      [
        "Ortalama net birim maliyet",
        "=AVERAGE(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B8>0.55,\"DURDUR\",IF(B9>3000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Vergi yükü yüksek ürünlerde fiyatlamayı yeniden değerlendir.",
      "CIF ve yurt içi masrafları ürün bazında doğrula.",
      "Tam sürümde gümrük, ÖTV, KDV kademeli hesabı ve yurt içi masraf dağıtımı açılır."
    ]
  },
  "kacirilan-sgk-tesvikleri-ve-gercek-iscilik-maliyeti-analizi": {
    "karar": "Yararlanılabilir SGK teşvik tutarını ve kaçırılan tutarı gösterir.",
    "girisBasliklari": [
      "Çalışan",
      "Brüt ücret (₺)",
      "SGK matrahı (₺)",
      "Teşvik oranı",
      "Kullanılıyor mu"
    ],
    "ornek": [
      [
        "Çalışan 1",
        42000,
        42000,
        0.05,
        "Evet"
      ],
      [
        "Çalışan 2",
        38000,
        38000,
        0.05,
        "Evet"
      ],
      [
        "Çalışan 3",
        51000,
        51000,
        0.035,
        "Hayır"
      ],
      [
        "Çalışan 4",
        36000,
        36000,
        0.05,
        "Evet"
      ],
      [
        "Çalışan 5",
        47000,
        47000,
        0.035,
        "Hayır"
      ]
    ],
    "metrikler": [
      [
        "Toplam SGK matrahı",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Yararlanılabilir teşvik",
        "=SUMPRODUCT(DEMO_GIRIS!C6:C25,DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Kullanılan teşvik",
        "=SUMPRODUCT((DEMO_GIRIS!E6:E25=\"Evet\")*DEMO_GIRIS!C6:C25*DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Kaçırılan teşvik",
        "=B7-B8",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9>B7*0.5,\"DURDUR\",IF(B9>B7*0.15,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Kullanılmayan teşvik kodlarını SGK beyanına ekle.",
      "Teşvik oranlarını dönem bazında doğrula.",
      "Tam sürümde çalışan bazlı gerçek işçilik maliyeti ve teşvik simülatörü açılır."
    ]
  },
  "kdv-iadesi-azami-alacak-hesabi-dosya-hazirlayici": {
    "karar": "Devreden KDV ve ihracat iadesi sınırına göre azami KDV iade alacağını gösterir.",
    "girisBasliklari": [
      "Belge",
      "Tutar (₺)",
      "KDV (₺)",
      "İade hakkı",
      "Azami iade (₺)"
    ],
    "ornek": [
      [
        "İhracat faturası 1",
        650000,
        78000,
        "Evet",
        65000
      ],
      [
        "İhracat faturası 2",
        380000,
        45600,
        "Evet",
        38000
      ],
      [
        "Yurt içi satış",
        240000,
        28800,
        "Hayır",
        0
      ],
      [
        "Gider pusulası",
        48000,
        5760,
        "Hayır",
        0
      ]
    ],
    "metrikler": [
      [
        "Toplam belge tutarı",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam KDV",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "İade hakkı doğuran tutar",
        "=SUMIF(DEMO_GIRIS!D6:D25,\"Evet\",DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Azami iade",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9<B8,\"UYGUN\",IF(B9=B8,\"İNCELE\",\"DURDUR\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Eksik belge varsa iade dosyasını tamamla.",
      "İade hakkı sınırını güncel mevzuatla doğrula.",
      "Tam sürümde belge takibi, sınır hesabı ve dosya kontrolü açılır."
    ]
  },
  "kdv-tevkifat-mahsup-iade-listesi": {
    "karar": "Tevkifat mahsup listesi ile iade satırlarını ayırır; MAHSUP/İADE/BEKLE kararını üretir.",
    "girisBasliklari": [
      "Belge",
      "Matrah (₺)",
      "Tevkifat (₺)",
      "Mahsup (₺)",
      "İade (₺)"
    ],
    "ornek": [
      [
        "F-1001",
        100000,
        10000,
        8000,
        2000
      ],
      [
        "F-1002",
        50000,
        5000,
        5000,
        0
      ],
      [
        "F-1003",
        200000,
        20000,
        0,
        20000
      ],
      [
        "F-1004",
        75000,
        7500,
        7500,
        0
      ]
    ],
    "metrikler": [
      [
        "Matrah toplam",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Tevkifat toplam",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Mahsup toplam",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "İade toplam",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9>15000,\"İADE\",IF(B8>0,\"MAHSUP\",\"BEKLE\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "İade satırlarını belge paketleriyle eşleştirin.",
      "Mahsup listesini KDV2 ile hizalayın.",
      "Tam sürümde kural yılı ve tebliğ vakası açılır."
    ]
  },
  "kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici": {
    "karar": "Çalışan çıkarmada kıdem ve ihbar tazminatı yükünü gösterir.",
    "girisBasliklari": [
      "Çalışan",
      "Hizmet yılı",
      "Brüt ücret (₺)",
      "Kıdem tazminatı (₺)",
      "İhbar tazminatı (₺)"
    ],
    "ornek": [
      [
        "Çalışan 1",
        2,
        38000,
        76000,
        38000
      ],
      [
        "Çalışan 2",
        1,
        36000,
        36000,
        36000
      ],
      [
        "Çalışan 3",
        3,
        42000,
        126000,
        42000
      ],
      [
        "Çalışan 4",
        1,
        34000,
        34000,
        34000
      ]
    ],
    "metrikler": [
      [
        "Toplam kıdem tazminatı",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Toplam ihbar tazminatı",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Toplam çıkarma maliyeti",
        "=B6+B7",
        "para"
      ],
      [
        "En yüksek tek çalışan yükü",
        "=MAX(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B8>500000,\"DURDUR\",IF(B8>200000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Uzun kıdemli çalışanlarda çıkarma maliyetini ayrıca değerlendir.",
      "İhbar süresi bildirimi ile ihbar tazminatını karşılaştır.",
      "Tam sürümde kıdem tavanı, tazminat hesabı ve toplu çıkarma senaryosu açılır."
    ]
  },
  "kira-avans-takip-dekont": {
    "karar": "Kira ve avans mahsup bakiyesini gösterir; yazdırılabilir dekont için karar üretir.",
    "girisBasliklari": [
      "Kiracı",
      "Kira (₺)",
      "Avans (₺)",
      "Mahsup (₺)",
      "Bakiye (₺)"
    ],
    "ornek": [
      [
        "Kiracı A",
        25000,
        50000,
        25000,
        25000
      ],
      [
        "Kiracı B",
        18000,
        0,
        0,
        18000
      ],
      [
        "Kiracı C",
        32000,
        32000,
        32000,
        0
      ],
      [
        "Kiracı D",
        22000,
        10000,
        10000,
        12000
      ]
    ],
    "metrikler": [
      [
        "Kira toplam",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Avans toplam",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Mahsup toplam",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Açık bakiye",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9>40000,\"KRİTİK\",IF(B9>10000,\"DİKKAT\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Açık bakiyesi olan kiracıya dekont kesin.",
      "Avans mahsup sırasını sözleşme ile doğrulayın.",
      "Tam sürümde durum makinesi ve yaşlandırma açılır."
    ]
  },
  "kkeg-ve-finansman-gider-kisitlamasi-vergi-savunma-seti": {
    "karar": "Yabancı kaynak/öz kaynak farkına göre kısıtlamaya tabi finansman giderini ve KKEG etkisini gösterir.",
    "girisBasliklari": [
      "Kalem",
      "Yabancı kaynak (₺)",
      "Öz kaynak (₺)",
      "Finansman gideri (₺)",
      "KKEG (₺)"
    ],
    "ornek": [
      [
        "Kısa vadeli borç",
        1850000,
        900000,
        96000,
        48000
      ],
      [
        "Uzun vadeli borç",
        960000,
        900000,
        72000,
        3600
      ],
      [
        "Kredi faizi",
        420000,
        900000,
        38400,
        0
      ],
      [
        "Kur farkı",
        780000,
        900000,
        28800,
        0
      ]
    ],
    "metrikler": [
      [
        "Toplam yabancı kaynak",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam öz kaynak",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Fark (aşan kısım)",
        "=B6-B7",
        "para"
      ],
      [
        "Toplam KKEG",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9>40000,\"DURDUR\",IF(B9>10000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "KKEG tutarını dönem matrahına yansıt.",
      "Yabancı kaynak ve öz kaynak kalemlerini belgeyle doğrula.",
      "Tam sürümde oran hesabı, senaryo ve savunma raporu açılır."
    ]
  },
  "kobi-finans-yonetim-paketi": {
    "karar": "Günlük finans operasyonunun ana risklerini tek dosyada birleştirir.",
    "girisBasliklari": [
      "Gösterge",
      "Tutar (₺)",
      "Kritik eşik (₺)",
      "Durum notu",
      "Kaynak"
    ],
    "ornek": [
      [
        "Kullanılabilir nakit",
        90000,
        100000,
        "",
        "Kasa"
      ],
      [
        "30 gün tahsilat",
        260000,
        220000,
        "",
        "Cari"
      ],
      [
        "30 gün ödeme",
        240000,
        250000,
        "",
        "Ödeme"
      ],
      [
        "Aylık taksit",
        80000,
        85000,
        "",
        "Kredi"
      ],
      [
        "Vergi+SGK+maaş",
        168000,
        160000,
        "",
        "Karşılık"
      ],
      [
        "Stok bağlanması",
        190000,
        200000,
        "",
        "Stok"
      ]
    ],
    "metrikler": [
      [
        "Toplam izlenen finansal yük",
        "=SUM(DEMO_GIRIS!B6:B11)",
        "para"
      ],
      [
        "Eşik toplamı",
        "=SUM(DEMO_GIRIS!C6:C11)",
        "para"
      ],
      [
        "Eşiği aşan gösterge",
        "=SUMPRODUCT(--(DEMO_GIRIS!B6:B11>DEMO_GIRIS!C6:C11))",
        "sayi"
      ],
      [
        "Genel baskı oranı",
        "=IFERROR(B6/B7,0)",
        "oran"
      ],
      [
        "Demo karar",
        "=IF(B8>=3,\"DURDUR\",IF(B8>=1,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Kasa, cari, borç, karşılık ve stok baskısını birlikte okuyun.",
      "Tek bir göstergenin iyi olması genel finans sağlığı için yeterli değildir.",
      "Tam sürüm diğer finans motorlarını ortak yönetim ve karar katmanında birleştirir."
    ]
  },
  "mutfak-kayip-kacak-hesaplayici": {
    "karar": "Teorik ve fiili hammadde tüketimini karşılaştırarak mutfak kayıp oranını ve tutarını gösterir.",
    "girisBasliklari": [
      "Hammadde",
      "Teorik (kg)",
      "Fiili (kg)",
      "Kayıp (kg)",
      "Kayıp oranı"
    ],
    "ornek": [
      [
        "Domates",
        420,
        510,
        90,
        "=D6/B6"
      ],
      [
        "Peynir",
        180,
        235,
        55,
        "=D7/B7"
      ],
      [
        "Et",
        260,
        335,
        75,
        "=D8/B8"
      ],
      [
        "Un",
        380,
        445,
        65,
        "=D9/B9"
      ]
    ],
    "metrikler": [
      [
        "Toplam teorik",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam fiili",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Toplam kayıp",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Toplam kayıp oranı",
        "=IFERROR(B8/B7,0)",
        "oran"
      ],
      [
        "Demo karar",
        "=IF(B9>0.08,\"DURDUR\",IF(B9>0.03,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Yüksek kayıplı hammaddeyi mutfak sürecinde denetle.",
      "Stok ve reçete verisini doğrula.",
      "Tam sürümde reçete bazlı tüketim, anomali ve tahmin modülü açılır."
    ]
  },
  "nakliye-maliyeti-hesaplayici": {
    "karar": "Sefer bazında yakıt, bakım, sürücü ve amortisman maliyetinden km/ton-km birim maliyetini ve kârı gösterir.",
    "girisBasliklari": [
      "Sefer",
      "Km",
      "Ton",
      "Gelir (₺)",
      "Maliyet (₺)"
    ],
    "ornek": [
      [
        "Sefer 1",
        620,
        18,
        38500,
        29400
      ],
      [
        "Sefer 2",
        380,
        12,
        21800,
        17200
      ],
      [
        "Sefer 3",
        840,
        22,
        51200,
        39800
      ],
      [
        "Sefer 4",
        460,
        15,
        25600,
        21400
      ]
    ],
    "metrikler": [
      [
        "Toplam gelir",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Toplam maliyet",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Net kâr",
        "=B6-B7",
        "para"
      ],
      [
        "Km başına maliyet",
        "=IFERROR(B8/SUM(DEMO_GIRIS!B6:B25),0)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B8<0,\"DURDUR\",IF(B8<B7*0.15,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Zararlı seferlerde fiyatlamayı yeniden değerlendir.",
      "Yakıt tüketimi ve bakım kalemlerini doğrula.",
      "Tam sürümde filo bazlı amortisman, senaryo ve duyarlılık analizi açılır."
    ]
  },
  "ortaklar-cari-ve-kasa-adat-faiz-faturasi-hesaplayici": {
    "karar": "Ortak cari bakiyeleri ve kasa adatları üzerinden dönem faizini ve fatura tutarını gösterir.",
    "girisBasliklari": [
      "Ortak",
      "Bakiye (₺)",
      "Gün",
      "Günlük faiz oranı",
      "Faiz (₺)"
    ],
    "ornek": [
      [
        "Ahmet Y.",
        850000,
        120,
        0.000411,
        41910
      ],
      [
        "Mehmet K.",
        420000,
        90,
        0.000411,
        15540
      ],
      [
        "Ayşe D.",
        1200000,
        150,
        0.000411,
        73980
      ],
      [
        "Ali S.",
        260000,
        60,
        0.000411,
        6412
      ]
    ],
    "metrikler": [
      [
        "Toplam bakiye",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam faiz",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "KDV (%20)",
        "=B7*0.20",
        "para"
      ],
      [
        "Fatura tutarı",
        "=B7+B8",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9>100000,\"DURDUR\",IF(B9>50000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Faiz faturasını ortak bazında düzenle.",
      "Gün ve oranları belgelerle doğrula.",
      "Tam sürümde adat gün hesabı, dönem takibi ve fatura şablonu açılır."
    ]
  },
  "pazaryeri-net-kar-ve-eksik-hakedis-yakalayici": {
    "karar": "Pazaryeri satışlarında komisyon, kargo, reklam sonrası net kârı ve eksik hakedişi gösterir.",
    "girisBasliklari": [
      "Kalem",
      "Satış (₺)",
      "Gider (₺)",
      "Beklenen hakediş (₺)",
      "Gerçek hakediş (₺)"
    ],
    "ornek": [
      [
        "Kalem A",
        45000,
        13500,
        31500,
        29800
      ],
      [
        "Kalem B",
        28000,
        8200,
        19800,
        19800
      ],
      [
        "Kalem C",
        61000,
        19700,
        41300,
        38900
      ],
      [
        "Kalem D",
        18000,
        5400,
        12600,
        12600
      ]
    ],
    "metrikler": [
      [
        "Toplam satış",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam gider",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Beklenen hakediş",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Fark (eksik tahsilat)",
        "=B8-SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9>5000,\"DURDUR\",IF(B9>0,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Eksik hakedişi pazaryeri ile mutabık kıl.",
      "Gider kalemlerini belgeyle doğrula.",
      "Tam sürümde komisyon kademesi, iade, reklam dağıtımı ve tahmin açılır."
    ]
  },
  "pos-komisyon-ve-net-tahsilat-kontrol-sistemi": {
    "karar": "Bankanın eksik yatırdığı veya komisyona giden tutarı gösterir.",
    "girisBasliklari": [
      "Gün",
      "Brüt POS (₺)",
      "Komisyon %",
      "Beklenen net (₺)",
      "Banka yatan (₺)"
    ],
    "ornek": [
      [
        "01.08.2026",
        35000,
        0.025,
        "=B6*(1-C6)",
        34050
      ],
      [
        "02.08.2026",
        42000,
        0.027,
        "=B7*(1-C7)",
        40700
      ],
      [
        "03.08.2026",
        28000,
        0.025,
        "=B8*(1-C8)",
        27250
      ],
      [
        "04.08.2026",
        51000,
        0.03,
        "=B9*(1-C9)",
        48900
      ]
    ],
    "metrikler": [
      [
        "Toplam brüt POS",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Beklenen net",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Banka yatan",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Mutabakat farkı",
        "=B8-B7",
        "para"
      ],
      [
        "Demo karar",
        "=IF(ABS(B9)>2500,\"DURDUR\",IF(ABS(B9)>500,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Banka yatan ile beklenen neti günlük karşılaştır.",
      "Fark büyüyorsa komisyon/valör/kesinti belgesini kontrol et.",
      "Tam sürümde banka bazlı mutabakat ve kök-neden ayrıştırması açılır."
    ]
  },
  "proje-ve-is-bazinda-gercek-karlilik-sistemi": {
    "karar": "Hangi işin para kazandırdığını ve hangisinin nakit tükettiğini gösterir.",
    "girisBasliklari": [
      "Proje / iş",
      "Gelir (₺)",
      "Direkt maliyet (₺)",
      "Personel (₺)",
      "Finansman/nakit (₺)"
    ],
    "ornek": [
      [
        "Proje A",
        320000,
        165000,
        72000,
        18000
      ],
      [
        "Proje B",
        210000,
        138000,
        54000,
        24000
      ],
      [
        "Proje C",
        480000,
        255000,
        98000,
        32000
      ],
      [
        "Proje D",
        145000,
        97000,
        41000,
        15000
      ]
    ],
    "metrikler": [
      [
        "Toplam proje geliri",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam gerçek maliyet",
        "=SUM(DEMO_GIRIS!C6:E25)",
        "para"
      ],
      [
        "Gerçek proje kârı",
        "=B6-B7",
        "para"
      ],
      [
        "Gerçek marj",
        "=IFERROR(B8/B6,0)",
        "yuzde"
      ],
      [
        "Demo karar",
        "=IF(B8<0,\"DURDUR\",IF(B9<0.10,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Gelir yerine gerçek proje marjını yönetin.",
      "Finansman ve işletme sermayesi etkisini maliyetten ayırmayın.",
      "Tam sürümde iş bazlı nakit tüketimi, senaryo ve teklif eşiği açılır."
    ]
  },
  "restoran-recete-maliyet-fire": {
    "karar": "Fire dahil porsiyon maliyetini ve bileşim yüzde 100 kontrolünü gösterir.",
    "girisBasliklari": [
      "Reçete",
      "Hammadde (₺)",
      "Fire (₺)",
      "Porsiyon (₺)",
      "Hedef fiyat (₺)"
    ],
    "ornek": [
      [
        "Menemen",
        42,
        6,
        48,
        95
      ],
      [
        "Köfte",
        78,
        10,
        88,
        175
      ],
      [
        "Salata",
        28,
        4,
        32,
        75
      ],
      [
        "Pizza",
        55,
        8,
        63,
        140
      ]
    ],
    "metrikler": [
      [
        "Ort. porsiyon",
        "=AVERAGE(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Ort. fire",
        "=AVERAGE(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Food cost oranı",
        "=B6/AVERAGE(DEMO_GIRIS!E6:E25)",
        "yuzde"
      ],
      [
        "Kalem sayısı",
        "=COUNTA(DEMO_GIRIS!A6:A25)",
        "adet"
      ],
      [
        "Demo karar",
        "=IF(B8>0.45,\"KRİTİK\",IF(B8>0.35,\"DİKKAT\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Fire oranı yüksek reçeteyi gözden geçirin.",
      "Bileşim yüzde 100 dışına çıkan satırları düzeltin.",
      "Tam sürümde birim dönüşüm ve varyans köprüsü açılır."
    ]
  },
  "sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli": {
    "karar": "Öz kaynak, sermaye kaybı ve TTK 376 eşiklerini (1/3, 2/3, borca batıklık) gösterir.",
    "girisBasliklari": [
      "Dönem",
      "Sermaye (₺)",
      "Öz kaynak (₺)",
      "Kayıp (₺)",
      "Kayıp oranı"
    ],
    "ornek": [
      [
        "Oca 2026",
        2000000,
        1750000,
        250000,
        "=D6/B6"
      ],
      [
        "Şub 2026",
        2000000,
        1480000,
        520000,
        "=D7/B7"
      ],
      [
        "Mar 2026",
        2000000,
        1260000,
        740000,
        "=D8/B8"
      ],
      [
        "Nis 2026",
        2000000,
        1080000,
        920000,
        "=D9/B9"
      ]
    ],
    "metrikler": [
      [
        "Toplam kayıp",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Son dönem kayıp oranı",
        "=IFERROR(DEMO_GIRIS!D9/DEMO_GIRIS!B9,0)",
        "oran"
      ],
      [
        "2/3 eşiği",
        "=2/3",
        "oran"
      ],
      [
        "1/3 eşiği",
        "=1/3",
        "oran"
      ],
      [
        "Demo karar",
        "=IF(B7>=B8,\"DURDUR\",IF(B7>=B9,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Kayıp oranı 2/3 eşiğini aştıysa genel kurul sürecini başlat.",
      "Bilanço kalemlerini belgelerle doğrula.",
      "Tam sürümde aylık izleme, borca batıklık ve yönetim kurulu raporu açılır."
    ]
  },
  "stok-satis-ve-nakit-baglanma-sistemi": {
    "karar": "Hangi stokun para tükettiğini ve hangisinin yeniden alınması gerektiğini gösterir.",
    "girisBasliklari": [
      "Ürün",
      "Stok maliyeti (₺)",
      "30 gün satış (₺)",
      "Brüt marj %",
      "Tedarik süresi (gün)"
    ],
    "ornek": [
      [
        "Ürün A",
        95000,
        45000,
        0.28,
        12
      ],
      [
        "Ürün B",
        42000,
        68000,
        0.34,
        8
      ],
      [
        "Ürün C",
        120000,
        15000,
        0.22,
        25
      ],
      [
        "Ürün D",
        28000,
        51000,
        0.31,
        5
      ],
      [
        "Ürün E",
        76000,
        38000,
        0.18,
        18
      ]
    ],
    "metrikler": [
      [
        "Toplam stok maliyeti",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "30 gün satış",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Stok / satış oranı",
        "=IFERROR(B6/B7,0)",
        "oran"
      ],
      [
        "Satışı çok düşük ürün",
        "=COUNTIF(DEMO_GIRIS!C6:C25,\"<20000\")",
        "sayi"
      ],
      [
        "Demo karar",
        "=IF(B8>2,\"DURDUR\",IF(B9>=1,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Satışı düşük ve stok maliyeti yüksek ürünlerde yeniden alımı durdur.",
      "Hızlı dönen ürünlerde tedarik süresini ayrıca izle.",
      "Tam sürümde stok gün sayısı, yeniden sipariş ve nakit bağlanma motoru açılır."
    ]
  },
  "sube-karlilik-ve-nakit-hesaplayici": {
    "karar": "Şube bazında gelir, gider ve nakit akışını toplayarak kârlılık ve nakit pozisyonunu gösterir.",
    "girisBasliklari": [
      "Şube",
      "Gelir (₺)",
      "Gider (₺)",
      "Net kâr (₺)",
      "Marj"
    ],
    "ornek": [
      [
        "İstanbul",
        940000,
        620000,
        320000,
        "=D6/B6"
      ],
      [
        "Ankara",
        480000,
        390000,
        90000,
        "=D7/B7"
      ],
      [
        "İzmir",
        520000,
        415000,
        105000,
        "=D8/B8"
      ],
      [
        "Bursa",
        260000,
        275000,
        -15000,
        "=D9/B9"
      ]
    ],
    "metrikler": [
      [
        "Toplam gelir",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam gider",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Toplam net kâr",
        "=B6-B7",
        "para"
      ],
      [
        "Genel marj",
        "=IFERROR(B8/B6,0)",
        "oran"
      ],
      [
        "Demo karar",
        "=IF(B9<0.05,\"DURDUR\",IF(B9<0.15,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Zararlı şubelerin gider yapısını denetle.",
      "Tahsilat ve ödeme planını gözden geçir.",
      "Tam sürümde nakit akışı, senaryo ve duyarlılık analizi açılır."
    ]
  },
  "taseron-hakedis-kesinti-mutabakati": {
    "karar": "Taşeron hakedişi ile yapılan ödeme arasındaki mutabakat farkını gösterir.",
    "girisBasliklari": [
      "Taşeron",
      "Hakediş (₺)",
      "Kesinti (₺)",
      "Ödenen (₺)",
      "Fark (₺)"
    ],
    "ornek": [
      [
        "Taşeron A",
        185000,
        25000,
        155000,
        "=B6-C6-D6"
      ],
      [
        "Taşeron B",
        96000,
        12000,
        80000,
        "=B7-C7-D7"
      ],
      [
        "Taşeron C",
        132000,
        18000,
        118000,
        "=B8-C8-D8"
      ],
      [
        "Taşeron D",
        74000,
        9000,
        68000,
        "=B9-C9-D9"
      ]
    ],
    "metrikler": [
      [
        "Toplam hakediş",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam kesinti",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Toplam ödenen",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Toplam mutabakat farkı",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(ABS(B9)>B6*0.05,\"DURDUR\",IF(ABS(B9)>1000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Fark tutarı taşeron hakedişinde işaretle.",
      "Kesinti kalemlerini sözleşme hükümleriyle doğrula.",
      "Tam sürümde taşeron bazlı mutabakat ve ihtilaf erken uyarısı açılır."
    ]
  },
  "tesvikli-bordro-avantajli-tesvik": {
    "karar": "Çalışan bazında en avantajlı SGK teşvik kodunu seçer ve kaçırılan tasarrufu gösterir.",
    "girisBasliklari": [
      "Çalışan",
      "Brüt (₺)",
      "Mevcut teşvik (₺)",
      "Potansiyel (₺)",
      "Kaçırılan (₺)"
    ],
    "ornek": [
      [
        "A. Yılmaz",
        45000,
        2250,
        6500,
        4250
      ],
      [
        "B. Demir",
        38000,
        1900,
        5200,
        3300
      ],
      [
        "C. Kaya",
        52000,
        0,
        7800,
        7800
      ],
      [
        "D. Çelik",
        41000,
        2050,
        2050,
        0
      ]
    ],
    "metrikler": [
      [
        "Toplam brüt",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Mevcut teşvik",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Potansiyel",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Kaçırılan",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9>10000,\"UYGULA\",IF(B9>0,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Kaçırılan tutarı yüksek çalışanlarda teşvik kodunu değiştirin.",
      "Bordro satırını dönem kapanışından önce kilitleyin.",
      "Tam sürümde kural yılı ve vaka doğrulaması açılır."
    ]
  },
  "trendyol-komisyon-sonrasi-net-kar": {
    "karar": "Trendyol komisyon, TY Plus, flash, reklam ve kargo sonrası net kârı ve SAT/ZAM/ÇEKİL kararını gösterir.",
    "girisBasliklari": [
      "SKU",
      "Satış (₺)",
      "Maliyet (₺)",
      "Komisyon+kesinti (₺)",
      "Hakediş (₺)"
    ],
    "ornek": [
      [
        "SKU-A",
        45000,
        18000,
        9000,
        36000
      ],
      [
        "SKU-B",
        28000,
        12000,
        5600,
        22400
      ],
      [
        "SKU-C",
        61000,
        25000,
        14000,
        45000
      ],
      [
        "SKU-D",
        18000,
        8000,
        4000,
        14000
      ]
    ],
    "metrikler": [
      [
        "Toplam satış",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam maliyet",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Toplam kesinti",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Net kâr",
        "=B6-B7-B8",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9<0,\"ÇEKİL\",IF(B9<5000,\"ZAM\",\"SAT\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Negatif net kârda SKUyu çekin veya fiyatı artırın.",
      "Kesinti kırılımını kategori komisyon tablosuyla doğrulayın.",
      "Tam sürümde TY Plus, flash ve desi kargo ayrı kolonlarda açılır."
    ]
  },
  "uretim-recetesi-ve-zam-yansitma-hesaplayici": {
    "karar": "Reçete hammadde zamlarının ürün maliyetine ve önerilen satış fiyatına etkisini gösterir.",
    "girisBasliklari": [
      "Ürün",
      "Eski maliyet (₺)",
      "Yeni maliyet (₺)",
      "Artış (₺)",
      "Zam oranı"
    ],
    "ornek": [
      [
        "Ürün A",
        85,
        104,
        19,
        "=D6/B6"
      ],
      [
        "Ürün B",
        42,
        51,
        9,
        "=D7/B7"
      ],
      [
        "Ürün C",
        130,
        168,
        38,
        "=D8/B8"
      ],
      [
        "Ürün D",
        65,
        76,
        11,
        "=D9/B9"
      ]
    ],
    "metrikler": [
      [
        "Toplam eski maliyet",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam yeni maliyet",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Toplam artış",
        "=B7-B6",
        "para"
      ],
      [
        "Ortalama zam oranı",
        "=IFERROR(B8/B6,0)",
        "oran"
      ],
      [
        "Demo karar",
        "=IF(B9>0.20,\"DURDUR\",IF(B9>0.08,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Yüksek artışlı ürünlerde fiyatı yeniden hesapla.",
      "Fire ve hammadde verisini doğrula.",
      "Tam sürümde reçete bazlı hedef marj ve önerilen fiyat açılır."
    ]
  },
  "vergi-sgk-borcunu-tecil-etmeli-miyim-kredi-mi-tecil-mi": {
    "karar": "Tecil (6183/48), kredi ve peşin ödeme seçeneklerinin toplam maliyetini karşılaştırır.",
    "girisBasliklari": [
      "Borç",
      "Asıl (₺)",
      "Zam (₺)",
      "Tecil taksit",
      "Toplam ödeme (₺)"
    ],
    "ornek": [
      [
        "Vergi borcu",
        480000,
        86400,
        24,
        597600
      ],
      [
        "SGK borcu",
        315000,
        47250,
        12,
        382725
      ],
      [
        "Kira tevkifatı",
        96000,
        12480,
        6,
        113280
      ],
      [
        "Motorlu taşıt",
        42000,
        5880,
        12,
        50148
      ]
    ],
    "metrikler": [
      [
        "Toplam asıl",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam zam",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Toplam ödeme",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Kredi maliyeti (örnek)",
        "=B6*1.18",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9<B8,\"UYGUN\",IF(B9<B8*1.05,\"İNCELE\",\"DURDUR\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "En düşük maliyetli seçeneği borç bazında değerlendir.",
      "Kredi faiz oranını güncel piyasa ile doğrula.",
      "Tam sürümde NBD, vade ve taksit planı senaryoları açılır."
    ]
  },
  "vergi-sgk-ve-maas-karsilik-ayirma-sistemi": {
    "karar": "Bugünkü paranın ne kadarının gerçekte harcanabilir olmadığını gösterir.",
    "girisBasliklari": [
      "Dönem",
      "Mevcut nakit (₺)",
      "Vergi karşılığı (₺)",
      "SGK+maaş (₺)",
      "Diğer zorunlu (₺)"
    ],
    "ornek": [
      [
        "Ağustos",
        280000,
        62000,
        118000,
        18000
      ],
      [
        "Eylül",
        210000,
        51000,
        112000,
        15000
      ],
      [
        "Ekim",
        260000,
        56000,
        116000,
        16000
      ],
      [
        "Kasım",
        240000,
        53000,
        110000,
        15000
      ],
      [
        "Aralık",
        230000,
        50000,
        108000,
        14000
      ]
    ],
    "metrikler": [
      [
        "Toplam görünen nakit",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam zorunlu karşılık",
        "=SUM(DEMO_GIRIS!C6:E25)",
        "para"
      ],
      [
        "Gerçek harcanabilir nakit",
        "=B6-B7",
        "para"
      ],
      [
        "Karşılık / nakit",
        "=IFERROR(B7/B6,0)",
        "yuzde"
      ],
      [
        "Demo karar",
        "=IF(B8<0,\"DURDUR\",IF(B9>0.8,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Vergi, SGK ve maaş karşılıklarını ayrı tut.",
      "Harcanabilir nakit negatifse discretionary ödemeyi durdur.",
      "Tam sürümde dönemsel tahakkuk ve ödeme takvimi detaylanır."
    ]
  },
  "yeniden-degerleme-yapmali-miyim-vergi-tasarruf-analizi": {
    "karar": "Kıymet bazında yeniden değerlemenin fon vergisi ve ek amortisman sonrası net vergi tasarrufunu gösterir.",
    "girisBasliklari": [
      "Kıymet",
      "Maliyet (₺)",
      "Katsayı",
      "Değer artışı (₺)",
      "Net tasarruf (₺)"
    ],
    "ornek": [
      [
        "Makine 1",
        850000,
        1.42,
        357000,
        42500
      ],
      [
        "Makine 2",
        420000,
        1.38,
        159600,
        18600
      ],
      [
        "Bina 1",
        2500000,
        1.51,
        1275000,
        152000
      ],
      [
        "Taşıt 1",
        380000,
        1.24,
        91200,
        10400
      ]
    ],
    "metrikler": [
      [
        "Toplam maliyet",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam değer artışı",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Fon vergisi (%2)",
        "=B7*0.02",
        "para"
      ],
      [
        "Toplam net tasarruf",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9>B8,\"UYGUN\",IF(B9>B8*0.7,\"İNCELE\",\"DURDUR\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Net tasarrufu pozitif kıymetlerde yeniden değerlemeyi değerlendir.",
      "Katsayı ve oranları güncel mevzuatla doğrula.",
      "Tam sürümde kıymet bazlı yıl ve ömür hesabı, senaryo ve karar motoru açılır."
    ]
  },
  "yillara-sari-insaat-stopaj-nakit-akis-planlayici": {
    "karar": "Dönem stopaj yükünü ve nakit akışı açığını gösterir.",
    "girisBasliklari": [
      "Dönem",
      "Hakediş (₺)",
      "Stopaj oranı",
      "Tahsilat (₺)",
      "Ödeme (₺)"
    ],
    "ornek": [
      [
        "2026 Q3",
        320000,
        0.05,
        290000,
        260000
      ],
      [
        "2026 Q4",
        410000,
        0.05,
        380000,
        350000
      ],
      [
        "2027 Q1",
        480000,
        0.05,
        300000,
        360000
      ],
      [
        "2027 Q2",
        520000,
        0.05,
        470000,
        430000
      ]
    ],
    "metrikler": [
      [
        "Toplam hakediş",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Toplam stopaj",
        "=SUMPRODUCT(DEMO_GIRIS!B6:B25,DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Net nakit (tahsilat-ödeme)",
        "=SUM(DEMO_GIRIS!D6:D25)-SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Stopaj / nakit oranı",
        "=IFERROR(B7/MAX(B8,1),0)",
        "oran"
      ],
      [
        "Demo karar",
        "=IF(B8<0,\"DURDUR\",IF(B9>0.5,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Stopaj ödeme tarihlerini tahsilatla eşleştir.",
      "Nakit açığı dönemlerinde finansman planı oluştur.",
      "Tam sürümde yıllık plan, vergi planı ve nakit köprüsü motoru açılır."
    ]
  },
  "kdv-iade-listesi-robotu-gib7": {
    "karar": "GİB 7 liste tutarlarını toplar; İADE/EKSİK BELGE/BEKLE kararını üretir.",
    "girisBasliklari": [
      "Liste",
      "Satır",
      "Tutar (₺)",
      "Eksik belge",
      "Durum"
    ],
    "ornek": [
      [
        "İndirilecek KDV",
        12,
        180000,
        0,
        "Hazır"
      ],
      [
        "Yüklenilen KDV",
        8,
        95000,
        1,
        "Eksik"
      ],
      [
        "GÇB",
        5,
        220000,
        0,
        "Hazır"
      ],
      [
        "Diğer",
        3,
        40000,
        0,
        "Hazır"
      ]
    ],
    "metrikler": [
      [
        "Toplam tutar",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Satır",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "adet"
      ],
      [
        "Eksik",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "adet"
      ],
      [
        "Liste sayısı",
        "=COUNTA(DEMO_GIRIS!A6:A25)",
        "adet"
      ],
      [
        "Demo karar",
        "=IF(B8>0,\"EKSİK BELGE\",IF(B6>0,\"İADE\",\"BEKLE\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Eksik belgeleri tamamlayın.",
      "7 listeyi GİB sırasıyla kontrol edin.",
      "Tam sürümde kural yılı ve kanıt raporu açılır."
    ]
  },
  "e-fatura-toplu-donusturucu": {
    "karar": "Toplu e-Fatura satırlarını muhasebe ile eşleştirir; PDF kanıt için karar üretir.",
    "girisBasliklari": [
      "UUID",
      "e-Fatura (₺)",
      "Muhasebe (₺)",
      "Fark (₺)",
      "Kova"
    ],
    "ornek": [
      [
        "U1",
        11800,
        11800,
        0,
        "Eşleşti"
      ],
      [
        "U2",
        23600,
        23000,
        600,
        "Fark"
      ],
      [
        "U3",
        5900,
        5900,
        0,
        "Eşleşti"
      ],
      [
        "U4",
        17700,
        0,
        17700,
        "Eşleşmedi"
      ]
    ],
    "metrikler": [
      [
        "e-Fatura",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Muhasebe",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Fark",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Kayıt",
        "=COUNTA(DEMO_GIRIS!A6:A25)",
        "adet"
      ],
      [
        "Demo karar",
        "=IF(B8>10000,\"DURDUR\",IF(B8>0,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "XML dosyası okunmaz; export satırı yapıştırılır.",
      "Eşleşmeyen UUID’leri işleyin.",
      "Tam sürümde 10.000 satır ölçeği açılır."
    ]
  },
  "ymm-tasdik-kontrol-robotu": {
    "karar": "YMM tasdik kontrol listesi ve eksik belge kuyruğundan TASDIK/EKSİK/DURDUR üretir.",
    "girisBasliklari": [
      "Kalem",
      "Zorunlu",
      "Durum",
      "Gecikme (gün)",
      "Risk"
    ],
    "ornek": [
      [
        "Bilanço",
        "Evet",
        "Tamam",
        0,
        "Düşük"
      ],
      [
        "Defter",
        "Evet",
        "Eksik",
        12,
        "Yüksek"
      ],
      [
        "Fatura örnek",
        "Evet",
        "İnceleme",
        3,
        "Orta"
      ],
      [
        "Sözleşme",
        "Hayır",
        "Tamam",
        0,
        "Düşük"
      ]
    ],
    "metrikler": [
      [
        "Kalem",
        "=COUNTA(DEMO_GIRIS!A6:A25)",
        "adet"
      ],
      [
        "Eksik",
        "=COUNTIF(DEMO_GIRIS!C6:C25,\"Eksik\")",
        "adet"
      ],
      [
        "Ort. gecikme",
        "=AVERAGE(DEMO_GIRIS!D6:D25)",
        "adet"
      ],
      [
        "Yüksek risk",
        "=COUNTIF(DEMO_GIRIS!E6:E25,\"Yüksek\")",
        "adet"
      ],
      [
        "Demo karar",
        "=IF(B9>=1,\"DURDUR\",IF(B7>0,\"EKSİK\",\"TASDIK\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Eksik zorunlu belgeleri tamamlayın.",
      "Geciken kalemleri önceliklendirin.",
      "Tam sürümde durum makinesi ve kanıt paketi açılır."
    ]
  },
  "tesvikli-bordro-optimizasyon": {
    "karar": "Kişi×teşvik matrisi ile kaçırılan tasarrufu ve UYGULA/İNCELE kararını gösterir.",
    "girisBasliklari": [
      "Çalışan",
      "Brüt (₺)",
      "Mevcut (₺)",
      "Potansiyel (₺)",
      "Kaçırılan (₺)"
    ],
    "ornek": [
      [
        "A",
        50000,
        2500,
        8000,
        5500
      ],
      [
        "B",
        42000,
        2100,
        6500,
        4400
      ],
      [
        "C",
        60000,
        0,
        9500,
        9500
      ],
      [
        "D",
        38000,
        3800,
        3800,
        0
      ]
    ],
    "metrikler": [
      [
        "Brüt",
        "=SUM(DEMO_GIRIS!B6:B25)",
        "para"
      ],
      [
        "Mevcut",
        "=SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Potansiyel",
        "=SUM(DEMO_GIRIS!D6:D25)",
        "para"
      ],
      [
        "Kaçırılan",
        "=SUM(DEMO_GIRIS!E6:E25)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9>15000,\"UYGULA\",IF(B9>0,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Kaçırılan tutarı yüksek çalışanlarda kod değiştirin.",
      "Matris setini dönem kapanışına kilitleyin.",
      "Tam sürümde M-SEN ve vaka doğrulaması açılır."
    ]
  },
  "konkordato-nakit-akis-on-projesi": {
    "karar": "3 yıllık nakit açık/fazlasından BAŞVUR/BEKLE/UYGUN DEĞİL kararını üretir.",
    "girisBasliklari": [
      "Yıl",
      "Gelir (₺)",
      "Gider (₺)",
      "Nakit (₺)",
      "Açık (₺)"
    ],
    "ornek": [
      [
        "2026",
        12000000,
        13500000,
        -1500000,
        1500000
      ],
      [
        "2027",
        14000000,
        13800000,
        200000,
        0
      ],
      [
        "2028",
        16000000,
        14500000,
        1500000,
        0
      ],
      [
        "Toplam",
        42000000,
        41800000,
        200000,
        1500000
      ]
    ],
    "metrikler": [
      [
        "Gelir",
        "=SUM(DEMO_GIRIS!B6:B8)",
        "para"
      ],
      [
        "Gider",
        "=SUM(DEMO_GIRIS!C6:C8)",
        "para"
      ],
      [
        "Net nakit",
        "=B6-B7",
        "para"
      ],
      [
        "Açık",
        "=SUM(DEMO_GIRIS!E6:E8)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(B9>2000000,\"UYGUN DEĞİL\",IF(B8<0,\"BEKLE\",\"BAŞVUR\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit açık yılları mahkeme raporunda vurgulayın.",
      "3 yıllık bilanço satırlarını kilitleyin.",
      "Tam sürümde durum makinesi ve kanıt raporu açılır."
    ]
  },
  "logo-sql-cari-yaslandirma-tahsilat-karar-motoru": {
    "karar": "Logo/SQL cari hareket örneğinden açık alacak, gecikme ve tahsilat önceliğini gösterir.",
    "girisBasliklari": [
      "Cari",
      "Borç (₺)",
      "Tahsilat (₺)",
      "Vade (gün)",
      "Gecikme (gün)"
    ],
    "ornek": [
      [
        "Atlas Yapı",
        320000,
        195000,
        30,
        107
      ],
      [
        "Bora Lojistik",
        210000,
        80000,
        30,
        126
      ],
      [
        "Cem Gıda",
        145000,
        0,
        30,
        18
      ],
      [
        "Delta Dış Ticaret",
        402000,
        167500,
        30,
        78
      ],
      [
        "Eksen Makine",
        420000,
        0,
        30,
        0
      ]
    ],
    "metrikler": [
      [
        "Açık alacak",
        "=SUM(DEMO_GIRIS!B6:B25)-SUM(DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "Vadesi geçmiş",
        "=SUMIF(DEMO_GIRIS!E6:E25,\">0\",DEMO_GIRIS!B6:B25)-SUMIF(DEMO_GIRIS!E6:E25,\">0\",DEMO_GIRIS!C6:C25)",
        "para"
      ],
      [
        "90+ riskli satır",
        "=COUNTIF(DEMO_GIRIS!E6:E25,\">90\")",
        "adet"
      ],
      [
        "Ortalama gecikme",
        "=IFERROR(SUMIF(DEMO_GIRIS!E6:E25,\">0\",DEMO_GIRIS!E6:E25)/COUNTIF(DEMO_GIRIS!E6:E25,\">0\"),0)",
        "adet"
      ],
      [
        "Demo karar",
        "=IF(B8>=2,\"ACİL\",IF(B7>100000,\"İNCELE\",\"UYGUN\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "90 günü aşan açık alacakları önce ele alın.",
      "Vadesi geçmiş tutarı müşteri bazında tahsilat kuyruğuna alın.",
      "Tam sürümde FIFO açık kalem, 7 kovalı yaşlandırma, 13 hafta ve VUK 323 ön eleği açılır."
    ]
  }
,
  "gayrimenkul-yatirim-fizibilite-ve-karlilik": {
    "karar": "Ticari/konut gayrimenkul yatırımlarında kira getirisi, Cap Rate ve IRR getiri analizi.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "ges-yatirim-fizibilite-ve-proje-finansmani": {
    "karar": "Güneş santrali yatırımlarında elektrik satışı, DSCR ve proje finansmanı amortismanı.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "otel-fizibilite-ve-karlilik": {
    "karar": "Otel yatırımlarında oda doluluk, RevPAR, GOP ve başabaş analiz kokpiti.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "uretim-kapasite-yatirim-ve-karlilik": {
    "karar": "Yeni makine/hat yatırımlarında OEE kapasite, birim maliyet ve ROI analizi.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "arsa-ve-gayrimenkul-proje-gelistirme-fizibilite": {
    "karar": "Arsa projelerinde KAKS/TAKS hesabı, hasılat paylaşımı ve müteahhit net kâr modeli.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "sirket-satin-alma-ve-ortaklik-devir-analizi": {
    "karar": "M&A süreçlerinde normalize EBITDA, net borç köprüsü ve hisse değerleme kokpiti.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "saas-finansal-planlama-runway-ve-yatirim": {
    "karar": "Abonelik modellerinde MRR, Churn, CAC/LTV ve nakit ömrü (runway) modeli.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "gayrimenkul-tut-sat-karar": {
    "karar": "Gayrimenkulü kirada tutmak ile satıp getiri fonuna yatırmayı kıyaslayan model.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "ortakli-gayrimenkul-yatirimi-kar-dagitim": {
    "karar": "Gayrimenkul ortaklıklarında Tercihli Getiri ve Waterfall kâr dağıtım mimarisi.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "borcla-sirket-satin-alma-ve-yatirim-getirisi": {
    "karar": "LBO modellerinde borç amortismanı ve kaldıraçlı özkaynak IRR hesabı.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "startup-yatirim-alma-ve-nakit-runway": {
    "karar": "Erken aşama girişimlerde yatırım sonrası değerleme, Cap Table ve burn rate modeli.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "emsal-sirket-carpanlariyla-degerleme": {
    "karar": "Emsal şirket çarpanları (EV/EBITDA, P/E) ile piyasa değeri ve adil hisse fiyatı.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "wacc-hesaplama-ve-sermaye-maliyeti": {
    "karar": "CAPM modeli ve borçlanma maliyetiyle ağırlıklı ortalama sermaye maliyeti (WACC).",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "profesyonel-hizmet-sirketi-karlilik-ve-kapasite": {
    "karar": "Danışmanlık/yazılımda faturalandırılabilir kapasite ve müşteri kârlılık matrisi.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "perakende-magaza-acilis-fizibilite": {
    "karar": "Mağaza açılışlarında fit-out capex, ciro kirası ve başabaş ciro eşiği.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "filo-yatirim-ve-arac-yenileme-fizibilite": {
    "karar": "Şirket araçlarında satın alma vs kiralama (Buy vs Lease) TCO maliyet analizi.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "franchise-yatirim-fizibilite": {
    "karar": "Franchise yatırımlarında royalty, isim hakkı, tedarik marjı ve payback hesabı.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "klinik-saglik-merkezi-fizibilite-ve-karlilik": {
    "karar": "Klinik yatırımlarında cihaz leasingi, hekim hakedişi ve hasta başı kârlılık.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  "restoran-kafe-yatirim-fizibilite-ve-karlilik": {
    "karar": "Restoranlarda masa devri, adisyon, food cost ve başabaş ciro kokpiti.",
    "girisBasliklari": [
      "Dönem / Kalem",
      "Gelir (₺)",
      "Gider / OPEX (₺)",
      "Net Nakit Akışı (₺)",
      "Kümülatif Getiri (₺)"
    ],
    "ornek": [
      [
        "Yıl 1",
        1200000,
        650000,
        "=B6-C6",
        "=D6"
      ],
      [
        "Yıl 2",
        1450000,
        720000,
        "=B7-C7",
        "=E6+D7"
      ],
      [
        "Yıl 3",
        1800000,
        800000,
        "=B8-C8",
        "=E7+D8"
      ],
      [
        "Yıl 4",
        2200000,
        890000,
        "=B9-C9",
        "=E8+D9"
      ]
    ],
    "ozet": [
      [
        "Toplam Net Getiri",
        "=E9",
        "para"
      ],
      [
        "Yıllık Ortalama Nakit",
        "=AVERAGE(D6:D9)",
        "para"
      ],
      [
        "Demo karar",
        "=IF(E9>=3000000,\"UYGUN\",IF(E9>=1500000,\"İNCELE\",\"RİSKLİ\"))",
        "metin"
      ]
    ],
    "aksiyonlar": [
      "Nakit projeksiyonunu ve kümülatif başabaş süresini kontrol edin.",
      "İskonto ve enflasyon parametrelerini şirket hedeflerine göre güncelleyin.",
      "Tam sürümde 60+ nokta bağımsız audit kernel, dinamik senaryo seçici ve kilitli karar kokpiti açılır."
    ]
  },
  'asgari-kurumlar-vergisi-simulasyon-motoru': {
      "karar": "İndirim ve istisnalar sonrası asgari kurumlar vergisi matrah farkını ve ek vergi yükünü gösterir.",
      "girisBasliklari": [
          "Dönem",
          "Ticari Bilanço Kârı (₺)",
          "İstisna & İndirim (₺)",
          "KKEG (₺)",
          "Asgari KV Farkı (₺)"
      ],
      "ornek": [
          [
              "1. Geçici",
              850000,
              240000,
              45000,
              "=MAX(0,(B6+D6-C6)*0.10-(B6-C6)*0.25*0.3)"
          ],
          [
              "2. Geçici",
              1120000,
              310000,
              62000,
              "=MAX(0,(B7+D7-C7)*0.10-(B7-C7)*0.25*0.3)"
          ],
          [
              "3. Geçici",
              1450000,
              420000,
              78000,
              "=MAX(0,(B8+D8-C8)*0.10-(B8-C8)*0.25*0.3)"
          ],
          [
              "4. Geçici",
              1980000,
              560000,
              95000,
              "=MAX(0,(B9+D9-C9)*0.10-(B9-C9)*0.25*0.3)"
          ],
          [
              "Yıllık Beyan",
              2400000,
              680000,
              120000,
              "=MAX(0,(B10+D10-C10)*0.10-(B10-C10)*0.25*0.3)"
          ],
          [
              "Revize Proj.",
              2750000,
              750000,
              135000,
              "=MAX(0,(B11+D11-C11)*0.10-(B11-C11)*0.25*0.3)"
          ]
      ],
      "metrikler": [
          [
              "Toplam ek vergi farkı",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Fark çıkan dönem",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">0\")",
              "sayi"
          ],
          [
              "Dönem matrah toplamı",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "para"
          ],
          [
              "Ek vergi baskı oranı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B7>1,\"DURDUR\",IF(B6>50000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "İstisna ve indirim kalemlerini Geçici Vergi dönemleri bazında yeniden sınıflandırın.",
          "7524 sayılı Kanun 32/C kapsamındaki indirim tavanlarını kontrol edin.",
          "Tam sürümde 32/A teşvik entegrasyonu, iştirak kazancı istisnası ve vergi kalkanı simülatörü birlikte çalışır."
      ],
      "ui": [
          "Asgari KV Radarı",
          "Matrah farkı görünümü",
          "Ek vergi baskısı"
      ]
  },
  'banka-kredi-covenant-erken-uyari': {
      "karar": "Banka kredi sözleşmelerindeki finansal taahhüt (covenant) ihlal riskini ve erken uyarı eşiklerini gösterir.",
      "girisBasliklari": [
          "Dönem",
          "FAVÖK (₺)",
          "Net Borç (₺)",
          "Faiz Gideri (₺)",
          "Borç / FAVÖK"
      ],
      "ornek": [
          [
              "2025/Q1",
              4200000,
              11500000,
              650000,
              "=C6/MAX(1,B6)"
          ],
          [
              "2025/Q2",
              4600000,
              13200000,
              720000,
              "=C7/MAX(1,B7)"
          ],
          [
              "2025/Q3",
              3900000,
              14800000,
              810000,
              "=C8/MAX(1,B8)"
          ],
          [
              "2025/Q4",
              4800000,
              15100000,
              890000,
              "=C9/MAX(1,B9)"
          ],
          [
              "2026/Q1",
              4500000,
              15900000,
              940000,
              "=C10/MAX(1,B10)"
          ],
          [
              "2026/Q2",
              4700000,
              16200000,
              980000,
              "=C11/MAX(1,B11)"
          ]
      ],
      "metrikler": [
          [
              "En yüksek Borç/FAVÖK",
              "=MAX(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "İhlal riski taşıyan çeyrek",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">3.5\")",
              "sayi"
          ],
          [
              "Toplam net borç",
              "=MAX(DEMO_GIRIS!C6:C25)",
              "para"
          ],
          [
              "Ortalama faiz karşılama",
              "=AVERAGE(DEMO_GIRIS!B6:B25)/MAX(1,AVERAGE(DEMO_GIRIS!D6:D25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B7>0,\"DURDUR\",IF(B6>3.2,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Borç/FAVÖK rasyosu 3.5 sınırına yaklaşan dönemlerde kısa vadeli borçlanmayı durdurun.",
          "Banka konsorsiyumuna bildirim öncesi FAVÖK düzeltmelerini kesinleştirin.",
          "Tam sürümde DSCR, ICR, Net Debt/EBITDA stres testi ve otomatik banka bildirim mektubu üretilir."
      ],
      "ui": [
          "Kredi Covenant Radarı",
          "Finansal rasyo görünümü",
          "İhlal erken uyarısı"
      ]
  },
  'dava-masraf-harc-hesaplama': {
      "karar": "Dava açılış harçlarını, nispi ve maktu gider avansını ve olası aleyhe vekalet ücreti riskini gösterir.",
      "girisBasliklari": [
          "Dava Konusu",
          "Dava Değeri (₺)",
          "Peşin Harç (₺)",
          "Gider Avansı (₺)",
          "Toplam Dava Maliyeti (₺)"
      ],
      "ornek": [
          [
              "Ticari Alacak A",
              450000,
              7688,
              2500,
              "=B6*0.06831/4+D6"
          ],
          [
              "Haksız Fesih B",
              180000,
              3074,
              1800,
              "=B7*0.06831/4+D7"
          ],
          [
              "Tazminat Davası C",
              850000,
              14516,
              3200,
              "=B8*0.06831/4+D8"
          ],
          [
              "Hissedar İhtilafı D",
              1200000,
              20493,
              4000,
              "=B9*0.06831/4+D9"
          ],
          [
              "Sözleşme Cezai Şart E",
              320000,
              5465,
              2100,
              "=B10*0.06831/4+D10"
          ],
          [
              "Fikri Mülkiyet F",
              600000,
              10247,
              2800,
              "=B11*0.06831/4+D11"
          ]
      ],
      "metrikler": [
          [
              "Toplam risk tutarı",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "para"
          ],
          [
              "Toplam yargılama harcı",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Yüksek bütçeli dava sayısı",
              "=COUNTIF(DEMO_GIRIS!B6:B25,\">500000\")",
              "sayi"
          ],
          [
              "Ortalama dava maliyet oranı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B6>3000000,\"DURDUR\",IF(B8>2,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Islah ve nispi peşin harç yükünü dava açılış öncesi finansman planına ekleyin.",
          "Arabuluculuk aşamasında anlaşma ihtimali yüksek dosyaları önceliklendirin.",
          "Tam sürümde AAÜT nispi vekalet baremleri, tebligat/bilirkişi gider avansı ve icra masraf motoru çalışır."
      ],
      "ui": [
          "Dava & Harç Radarı",
          "Yargılama masraf görünümü",
          "Vekalet ücreti riski"
      ]
  },
  'depo-doluluk-lokasyon': {
      "karar": "Depo hacimsel doluluk oranını, atıl alan maliyetini ve lokasyon bazında palet başı gideri gösterir.",
      "girisBasliklari": [
          "Depo/Koridor",
          "Kapasite (Palet)",
          "Mevcut Stok (Palet)",
          "Kira+Enerji (₺)",
          "Doluluk Oranı"
      ],
      "ornek": [
          [
              "Koridor A - Hızlı Tüketim",
              1200,
              1080,
              48000,
              "=C6/MAX(1,B6)"
          ],
          [
              "Koridor B - Soğuk Hava",
              600,
              550,
              72000,
              "=C7/MAX(1,B7)"
          ],
          [
              "Koridor C - Kuru Gıda",
              1500,
              920,
              42000,
              "=C8/MAX(1,B8)"
          ],
          [
              "Koridor D - Ambalaj",
              800,
              420,
              26000,
              "=C9/MAX(1,B9)"
          ],
          [
              "Koridor E - Kimyasal",
              400,
              360,
              38000,
              "=C10/MAX(1,B10)"
          ],
          [
              "Asma Kat - Yedek",
              500,
              180,
              16000,
              "=C11/MAX(1,B11)"
          ]
      ],
      "metrikler": [
          [
              "Ortalama depo doluluğu",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "Kritik doluluk koridoru (>%85)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">0.85\")",
              "sayi"
          ],
          [
              "Düşük verimli alan (<%60)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<0.60\")",
              "sayi"
          ],
          [
              "Toplam palet kapasitesi",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B7>2,\"DURDUR\",IF(B8>2,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "%85 üzeri koridorlarda operasyonel kilitlenmeyi önlemek için ikmal rotalarını revize edin.",
          "Düşük doluluklu alanlardaki sabit gider yükünü yüksek devirli ürünlere kaydırın.",
          "Tam sürümde 3D hacim analitiği, ABC stok yerleşimi, forklift rotalama ve palet başı birim maliyet hesaplanır."
      ],
      "ui": [
          "Depo Alan Radarı",
          "Hacimsel doluluk görünümü",
          "Atıl alan maliyet riski"
      ]
  },
  'dernek-kooperatif-mali-yonetim': {
      "karar": "Dernek ve kooperatiflerde aidat tahsilat performansını, bütçe sapmasını ve nakit karşılık güvencesini gösterir.",
      "girisBasliklari": [
          "Ay",
          "Tahakkuk Eden Aidat (₺)",
          "Tahsil Edilen (₺)",
          "Genel Gider (₺)",
          "Tahsilat Oranı"
      ],
      "ornek": [
          [
              "Ocak",
              145000,
              132000,
              118000,
              "=C6/MAX(1,B6)"
          ],
          [
              "Şubat",
              145000,
              126000,
              122000,
              "=C7/MAX(1,B7)"
          ],
          [
              "Mart",
              150000,
              138000,
              129000,
              "=C8/MAX(1,B8)"
          ],
          [
              "Nisan",
              150000,
              115000,
              134000,
              "=C9/MAX(1,B9)"
          ],
          [
              "Mayıs",
              155000,
              142000,
              130000,
              "=C10/MAX(1,B10)"
          ],
          [
              "Haziran",
              155000,
              120000,
              141000,
              "=C11/MAX(1,B11)"
          ]
      ],
      "metrikler": [
          [
              "Dönem net nakit fazlası",
              "=SUM(DEMO_GIRIS!C6:C25)-SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Ortalama tahsilat oranı",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "Gecikmeli tahsilat tutarı",
              "=SUM(DEMO_GIRIS!B6:B25)-SUM(DEMO_GIRIS!C6:C25)",
              "para"
          ],
          [
              "Kritik açık ayı sayısı",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<0.80\")",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B9>0,\"DURDUR\",IF(B7<0.85,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Tahsilat oranı %80 altına inen aylarda gecikme faizi ve icra takip bildirimlerini başlatın.",
          "Demirbaş ve fon karşılıklarını vadesiz hesap yerine nemalandırılan fonlarda tutun.",
          "Tam sürümde DERBİS uyumlu gelir-gider cetveli, genel kurul bütçe simülatörü ve üye cari defteri açılır."
      ],
      "ui": [
          "Kooperatif Finans Radarı",
          "Aidat tahsilat dengesi",
          "Bütçe sapma riski"
      ]
  },
  'e-belge-zorunluluk-radari': {
      "karar": "Şirketin ciro ve sektör kriterlerine göre e-Fatura, e-Arşiv ve e-Defter geçiş takvimini ve ceza riskini gösterir.",
      "girisBasliklari": [
          "Mali Yıl / Çeyrek",
          "Brüt Satış Hasılatı (₺)",
          "E-Ticaret Hasılatı (₺)",
          "Ciro Eşiği (₺)",
          "Zorunluluk Oranı"
      ],
      "ornek": [
          [
              "2024 Yılı",
              3400000,
              650000,
              3000000,
              "=B6/D6"
          ],
          [
              "2025/Q1",
              1100000,
              220000,
              3000000,
              "=(B7*4)/D7"
          ],
          [
              "2025/Q2",
              2300000,
              480000,
              3000000,
              "=(B8*2)/D8"
          ],
          [
              "2025/Q3",
              3600000,
              750000,
              3000000,
              "=B9/D9"
          ],
          [
              "2025/Q4",
              4900000,
              1100000,
              3000000,
              "=B10/D10"
          ],
          [
              "2026 Hedef",
              6200000,
              1400000,
              3000000,
              "=B11/D11"
          ]
      ],
      "metrikler": [
          [
              "En güncel yıllık ciro",
              "=MAX(DEMO_GIRIS!B6:B25)",
              "para"
          ],
          [
              "Eşik aşım katsayısı",
              "=MAX(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "Eşik aşan dönem sayısı",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">1.0\")",
              "sayi"
          ],
          [
              "E-Ticaret payı",
              "=SUM(DEMO_GIRIS!C6:C25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B8>0,\"DURDUR\",IF(B7>0.9,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Ciro eşiğini aşan dönemlerde 1 Temmuz e-Fatura ve müteakip 1 Ocak e-Defter takvimini başlatın.",
          "Özel entegratör veya GİB portal sözleşmelerini geçiş tarihinden 30 gün önce tamamlayın.",
          "Tam sürümde 509 Sıra No VUK Genel Tebliği sektörel radar, ceza hesaplayıcı ve otomatik başvuru kılavuzu çalışır."
      ],
      "ui": [
          "e-Belge Geçiş Radarı",
          "Ciro eşiği görünümü",
          "Ceza ve takvim riski"
      ]
  },
  'e-ticaret-gercek-karlilik-fiyatlama': {
      "karar": "Pazaryeri komisyonları, kargo, iade ve reklam sonrası e-ticarette gerçek net kârlılığı ve optimum fiyatı gösterir.",
      "girisBasliklari": [
          "Ürün Adı",
          "Satış Fiyatı (₺)",
          "Ürün Maliyeti (₺)",
          "Komisyon+Kargo+İade (₺)",
          "Net Kâr (₺)"
      ],
      "ornek": [
          [
              "Deri Cüzdan A",
              450,
              160,
              145,
              "=B6-C6-D6"
          ],
          [
              "Kablosuz Kulaklık B",
              890,
              420,
              240,
              "=B7-C7-D7"
          ],
          [
              "Termos Kupa C",
              320,
              110,
              125,
              "=B8-C8-D8"
          ],
          [
              "Sırt Çantası D",
              650,
              240,
              195,
              "=B9-C9-D9"
          ],
          [
              "Telefon Kılıfı E",
              180,
              45,
              95,
              "=B10-C10-D10"
          ],
          [
              "Akıllı Saat F",
              1250,
              680,
              310,
              "=B11-C11-D11"
          ]
      ],
      "metrikler": [
          [
              "Toplam net kâr",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Ortalama kâr marjı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Zarar eden ürün sayısı",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<0\")",
              "sayi"
          ],
          [
              "Düşük marjlı (<%15) ürün",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<50\")",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B8>0,\"DURDUR\",IF(B7<0.15,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Zarar eden veya marjı %15'in altına inen ürünlerde pazaryeri komisyon anlaşmalarını güncelleyin.",
          "Kargo ve iade kesintisi satış fiyatının %25'ini aşan kalemlerde sepet ortalamasını artırın.",
          "Tam sürümde Trendyol, Hepsiburada, Amazon komisyon matrisleri, kupon analitiği ve kâr koruma motoru çalışır."
      ],
      "ui": [
          "E-Ticaret Net Kâr Radarı",
          "Birim kârlılık dengesi",
          "Komisyon ve iade erozyonu"
      ]
  },
  'elektrik-faturasi-dogrulama': {
      "karar": "Sanayi ve ticarethane elektrik faturalarında aktif tüketim, reaktif ceza ve dağıtım kalemi tutarlılığını denetler.",
      "girisBasliklari": [
          "Dönem",
          "Aktif Tüketim (kWh)",
          "Reaktif Ceza (₺)",
          "Fatura Tutarı (₺)",
          "Reaktif / Aktif Oranı"
      ],
      "ornek": [
          [
              "Ocak",
              48000,
              0,
              185000,
              "=C6/MAX(1,B6)"
          ],
          [
              "Şubat",
              52000,
              3400,
              204000,
              "=C7/MAX(1,B7)"
          ],
          [
              "Mart",
              49000,
              8900,
              198000,
              "=C8/MAX(1,B8)"
          ],
          [
              "Nisan",
              55000,
              14200,
              226000,
              "=C9/MAX(1,B9)"
          ],
          [
              "Mayıs",
              61000,
              2200,
              238000,
              "=C10/MAX(1,B10)"
          ],
          [
              "Haziran",
              58000,
              0,
              224000,
              "=C11/MAX(1,B11)"
          ]
      ],
      "metrikler": [
          [
              "Toplam reaktif ceza tutarı",
              "=SUM(DEMO_GIRIS!C6:C25)",
              "para"
          ],
          [
              "Cezalı fatura sayısı",
              "=COUNTIF(DEMO_GIRIS!C6:C25,\">0\")",
              "sayi"
          ],
          [
              "Toplam elektrik gideri",
              "=SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Ortalama birim kWh maliyeti",
              "=SUM(DEMO_GIRIS!D6:D25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "para"
          ],
          [
              "Demo karar",
              "=IF(B7>2,\"DURDUR\",IF(B6>10000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Reaktif ceza görülen aylarda kompanzasyon panosu kondansatör ve kademe testini yaptırın.",
          "Endüktif ve kapasitif sayaç çarpanlarını dağıtım şirketiyle yazılı olarak teyit edin.",
          "Tam sürümde EPDK tarife denetimi, reaktif ceza kurtarma simülasyonu ve serbest tüketici sözleşme kontrolü açılır."
      ],
      "ui": [
          "Elektrik Fatura Denetimi",
          "Reaktif ceza radarı",
          "Birim kWh maliyet baskısı"
      ]
  },
  'emlak-pipeline-komisyon': {
      "karar": "Gayrimenkul portföyündeki satış/kiralama pipeline'ını, ağırlıklı komisyon potansiyelini ve nakit akışını gösterir.",
      "girisBasliklari": [
          "Portföy / Gayrimenkul",
          "Liste Fiyatı (₺)",
          "Kapanış Olasılığı",
          "Komisyon Oranı",
          "Beklenen Komisyon (₺)"
      ],
      "ornek": [
          [
              "Rezidans Daire A",
              8500000,
              0.7,
              0.02,
              "=B6*C6*D6"
          ],
          [
              "Ticari Plaza Katı B",
              24000000,
              0.4,
              0.02,
              "=B7*C7*D7"
          ],
          [
              "Villa Projesi C",
              16500000,
              0.5,
              0.02,
              "=B8*C8*D8"
          ],
          [
              "Sanayi Deposu D",
              32000000,
              0.3,
              0.02,
              "=B9*C9*D9"
          ],
          [
              "Cadde Mağaza E",
              14000000,
              0.8,
              0.02,
              "=B10*C10*D10"
          ],
          [
              "Arsa Parsel F",
              9500000,
              0.6,
              0.02,
              "=B11*C11*D11"
          ]
      ],
      "metrikler": [
          [
              "Toplam beklenen komisyon",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Portföy brüt değeri",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "para"
          ],
          [
              "Yüksek olasılıklı (>%60) portföy",
              "=COUNTIF(DEMO_GIRIS!C6:C25,\">0.60\")",
              "sayi"
          ],
          [
              "Ağırlıklı kapanış oranı",
              "=SUM(DEMO_GIRIS!E6:E25)/(SUM(DEMO_GIRIS!B6:B25)*0.02)",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B6<500000,\"DURDUR\",IF(B8<3,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Kapanış olasılığı %40 altında kalan yüksek değerli portföyler için fiyat ekspertizi isteyin.",
          "Yetki sözleşmesi süresi bitmek üzere olan gayrimenkullerde müşteri temasını hızlandırın.",
          "Tam sürümde danışman prim paylaşımı, tapu harcı/stopaj matrahı ve aylık ciro projeksiyonu çalışır."
      ],
      "ui": [
          "Emlak Komisyon Radarı",
          "Pipeline nakit görünümü",
          "Ağırlıklı tahsilat potansiyeli"
      ]
  },
  'fason-uretim-takip-sistemi': {
      "karar": "Fason atölyelere gönderilen hammadde, teslim edilen yarı mamul, fire sapması ve birim fason maliyetini denetler.",
      "girisBasliklari": [
          "Fasoncu / İş Emri",
          "Gönderilen Kumaş (m)",
          "Gelen Ürün (Adet)",
          "Birim Fason Ücreti (₺)",
          "Fiili Fire Oranı"
      ],
      "ornek": [
          [
              "Atölye Alfa - Model 101",
              3500,
              2100,
              48,
              "=1-(C6*1.5/B6)"
          ],
          [
              "Atölye Beta - Model 102",
              4200,
              2350,
              52,
              "=1-(C7*1.6/B7)"
          ],
          [
              "Atölye Gama - Model 103",
              2800,
              1620,
              45,
              "=1-(C8*1.55/B8)"
          ],
          [
              "Atölye Delta - Model 104",
              5100,
              2800,
              55,
              "=1-(C9*1.65/B9)"
          ],
          [
              "Atölye Epsilon - Model 105",
              1900,
              1080,
              50,
              "=1-(C10*1.58/B10)"
          ],
          [
              "Atölye Zeta - Model 106",
              3300,
              1920,
              46,
              "=1-(C11*1.52/B11)"
          ]
      ],
      "metrikler": [
          [
              "Ortalama fiili fire",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "Yüksek fireli atölye (>%8)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">0.08\")",
              "sayi"
          ],
          [
              "Toplam fason hakedişi",
              "=SUMPRODUCT(DEMO_GIRIS!C6:C25,DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Toplam üretilen adet",
              "=SUM(DEMO_GIRIS!C6:C25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B7>1,\"DURDUR\",IF(B6>0.06,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Standart fire payını (%5) aşan atölyelerde hakedişten hammadde bedeli kesintisi uygulayın.",
          "İş emri bazında numune kontrolü ve teslimat tutanağı olmadan ödeme yapmayın.",
          "Tam sürümde pastal planı entegrasyonu, iplik/kumaş denge mutabakatı ve fasoncu kârlılık karnesi çalışır."
      ],
      "ui": [
          "Fason İmalat Radarı",
          "Fire ve teslimat dengesi",
          "Hammadde kayıp riski"
      ]
  },
  'fesih-maliyeti-simulatoru': {
      "karar": "İş akdi feshinde kıdem, ihbar, yıllık izin, işe iade ve arabuluculuk masraflarının net maliyetini gösterir.",
      "girisBasliklari": [
          "Personel / Ünvan",
          "Brüt Ücret (₺)",
          "Kıdem Süresi (Yıl)",
          "İhbar Süresi (Hafta)",
          "Toplam Fesih Maliyeti (₺)"
      ],
      "ornek": [
          [
              "Üretim Müdürü A",
              65000,
              5.5,
              8,
              "=B6*C6+(B6/30*D6*7)"
          ],
          [
              "Muhasebe Uzmanı B",
              38000,
              3.2,
              6,
              "=B7*C7+(B7/30*D7*7)"
          ],
          [
              "Satış Temsilcisi C",
              42000,
              1.8,
              4,
              "=B8*C8+(B8/30*D8*7)"
          ],
          [
              "Vardiya Amiri D",
              46000,
              7.2,
              8,
              "=B9*C9+(B9/30*D9*7)"
          ],
          [
              "Depo Görevlisi E",
              28000,
              4.1,
              6,
              "=B10*C10+(B10/30*D10*7)"
          ],
          [
              "Operatör F",
              31000,
              2.4,
              6,
              "=B11*C11+(B11/30*D11*7)"
          ]
      ],
      "metrikler": [
          [
              "Toplam fesih yükü",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Yüksek maliyetli personel",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">250000\")",
              "sayi"
          ],
          [
              "Ortalama kıdem yılı",
              "=AVERAGE(DEMO_GIRIS!C6:C25)",
              "sayi"
          ],
          [
              "Ortalama kişi başı yük",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Demo karar",
              "=IF(B6>1500000,\"DURDUR\",IF(B7>1,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "İşe iade riski taşıyan personelde arabuluculuk anlaşma protokolünü fesih tarihinden önce hazırlayın.",
          "Kıdem tazminatı tavanını (GİB güncel) aşan ücretlerde vergi kesintisi kontrolü yapın.",
          "Tam sürümde giydirilmiş brüt ücret, boşta geçen süre tazminatı, 4+4 iş güvencesi tazminatı hesaplanır."
      ],
      "ui": [
          "Fesih Maliyet Radarı",
          "Kıdem ve ihbar görünümü",
          "İşe iade tazminat riski"
      ]
  },
  'filo-arac-maliyet-komutasi': {
      "karar": "Şirket araç filosunda yakıt, bakım, kasko, MTV ve km başı birim maliyeti ve araç değişim zamanlamasını gösterir.",
      "girisBasliklari": [
          "Plaka / Model",
          "Aylık Km",
          "Yakıt Gideri (₺)",
          "Bakım+Sigorta+MTV (₺)",
          "Km Başı Maliyet (₺)"
      ],
      "ornek": [
          [
              "34 ABC 101 - Sedan",
              4200,
              14500,
              6200,
              "=(C6+D6)/MAX(1,B6)"
          ],
          [
              "34 ABC 102 - SUV",
              3100,
              15800,
              8400,
              "=(C7+D7)/MAX(1,B7)"
          ],
          [
              "34 ABC 103 - Panelvan",
              5800,
              22400,
              7100,
              "=(C8+D8)/MAX(1,B8)"
          ],
          [
              "34 ABC 104 - Kamyonet",
              4900,
              24100,
              9500,
              "=(C9+D9)/MAX(1,B9)"
          ],
          [
              "34 ABC 105 - Hatchback",
              2600,
              8900,
              4800,
              "=(C10+D10)/MAX(1,B10)"
          ],
          [
              "34 ABC 106 - Panelvan",
              5400,
              21200,
              6900,
              "=(C11+D11)/MAX(1,B11)"
          ]
      ],
      "metrikler": [
          [
              "Aylık toplam filo gideri",
              "=SUM(DEMO_GIRIS!C6:C25)+SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Ortalama km başı maliyet",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Yüksek maliyetli araç (>7 ₺/km)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">7\")",
              "sayi"
          ],
          [
              "Toplam filo kilometresi",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B8>2,\"DURDUR\",IF(B7>6,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Km başı maliyeti 7 TL'yi aşan araçlarda yakıt tüketim anomalisi veya periyodik bakım gecikmesini inceleyin.",
          "150.000 km üzerindeki ticari araçlarda kiralama vs. satın alma karşılaştırması yapın.",
          "Tam sürümde TCO (Toplam Sahip Olma Maliyeti), amortisman erimesi ve optimum araç yenileme yılı çalışır."
      ],
      "ui": [
          "Filo Maliyet Radarı",
          "Km başı gider görünümü",
          "Araç verimsizlik riski"
      ]
  },
  'ges-uretim-performans': {
      "karar": "Güneş enerji santrallerinde (GES) ışınım, kurulu güç, üretilen kWh ve performans oranı (PR) sapmasını gösterir.",
      "girisBasliklari": [
          "Ay",
          "Kurulu Güç (kWp)",
          "Beklenen Üretim (kWh)",
          "Fiili Üretim (kWh)",
          "Performans Oranı (PR)"
      ],
      "ornek": [
          [
              "Ocak",
              1200,
              85000,
              78000,
              "=D6/MAX(1,C6)"
          ],
          [
              "Şubat",
              1200,
              102000,
              96000,
              "=D7/MAX(1,C7)"
          ],
          [
              "Mart",
              1200,
              142000,
              138000,
              "=D8/MAX(1,C8)"
          ],
          [
              "Nisan",
              1200,
              168000,
              159000,
              "=D9/MAX(1,C9)"
          ],
          [
              "Mayıs",
              1200,
              195000,
              189000,
              "=D10/MAX(1,C10)"
          ],
          [
              "Haziran",
              1200,
              210000,
              204000,
              "=D11/MAX(1,C11)"
          ]
      ],
      "metrikler": [
          [
              "Ortalama santral PR",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "Düşük verim ayı (<%92)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<0.92\")",
              "sayi"
          ],
          [
              "Toplam fiili elektrik üretimi",
              "=SUM(DEMO_GIRIS!D6:D25)",
              "sayi"
          ],
          [
              "Üretim kaybı (kWh)",
              "=SUM(DEMO_GIRIS!C6:C25)-SUM(DEMO_GIRIS!D6:D25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B7>1,\"DURDUR\",IF(B6<0.94,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Performans oranı %92 altına inen aylarda invertör arızası ve panel kirlilik temizliğini planlayın.",
          "Şebeke kesintisi veya trafo kayıplarını dağıtım şirketi scada kayıtlarıyla karşılaştırın.",
          "Tam sürümde degredasyon oranı, YEKDEM mahsuplaşması, arıza kök-neden ve aylık gelir kaybı motoru çalışır."
      ],
      "ui": [
          "GES Performans Radarı",
          "Üretim ve PR dengesi",
          "Gelir kayıp riski"
      ]
  },
  'ges-yatirim-fizibilite': {
      "karar": "Çatı ve arazi GES yatırımlarında CAPEX, elektrik üretim projeksiyonu, öz tüketim tasarrufu ve NPV/IRR gösterir.",
      "girisBasliklari": [
          "Yıl",
          "Yıllık Üretim (kWh)",
          "Elektrik Geliri/Tasarruf (₺)",
          "İşletme+Bakım OPEX (₺)",
          "Net Nakit Akışı (₺)"
      ],
      "ornek": [
          [
              "Yıl 1",
              1650000,
              6270000,
              280000,
              "=C6-D6"
          ],
          [
              "Yıl 2",
              1640000,
              6865000,
              310000,
              "=C7-D7"
          ],
          [
              "Yıl 3",
              1630000,
              7517000,
              340000,
              "=C8-D8"
          ],
          [
              "Yıl 4",
              1620000,
              8231000,
              375000,
              "=C9-D9"
          ],
          [
              "Yıl 5",
              1610000,
              9013000,
              415000,
              "=C10-D10"
          ],
          [
              "Yıl 6",
              1600000,
              9869000,
              460000,
              "=C11-D11"
          ]
      ],
      "metrikler": [
          [
              "Kümülatif 6 yıllık net nakit",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Yıllık ortalama tasarruf",
              "=AVERAGE(DEMO_GIRIS!C6:C25)",
              "para"
          ],
          [
              "Toplam üretilecek elektrik",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "İşletme gider oranı",
              "=SUM(DEMO_GIRIS!D6:D25)/MAX(1,SUM(DEMO_GIRIS!C6:C25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B6<25000000,\"DURDUR\",IF(B9>0.08,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Yatırım geri dönüş süresini (Payback) 4.2 yılın altında tutmak için panel tedarik tekliflerini kıyaslayın.",
          "Çağrı mektubu ve bağlantı anlaşması kapasite tahsis harçlarını CAPEX bütçesine dahil edin.",
          "Tam sürümde WACC, Vergi Öncesi/Sonrası IRR, DSCR borç servisi ve 25 yıllık panel degredasyon eğrisi çalışır."
      ],
      "ui": [
          "GES Fizibilite Radarı",
          "Yatırım geri dönüşü",
          "Elektrik tasarruf projeksiyonu"
      ]
  },
  'gida-lot-maliyet-izlenebilirlik': {
      "karar": "Gıda üretiminde parti (lot) bazında hammadde, fire, ambalaj, işçilik ve birim paket maliyetini izler.",
      "girisBasliklari": [
          "Parti (Lot) No",
          "Hammadde Girdisi (kg)",
          "Çıkan Mamul (kg)",
          "Toplam Lot Gideri (₺)",
          "Birim Maliyet (₺/kg)"
      ],
      "ornek": [
          [
              "LOT-2026-001 - Peynir",
              4500,
              680,
              85000,
              "=D6/MAX(1,C6)"
          ],
          [
              "LOT-2026-002 - Yoğurt",
              8000,
              7400,
              112000,
              "=D7/MAX(1,C7)"
          ],
          [
              "LOT-2026-003 - Tereyağı",
              6200,
              520,
              94000,
              "=D8/MAX(1,C8)"
          ],
          [
              "LOT-2026-004 - Kaşar",
              5000,
              510,
              89000,
              "=D9/MAX(1,C9)"
          ],
          [
              "LOT-2026-005 - Ayran",
              9500,
              9100,
              78000,
              "=D10/MAX(1,C10)"
          ],
          [
              "LOT-2026-006 - Lor",
              3800,
              890,
              42000,
              "=D11/MAX(1,C11)"
          ]
      ],
      "metrikler": [
          [
              "Ortalama birim kg maliyeti",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "En yüksek birim maliyetli lot",
              "=MAX(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Toplam işlenen hammadde",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Ortalama randıman oranı",
              "=SUM(DEMO_GIRIS!C6:C25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B7>180,\"DURDUR\",IF(B9<0.30,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Randımanı standart reçetenin %5 altına inen partilerde proses sıcaklık ve süzme kayıplarını denetleyin.",
          "SKT ve lot izlenebilirlik kayıtlarını dijital etiketleme sistemiyle eşleştirin.",
          "Tam sürümde HACCP izlenebilirlik zinciri, geri çağırma (recall) simülatörü ve parti kârlılık matrisi açılır."
      ],
      "ui": [
          "Gıda Lot Radarı",
          "Randıman ve birim maliyet",
          "Proses fire riski"
      ]
  },
  'hizmet-ihracati-100-indirim-transfer-takip': {
      "karar": "Yazılım, tasarım, veri analizi vb. hizmet ihracatında %80/%100 KV kazanç indirimi ve döviz transfer şartını denetler.",
      "girisBasliklari": [
          "Dönem / Fatura",
          "Yurt Dışı Fatura Tutarı (Döviz)",
          "Yurda Getirilen Bedel (₺)",
          "KVK 10/1-ğ Kazanç (₺)",
          "İndirim Tutarı (₺)"
      ],
      "ornek": [
          [
              "Fatura 2026-01 - Yazılım",
              45000,
              1620000,
              1150000,
              "=D6*0.80"
          ],
          [
              "Fatura 2026-02 - Danışmanlık",
              28000,
              1008000,
              720000,
              "=D7*0.80"
          ],
          [
              "Fatura 2026-03 - Tasarım",
              18500,
              666000,
              480000,
              "=D8*0.80"
          ],
          [
              "Fatura 2026-04 - Çağrı Mrk.",
              62000,
              2232000,
              1580000,
              "=D9*0.80"
          ],
          [
              "Fatura 2026-05 - Veri Analitiği",
              34000,
              1224000,
              890000,
              "=D10*0.80"
          ],
          [
              "Fatura 2026-06 - SaaS Lisans",
              52000,
              1872000,
              1340000,
              "=D11*0.80"
          ]
      ],
      "metrikler": [
          [
              "Toplam yararlanılan indirim",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Toplam döviz kazancı (₺)",
              "=SUM(DEMO_GIRIS!C6:C25)",
              "para"
          ],
          [
              "Vergi tasarrufu (%25 KV)",
              "=SUM(DEMO_GIRIS!E6:E25)*0.25",
              "para"
          ],
          [
              "Ortalama indirim oranı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!D6:D25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B6<1000000,\"DURDUR\",IF(B8<300000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Kazancın %50'sinin Türkiye'ye transfer şartını beyanname verilme tarihine kadar tamamlayın.",
          "Hizmetten yurt dışında faydalanıldığını gösteren sözleşme ve teslim kanıtlarını dosyaya ekleyin.",
          "Tam sürümde 7491 SK %80 indirim kuralı, transfer bildirim cetveli ve stopaj iade takibi çalışır."
      ],
      "ui": [
          "Hizmet İhracatı Radarı",
          "Kazanç indirimi görünümü",
          "Döviz transfer şartı riski"
      ]
  },
  'hukuk-burosu-dosya-vekalet': {
      "karar": "Hukuk bürolarında dosya bazlı tahsilat, masraf avansı, karşı yan vekalet ücreti ve avukat kârlılığını izler.",
      "girisBasliklari": [
          "Müvekkil / Dosya",
          "Anlaşılan Vekalet (₺)",
          "Tahsil Edilen (₺)",
          "Yapılan Masraf (₺)",
          "Tahsilat Oranı"
      ],
      "ornek": [
          [
              "Şirket A - Genel Danışmanlık",
              180000,
              150000,
              18000,
              "=C6/MAX(1,B6)"
          ],
          [
              "Holding B - Dava Takibi",
              320000,
              240000,
              42000,
              "=C7/MAX(1,B7)"
          ],
          [
              "Grup C - Tahkim Dosyası",
              450000,
              300000,
              68000,
              "=C8/MAX(1,B8)"
          ],
          [
              "Müvekkil D - İcra Takibi",
              95000,
              65000,
              24000,
              "=C9/MAX(1,B9)"
          ],
          [
              "Şirket E - Sözleşme Revizyon",
              120000,
              120000,
              8500,
              "=C10/MAX(1,B10)"
          ],
          [
              "Grup F - Ceza Dosyası",
              250000,
              175000,
              31000,
              "=C11/MAX(1,B11)"
          ]
      ],
      "metrikler": [
          [
              "Toplam tahsil edilen vekalet",
              "=SUM(DEMO_GIRIS!C6:C25)",
              "para"
          ],
          [
              "Bekleyen vekalet alacağı",
              "=SUM(DEMO_GIRIS!B6:B25)-SUM(DEMO_GIRIS!C6:C25)",
              "para"
          ],
          [
              "Ortalama tahsilat oranı",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "Gecikmeli dosya sayısı (<%75)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<0.75\")",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B9>1,\"DURDUR\",IF(B8<0.80,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Tahsilat oranı %75 altına inen dosyalarda ara hakediş ve masraf kapama talebini iletin.",
          "Gider avansı bakiyesi tükenen dosyalarda masraf ödemelerini durdurun.",
          "Tam sürümde avukat bazlı saatlik verimlilik, serbest meslek makbuzu (SMM) stopajı ve icra harç motoru çalışır."
      ],
      "ui": [
          "Hukuk Bürosu Radarı",
          "Vekalet tahsilat dengesi",
          "Masraf avansı riski"
      ]
  },
  'ihale-fiyat-farki-eskalasyon-pro': {
      "karar": "Kamu ve özel sektör ihalelerinde TUİK endeksleri, eskalasyon formülleri ve hak kaybı riskini hesaplar.",
      "girisBasliklari": [
          "Hakediş No / Ay",
          "Hakediş Tutarı (₺)",
          "Temel Endeks (Pn)",
          "Güncel Endeks (Po)",
          "Fiyat Farkı Tutarı (₺)"
      ],
      "ornek": [
          [
              "Hakediş 1 - Ocak",
              4200000,
              1250,
              1340,
              "=B6*(D6/C6-1)"
          ],
          [
              "Hakediş 2 - Şubat",
              4800000,
              1250,
              1395,
              "=B7*(D7/C7-1)"
          ],
          [
              "Hakediş 3 - Mart",
              5100000,
              1250,
              1460,
              "=B8*(D8/C8-1)"
          ],
          [
              "Hakediş 4 - Nisan",
              5600000,
              1250,
              1530,
              "=B9*(D9/C9-1)"
          ],
          [
              "Hakediş 5 - Mayıs",
              6200000,
              1250,
              1610,
              "=B10*(D10/C10-1)"
          ],
          [
              "Hakediş 6 - Haziran",
              6800000,
              1250,
              1690,
              "=B11*(D11/C11-1)"
          ]
      ],
      "metrikler": [
          [
              "Toplam fiyat farkı alacağı",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Kümülatif hakediş tutarı",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "para"
          ],
          [
              "En yüksek endeks artış oranı",
              "=MAX(DEMO_GIRIS!D6:D25)/MIN(DEMO_GIRIS!C6:C25)-1",
              "oran"
          ],
          [
              "Fiyat farkı / Hakediş oranı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B6<500000,\"DURDUR\",IF(B8>0.25,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "4734 sayılı KİK Fiyat Farkı Kararnamesi ağırlık katsayılarını (a1, a2, b, c) sözleşmeye göre teyit edin.",
          "İdarenin revize zeyilname ve ceza kesintisi uygulamalarını itirazi kayıtla imzalayın.",
          "Tam sürümde mazot, çimento, demir katsayıları, kriz kararnameleri ve gecikme cezası mahsubu çalışır."
      ],
      "ui": [
          "İhale Fiyat Farkı Radarı",
          "Eskalasyon alacak görünümü",
          "Hak kaybı riski"
      ]
  },
  'ihracat-siparis-karlilik-kur': {
      "karar": "Dövizli ihracat siparişlerinde kur riski, navlun, gümrük masrafı ve net TL kârlılık marjını gösterir.",
      "girisBasliklari": [
          "Sipariş / Ülke",
          "Döviz Tutarı ($/€)",
          "Döviz Kuru (₺)",
          "TL Üretim Maliyeti (₺)",
          "Net Kâr Marjı"
      ],
      "ornek": [
          [
              "Sipariş DE-101 - Almanya",
              85000,
              39.5,
              2450000,
              "=(B6*C6-D6)/(B6*C6)"
          ],
          [
              "Sipariş US-102 - ABD",
              120000,
              36.2,
              3100000,
              "=(B7*C7-D7)/(B7*C7)"
          ],
          [
              "Sipariş UK-103 - İngiltere",
              65000,
              46.8,
              2180000,
              "=(B8*C8-D8)/(B8*C8)"
          ],
          [
              "Sipariş FR-104 - Fransa",
              94000,
              39.5,
              2750000,
              "=(B9*C9-D9)/(B9*C9)"
          ],
          [
              "Sipariş IT-105 - İtalya",
              78000,
              39.5,
              2300000,
              "=(B10*C10-D10)/(B10*C10)"
          ],
          [
              "Sipariş ES-106 - İspanya",
              110000,
              39.5,
              3250000,
              "=(B11*C11-D11)/(B11*C11)"
          ]
      ],
      "metrikler": [
          [
              "Toplam ciro (₺)",
              "=SUMPRODUCT(DEMO_GIRIS!B6:B25,DEMO_GIRIS!C6:C25)",
              "para"
          ],
          [
              "Ortalama ihracat marjı",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "Kritik marjlı sipariş (<%20)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<0.20\")",
              "sayi"
          ],
          [
              "Toplam ihracat geliri (Döviz)",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B8>1,\"DURDUR\",IF(B7<0.25,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Kâr marjı %20 altına inen siparişlerde vadeli kur (forward) koruma kontratı yapın.",
          "Dahilde İşleme İzin Belgesi (DİİB) kapsamında hammadde alımlarını gümrük vergisiz kapatın.",
          "Tam sürümde döviz açık pozisyon stres testi, navlun optimizasyonu ve ihracat teşvik modülü çalışır."
      ],
      "ui": [
          "İhracat Kârlılık Radarı",
          "Döviz ve kur marjı",
          "Kur riski erozyonu"
      ]
  },
  'insaat-hakedis-yonetim-sistemi': {
      "karar": "İnşaat projelerinde kümülatif imalat, avans, teminat, stopaj kesintileri ve ödenecek net hakedişi hesaplar.",
      "girisBasliklari": [
          "Hakediş No",
          "Dönem İmalat Tutarı (₺)",
          "Avans Kesintisi (₺)",
          "Nakit Teminat %5 (₺)",
          "Ödenecek Net Tutar (₺)"
      ],
      "ornek": [
          [
              "Hakediş 01",
              3800000,
              760000,
              190000,
              "=B6-C6-D6-(B6*0.05)"
          ],
          [
              "Hakediş 02",
              4500000,
              900000,
              225000,
              "=B7-C7-D7-(B7*0.05)"
          ],
          [
              "Hakediş 03",
              5200000,
              1040000,
              260000,
              "=B8-C8-D8-(B8*0.05)"
          ],
          [
              "Hakediş 04",
              6100000,
              1220000,
              305000,
              "=B9-C9-D9-(B9*0.05)"
          ],
          [
              "Hakediş 05",
              5800000,
              1160000,
              290000,
              "=B10-C10-D10-(B10*0.05)"
          ],
          [
              "Hakediş 06",
              4900000,
              980000,
              245000,
              "=B11-C11-D11-(B11*0.05)"
          ]
      ],
      "metrikler": [
          [
              "Toplam ödenen net hakediş",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Kümülatif kesinti tutarı",
              "=SUM(DEMO_GIRIS!B6:B25)-SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Kümülatif brüt imalat",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "para"
          ],
          [
              "Ortalama net ödeme oranı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B6<10000000,\"DURDUR\",IF(B9<0.70,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Yıllara sari inşaat stopaj kesintisi (%5) mahsup evraklarını vergi dairesi sisteminden doğrulayın.",
          "Nakit teminat kesintileri yerine teminat mektubu vererek nakit blokajını kaldırın.",
          "Tam sürümde yeşil defter, metraj cetveli, fiyat farkı eskalasyonu ve alt yüklenici mutabakatı açılır."
      ],
      "ui": [
          "İnşaat Hakediş Radarı",
          "Kesinti ve nakit dengesi",
          "Nakit blokaj riski"
      ]
  },
  'isg-risk-degerlendirme-pro-6331': {
      "karar": "6331 sayılı Kanun kapsamında işyerindeki tehlikeleri, olasılık-şiddet matrisini ve acil aksiyonları derecelendirir.",
      "girisBasliklari": [
          "Bölüm / Tehlike Kaynağı",
          "Olasılık (1-5)",
          "Şiddet (1-5)",
          "Aksiyon Maliyeti (₺)",
          "Risk Skoru (OxŞ)"
      ],
      "ornek": [
          [
              "Pres Bölümü - Koruyucu Eksikliği",
              4,
              5,
              24000,
              "=B6*C6"
          ],
          [
              "Kimyasal Depo - Havalandırma",
              3,
              4,
              38000,
              "=B7*C7"
          ],
          [
              "Kaynak Atölyesi - Gaz Kaçağı",
              2,
              5,
              18000,
              "=B8*C8"
          ],
          [
              "Yüksekte Çalışma - İskele Güvenliği",
              4,
              4,
              32000,
              "=B9*C9"
          ],
          [
              "Elektrik Panosu - Açık İletken",
              3,
              5,
              12000,
              "=B10*C10"
          ],
          [
              "Forklift Yolu - Yaya Çizgisi",
              3,
              3,
              8500,
              "=B11*C11"
          ]
      ],
      "metrikler": [
          [
              "Kritik risk sayısı (Skor >=15)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">=15\")",
              "sayi"
          ],
          [
              "En yüksek risk skoru",
              "=MAX(DEMO_GIRIS!E6:E25)",
              "sayi"
          ],
          [
              "Gerekli önleyici yatırım bütçesi",
              "=SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Ortalama risk skoru",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B6>2,\"DURDUR\",IF(B7>=20,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Risk skoru 15 ve üzeri olan yüksek tehlikeli alanlarda derhal durdurma ve mühendislik önlemi alın.",
          "İş güvenliği uzmanı ve işyeri hekimi onaylı risk analiz raporunu İSG-KATİP sistemine kaydedin.",
          "Tam sürümde Fine-Kinney yöntemi, 5x5 L Matrisi, ramak kala bildirim takip ve yasal teftiş denetim listesi bulunur."
      ],
      "ui": [
          "İSG Risk Radarı",
          "6331 tehlike görünümü",
          "Yasal teftiş ve kaza riski"
      ]
  },
  'ithalat-landed-cost-motoru': {
      "karar": "İthalatta FOB bedel, navlun, sigorta, gümrük vergisi, ÖTV ve KKDF sonrası kapı teslim birim maliyeti gösterir.",
      "girisBasliklari": [
          "Ürün / GTİP",
          "FOB Fatura ($)",
          "Navlun+Sigorta ($)",
          "Gümrük+ÖTV+KKDF (₺)",
          "Kapı Teslim Birim (₺)"
      ],
      "ornek": [
          [
              "Elektronik Kart A - 8534.00",
              35000,
              4200,
              485000,
              "=((B6+C6)*36.5)+D6"
          ],
          [
              "Hidrolik Pompa B - 8413.60",
              48000,
              5800,
              620000,
              "=((B7+C7)*36.5)+D7"
          ],
          [
              "Optik Sensör C - 9031.80",
              22000,
              2600,
              295000,
              "=((B8+C8)*36.5)+D8"
          ],
          [
              "Paslanmaz Çelik D - 7219.33",
              64000,
              8900,
              840000,
              "=((B9+C9)*36.5)+D9"
          ],
          [
              "Rulman Grubu E - 8482.10",
              18500,
              2100,
              245000,
              "=((B10+C10)*36.5)+D10"
          ],
          [
              "Pnömatik Valf F - 8481.20",
              29000,
              3400,
              390000,
              "=((B11+C11)*36.5)+D11"
          ]
      ],
      "metrikler": [
          [
              "Toplam ithalat maliyeti (₺)",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Toplam ödenen vergi (₺)",
              "=SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Efektif vergi yükü oranı",
              "=SUM(DEMO_GIRIS!D6:D25)/MAX(1,SUM(DEMO_GIRIS!E6:E25))",
              "oran"
          ],
          [
              "Toplam FOB bedel ($)",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B8>0.35,\"DURDUR\",IF(B6>3000000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Gözetim ve referans fiyat uygulaması olan GTİP kodlarında ek teminat maliyetini hesaba katın.",
          "Vadeli ithalatta KKDF (%6) doğmaması için peşin akreditif alternatiflerini değerlendirin.",
          "Tam sürümde antrepo ardiye maliyeti, damping vergisi, ilave gümrük vergisi (İGV) ve koli başı maliyet çalışır."
      ],
      "ui": [
          "İthalat Maliyet Radarı",
          "Landed cost görünümü",
          "Vergi ve fon baskısı"
      ]
  },
  'kargo-desi-maliyet-optimizasyonu': {
      "karar": "E-ticaret ve kargo gönderilerinde faturalanan desi ile gerçek desi farkını, ambalaj kaybını ve maliyet sızıntısını bulur.",
      "girisBasliklari": [
          "Paket / Sipariş",
          "Gerçek Desi",
          "Faturalanan Desi",
          "Kargo Fatura Tutarı (₺)",
          "Desi Fark Maliyeti (₺)"
      ],
      "ornek": [
          [
              "Sipariş TR-101 - Ayakkabı",
              3.2,
              5,
              78,
              "=(C6-B6)*14.5"
          ],
          [
              "Sipariş TR-102 - Mont",
              6.5,
              9,
              142,
              "=(C7-B7)*14.5"
          ],
          [
              "Sipariş TR-103 - Takı Seti",
              0.8,
              2,
              48,
              "=(C8-B8)*14.5"
          ],
          [
              "Sipariş TR-104 - Küçük Ev Aleti",
              8.2,
              12,
              185,
              "=(C9-B9)*14.5"
          ],
          [
              "Sipariş TR-105 - Kozmetik Kutu",
              1.5,
              3,
              56,
              "=(C10-B10)*14.5"
          ],
          [
              "Sipariş TR-106 - Ev Tekstili",
              5,
              7.5,
              118,
              "=(C11-B11)*14.5"
          ]
      ],
      "metrikler": [
          [
              "Toplam desi fark kaybı",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Toplam kargo harcaması",
              "=SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Desi sızıntı oranı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!D6:D25))",
              "oran"
          ],
          [
              "Hatalı ölçümlenen paket sayısı",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">20\")",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B8>0.20,\"DURDUR\",IF(B6>150,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Desi farkı %20'yi aşan paketlerde kargo şirketiyle şube ölçüm kalibrasyon itirazı başlatın.",
          "Hacimsel ağırlığı düşürmek için ambalaj kutu ebatlarını optimize edin ve hava yastıklarını küçültün.",
          "Tam sürümde kargo barem denetimi, tazmin süreci mutabakatı ve çoklu taşıyıcı fiyat kıyaslama motoru açılır."
      ],
      "ui": [
          "Kargo Desi Radarı",
          "Fatura ve desi dengesi",
          "Ambalaj maliyet kaybı"
      ]
  },
  'kat-karsiligi-hasilat-paylasimi-simulator': {
      "karar": "Kentsel dönüşüm ve kat karşılığı projelerde arsa sahibi/müteahhit paylaşım oranını, hasılatı ve inşaat fizibilitesini test eder.",
      "girisBasliklari": [
          "Bağımsız Bölüm Grubu",
          "İnşaat Alanı (m²)",
          "Satış Değeri (₺/m²)",
          "İnşaat Maliyeti (₺/m²)",
          "Müteahhit Brüt Kârı (₺)"
      ],
      "ornek": [
          [
              "A Blok Konutlar (%50 Pay)",
              4500,
              65000,
              26000,
              "=B6*(C6*0.50-D6)"
          ],
          [
              "B Blok Konutlar (%50 Pay)",
              5200,
              68000,
              26000,
              "=B7*(C7*0.50-D7)"
          ],
          [
              "Zemin Cadde Dükkanlar",
              1800,
              145000,
              32000,
              "=B8*(C8*0.50-D8)"
          ],
          [
              "C Blok Ofis Katları",
              3200,
              85000,
              28000,
              "=B9*(C9*0.50-D9)"
          ],
          [
              "Kapalı Otopark & Sosyal",
              2400,
              25000,
              18000,
              "=B10*(C10*0.50-D10)"
          ],
          [
              "Çatı Dubleksler",
              1600,
              95000,
              30000,
              "=B11*(C11*0.50-D11)"
          ]
      ],
      "metrikler": [
          [
              "Toplam müteahhit brüt kârı",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Toplam proje hasılatı",
              "=SUMPRODUCT(DEMO_GIRIS!B6:B25,DEMO_GIRIS!C6:C25)",
              "para"
          ],
          [
              "Toplam inşaat maliyeti",
              "=SUMPRODUCT(DEMO_GIRIS!B6:B25,DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Proje kâr marjı",
              "=SUM(DEMO_GIRIS!E6:E25)/(SUMPRODUCT(DEMO_GIRIS!B6:B25,DEMO_GIRIS!C6:C25)*0.5)",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B9<0.15,\"DURDUR\",IF(B6<20000000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Müteahhit payı kâr marjı %20 altına iniyorsa arsa sahibi paylaşım oranını %45 seviyesine çekin.",
          "Ruhsat ve hafriyat gecikmelerine karşı sözleşmeye mücbir sebep ve eskalasyon maddesi ekleyin.",
          "Tam sürümde şerefiye puanlama matrisi, nakit hakediş takvimi ve hasılat paylaşımlı gelir dağıtım motoru çalışır."
      ],
      "ui": [
          "Hasılat Paylaşım Radarı",
          "Arsa ve inşaat dengesi",
          "Proje kârlılık riski"
      ]
  },
  'kdv-iadesi-tutar-surec-simulasyonu': {
      "karar": "İhracat ve indirimli oranda KDV iadesi potansiyelini, yüklenim listesi tenzil riskini ve mahsup süresini simüle eder.",
      "girisBasliklari": [
          "Dönem / İşlem Türü",
          "İade Hakkı Doğuran İşlem (₺)",
          "Yüklenilen KDV (₺)",
          "Tenzil Edilen KDV (₺)",
          "Talep Edilen Net İade (₺)"
      ],
      "ornek": [
          [
              "Ocak - Mal İhracatı",
              8500000,
              1150000,
              45000,
              "=C6-D6"
          ],
          [
              "Şubat - Mal İhracatı",
              9200000,
              1280000,
              62000,
              "=C7-D7"
          ],
          [
              "Mart - İndirimli Oran",
              6400000,
              780000,
              38000,
              "=C8-D8"
          ],
          [
              "Nisan - Mal İhracatı",
              11000000,
              1540000,
              85000,
              "=C9-D9"
          ],
          [
              "Mayıs - Hizmet İhracatı",
              4800000,
              580000,
              22000,
              "=C10-D10"
          ],
          [
              "Haziran - Mal İhracatı",
              12500000,
              1720000,
              94000,
              "=C11-D11"
          ]
      ],
      "metrikler": [
          [
              "Toplam talep edilen KDV iadesi",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Toplam tenzil riski tutarı",
              "=SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Kümülatif iade matrahı",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "para"
          ],
          [
              "Efektif iade oranı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B6<3000000,\"DURDUR\",IF(B7>200000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Tenzil riski %5'i aşan faturalarda alt tedarikçi KDV beyanname ve ödeme teyitlerini yapın.",
          "YMM KDV İadesi Tasdik Raporunu beyanname verme süresini takip eden ay içinde sisteme yükleyin.",
          "Tam sürümde GEKSİS hata kontrol simülatörü, yüklenim listesi doğrulama ve vergi borcu mahsup planlayıcı çalışır."
      ],
      "ui": [
          "KDV İadesi Radarı",
          "Yüklenim ve iade dengesi",
          "GEKSİS tenzil riski"
      ]
  },
  'kik-asiri-dusuk-savunma-sinir-deger': {
      "karar": "Kamu ihalelerinde sınır değer hesabını, yaklaşık maliyet katsayısını ve aşırı düşük teklif sorgu riskini gösterir.",
      "girisBasliklari": [
          "Teklif Sahibi / Firma",
          "Teklif Tutarı (₺)",
          "Yaklaşık Maliyet (₺)",
          "Sınır Değer (₺)",
          "Sınır Değer Farkı (₺)"
      ],
      "ornek": [
          [
              "Bizim Teklifimiz",
              18500000,
              24000000,
              19200000,
              "=B6-D6"
          ],
          [
              "Rakip Firma A",
              19800000,
              24000000,
              19200000,
              "=B7-D7"
          ],
          [
              "Rakip Firma B",
              21200000,
              24000000,
              19200000,
              "=B8-D8"
          ],
          [
              "Rakip Firma C",
              18900000,
              24000000,
              19200000,
              "=B9-D9"
          ],
          [
              "Rakip Firma D",
              22500000,
              24000000,
              19200000,
              "=B10-D10"
          ],
          [
              "Rakip Firma E",
              19100000,
              24000000,
              19200000,
              "=B11-D11"
          ]
      ],
      "metrikler": [
          [
              "Sınır değer altı teklif sayısı",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<0\")",
              "sayi"
          ],
          [
              "En düşük sınır değer farkı",
              "=MIN(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Bizim teklif sınır oranı",
              "=B6/D6",
              "oran"
          ],
          [
              "Ortalama geçerli teklif",
              "=AVERAGE(DEMO_GIRIS!B6:B25)",
              "para"
          ],
          [
              "Demo karar",
              "=IF(B8<0.98,\"DURDUR\",IF(B6>2,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Teklifiniz sınır değerin altında kalıyorsa analiz formatına uygun aşırı düşük savunma dosyasını 3 gün içinde hazırlayın.",
          "Çevre ve Şehircilik Bakanlığı pozları ile piyasa proforma faturalarını üçüncü kişi onaylı teyit edin.",
          "Tam sürümde KİK Tebliği Madde 45/79 şablonları, aritmetik ortalama simülatörü ve analiz savunma robotu bulunur."
      ],
      "ui": [
          "KİK Sınır Değer Radarı",
          "Aşırı düşük sorgu görünümü",
          "İhale elenme riski"
      ]
  },
  'kira-portfoyu-getiri-komutasi': {
      "karar": "Gayrimenkul kira portföyünde tahsilat performansı, boş kalma oranı, işletme masrafı ve net Cap Rate getirisini gösterir.",
      "girisBasliklari": [
          "Gayrimenkul / Kiracı",
          "Ekspertiz Değeri (₺)",
          "Yıllık Kira Geliri (₺)",
          "Aidat+Vergi+Bakım (₺)",
          "Net Getiri (Cap Rate)"
      ],
      "ornek": [
          [
              "Cadde Mağaza 1 - Banka",
              35000000,
              2450000,
              180000,
              "=(C6-D6)/B6"
          ],
          [
              "Ofis Katı 2 - Yazılım",
              18000000,
              1350000,
              120000,
              "=(C7-D7)/B7"
          ],
          [
              "Lojistik Depo 3 - Kargo",
              48000000,
              3600000,
              240000,
              "=(C8-D8)/B8"
          ],
          [
              "Konut Portföyü 4 - Rezidans",
              14000000,
              840000,
              95000,
              "=(C9-D9)/B9"
          ],
          [
              "AVM Mağaza 5 - Giyim",
              22000000,
              1650000,
              160000,
              "=(C10-D10)/B10"
          ],
          [
              "Müstakil Bina 6 - Klinik",
              28000000,
              2100000,
              190000,
              "=(C11-D11)/B11"
          ]
      ],
      "metrikler": [
          [
              "Portföy ağırlıklı net getiri",
              "=SUM(DEMO_GIRIS!C6:C25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Toplam yıllık net kira nakdi",
              "=SUM(DEMO_GIRIS!C6:C25)-SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Toplam portföy varlık değeri",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "para"
          ],
          [
              "Düşük verimli mülk (<%6)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<0.06\")",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B9>1,\"DURDUR\",IF(B6<0.065,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Cap Rate getirisi %6 altına inen konut ve ofis mülklerinde satış veya kentsel dönüşüm alternatifini inceleyin.",
          "TÜFE kira artış tavanı ve 5 yıllık kira tespit davası sürelerini portföy takvimine işleyin.",
          "Tam sürümde tahliye taahhüdü takibi, stopaj/KDV tevkifatı ve portföy getiri maksimizasyon simülatörü çalışır."
      ],
      "ui": [
          "Kira Portföy Radarı",
          "Cap Rate getiri dengesi",
          "Boş kalma ve getiri riski"
      ]
  },
  'konkordato-ttk376-kriz-paketi': {
      "karar": "Şirketin TTK 376 kapsamındaki sermaye kaybı ve borca batıklık derecesini, konkordato ön projeksiyonunu denetler.",
      "girisBasliklari": [
          "Dönem",
          "Ödenmiş Sermaye (₺)",
          "Geçmiş Yıl Zararları (₺)",
          "Dönem Net Zararı (₺)",
          "Sermaye Kayıp Oranı"
      ],
      "ornek": [
          [
              "2024 Yıl Sonu",
              15000000,
              4200000,
              3100000,
              "=(C6+D6)/B6"
          ],
          [
              "2025/Q1",
              15000000,
              7300000,
              1800000,
              "=(C7+D7)/B7"
          ],
          [
              "2025/Q2",
              15000000,
              9100000,
              2400000,
              "=(C8+D8)/B8"
          ],
          [
              "2025/Q3",
              15000000,
              11500000,
              2900000,
              "=(C9+D9)/B9"
          ],
          [
              "2025/Q4 Projeksiyon",
              15000000,
              14400000,
              3500000,
              "=(C10+D10)/B10"
          ],
          [
              "2026/Q1 Stres Test",
              15000000,
              17900000,
              4100000,
              "=(C11+D11)/B11"
          ]
      ],
      "metrikler": [
          [
              "En yüksek sermaye kayıp oranı",
              "=MAX(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "Borca batıklık (>%100) dönemi",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">1.0\")",
              "sayi"
          ],
          [
              "Kritik dönem (TTK 376/2 >%66)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">0.66\")",
              "sayi"
          ],
          [
              "Toplam kümülatif zarar",
              "=MAX(DEMO_GIRIS!C6:C25)+MAX(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Demo karar",
              "=IF(B7>0,\"DURDUR\",IF(B8>0,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Sermaye kaybı 2/3 eşiğini (%66) aştığında derhal genel kurulu toplayıp sermaye artırımı veya tamamlama kararı alın.",
          "Yeniden değerleme fonlarını özkaynaklara aktararak bilançoyu yasal olarak güçlendirin.",
          "Tam sürümde borca batıklık bilançosu, konkordato ön projesi nakit akışı ve alacaklı tenzilat tablosu açılır."
      ],
      "ui": [
          "TTK 376 Kriz Radarı",
          "Sermaye kaybı görünümü",
          "Borca batıklık riski"
      ]
  },
  'kvkk-veri-envanteri-verbis-uyum': {
      "karar": "Kişisel veri envanterini, VERBİS kayıt yükümlülüğünü, saklama sürelerini ve idari para cezası riskini gösterir.",
      "girisBasliklari": [
          "Veri Kategorisi / Süreç",
          "Kayıt Sayısı (Kişi)",
          "Yurt Dışı Aktarım",
          "Tedbir Uyum Puanı (1-100)",
          "Ceza Risk Skoru"
      ],
      "ornek": [
          [
              "Müşteri Finansal Verisi",
              45000,
              1,
              62,
              "=(100-D6)*1.5"
          ],
          [
              "Çalışan Özlük Dosyaları",
              280,
              0,
              78,
              "=(100-D7)*1.0"
          ],
          [
              "Kamera Güvenlik Kayıtları",
              12000,
              0,
              85,
              "=(100-D8)*0.8"
          ],
          [
              "Pazarlama Çerez & İletişim",
              185000,
              1,
              48,
              "=(100-D9)*1.8"
          ],
          [
              "Biyometrik Giriş Verisi",
              350,
              0,
              42,
              "=(100-D10)*2.0"
          ],
          [
              "Tedarikçi Yetkili Verisi",
              850,
              0,
              72,
              "=(100-D11)*1.0"
          ]
      ],
      "metrikler": [
          [
              "En yüksek risk skoru",
              "=MAX(DEMO_GIRIS!E6:E25)",
              "sayi"
          ],
          [
              "Yüksek riskli süreç sayısı",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">80\")",
              "sayi"
          ],
          [
              "Ortalama teknik tedbir puanı",
              "=AVERAGE(DEMO_GIRIS!D6:D25)",
              "oran"
          ],
          [
              "Toplam işlenen ilgili kişi sayısı",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B7>1,\"DURDUR\",IF(B8<70,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Açık rıza ve aydınlatma metni eksik olan süreçlerde kişisel veri işlemeyi derhal durdurun.",
          "Yurt dışı aktarım yapılan bulut yazılımlarda taahhütname veya standart sözleşme bildirimini yapın.",
          "Tam sürümde VERBİS kategori eşleştirme, saklama ve imha politikası ve Kurul güncel ceza hesaplayıcı çalışır."
      ],
      "ui": [
          "KVKK Uyum Radarı",
          "Envanter ve tedbir dengesi",
          "VERBİS idari para cezası riski"
      ]
  },
  'makine-bakim-kalibrasyon-durus-maliyeti': {
      "karar": "Üretim tesislerinde plansız makine duruşlarının, kalibrasyon sapmalarının ve kayıp ciro maliyetini hesaplar.",
      "girisBasliklari": [
          "Makine / Hat",
          "Plansız Duruş (Saat)",
          "Yedek Parça Maliyeti (₺)",
          "Saatlik Kayıp Ciro (₺)",
          "Toplam Duruş Kaybı (₺)"
      ],
      "ornek": [
          [
              "CNC İşleme Merkezi 1",
              18.5,
              45000,
              6500,
              "=B6*D6+C6"
          ],
          [
              "Enjeksiyon Hattı 2",
              24,
              68000,
              8200,
              "=B7*D7+C7"
          ],
          [
              "Robotik Kaynak Hücresi 3",
              12,
              28000,
              5400,
              "=B8*D8+C8"
          ],
          [
              "Otomatik Boya Hattı 4",
              32.5,
              95000,
              9800,
              "=B9*D9+C9"
          ],
          [
              "Paketleme Konveyörü 5",
              8,
              14000,
              4200,
              "=B10*D10+C10"
          ],
          [
              "Pres Hattı 6",
              16,
              38000,
              7500,
              "=B11*D11+C11"
          ]
      ],
      "metrikler": [
          [
              "Toplam duruş maliyeti",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Toplam kayıp saat",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Kritik duruş sayısı (>20 saat)",
              "=COUNTIF(DEMO_GIRIS!B6:B25,\">20\")",
              "sayi"
          ],
          [
              "Ortalama saatlik duruş kaybı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "para"
          ],
          [
              "Demo karar",
              "=IF(B6>800000,\"DURDUR\",IF(B8>1,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Duruş süresi 20 saati aşan kritik makinelerde kestirimci bakım ve titreşim analizini devreye alın.",
          "Kritik yedek parça emniyet stoğu seviyelerini tedarik süresine göre revize edin.",
          "Tam sürümde MTBF (Arızalar Arası Ortalama Süre), MTTR (Ortalama Onarım Süresi) ve OEE entegrasyonu bulunur."
      ],
      "ui": [
          "Bakım & Duruş Radarı",
          "Duruş maliyet görünümü",
          "Kayıp üretim riski"
      ]
  },
  'mini-mrp-bom-malzeme-kapasite': {
      "karar": "Ürün ağacı (BOM), net hammadde ihtiyacı, emniyet stoku ve tedarik sipariş emirlerini otomatik dengeler.",
      "girisBasliklari": [
          "Bileşen / Hammadde",
          "Brüt İhtiyaç (Adet/kg)",
          "Mevcut Stok",
          "Emniyet Stoku",
          "Net Satın Alma Emri"
      ],
      "ornek": [
          [
              "Alüminyum Profil 6063",
              4500,
              1800,
              500,
              "=MAX(0,B6-C6+D6)"
          ],
          [
              "Paslanmaz Civata M8",
              28000,
              14000,
              3000,
              "=MAX(0,B7-C7+D7)"
          ],
          [
              "Elektronik Kart MCU-01",
              1200,
              450,
              200,
              "=MAX(0,B8-C8+D8)"
          ],
          [
              "Plastik Gövde Enjeksiyon",
              1850,
              1900,
              300,
              "=MAX(0,B9-C9+D9)"
          ],
          [
              "Güç Kaynağı 24V 5A",
              950,
              320,
              150,
              "=MAX(0,B10-C10+D10)"
          ],
          [
              "Ambalaj Kolisi Standart",
              3200,
              800,
              600,
              "=MAX(0,B11-C11+D11)"
          ]
      ],
      "metrikler": [
          [
              "Sipariş verilecek kalem sayısı",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">0\")",
              "sayi"
          ],
          [
              "Toplam net sipariş adedi",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "sayi"
          ],
          [
              "Stok kapsama yeterliliği",
              "=SUM(DEMO_GIRIS!C6:C25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Emniyet stoku altı kalem",
              "=COUNTIF(DEMO_GIRIS!C6:C25,\"<D6\")",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B8<0.50,\"DURDUR\",IF(B6>4,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Net satın alma emri çıkan kalemlerde tedarikçi temin sürelerini kontrol ederek acil sipariş açın.",
          "Tedarik riski yüksek elektronik ve ithal bileşenlerde emniyet stoku gün sayısını artırın.",
          "Tam sürümde çok seviyeli BOM patlatma, iş merkezi kapasite çizelgeleme ve satın alma bütçe planlayıcı çalışır."
      ],
      "ui": [
          "Mini-MRP Radarı",
          "Malzeme ihtiyaç dengesi",
          "Hammadde yok satma riski"
      ]
  },
  'mizan-anomali-denetim-oncesi-kontrol': {
      "karar": "Geçici ve kesin mizanda ters bakiye veren hesapları, bakiye uyumsuzluklarını ve vergi inceleme risklerini yakalar.",
      "girisBasliklari": [
          "Hesap Kodu / Adı",
          "Borç Toplamı (₺)",
          "Alacak Toplamı (₺)",
          "Borç Kalan (₺)",
          "Anomali Puanı (1-100)"
      ],
      "ornek": [
          [
              "100 Kasa Hesabı",
              2450000,
              2480000,
              -30000,
              "=IF(D6<0,100,0)"
          ],
          [
              "102 Bankalar",
              18500000,
              18200000,
              300000,
              "=IF(D7<0,90,0)"
          ],
          [
              "120 Alıcılar",
              14200000,
              14900000,
              -700000,
              "=IF(D8<0,85,0)"
          ],
          [
              "320 Satıcılar",
              11800000,
              11200000,
              600000,
              "=IF(D9>0,85,0)"
          ],
          [
              "331 Ortaklara Borçlar",
              8500000,
              8900000,
              -400000,
              "=IF(D10<0,60,0)"
          ],
          [
              "391 Hesaplanan KDV",
              4800000,
              4800000,
              0,
              "=IF(D11<>0,95,0)"
          ]
      ],
      "metrikler": [
          [
              "Kritik anomali sayısı",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">=80\")",
              "sayi"
          ],
          [
              "Ters bakiye tutarı toplamı",
              "=SUMIF(DEMO_GIRIS!E6:E25,\">0\",DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "En yüksek anomali puanı",
              "=MAX(DEMO_GIRIS!E6:E25)",
              "sayi"
          ],
          [
              "İncelenen hesap adedi",
              "=COUNTA(DEMO_GIRIS!A6:A25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B6>0,\"DURDUR\",IF(B8>60,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "100 Kasa hesabının alacak bakiyesi vermesi inceleme gerekçesidir; ortaklar cari virmanını düzeltin.",
          "120 ve 320 ters bakiyelerini fatura veya avans hesaplarına (159/340) aktarın.",
          "Tam sürümde 100+ otomatik mizan denetim kuralı, KDV-gelir mutabakatı ve VUK ceza risk simülatörü açılır."
      ],
      "ui": [
          "Mizan Denetim Radarı",
          "Ters bakiye anomali görünümü",
          "Vergi inceleme riski"
      ]
  },
  'oee-durus-kok-neden-komutasi': {
      "karar": "Üretim hatlarında OEE (Genel Ekipman Verimliliği), Kullanılabilirlik, Performans ve Kalite kayıplarını kök-nedenle çözer.",
      "girisBasliklari": [
          "Hat / Vardiya",
          "Kullanılabilirlik (A)",
          "Performans (P)",
          "Kalite (Q)",
          "OEE Skoru"
      ],
      "ornek": [
          [
              "Hat 1 - Sabah Vardiyası",
              0.88,
              0.92,
              0.98,
              "=B6*C6*D6"
          ],
          [
              "Hat 1 - Akşam Vardiyası",
              0.78,
              0.85,
              0.96,
              "=B7*C7*D7"
          ],
          [
              "Hat 2 - Sabah Vardiyası",
              0.92,
              0.94,
              0.99,
              "=B8*C8*D8"
          ],
          [
              "Hat 2 - Akşam Vardiyası",
              0.82,
              0.88,
              0.97,
              "=B9*C9*D9"
          ],
          [
              "Hat 3 - Montaj Hattı",
              0.72,
              0.8,
              0.95,
              "=B10*C10*D10"
          ],
          [
              "Hat 4 - Paketleme Hattı",
              0.85,
              0.9,
              0.98,
              "=B11*C11*D11"
          ]
      ],
      "metrikler": [
          [
              "Tesis ortalama OEE",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "Dünya sınıfı OEE (>%85) hattı",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">0.85\")",
              "sayi"
          ],
          [
              "Kritik verimsiz hat (<%70)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<0.70\")",
              "sayi"
          ],
          [
              "En düşük OEE skoru",
              "=MIN(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B8>0,\"DURDUR\",IF(B6<0.80,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "OEE skoru %70 altına inen hatlarda SMED (kalıp değişim hızı) ve küçük duruş analizlerini başlatın.",
          "Vardiyalar arasındaki %10'u aşan verimlilik farkı için standart iş talimatlarını güncelleyin.",
          "Tam sürümde 6 büyük kayıp dökümü, Pareto duruş analitiği ve hat bazlı kayıp ciro simülatörü çalışır."
      ],
      "ui": [
          "OEE Verimlilik Radarı",
          "Hat bazlı APQ dengesi",
          "Üretim kapasite kayıp riski"
      ]
  },
  'on-uc-haftalik-nakit-odeme-onceligi': {
      "karar": "Kritik ödemeleri (çek, vergi, SGK, personel, kredi) haftalık serbest nakitle eşleştirerek temerrüt riskini önler.",
      "girisBasliklari": [
          "Hafta",
          "Beklenen Tahsilat (₺)",
          "Kritik Yasal Ödemeler (₺)",
          "Tedarikçi Ödemeleri (₺)",
          "Net Serbest Nakit (₺)"
      ],
      "ornek": [
          [
              "Hafta 1",
              850000,
              420000,
              380000,
              "=B6-C6-D6"
          ],
          [
              "Hafta 2",
              620000,
              580000,
              240000,
              "=B7-C7-D7"
          ],
          [
              "Hafta 3",
              940000,
              310000,
              450000,
              "=B8-C8-D8"
          ],
          [
              "Hafta 4",
              510000,
              680000,
              190000,
              "=B9-C9-D9"
          ],
          [
              "Hafta 5",
              1100000,
              450000,
              520000,
              "=B10-C10-D10"
          ],
          [
              "Hafta 6",
              750000,
              390000,
              410000,
              "=B11-C11-D11"
          ]
      ],
      "metrikler": [
          [
              "Nakit açığı veren hafta sayısı",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<0\")",
              "sayi"
          ],
          [
              "En derin haftalık açık",
              "=MIN(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Dönem toplam serbest nakit",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Yasal ödeme karşılama oranı",
              "=SUM(DEMO_GIRIS!B6:B25)/MAX(1,SUM(DEMO_GIRIS!C6:C25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B6>1,\"DURDUR\",IF(B7<0,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Nakit açığı veren 4. haftadaki tedarikçi vadelerini 5. haftaya öteleyin veya iskonto kredisi kullanın.",
          "Çek ve vergi gibi yasal öncelikli ödemeleri tahsilat garantili müşterilerle bloke edin.",
          "Tam sürümde 13 haftalık dinamik kurgu, senaryo stres testi ve otomatik ödeme takvimi raporu çalışır."
      ],
      "ui": [
          "Ödeme Öncelik Radarı",
          "Haftalık serbest nakit",
          "Yasal temerrüt riski"
      ]
  },
  'ortulu-sermaye-emsal-faiz-tf-paketi': {
      "karar": "Ortak ve ilişkili kuruluşlardan kullanılan borçlarda 3 katı eşiğini, örtülü sermaye faizini ve transfer fiyatlandırmasını denetler.",
      "girisBasliklari": [
          "İlişkili Kurum / Dönem",
          "Kullanılan Borç (₺)",
          "Dönem Başı Öz Sermaye (₺)",
          "3 Katı Eşiği (₺)",
          "Örtülü Sermaye Farkı (₺)"
      ],
      "ornek": [
          [
              "Ortak A - Cari Borç",
              18500000,
              5000000,
              "=C6*3",
              "=MAX(0,B6-D6)"
          ],
          [
              "Grup Şirketi B - Kredi",
              24000000,
              5000000,
              "=C7*3",
              "=MAX(0,B7-D7)"
          ],
          [
              "Yurt Dışı İştirak C",
              32000000,
              5000000,
              "=C8*3",
              "=MAX(0,B8-D8)"
          ],
          [
              "Ortak D - Finansman",
              14000000,
              5000000,
              "=C9*3",
              "=MAX(0,B9-D9)"
          ],
          [
              "Grup Şirketi E - Cari",
              16500000,
              5000000,
              "=C10*3",
              "=MAX(0,B10-D10)"
          ],
          [
              "Ortak F - Borç",
              12000000,
              5000000,
              "=C11*3",
              "=MAX(0,B11-D11)"
          ]
      ],
      "metrikler": [
          [
              "Örtülü sermaye sayılan borç",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Eşiği aşan işlem adedi",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">0\")",
              "sayi"
          ],
          [
              "Tahmini KKEG faiz tutarı",
              "=SUM(DEMO_GIRIS!E6:E25)*0.45",
              "para"
          ],
          [
              "Toplam ilişkili kurum borcu",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "para"
          ],
          [
              "Demo karar",
              "=IF(B7>1,\"DURDUR\",IF(B6>10000000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Öz sermayenin 3 katını aşan borçlanmalarda faiz ve kur farklarını KKEG olarak muhasebeleştirin.",
          "KKEG faizlerinin dönem sonunda kâr payı dağıtımı (stopaj %10) sayılacağını unutmayın.",
          "Tam sürümde emsal faiz oranı testi, transfer fiyatlandırması raporu ve örtülü kazanç simülatörü çalışır."
      ],
      "ui": [
          "Örtülü Sermaye Radarı",
          "Öz sermaye 3 katı görünümü",
          "KKEG ve stopaj riski"
      ]
  },
  'pazaryeri-hakedis-mutabakat-motoru': {
      "karar": "Pazaryeri raporlarındaki siparişler, komisyonlar, kargo kesintileri ve cezaları bankaya yatan tutarla kuruşu kuruşuna eşler.",
      "girisBasliklari": [
          "Hakediş Dönemi / Sipariş",
          "Brüt Satış Tutarı (₺)",
          "Komisyon+Hizmet (₺)",
          "Kargo+Ceza+İade (₺)",
          "Net Hakediş (₺)"
      ],
      "ornek": [
          [
              "Dönem 1 - Trendyol",
              485000,
              97000,
              68000,
              "=B6-C6-D6"
          ],
          [
              "Dönem 2 - Hepsiburada",
              320000,
              57600,
              44000,
              "=B7-C7-D7"
          ],
          [
              "Dönem 3 - Amazon TR",
              240000,
              36000,
              31000,
              "=B8-C8-D8"
          ],
          [
              "Dönem 4 - Trendyol",
              540000,
              108000,
              78000,
              "=B9-C9-D9"
          ],
          [
              "Dönem 5 - Çiçeksepeti",
              185000,
              37000,
              26000,
              "=B10-C10-D10"
          ],
          [
              "Dönem 6 - N11",
              140000,
              23800,
              19500,
              "=B11-C11-D11"
          ]
      ],
      "metrikler": [
          [
              "Toplam net hakediş",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Toplam komisyon ve kesinti",
              "=SUM(DEMO_GIRIS!C6:C25)+SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Ortalama kesinti oranı",
              "=(SUM(DEMO_GIRIS!C6:C25)+SUM(DEMO_GIRIS!D6:D25))/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Brüt satış toplamı",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "para"
          ],
          [
              "Demo karar",
              "=IF(B8>0.35,\"DURDUR\",IF(B6<1000000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Kesinti oranı %35'i aşan dönemlerde pazaryeri faturasındaki ceza ve kargo desilerini tek tek kontrol edin.",
          "Banka hesabına eksik yatan tutarlar için pazaryeri destek panelinden hakediş itiraz kaydı açın.",
          "Tam sürümde sipariş bazlı otomatik CSV/Excel eşleştirme, faturası kesilmemiş sipariş avı ve iade takibi çalışır."
      ],
      "ui": [
          "Pazaryeri Mutabakatı",
          "Brüt ve net hakediş dengesi",
          "Haksız kesinti ve ceza riski"
      ]
  },
  'proje-finansmani-dscr-llcr-paketi': {
      "karar": "Büyük ölçekli yatırımlarda DSCR (Borç Servisi Karşılama) ve LLCR (Kredi Ömrü Karşılama) rasyolarını denetler.",
      "girisBasliklari": [
          "Yıl / Dönem",
          "CFADS Nakit Akışı (₺)",
          "Anapara Borç Servisi (₺)",
          "Faiz Borç Servisi (₺)",
          "Dönem DSCR Oranı"
      ],
      "ornek": [
          [
              "2026/Yıl 1",
              18500000,
              8000000,
              4200000,
              "=B6/MAX(1,C6+D6)"
          ],
          [
              "2027/Yıl 2",
              21000000,
              9000000,
              3800000,
              "=B7/MAX(1,C7+D7)"
          ],
          [
              "2028/Yıl 3",
              24500000,
              10500000,
              3200000,
              "=B8/MAX(1,C8+D8)"
          ],
          [
              "2029/Yıl 4",
              26000000,
              12000000,
              2600000,
              "=B9/MAX(1,C9+D9)"
          ],
          [
              "2030/Yıl 5",
              28500000,
              13500000,
              1900000,
              "=B10/MAX(1,C10+D10)"
          ],
          [
              "2031/Yıl 6",
              31000000,
              15000000,
              1100000,
              "=B11/MAX(1,C11+D11)"
          ]
      ],
      "metrikler": [
          [
              "En düşük DSCR oranı",
              "=MIN(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "Kritik DSCR dönemi (<1.25)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<1.25\")",
              "sayi"
          ],
          [
              "Toplam kümülatif borç servisi",
              "=SUM(DEMO_GIRIS!C6:C25)+SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Toplam CFADS serbest nakit",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "para"
          ],
          [
              "Demo karar",
              "=IF(B7>0,\"DURDUR\",IF(B6<1.30,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "DSCR 1.25 eşiğinin altına inen dönemlerde borç servisi rezerv hesabı (DSRA) tutarını artırın.",
          "Kredi sendikasyonuna sunulacak nakit akışı projeksiyonunda duyarlılık stres testlerini tamamlayın.",
          "Tam sürümde LLCR, PLCR, Proje IRR, Özkaynak IRR ve kademeli faiz swap simülasyonu bulunur."
      ],
      "ui": [
          "Proje Finansman Radarı",
          "DSCR borç karşılama",
          "Temerrüt ve kovenant riski"
      ]
  },
  'puantaj-vardiya-fazla-mesai-kanit-sistemi': {
      "karar": "PDKS verileri, vardiya çizelgesi ve fazla mesai kayıtlarını yasal sınır ve iş mahkemesi ispat gücüyle doğrular.",
      "girisBasliklari": [
          "Personel / Sicil",
          "Normal Çalışma (Saat)",
          "Fazla Mesai %50 (Saat)",
          "Tatil Mesaisi %100 (Saat)",
          "Toplam Mesai Ücreti (₺)"
      ],
      "ornek": [
          [
              "Ahmet Yılmaz - 1001",
              180,
              24,
              8,
              "=(24*(35000/225*1.5))+(8*(35000/225*2))"
          ],
          [
              "Mehmet Demir - 1002",
              180,
              18,
              0,
              "=(18*(32000/225*1.5))+(0*(32000/225*2))"
          ],
          [
              "Ayşe Kaya - 1003",
              180,
              32,
              16,
              "=(32*(38000/225*1.5))+(16*(38000/225*2))"
          ],
          [
              "Fatma Çelik - 1004",
              180,
              12,
              0,
              "=(12*(30000/225*1.5))+(0*(30000/225*2))"
          ],
          [
              "Ali Öztürk - 1005",
              180,
              42,
              8,
              "=(42*(36000/225*1.5))+(8*(36000/225*2))"
          ],
          [
              "Hasan Şahin - 1006",
              180,
              20,
              8,
              "=(20*(31000/225*1.5))+(8*(31000/225*2))"
          ]
      ],
      "metrikler": [
          [
              "Toplam tahakkuk eden mesai",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Yasal sınır aşımı (>40 saat)",
              "=COUNTIF(DEMO_GIRIS!C6:C25,\">40\")",
              "sayi"
          ],
          [
              "Toplam fazla mesai saati",
              "=SUM(DEMO_GIRIS!C6:C25)+SUM(DEMO_GIRIS!D6:D25)",
              "sayi"
          ],
          [
              "Ortalama kişi başı mesai",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Demo karar",
              "=IF(B7>0,\"DURDUR\",IF(B6>50000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Aylık 40 saat ve yıllık 270 saat yasal fazla mesai sınırını aşan personelde vardiya rotasyonunu düzenleyin.",
          "İmzalı puantaj föyleri ile banka ödeme açıklamalarının birebir mutabakatını sağlayın.",
          "Tam sürümde arabuluculuk risk puanı, gece çalışma sınırı denetimi ve otomatik bordro bordro entegrasyonu açılır."
      ],
      "ui": [
          "Puantaj & Mesai Radarı",
          "Mesai tahakkuk dengesi",
          "İş mahkemesi dava riski"
      ]
  },
  'recete-maliyeti-menu-muhendisligi': {
      "karar": "Restoran ve kafelerde porsiyon maliyeti, fire oranı, brüt kâr marjı ve menü mühendisliği sınıflarını (Yıldız/Yük) gösterir.",
      "girisBasliklari": [
          "Menü Kalemi / Yemek",
          "Porsiyon Hammadde (₺)",
          "Fire & İşçilik (₺)",
          "Satış Fiyatı (₺)",
          "Brüt Porsiyon Kârı (₺)"
      ],
      "ornek": [
          [
              "Dana Antrikot 250g",
              185,
              25,
              420,
              "=D6-B6-C6"
          ],
          [
              "Kuzu İncik Fırın",
              160,
              20,
              360,
              "=D7-B7-C7"
          ],
          [
              "Tavuk Şinitzel",
              48,
              12,
              195,
              "=D8-B8-C8"
          ],
          [
              "Fettuccine Alfredo",
              32,
              8,
              165,
              "=D9-B9-C9"
          ],
          [
              "Somon Izgara",
              145,
              18,
              380,
              "=D10-B10-C10"
          ],
          [
              "Tiramisu Porsiyon",
              28,
              6,
              140,
              "=D11-B11-C11"
          ]
      ],
      "metrikler": [
          [
              "Ortalama porsiyon kârı",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Düşük marjlı (<%60) kalem",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<100\")",
              "sayi"
          ],
          [
              "Toplam porsiyon kâr potansiyeli",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Ortalama menü kâr marjı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!D6:D25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B7>2,\"DURDUR\",IF(B9<0.60,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Hammadde maliyet payı %35'i aşan ürünlerde porsiyon gramajı ve tedarikçi fiyatlarını revize edin.",
          "Düşük kârlı fakat popüler ürünleri menü mühendisliğinde Yıldız sınıfına taşımak için fiyat ayarlaması yapın.",
          "Tam sürümde dinamik reçete patlatma, Boston Matrix (Yıldız, At, Soru İşareti, Köpek) ve zam simülatörü çalışır."
      ],
      "ui": [
          "Menü Maliyet Radarı",
          "Porsiyon kârlılık dengesi",
          "Hammadde erozyon riski"
      ]
  },
  'sarj-istasyonu-yatirim': {
      "karar": "Elektrikli araç şarj istasyonu yatırımlarında soket başı günlük şarj süresi, elektrik marjı ve geri dönüşü gösterir.",
      "girisBasliklari": [
          "Lokasyon / İstasyon",
          "Soket Sayısı",
          "Günlük Şarj Süresi (Saat)",
          "Elektrik Marjı (₺/kWh)",
          "Aylık Net Gelir (₺)"
      ],
      "ornek": [
          [
              "Lokasyon A - AVM Otopark",
              4,
              6.5,
              4.2,
              "=B6*C6*30*50*D6*0.65"
          ],
          [
              "Lokasyon B - Otoyol Dinlenme",
              6,
              9,
              5.1,
              "=B7*C7*30*90*D7*0.65"
          ],
          [
              "Lokasyon C - Plaza Önü",
              2,
              4.5,
              3.8,
              "=B8*C8*30*40*D8*0.65"
          ],
          [
              "Lokasyon D - Otel Otopark",
              3,
              5,
              4,
              "=B9*C9*30*45*D9*0.65"
          ],
          [
              "Lokasyon E - Akaryakıt İstasyonu",
              4,
              8,
              4.8,
              "=B10*C10*30*80*D10*0.65"
          ],
          [
              "Lokasyon F - Şehir İçi Cadde",
              2,
              5.5,
              4.1,
              "=B11*C11*30*45*D11*0.65"
          ]
      ],
      "metrikler": [
          [
              "Aylık toplam şarj geliri",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Ortalama soket doluluk saati",
              "=AVERAGE(DEMO_GIRIS!C6:C25)",
              "sayi"
          ],
          [
              "Düşük verimli lokasyon (<5 saat)",
              "=COUNTIF(DEMO_GIRIS!C6:C25,\"<5\")",
              "sayi"
          ],
          [
              "Toplam şarj soketi sayısı",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B8>1,\"DURDUR\",IF(B6<250000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Günlük 5 saatin altında kalan istasyonlarda tarife indirimi ve filo anlaşmalarıyla kullanımı teşvik edin.",
          "Trafo gücü yetersiz lokasyonlarda EPDK lisans ve kapasite artış maliyetlerini gözden geçirin.",
          "Tam sürümde DC/AC soket karması, dinamik elektrik fiyatlama algoritması ve 10 yıllık fizibilite çalışır."
      ],
      "ui": [
          "Şarj İstasyonu Radarı",
          "Soket doluluk ve marj",
          "Yatırım amortisman riski"
      ]
  },
  'sekiz-d-duzeltici-faaliyet-dosya': {
      "karar": "Otomotiv ve sanayi üretiminde müşteri şikayetlerini 8D metodolojisiyle (D1-D8), kök-neden ve maliyetle kapatır.",
      "girisBasliklari": [
          "Şikayet No / Parça",
          "Hata Adedi",
          "Kök Neden Kategorisi",
          "Düzeltici Önlem Maliyeti (₺)",
          "Kapatma Skoru (1-100)"
      ],
      "ornek": [
          [
              "8D-2026-01 - Fren Kaliperi",
              45,
              1,
              18500,
              "=85"
          ],
          [
              "8D-2026-02 - Direksiyon Mili",
              12,
              2,
              34000,
              "=65"
          ],
          [
              "8D-2026-03 - Yağ Karteri Sızıntı",
              85,
              1,
              14200,
              "=90"
          ],
          [
              "8D-2026-04 - Süspansiyon Takozu",
              120,
              3,
              22000,
              "=55"
          ],
          [
              "8D-2026-05 - Egzoz Bağlantı Braketi",
              38,
              2,
              16000,
              "=75"
          ],
          [
              "8D-2026-06 - Yakıt Rayı Contası",
              18,
              1,
              42000,
              "=80"
          ]
      ],
      "metrikler": [
          [
              "Açık kritik dosya (Skor <70)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<70\")",
              "sayi"
          ],
          [
              "Toplam hata maliyeti",
              "=SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Ortalama 8D kapatma kalitesi",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "sayi"
          ],
          [
              "Toplam hatalı parça sayısı",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B6>1,\"DURDUR\",IF(B8<75,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Kapatma skoru 70'in altında olan dosyalarda Ishikawa (Kılçık) ve 5 Neden analizini yeniden yapın.",
          "Kalıcı önlem doğrulanmadan müşteriye onaylı D8 kapanış raporu göndermeyin.",
          "Tam sürümde IATF 16949 uyumlu 8D formatı, FMEA entegrasyonu ve kalite maliyeti raporu açılır."
      ],
      "ui": [
          "8D Kalite Radarı",
          "Kök neden ve kapatma dengesi",
          "Müşteri iade ve ceza riski"
      ]
  },
  'sera-kurulum-fizibilite': {
      "karar": "Modern sera yatırımlarında dekar başı kurulum, jeotermal/doğalgaz ısıtma, rekolte projeksiyonu ve kârlılığı gösterir.",
      "girisBasliklari": [
          "Yıl / Ürün Türü",
          "Üretim Alanı (Dekar)",
          "Yıllık Rekolte (Ton)",
          "Satış Geliri (₺)",
          "Net Faaliyet Kârı (₺)"
      ],
      "ornek": [
          [
              "Yıl 1 - Salkım Domates",
              50,
              1800,
              48500000,
              "=D6*0.32"
          ],
          [
              "Yıl 2 - Salkım Domates",
              50,
              1950,
              56500000,
              "=D7*0.34"
          ],
          [
              "Yıl 3 - Kokteyl Domates",
              50,
              1750,
              64200000,
              "=D8*0.36"
          ],
          [
              "Yıl 4 - Kokteyl Domates",
              50,
              1850,
              72800000,
              "=D9*0.36"
          ],
          [
              "Yıl 5 - Biber Grubu",
              50,
              1600,
              78500000,
              "=D10*0.35"
          ],
          [
              "Yıl 6 - Salkım Domates",
              50,
              2050,
              89000000,
              "=D11*0.37"
          ]
      ],
      "metrikler": [
          [
              "Kümülatif net kâr nakdi",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Ortalama dekar başı rekolte",
              "=AVERAGE(DEMO_GIRIS!C6:C25)/50",
              "sayi"
          ],
          [
              "Toplam brüt hasılat",
              "=SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Ortalama faaliyet marjı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!D6:D25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B6<35000000,\"DURDUR\",IF(B9<0.30,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Isıtma giderlerinin toplam işletme maliyetinin %40'ını aşmaması için jeotermal kuyu debisini teyit edin.",
          "Ziraat Bankası sübvansiyonlu yatırım kredisi ve kırsal kalkınma (IPARD) hibe takvimini takip edin.",
          "Tam sürümde iklimlendirme otomasyonu, gübre/ilaç sarfiyatı, fide döngüsü ve 10 yıllık nakit akışı çalışır."
      ],
      "ui": [
          "Sera Fizibilite Radarı",
          "Dekar rekolte ve kâr",
          "Yatırım geri dönüş riski"
      ]
  },
  'sevkiyat-fiyatlama-navlun': {
      "karar": "Lojistik ve taşımacılıkta rota, araç tipi, mazot, otoyol, köprü ve dönüş yükü sonrası kârlı navlun fiyatını hesaplar.",
      "girisBasliklari": [
          "Sevkiyat Rotası",
          "Mesafe (Km)",
          "Mazot+Geçiş Gideri (₺)",
          "Teklif Edilen Navlun (₺)",
          "Sefer Brüt Kârı (₺)"
      ],
      "ornek": [
          [
              "İstanbul - İzmir (Tır)",
              480,
              11500,
              18500,
              "=D6-C6"
          ],
          [
              "İstanbul - Ankara (Kırkayak)",
              450,
              9800,
              15500,
              "=D7-C7"
          ],
          [
              "Bursa - Adana (Tır)",
              890,
              19500,
              31000,
              "=D8-C8"
          ],
          [
              "İzmir - Antalya (Kamyonet)",
              460,
              8400,
              13800,
              "=D9-C9"
          ],
          [
              "Kocaeli - Gaziantep (Tır)",
              1120,
              24800,
              38500,
              "=D10-C10"
          ],
          [
              "İstanbul - Samsun (Tır)",
              740,
              16800,
              26000,
              "=D11-C11"
          ]
      ],
      "metrikler": [
          [
              "Toplam sefer kârı",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Ortalama sefer kâr marjı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!D6:D25))",
              "oran"
          ],
          [
              "Kritik düşük marjlı sefer (<%25)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<5000\")",
              "sayi"
          ],
          [
              "Toplam navlun cirosu",
              "=SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Demo karar",
              "=IF(B8>1,\"DURDUR\",IF(B7<0.30,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Dönüş yükü garantisi olmayan uzun mesafe rotalarda navlun teklifine boş dönüş amortisman payı ekleyin.",
          "Otoyol ve köprü maliyetlerinin toplam sefer giderinin %30'unu aştığı güzergahlarda alternatif rotaları kıyaslayın.",
          "Tam sürümde dinamik mazot zammı yansıtma, ton-km başı taban fiyat ve sürücü prim motoru bulunur."
      ],
      "ui": [
          "Navlun Fiyatlama Radarı",
          "Sefer kârlılık dengesi",
          "Dönüş yükü zarar riski"
      ]
  },
  'sgk-prim-tesvikleri-optimizasyon-motoru': {
      "karar": "Şirket bordrosunda 5510, 6111, 7103, engelli ve genç istihdam teşviklerini eşleştirerek maksimum prim tasarrufunu bulur.",
      "girisBasliklari": [
          "Departman / Şube",
          "Personel Sayısı",
          "Brüt Bordro Tutarı (₺)",
          "5 Puan Teşviki (₺)",
          "Toplam Teşvik Tasarrufu (₺)"
      ],
      "ornek": [
          [
              "Genel Merkez Bordrosu",
              85,
              3400000,
              170000,
              "=D6+(B6*1850)"
          ],
          [
              "Üretim Tesisi Bordrosu",
              140,
              4900000,
              245000,
              "=D7+(B7*2200)"
          ],
          [
              "AR-GE Merkezi",
              35,
              2100000,
              105000,
              "=D8+(B8*3800)"
          ],
          [
              "Lojistik & Depo",
              45,
              1575000,
              78750,
              "=D9+(B9*1650)"
          ],
          [
              "Bölge Satış Ofisleri",
              28,
              1260000,
              63000,
              "=D10+(B10*1900)"
          ],
          [
              "Çağrı Merkezi",
              60,
              1950000,
              97500,
              "=D11+(B11*2400)"
          ]
      ],
      "metrikler": [
          [
              "Aylık toplam teşvik tasarrufu",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Yıllık kümülatif tasarruf",
              "=SUM(DEMO_GIRIS!E6:E25)*12",
              "para"
          ],
          [
              "Ortalama kişi başı teşvik",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "para"
          ],
          [
              "Toplam istihdam sayısı",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B6<500000,\"DURDUR\",IF(B8<2000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Ortalama sigortalı sayısına ilave istihdam edilen personelde 6111 teşvik sürelerini her ay kontrol edin.",
          "SGK borcu veya gecikme zammı nedeniyle 5 puanlık indirimi (%5) kaybetmemek için ödemeleri otomatik talimata bağlayın.",
          "Tam sürümde TC kimlik bazlı teşvik sorgu eşleştirme, İŞKUR kayıt entegrasyonu ve geriye dönük teşvik raporu çalışır."
      ],
      "ui": [
          "SGK Teşvik Radarı",
          "Bordro prim tasarrufu",
          "Kaçırılan teşvik riski"
      ]
  },
  'site-apartman-yonetim-sistemi': {
      "karar": "Toplu konut ve sitelerde aidat tahakkuku, ortak gider dağıtımı, gecikme faizi ve işletme bütçe dengesini sağlar.",
      "girisBasliklari": [
          "Blok / Daire No",
          "Aylık Aidat (₺)",
          "Ödenen Tutar (₺)",
          "Gecikme Faizi %5 (₺)",
          "Kalan Bakiye Borç (₺)"
      ],
      "ornek": [
          [
              "A Blok Daire 01",
              2400,
              2400,
              0,
              "=B6-C6+D6"
          ],
          [
              "A Blok Daire 02",
              2400,
              0,
              120,
              "=B7-C7+D7"
          ],
          [
              "A Blok Daire 03",
              2400,
              2400,
              0,
              "=B8-C8+D8"
          ],
          [
              "B Blok Daire 04",
              2800,
              1400,
              70,
              "=B9-C9+D9"
          ],
          [
              "B Blok Daire 05",
              2800,
              0,
              140,
              "=B10-C10+D10"
          ],
          [
              "B Blok Daire 06",
              2800,
              2800,
              0,
              "=B11-C11+D11"
          ]
      ],
      "metrikler": [
          [
              "Toplam tahsil edilmeyen bakiye",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Borçlu bağımsız bölüm sayısı",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\">0\")",
              "sayi"
          ],
          [
              "Toplam tahsilat tutarı",
              "=SUM(DEMO_GIRIS!C6:C25)",
              "para"
          ],
          [
              "Site tahsilat başarı oranı",
              "=SUM(DEMO_GIRIS!C6:C25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B7>2,\"DURDUR\",IF(B9<0.85,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Üst üste 2 ay aidat ödemeyen daireler için Kat Mülkiyeti Kanunu Madde 20 gereği icra takip ihtarı çekin.",
          "Asansör revizyonu ve dış cephe gibi demirbaş harcamalarını işletme bütçesinden değil yatırım avansından karşılayın.",
          "Tam sürümde arsa payı bazlı gider dağıtımı, kıdem tazminatı fonu karşılığı ve denetim kurulu raporu açılır."
      ],
      "ui": [
          "Site Yönetim Radarı",
          "Aidat tahsilat dengesi",
          "Bütçe açığı ve icra riski"
      ]
  },
  'spc-proses-yetenek-analizi': {
      "karar": "Seri imalatta ölçüm verilerini, kontrol sınırlarını (UCL/LCL) ve Cp/Cpk yeterlilik katsayılarını denetler.",
      "girisBasliklari": [
          "Numune / Parça",
          "Ölçülen Değer (mm)",
          "Nominal Değer",
          "Üst Tolerans (USL)",
          "Alt Tolerans (LSL)"
      ],
      "ornek": [
          [
              "Numune Grubu 1",
              25.04,
              25,
              25.1,
              24.9
          ],
          [
              "Numune Grubu 2",
              25.02,
              25,
              25.1,
              24.9
          ],
          [
              "Numune Grubu 3",
              25.07,
              25,
              25.1,
              24.9
          ],
          [
              "Numune Grubu 4",
              24.96,
              25,
              25.1,
              24.9
          ],
          [
              "Numune Grubu 5",
              25.03,
              25,
              25.1,
              24.9
          ],
          [
              "Numune Grubu 6",
              25.08,
              25,
              25.1,
              24.9
          ]
      ],
      "metrikler": [
          [
              "Ölçüm ortalaması",
              "=AVERAGE(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Standart sapma (Sigma)",
              "=STDEV(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Proses yeterliliği (Cp)",
              "=(25.10-24.90)/(6*MAX(0.001,STDEV(DEMO_GIRIS!B6:B25)))",
              "oran"
          ],
          [
              "Proses yetenek indisi (Cpk)",
              "=MIN((25.10-AVERAGE(DEMO_GIRIS!B6:B25))/(3*MAX(0.001,STDEV(DEMO_GIRIS!B6:B25))),(AVERAGE(DEMO_GIRIS!B6:B25)-24.90)/(3*MAX(0.001,STDEV(DEMO_GIRIS!B6:B25))))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B9<1.0,\"DURDUR\",IF(B9<1.33,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Cpk değeri 1.33 altına inen kritik operasyonlarda takım aşınması ve fikstür boşluklarını sıfırlayın.",
          "Kontrol sınırları dışına çıkan ölçümlerde partiyi karantinaya alarak %100 ölçüm uygulayın.",
          "Tam sürümde X-bar R kartları, histogram normal dağılım eğrisi, Ga-R&R ölçüm yeterliliği çalışır."
      ],
      "ui": [
          "SPC Proses Radarı",
          "Cpk yetenek görünümü",
          "Hatalı parça fire riski"
      ]
  },
  'stok-optimizasyon-abc-olu-stok-nakit': {
      "karar": "Stoktaki ürünleri ABC/XYZ sınıflarına ayırarak ölü stokta kilitli kalan nakdi ve sipariş emniyet seviyelerini bulur.",
      "girisBasliklari": [
          "Stok Kodu / Ürün",
          "Yıllık Satış Tutarı (₺)",
          "Eldeki Stok Değeri (₺)",
          "Son Hareket Günü",
          "Ölü Stok Tutarı (₺)"
      ],
      "ornek": [
          [
              "SKU-1001 - Hidrolik Valf",
              850000,
              120000,
              12,
              "=IF(D6>90,C6,0)"
          ],
          [
              "SKU-1002 - Bakır Boru",
              620000,
              95000,
              24,
              "=IF(D7>90,C7,0)"
          ],
          [
              "SKU-1003 - Özel Flanş",
              45000,
              185000,
              140,
              "=IF(D8>90,C8,0)"
          ],
          [
              "SKU-1004 - Rulman Seri B",
              1200000,
              160000,
              18,
              "=IF(D9>90,C9,0)"
          ],
          [
              "SKU-1005 - Eski Model Conta",
              15000,
              78000,
              210,
              "=IF(D10>90,C10,0)"
          ],
          [
              "SKU-1006 - PLC Modülü",
              420000,
              65000,
              45,
              "=IF(D11>90,C11,0)"
          ]
      ],
      "metrikler": [
          [
              "Toplam ölü stokta bağlı nakit",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Toplam stok değeri",
              "=SUM(DEMO_GIRIS!C6:C25)",
              "para"
          ],
          [
              "Ölü stok oranı",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!C6:C25))",
              "oran"
          ],
          [
              "90+ gün hareketsiz kalem",
              "=COUNTIF(DEMO_GIRIS!D6:D25,\">90\")",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B8>0.25,\"DURDUR\",IF(B6>150000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "90 günden uzun süredir satılmayan C grubu ürünlerde kampanya veya hurda satışı ile nakde dönün.",
          "Ciroya katkısı %80 olan A grubu ürünlerde stok tükenmesini önleyecek otomatik sipariş eşiği kurun.",
          "Tam sürümde Wilson EOQ (Ekonomik Sipariş Miktarı), stok devir hızı ve depo taşıma maliyeti simülatörü çalışır."
      ],
      "ui": [
          "Stok ABC Radarı",
          "Ölü stok nakit görünümü",
          "Bağlı sermaye riski"
      ]
  },
  'sut-surusu-yonetim': {
      "karar": "Süt hayvancılığı işletmelerinde sağmal adet, günlük süt verimi, rasyon gideri ve süt/yem paritesini takip eder.",
      "girisBasliklari": [
          "Grup / Padok",
          "Sağmal Hayvan",
          "Günlük Süt (Litre)",
          "Günlük Yem Maliyeti (₺)",
          "Süt / Yem Paritesi"
      ],
      "ornek": [
          [
              "Padok A - Yüksek Verim",
              45,
              1530,
              16200,
              "=(1530*15.5)/MAX(1,D6)"
          ],
          [
              "Padok B - Orta Verim",
              55,
              1485,
              17050,
              "=(1485*15.5)/MAX(1,D7)"
          ],
          [
              "Padok C - İlk Doğum",
              35,
              875,
              11550,
              "=(875*15.5)/MAX(1,D8)"
          ],
          [
              "Padok D - Kuru Dönem",
              25,
              0,
              4750,
              "=0"
          ],
          [
              "Padok E - Tedavi/Revir",
              12,
              180,
              3360,
              "=(180*15.5)/MAX(1,D10)"
          ],
          [
              "Padok F - Yeni Doğan",
              28,
              700,
              8960,
              "=(700*15.5)/MAX(1,D11)"
          ]
      ],
      "metrikler": [
          [
              "Ortalama sürü süt/yem paritesi",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "sayi"
          ],
          [
              "Toplam günlük süt üretimi (L)",
              "=SUM(DEMO_GIRIS!C6:C25)",
              "sayi"
          ],
          [
              "Günlük net süt geliri (₺)",
              "=(SUM(DEMO_GIRIS!C6:C25)*15.5)-SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Kritik düşük parite padoku (<1.3)",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<1.3\")",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B9>2,\"DURDUR\",IF(B6<1.4,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Süt/yem paritesi 1.30 altına inen padoklarda rasyon protein-enerji dengesini revize edin.",
          "Kuru dönem beslemesi ve laktasyon eğrisi sapmalarını bireysel süt ölçüm kayıtlarıyla izleyin.",
          "Tam sürümde buzağılama aralığı, tohumlama başarısı, mastitis takip ve 5 yıllık sürü büyüme simülatörü çalışır."
      ],
      "ui": [
          "Süt Sürüsü Radarı",
          "Süt/yem parite dengesi",
          "Yem maliyet erozyonu"
      ]
  },
  'tahsilat-riski-vade-komuta-paneli': {
      "karar": "Müşteri cari hesaplarında vadesi geçen alacakları, gecikme günlerini ve şüpheli alacak karşılık riskini yönetir.",
      "girisBasliklari": [
          "Cari Ünvan",
          "Toplam Açık Bakiye (₺)",
          "Vadesi Geçen Tutar (₺)",
          "Ortalama Gecikme (Gün)",
          "Karşılık Riski (₺)"
      ],
      "ornek": [
          [
              "Müşteri Alfa Ltd.",
              850000,
              420000,
              45,
              "=IF(D6>30,C6*0.3,0)"
          ],
          [
              "Müşteri Beta A.Ş.",
              1450000,
              1100000,
              75,
              "=IF(D7>60,C7*0.5,0)"
          ],
          [
              "Müşteri Gama Ltd.",
              620000,
              0,
              0,
              "=0"
          ],
          [
              "Müşteri Delta İnşaat",
              2100000,
              1850000,
              120,
              "=IF(D9>90,C9*1.0,0)"
          ],
          [
              "Müşteri Epsilon Sanayi",
              940000,
              310000,
              20,
              "=0"
          ],
          [
              "Müşteri Zeta Otomotiv",
              780000,
              520000,
              65,
              "=IF(D11>60,C11*0.5,0)"
          ]
      ],
      "metrikler": [
          [
              "Toplam tahsilat risk karşılığı",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Vadesi geçmiş alacak toplamı",
              "=SUM(DEMO_GIRIS!C6:C25)",
              "para"
          ],
          [
              "Vadesi geçmiş alacak oranı",
              "=SUM(DEMO_GIRIS!C6:C25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "60+ gün gecikmiş müşteri",
              "=COUNTIF(DEMO_GIRIS!D6:D25,\">60\")",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B8>0.50,\"DURDUR\",IF(B6>500000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "60 günü aşan müşterilere derhal sevkiyat blokajı koyun ve teminat çeki/ipotek talep edin.",
          "Dava ve icra aşamasına gelen alacaklarda VUK 323 şüpheli ticari alacak karşılığı ayırın.",
          "Tam sürümde dinamik yaşlandırma (Aging 0-30-60-90+), kredi limiti kontrolü ve otomatik mutabakat mektubu çalışır."
      ],
      "ui": [
          "Tahsilat Riski Radarı",
          "Vade yaşlandırma dengesi",
          "Şüpheli alacak batık riski"
      ]
  },
  'tarimsal-destek-uygunluk': {
      "karar": "Tarımsal işletmelerde ÇKS kayıtları, mazot-gübre, prim, organik tarım ve IPARD destek hakedişlerini hesaplar.",
      "girisBasliklari": [
          "Ürün / Parsel",
          "Alan (Dekar)",
          "Mazot+Gübre Desteği (₺)",
          "Fark Ödemesi Primi (₺)",
          "Toplam Devlet Desteği (₺)"
      ],
      "ornek": [
          [
              "Buğday - Parsel 101",
              120,
              22200,
              18000,
              "=C6+D6"
          ],
          [
              "Mısır - Parsel 102",
              85,
              10200,
              12750,
              "=C7+D7"
          ],
          [
              "Ayçiçeği - Parsel 103",
              95,
              16340,
              28500,
              "=C8+D8"
          ],
          [
              "Pamuk - Parsel 104",
              60,
              21000,
              42000,
              "=C9+D9"
          ],
          [
              "Arpa - Parsel 105",
              140,
              25900,
              14000,
              "=C10+D10"
          ],
          [
              "Çeltik - Parsel 106",
              45,
              18000,
              15750,
              "=C11+D11"
          ]
      ],
      "metrikler": [
          [
              "Toplam tarımsal destek hakedişi",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Toplam kayıtlı arazi büyüklüğü",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Ortalama dekar başı destek",
              "=SUM(DEMO_GIRIS!E6:E25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "para"
          ],
          [
              "Yüksek primli ürün parseli",
              "=COUNTIF(DEMO_GIRIS!D6:D25,\">20000\")",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B6<100000,\"DURDUR\",IF(B8<350,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "ÇKS güncelleme ve havza bazlı destekleme son başvuru tarihlerini kaçırmayın.",
          "Sertifikalı tohum ve organik tarım faturalarını il/ilçe tarım müdürlüğü sistemine zamanında ibraz edin.",
          "Tam sürümde havza bazlı destek matrisi, TARSİM sigorta prim indirimi ve Ziraat sübvansiyon hesabı çalışır."
      ],
      "ui": [
          "Tarımsal Destek Radarı",
          "ÇKS prim ve hakediş",
          "Eksik başvuru hak kaybı"
      ]
  },
  'uretim-kari-125-kv-optimizasyonu': {
      "karar": "Sanayi sicil belgeli imalatçılarda üretim faaliyetinden elde edilen kazançlara uygulanan 1 puanlık KV indirimini optimize eder.",
      "girisBasliklari": [
          "Dönem",
          "İmalat Satış Hasılatı (₺)",
          "İmalat Maliyeti (₺)",
          "Üretim Faaliyeti Kârı (₺)",
          "1 Puan İndirimli Vergi (₺)"
      ],
      "ornek": [
          [
              "1. Geçici Vergi",
              12500000,
              8900000,
              3600000,
              "=D6*0.24"
          ],
          [
              "2. Geçici Vergi",
              16800000,
              11800000,
              5000000,
              "=D7*0.24"
          ],
          [
              "3. Geçici Vergi",
              19500000,
              13700000,
              5800000,
              "=D8*0.24"
          ],
          [
              "4. Geçici Vergi",
              24000000,
              16900000,
              7100000,
              "=D9*0.24"
          ],
          [
              "Yıllık Kurumlar",
              28500000,
              19800000,
              8700000,
              "=D10*0.24"
          ],
          [
              "Revize Projeksiyon",
              32000000,
              22100000,
              9900000,
              "=D11*0.24"
          ]
      ],
      "metrikler": [
          [
              "Toplam üretim faaliyeti kazancı",
              "=SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "1 Puan KV tasarruf tutarı",
              "=SUM(DEMO_GIRIS!D6:D25)*0.01",
              "para"
          ],
          [
              "Ödenecek indirimli kurumlar vergisi",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Ortalama imalat kâr marjı",
              "=SUM(DEMO_GIRIS!D6:D25)/MAX(1,SUM(DEMO_GIRIS!B6:B25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B7<200000,\"DURDUR\",IF(B9<0.25,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Üretim kazancı ile ticari alım-satım ve finansman kazançlarını muhasebede ayrı alt hesaplarda izleyin.",
          "Sanayi Sicil Belgesi vize tarihini ve kapasite raporu geçerliliğini kontrol edin.",
          "Tam sürümde ihracat 5 puan indirimi, Ar-Ge 5746 indirimi ve müşterek genel gider dağıtım anahtarı çalışır."
      ],
      "ui": [
          "Üretim KV İndirim Radarı",
          "İmalat kazancı görünümü",
          "Vergi kalkanı tasarrufu"
      ]
  },
  'yem-rasyonu-maliyet': {
      "karar": "Büyükbaş ve küçükbaş işletmelerinde hammadde fiyat dalgalanmalarına karşı en düşük maliyetli dengeli rasyonu çözer.",
      "girisBasliklari": [
          "Hammadde / Yem Türü",
          "Rasyon Miktarı (kg)",
          "Birim Fiyat (₺/kg)",
          "Ham Protein Oranı",
          "Rasyon Maliyet Payı (₺)"
      ],
      "ornek": [
          [
              "Mısır Silajı",
              22,
              2.8,
              0.08,
              "=B6*C6"
          ],
          [
              "Yonca Otu Kuru",
              4.5,
              7.5,
              0.18,
              "=B7*C7"
          ],
          [
              "Soya Küspesi %46",
              2.2,
              18.5,
              0.46,
              "=B8*C8"
          ],
          [
              "Arpa Ezme",
              5,
              9.2,
              0.11,
              "=B9*C9"
          ],
          [
              "Kepek Buğday",
              2,
              6.4,
              0.15,
              "=B10*C10"
          ],
          [
              "Mineral & Premiks",
              0.3,
              45,
              0,
              "=B11*C11"
          ]
      ],
      "metrikler": [
          [
              "Günlük hayvan başı rasyon maliyeti",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Toplam kuru madde tüketimi (kg)",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Ortalama rasyon ham proteini",
              "=SUMPRODUCT(DEMO_GIRIS!B6:B25,DEMO_GIRIS!D6:D25)/SUM(DEMO_GIRIS!B6:B25)",
              "oran"
          ],
          [
              "Konsantre yem maliyet oranı",
              "=(E8+E9+E10)/MAX(1,SUM(DEMO_GIRIS!E6:E25))",
              "oran"
          ],
          [
              "Demo karar",
              "=IF(B6>220,\"DURDUR\",IF(B8<0.15,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Hayvan başı günlük rasyon maliyeti 200 TL'yi aştığında alternatif protein kaynaklarını (kanola/ayçiçeği) formüle edin.",
          "Kaba yem oranını kuru maddede %45'in altına düşürmeyerek asidoz riskini engelleyin.",
          "Tam sürümde Linear Programming en düşük maliyet optimizasyonu, süt/besi rasyon motoru ve stok takip bulunur."
      ],
      "ui": [
          "Yem Rasyon Radarı",
          "Protein ve maliyet dengesi",
          "Besleme gider baskısı"
      ]
  },
  'yeniden-degerleme-komuta-merkezi': {
      "karar": "VUK 298/Ç ve Geçici 32 kapsamında amortismana tabi iktisadi kıymetlerin değer artışını ve vergi tasarrufunu yönetir.",
      "girisBasliklari": [
          "Varlık Grubu / Sabit Kıymet",
          "Maliyet Bedeli (₺)",
          "Birikmiş Amortisman (₺)",
          "Yİ-ÜFE Endeks Artışı",
          "Net Değer Artış Fonu (₺)"
      ],
      "ornek": [
          [
              "Fabrika Binası",
              45000000,
              9000000,
              1.45,
              "=(B6-C6)*(D6-1)"
          ],
          [
              "Üretim Hattı Makineleri",
              28000000,
              14000000,
              1.45,
              "=(B7-C7)*(D7-1)"
          ],
          [
              "Nakil Vasıtaları",
              8500000,
              4250000,
              1.45,
              "=(B8-C8)*(D8-1)"
          ],
          [
              "Depo Tesisleri",
              16000000,
              3200000,
              1.45,
              "=(B9-C9)*(D9-1)"
          ],
          [
              "Demirbaş ve Laboratuvar",
              4200000,
              2100000,
              1.45,
              "=(B10-C10)*(D10-1)"
          ],
          [
              "Kalıp ve Aparatlar",
              6500000,
              3900000,
              1.45,
              "=(B11-C11)*(D11-1)"
          ]
      ],
      "metrikler": [
          [
              "Toplam net değer artış fonu",
              "=SUM(DEMO_GIRIS!E6:E25)",
              "para"
          ],
          [
              "Ek amortisman vergi kalkanı",
              "=SUM(DEMO_GIRIS!E6:E25)*0.25",
              "para"
          ],
          [
              "Geçici 32 %2 vergi maliyeti",
              "=SUM(DEMO_GIRIS!E6:E25)*0.02",
              "para"
          ],
          [
              "Net vergi avantajı",
              "=SUM(DEMO_GIRIS!E6:E25)*0.23",
              "para"
          ],
          [
              "Demo karar",
              "=IF(B9<0,\"DURDUR\",IF(B6<5000000,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Özkaynaklarda pasif özel fona aktarılan değer artışını 3 yıl boyunca sermayeye ilave dışında çekmeyin.",
          "Yeniden değerlenmiş tutarlar üzerinden cari yılda amortisman ayırarak kurumlar vergisi matrahını düşürün.",
          "Tam sürümde Geçici 32 %2 vergi beyannamesi, enflasyon düzeltmesi (VUK Geç.33) geçiş matrisi ve ATİK takip çalışır."
      ],
      "ui": [
          "Yeniden Değerleme Radarı",
          "ATİK fon artış görünümü",
          "Amortisman vergi kalkanı"
      ]
  },
  'yurt-disi-yapilanma-vergi-simulatoru': {
      "karar": "Yurt dışı şirket ve iştiraklerde (İngiltere, Hollanda, Estonya, Dubai) efektif vergi yükünü ve temettü transferini simüle eder.",
      "girisBasliklari": [
          "Yargı Bölgesi / Ülke",
          "Yıllık Ciro (Döviz)",
          "Yerel Kurumlar Vergisi (₺)",
          "ÇVÖA Stopaj Vergisi (₺)",
          "Efektif Konsolide Vergi"
      ],
      "ornek": [
          [
              "İngiltere (UK Limited)",
              850000,
              5312500,
              0,
              "=(C6+D6)/(B6*36.5)"
          ],
          [
              "Estonya (OÜ - Dağıtılmayan)",
              620000,
              0,
              0,
              "=0"
          ],
          [
              "Dubai (BAE Freezone %0/9)",
              1400000,
              3066000,
              0,
              "=(C8+D8)/(B8*36.5)"
          ],
          [
              "Hollanda (BV Holding)",
              2100000,
              19162500,
              1825000,
              "=(C9+D9)/(B9*36.5)"
          ],
          [
              "ABD (Delaware C-Corp)",
              950000,
              7281750,
              1733750,
              "=(C10+D10)/(B10*36.5)"
          ],
          [
              "Almanya (GmbH)",
              1150000,
              12592500,
              1050000,
              "=(C11+D11)/(B11*36.5)"
          ]
      ],
      "metrikler": [
          [
              "Toplam konsolide vergi yükü (₺)",
              "=SUM(DEMO_GIRIS!C6:C25)+SUM(DEMO_GIRIS!D6:D25)",
              "para"
          ],
          [
              "Ortalama efektif vergi oranı",
              "=AVERAGE(DEMO_GIRIS!E6:E25)",
              "oran"
          ],
          [
              "Vergiden muaf / avantajlı bölge",
              "=COUNTIF(DEMO_GIRIS!E6:E25,\"<0.15\")",
              "sayi"
          ],
          [
              "Toplam yurt dışı ciro (Döviz)",
              "=SUM(DEMO_GIRIS!B6:B25)",
              "sayi"
          ],
          [
              "Demo karar",
              "=IF(B7>0.25,\"DURDUR\",IF(B8<2,\"İNCELE\",\"UYGUN\"))",
              "metin"
          ]
      ],
      "aksiyonlar": [
          "Kontrol Edilen Yabancı Kurum (KEYK - KVK 7) şartlarının doğmaması için aktif ticari faaliyet ispatlarını oluşturun.",
          "Çifte Vergilendirmeyi Önleme Anlaşması (ÇVÖA) mukimlik belgesini her mali yıl başında temin edin.",
          "Tam sürümde Transfer Fiyatlandırması emsal faiz, CFC kuralları, PE işyeri riski ve temettü istisna motoru bulunur."
      ],
      "ui": [
          "Yurt Dışı Vergi Radarı",
          "Efektif vergi yükü görünümü",
          "KEYK ve ÇVÖA riski"
      ]
  }
});

function qualifyLegacyInputFormula(formula) {
  if (typeof formula !== 'string' || !formula.startsWith('=')) return formula;
  if (formula.includes('DEMO_GIRIS!')) return formula;
  return formula.replace(
    /(^|[^A-Z0-9_!])((?:\$?[A-Z]{1,3}\$?\d+)(?::(?:\$?[A-Z]{1,3}\$?\d+))?)/g,
    (_match, prefix, ref) => `${prefix}DEMO_GIRIS!${ref}`,
  );
}

function normalizeLegacySpec(spec) {
  if (!spec || Array.isArray(spec.metrikler)) return spec;
  if (!Array.isArray(spec.ozet) || spec.ozet.length < 3) return spec;

  const numeric = spec.ozet.slice(0, -1).map(([label, formula, type]) => [
    label,
    qualifyLegacyInputFormula(formula),
    type,
  ]);
  const decision = spec.ozet[spec.ozet.length - 1];

  while (numeric.length < 4) {
    if (numeric.length === 2) {
      numeric.push(['Dolu kayıt', '=COUNTA(DEMO_GIRIS!A6:A25)', 'sayi']);
    } else {
      numeric.push(['Veri doluluk oranı', '=COUNTA(DEMO_GIRIS!A6:E25)/(20*5)', 'oran']);
    }
  }

  return {
    ...spec,
    metrikler: [
      ...numeric.slice(0, 4),
      [decision[0], qualifyLegacyInputFormula(decision[1]), decision[2]],
    ],
  };
}

function getProofDemoSpec(slug) {
  return normalizeLegacySpec(SPECS[slug] ?? null);
}

module.exports = { SPECS, getProofDemoSpec, normalizeLegacySpec };
