# PARAMETRE KAYNAKLARI — ÇEK SENET VADE VE KARŞILIK RİSK SİSTEMİ

Dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz. Aşağıda AYARLAR sayfasındaki her
parametrenin kaynağı listelenir.

## Mevzuat kaynakları

| Parametre | Kaynak |
|---|---|
| Çek ve senet kayıt düzeni, ciro ve ibraz süreleri | TTK 6102 sayılı Kanun md.772–778 (kambiyo senetlerinde ödeme, ibraz) |
| Karşılıksız çek ve ibraz süresi (10 gün) | Çek Kanunu 5941 sayılı Kanun md.3–5 (ibraz ve karşılık) |
| Karşılıksız çekte sorumluluk ve hukuki takip | Çek Kanunu 5941 sayılı Kanun md.5–8 |
| Mükerrer kayıt ve sahtecilik kontrolü | Bankacılık uygulamaları ve iç kontrol standartları |

## Kullanıcı tarafından girilen parametreler

| Parametre | Birim | Kaynak |
|---|---|---|
| Firma Unvanı | Metin | Kullanıcı |
| Vergi Numarası (opsiyonel) | Metin | Kullanıcı |
| Rapor Tarihi | Tarih | Kullanıcı |
| Serbest Nakit (Banka Mevduatı) | TL | Kullanıcı |
| Risk katsayıları (DÜŞÜK/ORTA/YÜKSEK/KARŞILIKSIZ/VERİ YETERSİZ) | Oran | Kullanıcı |

## Varsayım etiketli parametreler (kaynak = Varsayım)

Bu parametreler ürünün kurulumunda makul işletme varsayımlarıyla doldurulmuştur; kullanıcı kendi
sektör/portföyüne göre AYARLAR'dan değiştirmelidir.

| Parametre | Varsayılan değer | Açıklama |
|---|---|---|
| Haftalık Verilen Belge Eşiği | 750.000 TL | Tek haftada vadesi gelen verilen belge üst sınırı |
| Karşılık Oranı Eşiği | %100 | Kaynak/yükümlülük oranı altındaysa riskli kabul edilir |
| Riskli Müşteri Skor Eşiği | 60 | Tahsilat skoru bu değerin altındaki müşteriler riskli sayılır |
| Vade Geçmişi Uyarı Günü | 7 gün | Vadesi geçmiş belge operasyon uyarısı eşiği |
| Yuvarlama Toleransı | 1 TL | Çapraz toplam kontrolünde kabul edilebilir fark |
| Tek Müşteri Yoğunlaşma Eşiği | %40 | Tek müşterinin toplam alınan belge içindeki azami payı |
| Düşük/Yüksek Risk Skor Eşiği | 80 / 40 | Tahsilat skoru risk sınıfı sınırları |
| Veri Kalite Yüksek/Orta Eşiği | 80 / 50 | Veri kalite skoru güven düzeyi sınırları |
| Haftalık Yoğunluk Uyarı Katsayısı | 0,80 | Haftalık eşiğin riskli işaretlenme katı |
| Yeni Çek Kapasite Oranı Eşiği | 1,20 | Karşılık oranı üzerindeyse yeni çek kapasitesi NORMAL |
| Hafta Uzunluğu | 7 gün | Haftalık planlama takvim genişliği |
| Hata / Uyarı Ağırlığı | 5 / 10 puan | Veri kalite skoru ceza puanları |
| Maksimum Belge Tutarı | 1.000.000.000 TL | Aşırı tutar senaryosu eşiği |
| Finansman Maliyeti Oranı | %30 | Gecikme maliyetinde kullanılan yıllık fon maliyeti |

## Not

Çek karşılıksız çıkması ve ibraz süreleri banka bildirimlerine bağlıdır; sistem banka bildiriminden
sonra kayıt güncellemesini kullanıcı girişiyle alır. Hesaplar mali danışmanlık görüşünün yerine geçmez.
