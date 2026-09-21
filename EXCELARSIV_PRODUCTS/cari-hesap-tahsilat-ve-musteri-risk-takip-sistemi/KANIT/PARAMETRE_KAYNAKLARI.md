# PARAMETRE KAYNAKLARI — CARİ HESAP, TAHSİLAT VE MÜŞTERİ RİSK TAKİP SİSTEMİ

Dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz. Aşağıda AYARLAR sayfasındaki her
parametrenin kaynağı listelenir.

## Mevzuat kaynakları

| Parametre | Kaynak |
|---|---|
| Cari hesap ve fatura kayıt düzeni | VUK 213 sayılı Kanun md.182–198 (defter tutma ve kayıt düzeni) |
| Defter tutma yükümlülüğü | TTK 6102 sayılı Kanun md.64–88 |
| Teminat ve rehin uygulamaları | TBK 6098 sayılı Kanun md.121–141 (teminat, kefalet) |
| İcra takibi ve tahsilat süreci | İİK 2004 sayılı Kanun (takip yolları) |

## Kullanıcı tarafından girilen parametreler

| Parametre | Birim | Kaynak |
|---|---|---|
| Firma Unvanı | Metin | Kullanıcı |
| Rapor Tarihi | Tarih | Kullanıcı |
| Rapor Hazırlayan | Metin | Kullanıcı |
| Analiz Yılı / Analiz Ayı 1–4 | Yıl/Ay | Kullanıcı |

## Varsayım etiketli parametreler (kaynak = Varsayım)

Bu parametreler ürünün kurulumunda makul işletme varsayımlarıyla doldurulmuştur; kullanıcı kendi
sektör/portföyüne göre AYARLAR'dan değiştirmelidir.

| Parametre | Varsayılan değer | Açıklama |
|---|---|---|
| Yaslama Dilim 1 / 2 / 3 | 15 / 30 / 60 gün | Yaşlandırma dilim sınırları |
| Eşik Devam / İzle / Limit Azalt / Peşin / Kritik | 24 / 49 / 69 / 84 / 85 | Risk skoru karar eşikleri |
| Eşik Yüksek | 70 | Riskli seviye eşiği |
| Gecikme / Vade / Limit / Teminat / İtiraz Maks Puan | 25 / 20 / 20 / 20 / 10 | Risk skoru bileşen ağırlıkları |
| Statü Askıda / İzleme Puan | 5 / 3 | Statü bileşen puanları |
| Güven Yüksek / Orta | 90 / 70 | Güven seviyesi sınırları |
| Duyarlılık Oranı | %10 | Senaryo değişim oranı |
| Zorunlu Alan Sayısı | 17 | Müşteri kartı tamlık kontrolü |
| Zorunlu Teminat Alan Sayısı | 10 | Teminat kaydı tamlık kontrolü |
| Tolerans Eşiği Çarpanı | 1,25 | Limit sapma toleransı |
| Yenileme Uyarı Günü | 30 | Teminat geçerlilik uyarısı |
| Sağlık Eşik Yüksek / Orta | 80 / 50 | Portföy sağlık seviyeleri |
| Veri Kalitesi Eksik Puan | 50 | Eksik kayıt ceza puanı |
| Finansman Maliyeti Oranı | %30 | Yıllık fon maliyeti |
| Aksiyon Günü Orta / Uzun | 3 / 7 | Aksiyon son tarih günleri |
| Karar Günü | 3 | KARAR_MOTORU son tarih günü |
| Vade Gecikme Eşiği | %25 | Vadesi geçmiş kırılım eşiği |
| Kalite Skor Eşiği | 90 | Yüksek güven veri kalitesi eşiği |
| P90 Oranı | 0,90 | Yüzdelik hesabı oranı |
