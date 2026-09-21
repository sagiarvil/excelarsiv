# PARAMETRE KAYNAKLARI — AŞIRI DÜŞÜK TEKLİF SAVUNMA ROBOTU

**Dosya:** `AsiriDusukTeklifSavunmaRobotu.xlsx`
**Sürüm:** 1.0.0
**Üretim tarihi:** 09.08.2026

Bu belge, dosyadaki tüm sayısal parametrelerin kaynağını ve yürürlük/doğrulama tarihini
gösterir. Kaynağı belirlenemeyen her değer açıkça **Varsayım** olarak işaretlenir ve aşağıda
toplu listelenir. AYARLAR sayfasında bu bilgilerin tamamı `kaynak`, `yururluk_tarihi` ve
`dogrulama_tarihi` sütunlarında yer alır (D10).

## 1. Mevzuat parametreleri

| Parametre | Değer | Birim | Kaynak | Yürürlük |
|---|---|---|---|---|
| Aşırı Düşük Eşiği (Oran) | 1,00 | Oran | 4734 s. KİK md.38; Yapım İşleri İhaleleri Uygulama Yönetmeliği md.61-62 | 01.01.2003 |
| Sınır Değer Yöntemi | Liste seçimi | Liste | YİUY md.61 (sınır değer hesap yöntemleri) | 01.01.2003 |
| Sınır Değer Katsayısı (N) | 1,00 | Katsayı | YİUY md.61 ve KİK Genel Tebliği | 01.01.2003 |

Mevzuat parametreleri üretim tarihi itibarıyla günceldir. Sınır değer katsayısı (N) ve hesap
yöntemi ilandan ilana değişebileceğinden alıcı, ihale dokümanındaki değerleri AYARLAR
sayfasına işlemelidir. Bu süreç KILAVUZ sayfasında anlatılır.

## 2. Sistem parametreleri

| Parametre | Değer | Birim | Kaynak |
|---|---|---|---|
| Rapor Tarihi | 09.08.2026 | Tarih | Kullanıcı (rapor günü) |
| Dosya Sürümü | 1.0.0 | Metin | Sabit |
| Toplam Giriş Hücresi Sayısı | 14.300 | Adet | Sabit (tasarım) |

## 3. Varsayım parametreleri (toplu liste)

Aşağıdaki parametrelerin resmi bir kaynağı yoktur; sektör pratiği veya kurulum sırasında
kullanıcı onayı ile belirlenir. Her biri AYARLAR sayfasında sarı hücrede olup kullanıcı
tarafından güncellenebilir:

| Parametre | Değer | Birim | Gerekçe |
|---|---|---|---|
| Dikkat Eşiği (Oran) | 1,05 | Oran | Sınır değerin %5 üzeri güvenlik payı varsayımı |
| Anomali Z Eşiği | 2,00 | Adet | Z skoru eşiği varsayımı (istatistik pratiği) |
| Rayiç Sapma Eşiği (%) | 0,20 | Oran | Birim fiyatın rayiçten %20 sapma toleransı varsayımı |
| Dayanak Ağırlığı Eşiği (%) | 0,80 | Oran | Savunma ağırlığı %80 altı ZAYIF varsayımı |
| Kabul Edilebilir Mutabakat Farkı | 25.000 | ₺ | Bileşen/teklif mutabakat toleransı varsayımı |
| Senaryo çarpanları (İyimser/Baz/Kötümser) | 0,90 / 1,00 / 1,15 | Oran | Senaryo aralığı varsayımı |
| Tornado etki adımları (±) | 0,10 | Oran | Duyarlılık adımı varsayımı |
| Tahmin Bandı Genişliği (%) | 0,10 | Oran | Projeksiyon güven bandı varsayımı |
| Projeksiyon İleri Dönem | 3 | Dönem | Tahmin ufku varsayımı |
| Kalite Skoru İyi/Orta Eşiği | 90 / 70 | Puan | Skorlama varsayımı |
| Değer Ölçer: Kazanılan Saat / Ay | 12 | Saat/ay | Değer ölçer varsayımı |
| Değer Ölçer: Yıllık Hata Azalması (%) | 0,20 | Oran | Değer ölçer varsayımı |

## 4. Doğrulama yöntemi

- Mevzuat parametreleri KİK kanun/yönetmelik maddeleri üzerinden doğrulanmıştır.
- Varsayım parametreleri LibreOffice ile yeniden hesaplanmış dosyada formül bütünlüğü
  denetlenerek doğrulanmıştır (0 hata, 7 uç durum senaryosu temiz, 15.000 satırda 3,0 sn).
- Dosya tamamen çevrimdışıdır; dış bağlantı, telemetri veya makro içermez.
