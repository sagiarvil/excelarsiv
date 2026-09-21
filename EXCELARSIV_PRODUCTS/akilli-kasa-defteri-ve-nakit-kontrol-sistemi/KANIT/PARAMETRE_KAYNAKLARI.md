# PARAMETRE KAYNAKLARI — AKILLI KASA DEFTERİ VE NAKİT KONTROL SİSTEMİ

**Dosya:** `AkilliKasaDefteriVeNakitKontrolSistemi.xlsx`
**Sürüm:** 1.0
**Üretim tarihi:** 10.08.2026

Bu belge, dosyadaki tüm sayısal parametrelerin kaynağını ve yürürlük/doğrulama tarihini
gösterir. Kaynağı belirlenemeyen her değer açıkça **Varsayım** olarak işaretlenir ve aşağıda
toplu listelenir. AYARLAR sayfasında bu bilgilerin tamamı `kaynak`, `yururluk_tarihi` ve
`dogrulama_tarihi` sütunlarında yer alır (D10).

## 1. Mevzuat parametreleri

| Parametre | Değer | Birim | Kaynak | Yürürlük |
|---|---|---|---|---|
| Kasa defteri tutma ve kayıt düzeni | VUK 213 s. md.185–198 | — | Vergi Usul Kanunu (VUK) | 01.01.1985 |
| Defter tutma ve envanter esasları | TTK 6102 s. md.64 | — | Türk Ticaret Kanunu (TTK) | 01.07.2012 |
| Fiziki kasa sayımı uygulaması | Sayım-dönem kontrolü | — | VUK md.185–198 uygulaması | 01.01.1985 |

Mevzuat parametreleri dönemsel olarak değişebilir; üretim tarihi itibarıyla güncel resmi
metinler kullanılmıştır. Alıcı, defter tutma yükümlülüklerinin güncelliğini mali müşaviri
üzerinden teyit etmelidir.

## 2. Sistem parametreleri

| Parametre | Değer | Birim | Kaynak |
|---|---|---|---|
| Rapor tarihi | 10.08.2026 | Tarih | Sistem (kullanıcı değiştirir) |
| Para birimi | TRY | — | Sistem |
| Hafta uzunluğu | 7 | Gün | Takvim |

## 3. Varsayım parametreleri (toplu liste)

Aşağıdaki parametrelerin resmi bir kaynağı yoktur; sektör ortalaması, işletme politika eşiği
veya kurulum sırasında kullanıcı onayı ile belirlenir. Her biri AYARLAR sayfasında sarı
hücrede olup kullanıcı tarafından güncellenebilir:

| Parametre | Değer | Birim | Gerekçe |
|---|---|---|---|
| Kasa başlangıç bakiyesi | 500.000 | ₺ | Dönem başı kasa beyanı varsayımı |
| Asgari nakit tamponu | 100.000 | ₺ | İşletme karar eşiği varsayımı |
| Kritik ödeme ufku | 30 | Gün | Nakit planlama ufku varsayımı |
| Kritik risk eşiği | 50.000 | ₺ | Kullanıcı karar eşiği varsayımı |
| Anomali katsayısı | 2,00 | Kat | Ortalama-üstü çıkış eşiği varsayımı |
| Mutabakat toleransı | 100 | ₺ | Fiziki sayım kabul toleransı varsayımı |
| Maksimum hareket tutarı | 100.000.000 | ₺ | Aşırı tutar kontrol eşiği varsayımı |
| İyimser tahsilat oranı | 0,10 | Oran | Senaryo varsayımı |
| Kötümser ek ödeme | 100.000 | ₺ | Senaryo varsayımı |
| Duyarlılık değişim oranı | 0,10 | Oran | Tornado duyarlılık varsayımı |
| Veri kalitesi yüksek/orta eşiği | 80 / 50 | Puan | Skorlama varsayımı |
| Hata cezası / uyarı cezası ağırlığı | 8 / 4 | Puan | Skorlama varsayımı |

## 4. Doğrulama yöntemi

- Mevzuat parametreleri VUK ve TTK resmi metinleri üzerinden doğrulanmıştır.
- Varsayım parametreleri LibreOffice ile yeniden hesaplanmış dosyada formül bütünlüğü
  denetlenerek doğrulanmıştır (0 hata, 7 uç durum senaryosu temiz).
- Dosya tamamen çevrimdışıdır; dış bağlantı, telemetri veya makro içermez.
