# PARAMETRE KAYNAKLARI — BANKA, KREDİ VE TAKSİT TAKİP SİSTEMİ

**Dosya:** `BankaKrediVeTaksitTakipSistemi.xlsx`
**Sürüm:** 1.0
**Üretim tarihi:** 10.08.2026

Bu belge, dosyadaki tüm sayısal parametrelerin kaynağını ve yürürlük/doğrulama tarihini
gösterir. Kaynağı belirlenemeyen her değer açıkça **Varsayım** olarak işaretlenir ve aşağıda
toplu listelenir. AYARLAR sayfasında bu bilgilerin tamamı `kaynak`, `yururluk_tarihi` ve
`dogrulama_tarihi` sütunlarında yer alır (D10).

## 1. Mevzuat parametreleri

| Parametre | Değer | Birim | Kaynak | Yürürlük |
|---|---|---|---|---|
| Borç ve alacakların kaydı, defter tutma ve kayıt düzeni | VUK 213 s. md.185–198 | — | Vergi Usul Kanunu (VUK) | 01.01.1985 |
| Defter tutma ve finansal tablo esasları | TTK 6102 s. md.64 | — | Türk Ticaret Kanunu (TTK) | 01.07.2012 |

Mevzuat parametreleri dönemsel olarak değişebilir; üretim tarihi itibarıyla güncel resmi
metinler kullanılmıştır. Alıcı, borç kayıt düzenine ilişkin güncel yükümlülükleri mali
müşaviri üzerinden teyit etmelidir.

## 2. Sistem parametreleri

| Parametre | Değer | Birim | Kaynak |
|---|---|---|---|
| Rapor tarihi | 10.08.2026 | Tarih | Sistem (kullanıcı değiştirir) |
| Para birimi | TRY | — | Sistem |
| Ay pencereleri (30–180 gün) | 6 pencere | Gün | Ödeme planlama ufku |

## 3. Varsayım parametreleri (toplu liste)

Aşağıdaki parametrelerin resmi bir kaynağı yoktur; işletme politika eşiği veya kurulum
sırasında kullanıcı onayı ile belirlenir. Her biri AYARLAR sayfasında sarı hücrede olup
kullanıcı tarafından güncellenebilir:

| Parametre | Değer | Birim | Gerekçe |
|---|---|---|---|
| Serbest nakit | 1.800.000 | ₺ | Kullanılabilir nakit beyanı varsayımı |
| 30 günlük yük eşiği | 500.000 | ₺ | Ödeme baskısı eşiği varsayımı |
| Kritik kur farkı eşiği | 0,05 | Oran | Döviz riski uyarı eşiği varsayımı |
| Likidite eşiği | 0,85 | Oran | Nakit/borç servisi karşılama eşiği varsayımı |
| Kaldıraç eşiği | 3,00 | Kat | Borç/kâr eşiği varsayımı |
| Veri kalitesi yüksek/orta eşiği | 80 / 50 | Puan | Skorlama varsayımı |
| Anomali katsayısı | 2,00 | Kat | Ortalama-üstü taksit eşiği varsayımı |
| Duyarlılık değişim oranı | 0,10 | Oran | Tornado duyarlılık varsayımı |
| P90 karşılaştırma çarpanı | 1,50 | Kat | Tahmin aralığı eşiği varsayımı |
| Kredi anaparası üst sınırı | 100.000.000 | ₺ | Aşırı değer kontrol eşiği varsayımı |

## 4. Doğrulama yöntemi

- Mevzuat parametreleri VUK ve TTK resmi metinleri üzerinden doğrulanmıştır.
- Varsayım parametreleri LibreOffice ile yeniden hesaplanmış dosyada formül bütünlüğü
  denetlenerek doğrulanmıştır (0 hata, 7 uç durum senaryosu temiz, Ö1 1,8 sn).
- Dosya tamamen çevrimdışıdır; dış bağlantı, telemetri veya makro içermez.
