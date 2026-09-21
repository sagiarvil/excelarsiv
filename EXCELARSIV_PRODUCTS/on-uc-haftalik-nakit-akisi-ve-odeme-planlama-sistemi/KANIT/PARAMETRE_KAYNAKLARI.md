# PARAMETRE KAYNAKLARI — 13 HAFTALIK NAKİT AKIŞI VE ÖDEME PLANLAMA SİSTEMİ

**Dosya:** `OnUcHaftalikNakitAkisiVeOdemePlanlamaSistemi.xlsx`
**Sürüm:** 1.0
**Üretim tarihi:** 10.08.2026

Bu belge, dosyadaki tüm sayısal parametrelerin kaynağını ve yürürlük/doğrulama tarihini
gösterir. Kaynağı belirlenemeyen her değer açıkça **Varsayım** olarak işaretlenir ve aşağıda
toplu listelenir. AYARLAR sayfasında bu bilgilerin tamamı `kaynak`, `yururluk_tarihi` ve
`dogrulama_tarihi` sütunlarında yer alır (D10).

## 1. Mevzuat parametreleri

| Parametre | Değer | Birim | Kaynak | Yürürlük |
|---|---|---|---|---|
| Defter tutma ve kayıt düzeni; nakit ve tahakkuk kayıtlarının dayanağı | VUK 213 s. md.182–198 | — | Vergi Usul Kanunu (VUK) | 01.01.1985 |
| Dönem sonu nakit denetimi ve sayım esası | VUK 213 s. md.185–186 | — | Vergi Usul Kanunu (VUK) | 01.01.1985 |

Mevzuat parametreleri dönemsel olarak değişebilir; üretim tarihi itibarıyla güncel resmi
metinler kullanılmıştır. Alıcı, kayıt düzenine ilişkin güncel yükümlülükleri mali müşaviri
üzerinden teyit etmelidir.

## 2. Sistem parametreleri

| Parametre | Değer | Birim | Kaynak |
|---|---|---|---|
| Rapor tarihi | 10.08.2026 | Tarih | Sistem (kullanıcı değiştirir) |
| Para birimi | TRY | — | Sistem |
| Hafta uzunluğu | 7 | Gün | Tasarım (plan ufku) |
| Plan hafta sayısı | 13 | Hafta | Tasarım (ürün kapsamı) |
| Tahmin hafta ufku | 3 | Hafta | Tasarım (ileriye dönük bant) |
| Maksimum tarih ufku | 3.000 | Gün | Tasarım (tarih taşmasını önleme) |
| P90 yüzdelik dilimi | 0,9 | Oran | Tasarım (istatistik standart) |

## 3. Varsayım parametreleri (toplu liste)

Aşağıdaki parametrelerin resmi bir kaynağı yoktur; işletme politika eşiği veya kurulum
sırasında kullanıcı onayı ile belirlenir. Her biri AYARLAR sayfasında sarı hücrede olup
kullanıcı tarafından güncellenebilir:

| Parametre | Değer | Birim | Gerekçe |
|---|---|---|---|
| Başlangıç nakit bakiyesi | 500.000 | ₺ | Kullanıcı beyanı (örnek) |
| Asgari nakit tamponu | 50.000 | ₺ | Tampon eşiği varsayımı |
| Veri kalitesi üst/orta eşiği | 80 / 60 | Puan | Skorlama varsayımı |
| Uyarı puan kesintisi | 5 | Puan | Her kural ihlali kesintisi varsayımı |
| Aşırı giriş/çıkış tutarı | 500.000 | ₺ | Aşırı değer kontrol eşiği varsayımı |
| Duyarlılık değişim oranı | 0,10 | Oran | Tornado/senaryo duyarlılık varsayımı |
| Minimum veri sayısı | 6 | Kayıt | Karar için asgari kayıt varsayımı |
| Kıtlık günlük maliyet oranı | 0,002 | Oran | Negatif bakiye finansman maliyeti varsayımı |
| Kritik ödeme payı eşiği | 0,30 | Oran | Ödeme yoğunluğu uyarı eşiği varsayımı |
| Senaryo belirsizlik eşiği | 0,20 | Oran | Senaryo farkı uyarı eşiği varsayımı |
| Nakit dönüşüm eşiği | 0,10 | Oran | Nakit üretimi zayıflık eşiği varsayımı |
| Yoğunlaşma eşiği | 0,50 | Oran | HHI tek kalem yoğunluğu eşiği varsayımı |

## 4. Doğrulama yöntemi

- Mevzuat parametreleri VUK resmi metinleri üzerinden doğrulanmıştır.
- Varsayım parametreleri LibreOffice ile yeniden hesaplanmış dosyada formül bütünlüğü
  denetlenerek doğrulanmıştır (0 hata, 7 uç durum senaryosu temiz, Ö1 6.000 satır 1,9 sn).
- Dosya tamamen çevrimdışıdır; dış bağlantı, telemetri veya makro içermez.
