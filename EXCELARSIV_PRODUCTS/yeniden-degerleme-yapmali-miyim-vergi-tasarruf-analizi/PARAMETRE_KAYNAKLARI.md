# PARAMETRE KAYNAKLARI — YENİDEN DEĞERLEME YAPMALI MIYIM? (VERGİ TASARRUF ANALİZİ)

AYARLAR sayfasındaki her parametrenin kaynağı ve yürürlük tarihi burada özetlenir. Dosyada `anahtar | deger | birim | alt | ust | aciklama | kaynak | yururluk_tarihi | dogrulama_tarihi` sütunlarıyla saklanır (D10).

| Parametre | Varsayılan | Kaynak |
|---|---|---|
| Firma Unvanı | Örnek Firma A.Ş. | Kullanıcı |
| Rapor Tarihi | 10.08.2026 | Kullanıcı |
| Raporu Hazırlayan | Mali İşler Direktörü | Kullanıcı |
| Dosya Sürümü | 1.0.0 | Kullanıcı |
| Fon Oranı | %2 | Mevzuat — VUK Geçici 31. madde fon hesabı oranı (değer artışı üzerinden %2 fon payı ayrılır) |
| Kurumlar Vergisi Oranı | %25 | Mevzuat — Kurumlar Vergisi Kanunu (KVK) 32. madde, 2026 dönemi |
| İskonto Oranı (baz) | %15 | Varsayım — kullanıcı kendi sermaye maliyetiyle günceller |
| İyimser / Baz / Kötümser / Kritik İskonto | %10 / %15 / %22 / %30 | Varsayım — senaryo bant aralığı oranları |
| NBD Eşiği | 100.000 ₺ | Varsayım — net fayda karar eşiği |
| Değer Artışı Eşiği | 50.000 ₺ | Varsayım — değer artışı karar eşiği |
| Fon Eşiği | 200.000 ₺ | Varsayım — fon vergisi karar eşiği |
| Tornado İskonto / Fon Şoku | %1 / %1 | Varsayım — duyarlılık adımı |
| Tahmin Çarpanı | 1,645 | Varsayım (normal dağılım %90 güven aralığı) |
| Kalite Skoru İyi / Orta Eşiği | 90 / 70 | Varsayım |
| Aksiyon 1–3 Açıklamaları | Metin | Varsayım |
| P90 Yüzdelik Oranı | 0,90 | Varsayım |
| Temel Giriş Kolonu Sayısı | 5 | Varsayım — veri kalite skoru temeli |
| Panel Satırı 1–12 | 1–12 | Kullanıcı |

**Not:** Fon oranı VUK Geçici 31. maddeye göre mevzuat kaynaklıdır; kurumlar vergisi oranı KVK 32. madde kaynaklıdır ve dönemsel güncellenir. Diğer eşik ve çarpanlar kullanıcı varsayımıdır; değişen değerler AYARLAR sayfasından tüm hesaplamalara yansır. Dosya karar destek aracıdır; mali müşavir görüşü yerine geçmez.

## Mevzuat dayanağı

- Vergi Usul Kanunu (213 s.) Geçici 31. madde: enflasyon düzeltmesi ile yeniden değerleme; değer artışı üzerinden %2 fon hesabı ayrılması ve ek amortisman ayırma imkânı.
- Vergi Usul Kanunu (213 s.) Md. 283/2 (yeniden değerleme öncesi hükümler) ve amortisman esasları (Md. 313–321).
- Kurumlar Vergisi Kanunu (5520 s.) Md. 32: kurumlar vergisi oranı.
- Bakanlar Kurulu/Cumhurbaşkanı Kararı ile belirlenen yeniden değerleme katsayıları (her yıl ilan edilir; kullanıcı AYARLAR'dan günceller).

**Veri gizliliği:** Dosya tamamen çevrimdışıdır; verileriniz cihazınızdan çıkmaz (G20).
