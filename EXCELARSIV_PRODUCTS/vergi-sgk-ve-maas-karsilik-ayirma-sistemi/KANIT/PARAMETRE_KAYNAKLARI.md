# PARAMETRE KAYNAKLARI — VERGİ, SGK VE MAAŞ KARŞILIK AYIRMA SİSTEMİ

AYARLAR sayfasındaki her parametrenin kaynağı ve yürürlük tarihi burada özetlenir. Dosyada `anahtar | deger | birim | alt | ust | aciklama | kaynak | yururluk_tarihi | dogrulama_tarihi` sütunlarıyla saklanır (D10).

| Parametre | Varsayılan | Kaynak |
|---|---|---|
| Firma Ünvanı | Örnek Firma A.Ş. | Kullanıcı |
| Rapor Tarihi | 08.08.2026 | Kullanıcı |
| Raporu Hazırlayan | Finans Direktörü | Kullanıcı |
| Dosya Sürümü | 1.0.0 | Kullanıcı |
| KDV Oranı (Satış) | %20 | Mevzuat: KDVK 28. m. + Cumhurbaşkanı Kararı |
| Kurumlar Vergisi Oranı | %25 | Mevzuat: KVK 32. m. (5520 s.) |
| SGK İşçi Oranı | %14 | Mevzuat: 5510 s. SGK K. 81. m. |
| SGK İşveren Oranı | %20,5 | Mevzuat: 5510 s. SGK K. 81. m. |
| İşsizlik İşçi Oranı | %1 | Mevzuat: 4447 s. İşsizlik Sigortası K. |
| İşsizlik İşveren Oranı | %2 | Mevzuat: 4447 s. İşsizlik Sigortası K. |
| Gelir Vergisi Stopaj Oranı | %15 | Mevzuat: GVK 103. m. (2026 tarifesi) |
| Damga Oranı | %0,759 | Mevzuat: DVK 528 sayılı Karar |
| Nakit Karşılık Oranı | %25 | Varsayım |
| Karşılık Üst Oran (Risk Eşiği) | %40 | Varsayım |
| Ay Başı Kasa (₺) | 500.000 | Kullanıcı |
| Nakit Eşiği (₺) | 100.000 | Varsayım |
| Zorunlu Nakit Eşiği (₺) | 150.000 | Varsayım |
| İyimser/Baz/Kötümser Senaryo Çarpanı | 0,90 / 1,00 / 1,20 | Varsayım |
| Tornado KDV/KV/SGK Etkisi | %1 / %1 / %1 | Varsayım |
| Tahmin Çarpanı | 1,645 | Varsayım (normal dağılım %90 güven aralığı) |
| Kalite Skoru İyi/Orta Eşiği | 90 / 70 | Varsayım |
| P90 Yüzdelik Oranı | 0,90 | Varsayım |
| Giriş Kolonu Sayıları (Dönem/Bordro/Nakit/Aylık) | 5 / 3 / 3 / 1 | Varsayım |
| Aksiyon 1-3 Açıklamaları | Metin | Varsayım |
| Panel Satırı 1-12 | 1-12 | Kullanıcı |

**Not:** Mevzuat kaynaklı oranlar `kaynak` sütununda kurum + kanun/karar numarası ile saklanır; yürürlük tarihi `yururluk_tarihi` kolonundadır. Oran değişikliklerinde AYARLAR tek güncelleme noktasıdır; formüllere dokunulmaz.

## Mevzuat dayanağı

- KDV: 3065 s. KDVK 28. m. + Cumhurbaşkanı Kararı (satış KDV oranı).
- Kurumlar vergisi: 5520 s. KVK 32. m.
- SGK primleri: 5510 s. SGK K. 81. m. (işçi %14, işveren %20,5).
- İşsizlik sigortası: 4447 s. İşsizlik Sigortası K. (işçi %1, işveren %2).
- Gelir vergisi stopajı: 193 s. GVK 103. m. (ücret tarifesi).
- Damga vergisi: 488 s. DVK + 528 sayılı Karar (bordro damga oranı).

**Veri gizliliği:** Dosya tamamen çevrimdışıdır; verileriniz cihazınızdan çıkmaz (G20).
