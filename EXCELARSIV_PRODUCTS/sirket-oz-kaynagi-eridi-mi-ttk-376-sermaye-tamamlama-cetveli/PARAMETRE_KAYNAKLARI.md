# PARAMETRE KAYNAKLARI — ŞİRKET ÖZ KAYNAĞI ERİDİ Mİ? (TTK 376 SERMAYE TAMAMLAMA CETVELİ)

AYARLAR sayfasındaki her parametrenin kaynağı ve yürürlük tarihi burada özetlenir. Dosyada `anahtar | deger | birim | alt | ust | aciklama | kaynak | yururluk_tarihi | dogrulama_tarihi` sütunlarıyla saklanır (D10).

| Parametre | Varsayılan | Kaynak |
|---|---|---|
| Firma Unvanı | Örnek Firma A.Ş. | Kullanıcı |
| Rapor Tarihi | 08.08.2026 | Kullanıcı |
| Raporu Hazırlayan | Finans Direktörü | Kullanıcı |
| Dosya Sürümü | 1.0.0 | Kullanıcı |
| TTK 376 1/3 Sermaye Kaybı Eşiği | %33,3333 | Mevzuat (TTK 376/1-a) |
| TTK 376 2/3 Sermaye Kaybı Eşiği | %66,6667 | Mevzuat (TTK 376/1-b) |
| Tahmin Çarpanı | 1,645 | Varsayım (normal dağılım %90 güven aralığı) |
| İyimser Senaryo Çarpanı | 0,60 | Varsayım |
| Baz Senaryo Çarpanı | 1,00 | Varsayım |
| Kötümser Senaryo Çarpanı | 1,50 | Varsayım |
| Tornado Zarar Etkisi | %1 | Varsayım |
| Kalite Skoru İyi/Orta Eşiği | 90 / 70 | Varsayım |
| Aksiyon 1-3 Açıklamaları | Metin | Varsayım |
| P90 Yüzdelik Oranı | 0,90 | Varsayım |
| Azami Bilanço Tutarı | 100.000.000.000 ₺ | Varsayım (gerçekçilik üst sınırı; aşılırsa karar İNCELE kalır) |
| Panel Satırı 1-12 | 1-12 | Kullanıcı |
| Dönem Bilanço Giriş Kolonu | 6 | Kullanıcı |
| Borç Listesi Giriş Kolonu | 6 | Kullanıcı |

**Not:** 1/3 ve 2/3 eşikleri Türk Ticaret Kanunu 376. maddeden mevzuat kaynaklıdır; diğer eşikler ve çarpanlar kullanıcı varsayımıdır. Dosya karar destek aracıdır; avukat veya mali müşavir görüşü yerine geçmez.

## Mevzuat dayanağı

- Türk Ticaret Kanunu (6102 s.) Md. 376: sermayenin kaybı, borca batıklık ve şirketin durumu; 1/3 ve 2/3 kayıp eşiklerinde genel kurul toplantı ve sermaye tamamlama/azaltım/infisah yükümlülükleri.
- Türk Ticaret Kanunu (6102 s.) Md. 376/3: borca batıklık durumunda yönetim kurulunun mahkemeye bildirim yükümlülüğü.
- Vergi Usul Kanunu (213 s.): aktif/bilanço değerleme esasları (kayıt kalitesi kapsamı).

**Veri gizliliği:** Dosya tamamen çevrimdışıdır; verileriniz cihazınızdan çıkmaz (G20).
