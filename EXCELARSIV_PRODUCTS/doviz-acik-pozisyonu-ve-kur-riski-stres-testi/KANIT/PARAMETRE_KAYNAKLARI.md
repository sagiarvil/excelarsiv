# PARAMETRE KAYNAKLARI — DÖVİZ AÇIK POZİSYONU VE KUR RİSKİ STRES TESTİ

AYARLAR sayfasındaki her parametrenin kaynağı ve yürürlük tarihi burada özetlenir. Dosyada `anahtar | deger | birim | alt | ust | aciklama | kaynak | yururluk_tarihi | dogrulama_tarihi` sütunlarıyla saklanır (D10).

| Parametre | Varsayılan | Kaynak |
|---|---|---|
| Firma Unvanı | Örnek Firma A.Ş. | Kullanıcı |
| Rapor Tarihi | 08.08.2026 | Kullanıcı |
| Raporu Hazırlayan | Finans Direktörü | Kullanıcı |
| Dosya Sürümü | 1.0.0 | Kullanıcı |
| Açık Pozisyon TL Eşiği | 2.000.000 ₺ | Varsayım |
| Kötümser Senaryo Kur Şoku | %20 | Varsayım |
| Stres Etki Eşiği | 400.000 ₺ | Varsayım |
| Kur Sapması Risk Eşiği | 0,05 | Varsayım |
| İyimser Senaryo Kur Şoku | %5 | Varsayım |
| Baz Senaryo Kur Şoku | %0 | Varsayım |
| Kötümser Senaryo Kur Şoku | %10 | Varsayım |
| Kritik Senaryo Kur Şoku | %30 | Varsayım |
| Tornado Kur/Pozisyon Etkisi | %1 / %1 | Varsayım |
| Tahmin Çarpanı | 1,645 | Varsayım (normal dağılım %90 güven aralığı) |
| Kalite Skoru İyi/Orta Eşiği | 90 / 70 | Varsayım |
| P90 Yüzdelik Oranı | 0,90 | Varsayım |
| İşlem Hacmi Eşiği | 10.000.000 ₺ | Varsayım |
| Aksiyon 1-3 Açıklamaları | Metin | Varsayım |
| Panel Satırı 1-12 | 1-12 | Kullanıcı |

**Not:** Bu dosyada vergi, SGK, faiz veya mevzuat parametresi sabitlenmemiştir; tüm eşikler kullanıcı tarafından değiştirilebilir varsayımlardır. Döviz işlemleri, kur şok senaryoları ve eşikler işletmeye göre güncellenmelidir.

## Mevzuat dayanağı

- Vergi Usul Kanunu (213 s.): yabancı paraların değerlemesi (VUK Md. 280) ve kur farklarının kaydı.
- Türk Ticaret Kanunu (6102 s.): düzenli ve gerçeğe uygun kayıt ilkeleri.
- TCMB: döviz kurlarının resmî kaynağı; kurlar kullanıcı tarafından işlem kuru olarak girilir.

**Veri gizliliği:** Dosya tamamen çevrimdışıdır; verileriniz cihazınızdan çıkmaz (G20).
