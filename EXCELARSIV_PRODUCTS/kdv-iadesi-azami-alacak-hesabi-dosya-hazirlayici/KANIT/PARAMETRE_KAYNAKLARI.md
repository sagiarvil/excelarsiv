# PARAMETRE KAYNAKLARI — KDV İADESİ AZAMİ ALACAK HESABI & DOSYA HAZIRLAYICI

AYARLAR sayfasındaki her parametrenin kaynağı ve yürürlük tarihi burada özetlenir. Dosyada `anahtar | deger | birim | alt | ust | aciklama | kaynak | yururluk_tarihi | dogrulama_tarihi` sütunlarıyla saklanır (D10).

| Parametre | Varsayılan | Kaynak |
|---|---|---|
| Firma Unvanı | Örnek Firma A.Ş. | Kullanıcı |
| Rapor Tarihi | 08.08.2026 | Kullanıcı |
| Raporu Hazırlayan | Finans Direktörü | Kullanıcı |
| Dosya Sürümü | 1.0.0 | Kullanıcı |
| Azami İade Tutarı Eşiği | 1.500.000 ₺ | Varsayım |
| Toplam Belge Matrahı Eşiği | 50.000.000 ₺ | Varsayım |
| İade Kullanım Oranı Risk Eşiği | %50 | Varsayım |
| İyimser Senaryo Çarpanı | 1,10 | Varsayım |
| Baz Senaryo Çarpanı | 1,00 | Varsayım |
| Kötümser Senaryo Çarpanı | 0,90 | Varsayım |
| Tornado Devreden Etkisi | %1 | Varsayım |
| Tornado İhracat Etkisi | %1 | Varsayım |
| Tahmin Çarpanı | 1,645 | Varsayım (normal dağılım %90 güven aralığı) |
| Kalite Skoru İyi/Orta Eşiği | 90 / 70 | Varsayım |
| P90 Yüzdelik Oranı | 0,90 | Varsayım |
| Aksiyon 1-3 Açıklamaları | Metin | Varsayım |
| Panel Satırı 1-12 | 1-12 | Kullanıcı |

**Not:** Bu dosyada iade eşikleri ve senaryo çarpanları kullanıcı tarafından değiştirilebilir varsayımlardır; KDV oranları (1%, 10%, 20%) güncel Türk KDV mevzuatındaki indirimli/normal oranlardır ve liste kullanıcı tarafından genişletilebilir. Azami iade tutarı vergi dairesinin inceleme sonucuna bağlıdır; dosya karar destek aracıdır, KDV beyannamesi veya mali müşavir onayı yerine geçmez.

## Mevzuat dayanağı

- Katma Değer Vergisi Kanunu (3065 s.): indirilecek KDV (Md. 29), hesaplanan KDV (Md. 28/30), ihracat istisnası ve iade hakkı (Md. 32, KDV Genel Uygulama Tebliği).
- Vergi Usul Kanunu (213 s.): belge düzeni ve defter tutma yükümlülükleri.
- KDV Genel Uygulama Tebliği: iade hakkı doğuran işlemler, iade sınırı ve belge koşulları; güncel tebliğ tarihi kullanıcı tarafından doğrulanmalıdır.

**Veri gizliliği:** Dosya tamamen çevrimdışıdır; verileriniz cihazınızdan çıkmaz (G20).
