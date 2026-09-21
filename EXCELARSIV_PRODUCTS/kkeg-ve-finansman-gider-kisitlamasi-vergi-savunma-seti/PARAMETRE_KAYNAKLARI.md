# PARAMETRE KAYNAKLARI — KKEG VE FİNANSMAN GİDER KISITLAMASI VERGİ SAVUNMA SETİ

AYARLAR sayfasındaki her parametrenin kaynağı ve yürürlük tarihi burada özetlenir. Dosyada `anahtar | deger | birim | alt | ust | aciklama | kaynak | yururluk_tarihi | dogrulama_tarihi` sütunlarıyla saklanır (D10).

| Parametre | Varsayılan | Kaynak |
|---|---|---|
| Firma Unvanı | Örnek Firma A.Ş. | Kullanıcı |
| Rapor Tarihi | 08.08.2026 | Kullanıcı |
| Raporu Hazırlayan | Finans Direktörü | Kullanıcı |
| Dosya Sürümü | 1.0.0 | Kullanıcı |
| KKEG Eşiği | 500.000 ₺ | Varsayım (karar eşiği) |
| Kapsamdaki Gider Eşiği | 10.000.000 ₺ | Varsayım (karar eşiği) |
| Fark Oranı Eşiği | %30 | Varsayım (örtülü sermaye/KKEG risk izleme eşiği) |
| Finansman Gider Kısıtlama Oranı | %25 | Mevzuat (2023 sonrası kurumlar vergisi uygulaması; kullanıcı güncel tebliği doğrulamalıdır) |
| Kurumlar Vergisi Oranı | %25 | Mevzuat (kullanıcı güncel oranı doğrulamalıdır) |
| Muafiyet Eşiği | 0 ₺ | Varsayım (uygulanmayacaksa 0) |
| İyimser/Baz/Kötümser Senaryo Çarpanı | 1,10 / 1,00 / 0,90 | Varsayım |
| Tornado Yabancı Kaynak/Oran Etkisi | %1 | Varsayım |
| Tahmin Çarpanı | 1,645 | Varsayım (normal dağılım %90 güven aralığı) |
| Kalite Skoru İyi/Orta Eşiği | 90 / 70 | Varsayım |
| P90 Yüzdelik Oranı | 0,90 | Varsayım |
| Aksiyon 1-3 Açıklamaları | Metin | Varsayım |
| Panel Satırı 1-12 | 1-12 | Kullanıcı |
| Dönem Kaynak Giriş Kolonu | 6 | Kullanıcı |
| Kredi Listesi Giriş Kolonu | 9 | Kullanıcı |

**Not:** Kısıtlama oranı (%25) ve kurumlar vergisi oranı mevzuat kaynaklı olup dönemler itibarıyla değişebilir; kullanıcı güncel tebliğ, kanun ve Gelir İdaresi duyurularını takip ederek AYARLAR'daki oranları doğrulamalıdır. Eşikler ve çarpanlar kullanıcı varsayımıdır. Dosya karar destek aracıdır; mali müşavir veya vergi danışmanı görüşü yerine geçmez.

## Mevzuat dayanağı

- Kurumlar Vergisi Kanunu (5520 s.): finansman gider kısıtlaması ve örtülü sermaye (Md. 11/1-i — finansman gider kısıtlaması; Md. 12 — örtülü sermaye).
- Vergi Usul Kanunu (213 s.): defter ve belge düzeni (karar destek kapsamı).
- Güncel kısıtlama oranları: Gelir İdaresi Başkanlığı tebliğ ve duyuruları (ör. %25 oranı); kullanıcı yürürlük tarihini doğrulamalıdır.

**Veri gizliliği:** Dosya tamamen çevrimdışıdır; verileriniz cihazınızdan çıkmaz (G20).
