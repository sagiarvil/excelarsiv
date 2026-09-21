# PARAMETRE KAYNAKLARI — ORTAKLAR CARİ & KASA ADAT FAİZ FATURASI HESAPLAYICI

AYARLAR sayfasındaki her parametrenin kaynağı ve yürürlük tarihi burada özetlenir. Dosyada `anahtar | deger | birim | alt | ust | aciklama | kaynak | yururluk_tarihi | dogrulama_tarihi` sütunlarıyla saklanır (D10).

| Parametre | Varsayılan | Kaynak |
|---|---|---|
| Firma Unvanı | Örnek Firma A.Ş. | Kullanıcı |
| Rapor Tarihi | 08.08.2026 | Kullanıcı |
| Raporu Hazırlayan | Finans Direktörü | Kullanıcı |
| Dosya Sürümü | 1.0.0 | Kullanıcı |
| Adat Faiz Oranı | %19 | Varsayım (dönemin emsal mevduat oranına göre kullanıcı günceller) |
| Yılın Gün Bazı | 36.500 | Varsayım (faiz = adat × oran / gün bazı) |
| Brüt Faiz Tutarı Eşiği | 250.000 ₺ | Varsayım |
| Şirket Öz Kaynağı | 10.000.000 ₺ | Kullanıcı |
| Örtülü Sermaye Oranı Eşiği | %30 | Varsayım (KKEG risk izleme eşiği) |
| İyimser/Baz/Kötümser Senaryo Çarpanı | 1,10 / 1,00 / 0,90 | Varsayım |
| Tornado Oran/Adat Etkisi | %1 | Varsayım |
| Tahmin Çarpanı | 1,645 | Varsayım (normal dağılım %90 güven aralığı) |
| Kalite Skoru İyi/Orta Eşiği | 90 / 70 | Varsayım |
| P90 Yüzdelik Oranı | 0,90 | Varsayım |
| Aksiyon 1-3 Açıklamaları | Metin | Varsayım |
| Panel Satırı 1-12 | 1-12 | Kullanıcı |

**Not:** Faiz oranı ve eşikler kullanıcı tarafından değiştirilebilir varsayımlardır. Kullanıcı, dönemin emsal mevduat/avans faiz oranını güncel kaynaklardan (TCMB, banka ilanları) doğrulamalıdır; dosya oranların mevzuata uygunluğunu taahhüt etmez. Dosya karar destek aracıdır; mali müşavir veya vergi danışmanı görüşü yerine geçmez.

## Mevzuat dayanağı

- Kurumlar Vergisi Kanunu (5520 s.): örtülü sermaye ve faiz gideri kısıtlaması (Md. 12 — örtülü sermaye kapsamında ortaklardan sağlanan borçlara ödenen faizlerin KKEG sayılması).
- Türk Ticaret Kanunu (6102 s.): ortak cari hesabı düzenine ilişkin genel çerçeve (karar destek kapsamı).
- TCMB ve banka emsal faiz oranları: adat faiz oranının güncellenmesinde referans; oran kullanıcı tarafından doğrulanmalıdır.

**Veri gizliliği:** Dosya tamamen çevrimdışıdır; verileriniz cihazınızdan çıkmaz (G20).
