# PARAMETRE KAYNAKLARI — PAZARYERİ NET KÂR & EKSİK HAKEDİŞ YAKALAYICI

AYARLAR sayfasındaki her parametrenin kaynağı ve yürürlük tarihi burada özetlenir. Dosyada `anahtar | deger | birim | alt | ust | aciklama | kaynak | yururluk_tarihi | dogrulama_tarihi` sütunlarıyla saklanır (D10).

| Parametre | Varsayılan | Kaynak |
|---|---|---|
| Firma Unvanı | Örnek Firma A.Ş. | Kullanıcı |
| Rapor Tarihi | 10.08.2026 | Kullanıcı |
| Raporu Hazırlayan | Pazaryeri Operasyon Müdürü | Kullanıcı |
| Dosya Sürümü | 1.0.0 | Kullanıcı |
| Eksik Hakediş Fark Oranı Eşiği | %1 | Varsayım — üzerindeki fark oranı eksik hakediş sayılır |
| Eksik Hakediş Asgari Tutar Eşiği | 50 ₺ | Varsayım — fark tutarının ihmal edilmeyeceği asgari TL |
| Net Kâr Marjı Hedefi | %10 | Varsayım — GÜÇLÜ kararı için genel marj hedefi |
| Net Kâr Marjı Kritik Eşiği | %0 | Varsayım — altına inilirse İNCELE kararı |
| Kargo/Ops. Varsayılan Oran | %5 | Varsayım — kalem girişinde önerilen oran |
| Reklam Varsayılan Oran | %5 | Varsayım |
| İade Varsayılan Oran | %5 | Varsayım |
| Senaryo Çarpanları (İyi/Baz/Kötü/Kritik) | %5 / %0 / −%5 / −%10 | Varsayım — satış fiyatı belirsizliği bant aralığı |
| Tahmin Çarpanı | 1,645 | Varsayım (normal dağılım %90 güven aralığı) |
| Kalite Skoru İyi / Orta Eşiği | 90 / 70 | Varsayım |
| Toplam Hakediş Fark Eşiği | 500 ₺ | Varsayım — toplam fark kontrolü |
| P90 Yüzdelik Oranı | 0,90 | Varsayım |
| Temel Giriş Kolonu Sayısı | 7 | Varsayım — veri kalite skoru temeli |
| Tornado Komisyon Adımı | +1 puan | Varsayım — komisyon duyarlılık adımı |
| Tornado Satış Fiyatı Adımı | −%5 | Varsayım — satış fiyatı duyarlılık adımı |
| Yüksek Toplam Satış Eşiği | 1.000.000 ₺ | Varsayım — kontrol eşiği |
| Aksiyon 1–3 Metinleri | — | Varsayım — karar kapısının ürettiği öneri metinleri |
| Panel Satırı 1–12 | 1–12 | Kullanıcı |

**Not:** Tüm oranlar ve eşikler kullanıcı varsayımıdır; pazaryeri komisyon/kargo oranları sözleşmenize göre AYARLAR'dan güncellenir. Dosya karar destek aracıdır; mali müşavir görüşü yerine geçmez.

## Mevzuat dayanağı

- Pazaryeri komisyon, kargo ve reklam uygulamaları pazaryeri sözleşmelerine tabidir; üründe varsayılan oranlar sözleşmeye göre güncellenir.
- Eksik hakediş tespiti, satıcı ile pazaryeri arasındaki mutabakatın bir parçasıdır; Ticaret Kanunu (6102 s.) kapsamında tacirler arası mutabakat ve cari hesap düzenlemelerine dayanır.
- Kâr marjı ve net kâr hesabı, Tekdüzen Hesap Planı ve Vergi Usul Kanunu (213 s.) gelir-gider esasına uygundur.

**Veri gizliliği:** Dosya tamamen çevrimdışıdır; verileriniz cihazınızdan çıkmaz (G20).
