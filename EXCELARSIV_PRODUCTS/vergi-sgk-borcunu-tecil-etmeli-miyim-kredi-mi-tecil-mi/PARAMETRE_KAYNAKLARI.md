# PARAMETRE KAYNAKLARI — VERGİ/SGK BORCUNU TECİL ETMELİ MİYİM? (KREDİ Mİ TECİL Mİ?)

AYARLAR sayfasındaki her parametrenin kaynağı ve yürürlük tarihi burada özetlenir. Dosyada `anahtar | deger | birim | alt | ust | aciklama | kaynak | yururluk_tarihi | dogrulama_tarihi` sütunlarıyla saklanır (D10).

| Parametre | Varsayılan | Kaynak |
|---|---|---|
| Firma Unvanı | Örnek Firma A.Ş. | Kullanıcı |
| Rapor Tarihi | 10.08.2026 | Kullanıcı |
| Raporu Hazırlayan | Mali İşler Direktörü | Kullanıcı |
| Dosya Sürümü | 1.0.0 | Kullanıcı |
| Tecil Peşinat Oranı | %10 | Varsayım — tecil başvurusunda peşin ödenen kısım oranı; idare kararına göre güncellenir |
| Tecil Taksit Sayısı | 12 | Varsayım — 6183/48 kapsamında en çok 48 aya kadar; kullanıcı günceller |
| Tecil Faizi (yıllık) | %18 | Varsayım — 6183 Md. 48 kapsamında belirlenen tecil faizi; resmi oran yayımlandıkça güncellenir |
| Kredi Faizi (yıllık) | %35 | Varsayım — borcu kapatacak banka kredisinin kullanıcı tarafından girilen oranı |
| Kredi Komisyon Oranı | %2 | Varsayım — kredi tahsis komisyonu |
| Kredi Taksit Sayısı | 12 | Varsayım — kullanıcı günceller |
| Kredi Onayı Var | Evet | Kullanıcı — Hayır ise karar kapısı kredi seçeneğini dışarıda bırakır |
| Peşin Terkin Oranı | %50 | Varsayım — peşin ödemede gecikme zammından terkin edilen oran (yapılandırma düzenlemeleri); ilgili tebliğe göre güncellenir |
| İskonto Oranı (NBD) | %30 | Varsayım — kullanıcı kendi sermaye maliyetiyle günceller |
| Senaryo Tecil Faizleri (İyi/Baz/Kötü/Kritik) | %12 / %18 / %25 / %35 | Varsayım — senaryo bant aralığı oranları |
| Tahmin Çarpanı | 1,645 | Varsayım (normal dağılım %90 güven aralığı) |
| Kalite Skoru İyi / Orta Eşiği | 90 / 70 | Varsayım |
| Aylık Nakit Akışı | 100.000 ₺ | Varsayım — tecil taksitinin nakit akışına oranı kontrol edilir |
| Yüksek Borç Eşiği | 500.000 ₺ | Varsayım — toplam borç kontrol eşiği |
| P90 Yüzdelik Oranı | 0,90 | Varsayım |
| Temel Giriş Kolonu Sayısı | 5 | Varsayım — veri kalite skoru temeli |
| Tornado Faiz Çarpanı (tecil/kredi) | %1 / %1 | Varsayım — duyarlılık adımı |
| Panel Satırı 1–12 | 1–12 | Kullanıcı |

**Not:** Tecil faizi ve terkin oranı idari/mevzuat düzenlemelerine göre değişir; dosya bu değerleri resmi yayımlandığı gibi getirmez, kullanıcı AYARLAR'dan günceller. Kredi faizi ve komisyon banka teklifine göre girilir. Diğer eşik ve çarpanlar kullanıcı varsayımıdır. Dosya karar destek aracıdır; mali müşavir görüşü yerine geçmez.

## Mevzuat dayanağı

- Amme Alacaklarının Tahsil Usulü Hakkında Kanun (6183 s.) Md. 48: amme alacaklarının tecili; tecil faizi ve ödeme planı esasları.
- 6183 s. Kanun Md. 51: gecikme zammı oranı ve hesaplanması (aylık oran aylık zam oranı olarak kullanılır).
- Vergi Usul Kanunu (213 s.) ve SGK 5510 s. Kanun: vergi/SGK alacaklarının takip esasları.
- Yapılandırma düzenlemeleri (ör. 7326, 7440 sayılı Kanunlar): peşin ödemede gecikme zammından terkin oranları — ilgili tebliğe göre güncellenir.

**Veri gizliliği:** Dosya tamamen çevrimdışıdır; verileriniz cihazınızdan çıkmaz (G20).
