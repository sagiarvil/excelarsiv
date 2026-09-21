# PARAMETRE KAYNAKLARI — AMORTİSMAN & SABİT KIYMET SATIŞ ZAMANLAMA STRATEJİSTİ

AYARLAR sayfasındaki her parametrenin kaynağı ve yürürlük tarihi burada özetlenir. Dosyada `anahtar | deger | birim | alt | ust | aciklama | kaynak | yururluk_tarihi | dogrulama_tarihi` sütunlarıyla saklanır (D10).

| Parametre | Varsayılan | Kaynak |
|---|---|---|
| Firma Unvanı | Örnek Firma A.Ş. | Kullanıcı |
| Rapor Tarihi | 10.08.2026 | Kullanıcı |
| Raporu Hazırlayan | Mali İşler Direktörü | Kullanıcı |
| Dosya Sürümü | 1.0.0 | Kullanıcı |
| Kurumlar Vergisi Oranı | %25 | Mevzuat — Kurumlar Vergisi Kanunu Md. 32; resmi oran değişikliklerinde kullanıcı günceller |
| İskonto Oranı (NBD) | %20 | Varsayım — kullanıcı kendi sermaye maliyetiyle günceller |
| Azalan Bakiye Çarpanı | 2,00 | Mevzuat — VUK Md. 315 azalan bakiyede normal oranın en çok 2 katı |
| Kıst Çeyrek Oranları (1Ç/2Ç/3Ç/4Ç) | %25 / %50 / %75 / %100 | Varsayım — yıl içi satışta ayrılan amortisman payı (3/12, 6/12, 9/12, 12/12) |
| Senaryo Satış Çarpanları (İyi/Baz/Kötü/Kritik) | %10 / %0 / −%15 / −%30 | Varsayım — satış bedeli belirsizliği bant aralığı |
| Tahmin Çarpanı | 1,645 | Varsayım (normal dağılım %90 güven aralığı) |
| Kalite Skoru İyi / Orta Eşiği | 90 / 70 | Varsayım |
| Net Nakit Eşiği | 100.000 ₺ | Varsayım — en iyi çeyrek net nakit kontrolü |
| P90 Yüzdelik Oranı | 0,90 | Varsayım |
| Temel Giriş Kolonu Sayısı | 5 | Varsayım — veri kalite skoru temeli |
| Tornado Oran Adımı | %1 | Varsayım — amortisman oranı duyarlılık adımı |
| Yüksek NDD Eşiği | 2.000.000 ₺ | Varsayım — net defter değeri kontrol eşiği |
| Aksiyon 1–3 Metinleri | — | Varsayım — karar kapısının ürettiği öneri metinleri |
| Panel Satırı 1–12 | 1–12 | Kullanıcı |

**Not:** Kurumlar vergisi oranı ve azalan bakiye çarpanı mevzuat dayanaklıdır; resmi oranlar değiştikçe kullanıcı AYARLAR'dan günceller. Kıst oranlar, senaryo çarpanları ve eşikler kullanıcı varsayımıdır. Dosya karar destek aracıdır; mali müşavir görüşü yerine geçmez.

## Mevzuat dayanağı

- Vergi Usul Kanunu (213 s.) Md. 313–320: amortisman esasları, uygulama ve ayırma usulleri.
- VUK Md. 315: amortisman oranları; azalan bakiyede normal oranın en çok 2 katı oran uygulanabilir.
- VUK Md. 328: amortismana tabi iktisadi kıymetlerin satışından doğan kâr/zararın hesabı (satış bedeli ile net defter değeri farkı).
- Kurumlar Vergisi Kanunu (5520 s.) Md. 32: kurumlar vergisi oranı (%25).

**Veri gizliliği:** Dosya tamamen çevrimdışıdır; verileriniz cihazınızdan çıkmaz (G20).
