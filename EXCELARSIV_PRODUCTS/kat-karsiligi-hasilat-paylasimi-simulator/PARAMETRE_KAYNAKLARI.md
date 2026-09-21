# PARAMETRE KAYNAKLARI — Kat Karşılığı / Hasılat Paylaşımı Simülatör

| Anahtar | Değer | Kaynak | Yürürlük |
|---|---|---|---|
| hesapYili | 2026 | Kullanıcı / GIRDI | 01.01.2025 |
| raporTarihi | 11.08.2026 | Üretim (TODAY yok) | 01.01.2025 |
| zincirTolerans | 50.000 TL | İş kuralı — zincir sapma | 01.01.2025 |
| KKH-001 pay_toplam_hedef | 1,0 | Sözleşme modeli | 01.01.2025 |
| KKH-002 kdv_orani | %20 | KDVK — brüt→net | 01.01.2025 |
| KKH-003 dikkat_marj_esik | %5 | İç politika — DİKKAT | 01.01.2025 |
| KKH-004 min_kar_esik | 0 TL | İç politika — UYGUN DEĞİL | 01.01.2025 |
| KKH-005 yuvarlama | 2 | Uygulama notu | 01.01.2025 |
| kkh_senaryoIyi / Kotu | 1,10 / 0,85 | Senaryo motoru | 01.01.2025 |
| kkh_yorumBCarpan | %5 | Hasılat paylaşımı yorum B | 01.01.2025 |
| kkh_olcekHedef | 20.000 | Manda A3 Ö1 | 01.01.2025 |

Durum haritası SPEC `akis.durum_haritasi` ile AYARLAR `tblDurumHaritasi` tablosundan gelir. KDV ve eşikler MOTOR'a sabit yazılmaz; `tblKurallar` INDEX/MATCH ile çekilir. Dosya oranların güncelliğini taahhüt etmez.
