# PARAMETRE KAYNAKLARI — 13 Haftalık Nakit + Ödeme Önceliği

| Anahtar | Değer | Kaynak | Yürürlük |
|---|---|---|---|
| raporTarihi | 10.08.2026 | Kullanıcı / dönem kapanışı | 01.01.2026 |
| zincirTolerans | 10.000 TL | İş kuralı — plan/onay sapma | 01.01.2026 |
| onh_esikGecikmeGun | 30 gün | İç politika | 01.01.2026 |
| onh_esikGeciken | 5 adet | İç politika — karar eşiği | 01.01.2026 |
| onh_gecikmeOranGun | %0,05 | Finans varsayımı (örnek) | 01.01.2026 |
| onh_senaryoTemkinli | %85 | Senaryo motoru | 01.01.2026 |
| onh_senaryoBaz | %100 | Senaryo motoru | 01.01.2026 |
| onh_senaryoIyimser | %115 | Senaryo motoru | 01.01.2026 |
| onh_tornadoOran | %8 | Duyarlılık | 01.01.2026 |
| onh_yuzdelikOran | %90 | İstatistik kuralı (PERCENTILE) | 01.01.2026 |
| onh_olcekHedef | 20.000 | Manda A3 Ö1 | 01.01.2026 |

Durum haritası SPEC `akis.durum_haritasi` ile AYARLAR `tblDurumHaritasi` tablosundan gelir. Gecikme oranı örnek finans varsayımıdır; gerçek oranı kullanıcı AYARLAR’da günceller. Dosya oranların güncelliğini taahhüt etmez.
