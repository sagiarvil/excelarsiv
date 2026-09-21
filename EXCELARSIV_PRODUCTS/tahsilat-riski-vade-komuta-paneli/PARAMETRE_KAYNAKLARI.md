# PARAMETRE KAYNAKLARI — Tahsilat Riski & Vade Komuta Paneli

| Anahtar | Değer | Kaynak | Yürürlük |
|---|---|---|---|
| raporTarihi | 11.08.2026 | Kullanıcı / dönem kapanışı | 01.01.2026 |
| zincirTolerans | 10.000 TL | İş kuralı — fatura/ödeme sapma | 01.01.2026 |
| trv_esikGecikmeGun | 30 gün | İç politika | 01.01.2026 |
| trv_esikGeciken | 5 adet | İç politika — karar eşiği | 01.01.2026 |
| trv_kovaEsik1/2/3 | 30/60/90 gün | Yaşlandırma kovaları | 01.01.2026 |
| trv_olasilikK1–K4 | %95/%75/%50/%25 | Tahsilat modeli (örnek) | 01.01.2026 |
| trv_riskTaban | %50 | Risk çarpan tabanı | 01.01.2026 |
| trv_senaryoTemkinli | %85 | Senaryo motoru | 01.01.2026 |
| trv_senaryoBaz | %100 | Senaryo motoru | 01.01.2026 |
| trv_senaryoIyimser | %115 | Senaryo motoru | 01.01.2026 |
| trv_tornadoOran | %8 | Duyarlılık | 01.01.2026 |
| trv_yuzdelikOran | %90 | İstatistik kuralı (PERCENTILE) | 01.01.2026 |
| trv_olcekHedef | 20.000 | Manda A3 Ö1 | 01.01.2026 |

Durum haritası SPEC `akis.durum_haritasi` ile AYARLAR `tblDurumHaritasi` tablosundan gelir. Olasılık oranları örnek modeldir; gerçek oranı kullanıcı AYARLAR’da günceller. Dosya oranların güncelliğini taahhüt etmez.
