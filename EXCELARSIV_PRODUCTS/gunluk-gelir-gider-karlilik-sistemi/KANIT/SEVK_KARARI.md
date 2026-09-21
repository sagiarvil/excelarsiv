# SEVK KARARI — Birleşik Kapı Raporu

- Dosya: `/Users/macair1/Desktop/excelarsiv-automation/urunler/gunluk-gelir-gider-karlilik-sistemi/GunlukGelirGiderKarlilikSistemi.xlsx`
- SHA-256: `028b0c7365a4c2dd72d5b59eb9d5c13c682775876a9580ac7310c12a4196600d`
- Arketip: **A2** · SPEC: `/Users/macair1/Desktop/excelarsiv-automation/urunler/gunluk-gelir-gider-karlilik-sistemi/SPEC.yaml` (TAM)
- Tarih: 2026-08-12 13:27 · Orkestratör: kapilar.py v2.0.0 (EXCELARSİV Üretim Mandası v6.0)
- Ortam: Python 3.14.6 · Darwin 27.0.0

## NİHAİ KARAR: **SEVK EDİLEBİLİR** (0 KALDI / 48 GEÇTİ)

## Katman tablosu

| Katman | Ad | Durum | KALDI | GEÇTİ | Süre (sn) | Gerekçe | Betik SHA-256 (ilk 16) |
|---|---|---|---|---|---|---|---|
| G | İşçilik | TEMİZ | 0 | 25 | 2.5 | her ürün | `196c027ccde477a0` |
| Ö | Dayanıklılık | TEMİZ | 0 | 1 | 59.7 | her ürün | `690aa76dda858f47` |
| D | Değer | TEMİZ | 0 | 15 | 14.2 | her ürün | `ed3f6da66984ed56` |
| M | Mevzuat | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| E | Eşleştirme | TEMİZ | 0 | 7 | 1.2 | arketip A2 | `e18eb5c601170e23` |
| A | İş Akışı | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| F | Fizibilite | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| S | Saha | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| R | Reçete | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |

## Notlar

- Koşullu katmanlar (M/E/A/F/S/R) yalnızca arketip gerektirdiğinde koşar;
  gerektiği halde betik yoksa bu rapor SEVK EDİLEBİLİR diyemez.
- Betik parmak izleri, denetimi yapan kodun sevk anındaki halini sabitler;
  kapı betiği değişirse bu raporun yeniden üretilmesi zorunludur.
- Katman raporları aynı dizindedir: RAPOR_DENETIM / RAPOR_OLCEK / RAPOR_DEGER
  (+ varsa RAPOR_MEVZUAT / RAPOR_ESLESTIRME / RAPOR_AKIS / RAPOR_FIZIBILITE /
  RAPOR_SAHA / RAPOR_RECETE).
