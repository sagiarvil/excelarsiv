# SEVK KARARI — Birleşik Kapı Raporu

- Dosya: `urunler/ithalat-landed-cost-motoru/IthalatLandedCostMotoru.xlsx`
- SHA-256: `97e60ea5b11e7c3a52ebebae4bb47a9d45811b23b24c5ecbe18806f6aa571961`
- Arketip: **A1+A6** · SPEC: `urunler/ithalat-landed-cost-motoru/SPEC.yaml` (TAM)
- Tarih: 2026-08-11 16:41 · Orkestratör: kapilar.py v2.0.0 (EXCELARSİV Üretim Mandası v6.0)
- Ortam: Python 3.14.6 · Darwin 27.0.0

## NİHAİ KARAR: **SEVK EDİLEBİLİR** (0 KALDI / 52 GEÇTİ)

## Katman tablosu

| Katman | Ad | Durum | KALDI | GEÇTİ | Süre (sn) | Gerekçe | Betik SHA-256 (ilk 16) |
|---|---|---|---|---|---|---|---|
| G | İşçilik | TEMİZ | 0 | 24 | 0.4 | her ürün | `196c027ccde477a0` |
| Ö | Dayanıklılık | TEMİZ | 0 | 1 | 17.0 | her ürün | `690aa76dda858f47` |
| D | Değer | TEMİZ | 0 | 16 | 2.8 | her ürün | `ed3f6da66984ed56` |
| M | Mevzuat | TEMİZ | 0 | 6 | 0.2 | arketip A1 | `ff90b3e997161a07` |
| E | Eşleştirme | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| A | İş Akışı | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| F | Fizibilite | TEMİZ | 0 | 5 | 0.2 | arketip A6 | `3e1e1fb5f288ec9a` |
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
