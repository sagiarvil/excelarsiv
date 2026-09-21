# SEVK KARARI — Birleşik Kapı Raporu

- Dosya: `urunler/dernek-kooperatif-mali-yonetim/DernekKooperatifMaliYonetim.xlsx`
- SHA-256: `6dd6c5ebe73ea553a5a2c204bd6c737fb28dd57fe3bb7ca50de27c0e508fe305`
- Arketip: **A8** · SPEC: `urunler/dernek-kooperatif-mali-yonetim/SPEC.yaml` (TAM)
- Tarih: 2026-08-11 22:20 · Orkestratör: kapilar.py v2.0.0 (EXCELARSİV Üretim Mandası v6.0)
- Ortam: Python 3.14.6 · Darwin 27.0.0

## NİHAİ KARAR: **SEVK EDİLEBİLİR** (0 KALDI / 54 GEÇTİ)

## Katman tablosu

| Katman | Ad | Durum | KALDI | GEÇTİ | Süre (sn) | Gerekçe | Betik SHA-256 (ilk 16) |
|---|---|---|---|---|---|---|---|
| G | İşçilik | TEMİZ | 0 | 24 | 0.3 | her ürün | `196c027ccde477a0` |
| Ö | Dayanıklılık | TEMİZ | 0 | 1 | 16.9 | her ürün | `690aa76dda858f47` |
| D | Değer | TEMİZ | 0 | 16 | 2.7 | her ürün | `ed3f6da66984ed56` |
| M | Mevzuat | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| E | Eşleştirme | TEMİZ | 0 | 7 | 0.2 | arketip A8 | `e18eb5c601170e23` |
| A | İş Akışı | TEMİZ | 0 | 6 | 0.2 | arketip A8 | `7c4979b5087b5847` |
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
