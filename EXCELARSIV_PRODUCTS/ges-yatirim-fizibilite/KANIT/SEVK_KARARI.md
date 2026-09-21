# SEVK KARARI — Birleşik Kapı Raporu

- Dosya: `urunler/ges-yatirim-fizibilite/GesYatirimFizibilite.xlsx`
- SHA-256: `7eca453300f5b952f419c3f0725530dfd963878bb4d97fda83fca4a0f591c8dd`
- Arketip: **A6** · SPEC: `urunler/ges-yatirim-fizibilite/SPEC.yaml` (TAM)
- Tarih: 2026-08-11 20:20 · Orkestratör: kapilar.py v2.0.0 (EXCELARSİV Üretim Mandası v6.0)
- Ortam: Python 3.14.6 · Darwin 27.0.0

## NİHAİ KARAR: **SEVK EDİLEBİLİR** (0 KALDI / 46 GEÇTİ)

## Katman tablosu

| Katman | Ad | Durum | KALDI | GEÇTİ | Süre (sn) | Gerekçe | Betik SHA-256 (ilk 16) |
|---|---|---|---|---|---|---|---|
| G | İşçilik | TEMİZ | 0 | 24 | 0.4 | her ürün | `196c027ccde477a0` |
| Ö | Dayanıklılık | TEMİZ | 0 | 1 | 16.5 | her ürün | `690aa76dda858f47` |
| D | Değer | TEMİZ | 0 | 16 | 2.8 | her ürün | `ed3f6da66984ed56` |
| M | Mevzuat | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| E | Eşleştirme | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| A | İş Akışı | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| F | Fizibilite | TEMİZ | 0 | 5 | 0.3 | arketip A6 | `3e1e1fb5f288ec9a` |
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
