# SEVK KARARI — Birleşik Kapı Raporu

- Dosya: `urunler/uretim-kari-125-kv-optimizasyonu/UretimKari125KvOptimizasyonu.xlsx`
- SHA-256: `ed8da2e36ae1c2bc189efbccd2605e77afe60f931f7aee6c0355d4e5780f0d90`
- Arketip: **A1** · SPEC: `urunler/uretim-kari-125-kv-optimizasyonu/SPEC.yaml` (TAM)
- Tarih: 2026-08-11 13:10 · Orkestratör: kapilar.py v2.0.0 (EXCELARSİV Üretim Mandası v6.0)
- Ortam: Python 3.14.6 · Darwin 27.0.0

## NİHAİ KARAR: **SEVK EDİLEBİLİR** (0 KALDI / 47 GEÇTİ)

## Katman tablosu

| Katman | Ad | Durum | KALDI | GEÇTİ | Süre (sn) | Gerekçe | Betik SHA-256 (ilk 16) |
|---|---|---|---|---|---|---|---|
| G | İşçilik | TEMİZ | 0 | 24 | 0.4 | her ürün | `196c027ccde477a0` |
| Ö | Dayanıklılık | TEMİZ | 0 | 1 | 34.4 | her ürün | `690aa76dda858f47` |
| D | Değer | TEMİZ | 0 | 16 | 5.3 | her ürün | `ed3f6da66984ed56` |
| M | Mevzuat | TEMİZ | 0 | 6 | 0.2 | arketip A1 | `ff90b3e997161a07` |
| E | Eşleştirme | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
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
