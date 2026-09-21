# SEVK KARARI — Birleşik Kapı Raporu

- Dosya: `/Users/macair1/Desktop/excelarsiv-automation/urunler/uretim-recetesi-ve-zam-yansitma-hesaplayici/UretimRecetesiVeZamYansitmaHesaplayici.xlsx`
- SHA-256: `4a332910d3a585eb16f3afacc11eceeb9a07530998f8c86c6fe5b8195c4eaf4c`
- Arketip: **A7** · SPEC: `/Users/macair1/Desktop/excelarsiv-automation/urunler/uretim-recetesi-ve-zam-yansitma-hesaplayici/SPEC.yaml` (TAM)
- Tarih: 2026-08-12 14:30 · Orkestratör: kapilar.py v2.0.0 (EXCELARSİV Üretim Mandası v6.0)
- Ortam: Python 3.14.6 · Darwin 27.0.0

## NİHAİ KARAR: **SEVK EDİLEBİLİR** (0 KALDI / 46 GEÇTİ)

## Katman tablosu

| Katman | Ad | Durum | KALDI | GEÇTİ | Süre (sn) | Gerekçe | Betik SHA-256 (ilk 16) |
|---|---|---|---|---|---|---|---|
| G | İşçilik | TEMİZ | 0 | 24 | 1.3 | her ürün | `196c027ccde477a0` |
| Ö | Dayanıklılık | TEMİZ | 0 | 1 | 71.8 | her ürün | `690aa76dda858f47` |
| D | Değer | TEMİZ | 0 | 16 | 12.5 | her ürün | `ed3f6da66984ed56` |
| M | Mevzuat | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| E | Eşleştirme | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| A | İş Akışı | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| F | Fizibilite | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| S | Saha | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| R | Reçete | TEMİZ | 0 | 5 | 1.1 | arketip A7 | `16119e1cf432676c` |

## Notlar

- Koşullu katmanlar (M/E/A/F/S/R) yalnızca arketip gerektirdiğinde koşar;
  gerektiği halde betik yoksa bu rapor SEVK EDİLEBİLİR diyemez.
- Betik parmak izleri, denetimi yapan kodun sevk anındaki halini sabitler;
  kapı betiği değişirse bu raporun yeniden üretilmesi zorunludur.
- Katman raporları aynı dizindedir: RAPOR_DENETIM / RAPOR_OLCEK / RAPOR_DEGER
  (+ varsa RAPOR_MEVZUAT / RAPOR_ESLESTIRME / RAPOR_AKIS / RAPOR_FIZIBILITE /
  RAPOR_SAHA / RAPOR_RECETE).
