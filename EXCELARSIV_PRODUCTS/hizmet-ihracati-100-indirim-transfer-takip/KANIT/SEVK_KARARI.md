# SEVK KARARI — Birleşik Kapı Raporu

- Dosya: `urunler/hizmet-ihracati-100-indirim-transfer-takip/HizmetIhracati100IndirimTransferTakip.xlsx`
- SHA-256: `c14e37e0b3c06d05acc2877e9acc1bf1975b800cc3062c83304c0b8158067872`
- Arketip: **A1+A2** · SPEC: `urunler/hizmet-ihracati-100-indirim-transfer-takip/SPEC.yaml` (TAM)
- Tarih: 2026-08-11 01:22 · Orkestratör: kapilar.py v2.0.0 (EXCELARSİV Üretim Mandası v6.0)
- Ortam: Python 3.14.6 · Darwin 27.0.0

## NİHAİ KARAR: **SEVK EDİLEBİLİR** (0 KALDI / 54 GEÇTİ)

## Katman tablosu

| Katman | Ad | Durum | KALDI | GEÇTİ | Süre (sn) | Gerekçe | Betik SHA-256 (ilk 16) |
|---|---|---|---|---|---|---|---|
| G | İşçilik | TEMİZ | 0 | 24 | 1.1 | her ürün | `196c027ccde477a0` |
| Ö | Dayanıklılık | TEMİZ | 0 | 1 | 37.8 | her ürün | `690aa76dda858f47` |
| D | Değer | TEMİZ | 0 | 16 | 6.7 | her ürün | `ed3f6da66984ed56` |
| M | Mevzuat | TEMİZ | 0 | 6 | 0.3 | arketip A1 | `ff90b3e997161a07` |
| E | Eşleştirme | TEMİZ | 0 | 7 | 0.3 | arketip A2 | `e18eb5c601170e23` |
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
