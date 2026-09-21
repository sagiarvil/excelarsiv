# SEVK KARARI — Birleşik Kapı Raporu

- Dosya: `urunler/ihracat-siparis-karlilik-kur/IhracatSiparisKarlilikKur.xlsx`
- SHA-256: `719cde12e29215eff7cf93dd6b02ec394b8f4e04487af0346dba426b6e41ed0a`
- Arketip: **A3+A6** · SPEC: `urunler/ihracat-siparis-karlilik-kur/SPEC.yaml` (TAM)
- Tarih: 2026-08-11 17:02 · Orkestratör: kapilar.py v2.0.0 (EXCELARSİV Üretim Mandası v6.0)
- Ortam: Python 3.14.6 · Darwin 27.0.0

## NİHAİ KARAR: **SEVK EDİLEBİLİR** (0 KALDI / 53 GEÇTİ)

## Katman tablosu

| Katman | Ad | Durum | KALDI | GEÇTİ | Süre (sn) | Gerekçe | Betik SHA-256 (ilk 16) |
|---|---|---|---|---|---|---|---|
| G | İşçilik | TEMİZ | 0 | 24 | 1.0 | her ürün | `196c027ccde477a0` |
| Ö | Dayanıklılık | TEMİZ | 0 | 1 | 51.4 | her ürün | `690aa76dda858f47` |
| D | Değer | TEMİZ | 0 | 16 | 8.8 | her ürün | `ed3f6da66984ed56` |
| M | Mevzuat | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| E | Eşleştirme | ATLANDI | 0 | 0 | 0.0 | arketip gerektirmiyor | `—` |
| A | İş Akışı | TEMİZ | 0 | 7 | 0.5 | arketip A3 | `7c4979b5087b5847` |
| F | Fizibilite | TEMİZ | 0 | 5 | 0.5 | arketip A6 | `3e1e1fb5f288ec9a` |
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
