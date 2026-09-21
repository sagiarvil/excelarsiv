# SEVK KARARI — Birleşik Kapı Raporu

- Dosya: `/Users/macair1/Desktop/excelarsiv-automation/urunler/kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici/KidemIhbarYukuVePersonelCikarmaMaliyetiHesaplayici.xlsx`
- SHA-256: `6df3b67e728826c94303bb8776f6a28992982c0b7e622981fe91afc422f34751`
- Arketip: **A1** · SPEC: `/Users/macair1/Desktop/excelarsiv-automation/urunler/kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici/SPEC.yaml` (TAM)
- Tarih: 2026-08-12 02:55 · Orkestratör: kapilar.py v2.0.0 (EXCELARSİV Üretim Mandası v6.0)
- Ortam: Python 3.14.6 · Darwin 27.0.0

## NİHAİ KARAR: **SEVK EDİLEBİLİR** (0 KALDI / 47 GEÇTİ)

## Katman tablosu

| Katman | Ad | Durum | KALDI | GEÇTİ | Süre (sn) | Gerekçe | Betik SHA-256 (ilk 16) |
|---|---|---|---|---|---|---|---|
| G | İşçilik | TEMİZ | 0 | 24 | 0.9 | her ürün | `196c027ccde477a0` |
| Ö | Dayanıklılık | TEMİZ | 0 | 1 | 39.9 | her ürün | `690aa76dda858f47` |
| D | Değer | TEMİZ | 0 | 16 | 7.9 | her ürün | `ed3f6da66984ed56` |
| M | Mevzuat | TEMİZ | 0 | 6 | 0.7 | arketip A1 | `ff90b3e997161a07` |
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
