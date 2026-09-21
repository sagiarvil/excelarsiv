# SÜRÜM NOTLARI — ORTAKLAR CARİ & KASA ADAT FAİZ FATURASI HESAPLAYICI

## 1.0.0 — 10.08.2026

### Bu sürümde olanlar
- Ortak cari hareket tablosu (tarih, ortak adı, işlem türü, tutar, işaretli tutar, bakiye, gün sayısı, adat, faiz, açıklama).
- Kasa adat tablosu; ortak ve kasa adatlarının tek motorda birleştirilmesiyle brüt faiz faturası tutarı.
- UYGUN / İNCELE / DURDUR karar kapısı ve gerekçesi; örtülü sermaye (KKEG) risk izleme eşiği.
- İyimser, baz, kötümser senaryolar; faiz oranı/adat tornado analizi; gelecek dönem tahmini + güven bandı; P90 yüzdelik.
- Ortak adı/faiz yoğunlaşması (HHI), veri kalite skoru ve makine üretimi yorum katmanı (analitik motor).
- 8 grafikli yönetim paneli, yönetici özeti raporu ve 1234 koruma şifresi.

### Denetim sonucu
- G kapıları (G01–G24): 0 KALDI
- Ö kapıları (Ö1, Ö2): 0 KALDI
- D kapıları (D01–D14): 0 KALDI
- SHA-256: `9c44aa98fb2b4d65caa294bc794acc9f0ea1862d512f0fbed8926f19a115e87e`

### Değişen girdi / çıktı
İlk üretim sürümüdür; altın dosya `KANIT/altin_cikti_ortak.json` ile aynı girdide aynı çıktıyı üretmesi doğrulanır.
