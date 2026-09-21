# SÜRÜM NOTLARI — DÖVİZ AÇIK POZİSYONU VE KUR RİSKİ STRES TESTİ

## 1.0.0 — 10.08.2026

### Bu sürümde olanlar
- Döviz pozisyon tablosu (tarih, döviz cinsi, işlem türü, tutar, işlem kuru, vade, TL değer, kayıt açıklaması).
- Döviz cinsi bazında varlık, borç, net pozisyon, ortalama kur, TL değer ve kur riski motoru.
- Kötümser ve kritik kur şoklarını simüle eden stres testi; UYGUN / İNCELE / DURDUR karar kapısı ve gerekçesi.
- İyimser, baz, kötümser ve kritik senaryolar; kur/pozisyon tornado analizi; tahmin + güven bandı.
- Döviz yoğunlaşması (HHI), veri kalite skoru ve makine üretimi yorum katmanı (analitik motor).
- 8 grafikli yönetim paneli, yönetici özeti raporu ve 1234 koruma şifresi.

### Denetim sonucu
- G kapıları (G01–G24): 0 KALDI
- Ö kapıları (Ö1, Ö2): 0 KALDI
- D kapıları (D01–D14): 0 KALDI
- SHA-256: `b57561245c8164c51c01fd410a62283df420886fe70bb572ec2451edc145e57e`

### Değişen girdi / çıktı
İlk üretim sürümüdür; altın dosya `KANIT/altin_cikti_doviz.json` ile aynı girdide aynı çıktıyı üretmesi doğrulanır.
