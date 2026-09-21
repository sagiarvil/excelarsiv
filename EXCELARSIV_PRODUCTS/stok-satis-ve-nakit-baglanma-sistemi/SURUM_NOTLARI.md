# SÜRÜM NOTLARI — STOK, SATIŞ VE NAKİT BAĞLANMA SİSTEMİ

## 1.0.0 — 10.08.2026

### Bu sürümde olanlar
- Stok hareket tablosu (tarih, ürün, hareket türü, adet, birim alış fiyatı, depo, tutar, kayıt açıklaması).
- Ürün bazında kalan stok, ortalama maliyet, stok değeri, devir hızı ve negatif stok motoru.
- Stoğa bağlanan nakit hesabı; UYGUN / İNCELE / DURDUR karar kapısı ve gerekçesi.
- İyimser, baz, kötümser senaryolar; fiyat/adet tornado analizi; tahmin + güven bandı.
- Ürün yoğunlaşması (HHI), veri kalite skoru ve makine üretimi yorum katmanı (analitik motor).
- 8 grafikli yönetim paneli, yönetici özeti raporu ve 1234 koruma şifresi.

### Denetim sonucu
- G kapıları (G01–G24): 0 KALDI
- Ö kapıları (Ö1, Ö2): 0 KALDI
- D kapıları (D01–D14): 0 KALDI
- SHA-256: `12b97668c5063cf5cf5bb8344ed635ccb8edca3a5a97995d1aa34d3b41cc965e`

### Değişen girdi / çıktı
İlk üretim sürümüdür; altın dosya `ornek/altin_stok_nakit.json` ile aynı girdide aynı çıktıyı üretmesi doğrulanır.
