# SÜRÜM NOTLARI — Stok Optimizasyon + ABC + Ölü Stok Nakit

## 1.0.0 — 11.08.2026

- İlk sevk: A3 iş akışı / stok ABC + ölü stok nakit
- Durum haritası: SAYIM → ABC → AKSIYON → KAPALI
- KARTLAR + AKIS kimlik disiplini; yetim SKU kaydı bayrağı
- Geçersiz geçiş uyarısı (engellemez, görünür kılar)
- Nakit zinciri: Stok adet × birim maliyet ↔ kontrol nakit — ZİNCİR KIRIK
- KUYRUKLAR: Sayım / ABC / Aksiyon / Ölü Stok; kapalılar düşer
- Yaşlandırma kovaları 0-7 / 8-30 / 31-90 / 90+
- Ölü stok nakit + önerilen sipariş; karar: SİPARİŞ / BEKLE / ÖLÜ STOK ERİT / VERİ YOK
- Kurallar: SOA-001..005
- Analitik: T1, T2, O1, O2, O3, O6, O8, M_SEN, I1, I2 (soa_*)
- Ölçek sözleşmesi: 20.000 satır (Ö1); dosya içi tablo kapasitesi 250
- Koruma şifresi: 1234
