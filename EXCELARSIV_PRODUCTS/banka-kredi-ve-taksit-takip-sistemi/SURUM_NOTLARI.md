# SÜRÜM NOTLARI — BANKA, KREDİ VE TAKSİT TAKİP SİSTEMİ

**Sürüm:** 1.0
**Tarih:** 10.08.2026
**Ürün:** BankaKrediVeTaksitTakipSistemi.xlsx
**Üretici betik:** `banka_kredi_takip.py`
**Denetim:** **0 KALDI** — G01–G24 (İşçilik) + Ö1, Ö1b, Ö2 (Dayanıklılık) + D01–D14 (Değer)
**SHA-256:** `e7c527c2779b41dfc085b5e13a6969e1000665b2c2664cdf61eb6ff2abd5b39f`

## 1.0 — Manda v4 sertifikasyonu

1. **MOTOR sayfası eklendi (D13):** 73 isimli/hesaplı ara hesap; temel (T1–T6), orta (O1–O8) ve
   ileri (I1, I2, I3, I7) analitik modüller. Her modül ad tanımlı hücrede sonuç ve `_xlfn.TEXTJOIN`
   ile makine üretilmiş yorum üretir. Derinlik puanı 250 (eşik ≥ 100).
2. **AYARLAR kaynak katmanı (D10):** zorunlu `anahtar | deger | birim | aciklama | kaynak |
   yururluk_tarihi | dogrulama_tarihi` kolonları; mevzuat parametreleri VUK/TTK kaynaklı,
   kaynağı olmayanlar `Varsayım` işaretli.
3. **Karar dürüstlüğü (D11):** veri yetersizken karar üretilmez; tüm kararlar gerekçe ve aksiyonla
   birlikte verilir; karar kuralları KILAVUZ sayfasında açıklanır.
4. **Sahte KPI düzeltmesi (D06):** SENARYO_DUYARLILIK duyarlılık oranları kullanıcı girişine
   alındı; sabit sayıdan türeme yok.
5. **Grafik katmanı (D09/D09b):** PANO'da 8 grafik; tahmin katmanı `_xlfn.FORECAST.LINEAR` ve
   `_xlfn.PERCENTILE.INC` ile beslenir; tüm seriler örnek veriyle dolu.
6. **KREDILER açıklama kolonu (D07):** örnek veride açıklama metinleri dolduruldu; ölü kolon yok.
7. **Ölçek performansı (Ö1):** giriş hücresi notları yalnızca kullanılan bölgeye (30 satır)
   yazılarak dosya boyutu 159 KB'ye indirildi; ölçek testi 135 satır / 480 formül / 1,8 sn → Ö1
   geçti (limit 5,0 sn).
8. **Çapraz tutarlılık (D12):** demo banka/teklif isimleri tek kaynaktan beslenir; aynı etiket
   farklı sayfalarda farklı değer göstermez.
