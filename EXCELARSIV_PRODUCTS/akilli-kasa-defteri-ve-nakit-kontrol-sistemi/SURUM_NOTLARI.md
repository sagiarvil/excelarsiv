# SÜRÜM NOTLARI — AKILLI KASA DEFTERİ VE NAKİT KONTROL SİSTEMİ

**Sürüm:** 1.0
**Tarih:** 10.08.2026
**Ürün:** AkilliKasaDefteriVeNakitKontrolSistemi.xlsx
**Üretici betik:** `akilli_kasa_defteri.py`
**Denetim:** **0 KALDI** — G01–G24 (İşçilik) + Ö1, Ö1b, Ö2 (Dayanıklılık) + D01–D14 (Değer)
**SHA-256:** `90997bc32be17d8e7969e094cbe49e8e711a3eb5f42d93bad522b9f73c3264d6`

## 1.0 — Manda v4 sertifikasyonu

1. **MOTOR sayfası eklendi (D13):** 88 isimli/heraplı adım; temel (T1–T6), orta (O1–O8) ve
   ileri (I1, I2, I7) analitik modüller; her modül ad tanımlı hücrede sonuç ve `_xlfn.TEXTJOIN`
   ile makine üretilmiş yorum üretir. Derinlik puanı 210 (eşik ≥ 100).
2. **KAPAK ve HIZLI_BASLANGIC sayfaları eklendi (G01):** Premium kapak ve 3 adımda başlangıç
   rehberi; sayfa sırası KAPAK → ... → KILAVUZ/AYARLAR/LISTELER en sonda.
3. **AYARLAR kaynak katmanı (D10):** zorunlu `anahtar | deger | birim | aciklama | kaynak |
   yururluk_tarihi | dogrulama_tarihi` kolonları; mevzuat parametreleri VUK/TTK kaynaklı,
   kaynağı olmayanlar `Varsayım` işaretli.
4. **Karar dürüstlüğü (D11):** boş dosyada tüm kararlar `VERİ YETERSİZ`; gerekçesiz karar yok;
   aksiyonlar türetilir.
5. **Sahte KPI düzeltmesi (D06):** SENARYO_DUYARLILIK iyimser/kötümser parametreleri kullanıcı
   girişine alındı; sabit sayıdan türeme yok.
6. **Grafik serileri (D09/D09b):** PANO'da 8 grafik; NAKIT_PROJEKSIYONU ve GUNLUK_OZET
   hesaplanan kolon formülleri hücrelere yazılarak LO önbelleğinde seri değerleri garanti edildi.
7. **Performans (Ö1):** tablo aralıkları 804 satıra indirildi; MOTOR/KONTROLLER/KARAR_MOTORU/
   GUNLUK_OZET/NAKIT_PROJEKSIYONU/ODEME formülleri tam sayı sütunu referansları yerine
   yapısal tablo referanslarına (`tblKasaHareket[Tutar (₺)]` vb.) çevrildi. Ölçek testi 90 satır /
   165 formül / 3,6 sn → Ö1 geçti (limit 5,0 sn).
8. **Üretim boru hattı:** `excel_uretim/lo_geri_islem.py` ile LO yeniden hesap sonrası tablo
   hesaplanan kolon formülleri ve giriş rengi geri yüklenir (G03, G08 korunur).
