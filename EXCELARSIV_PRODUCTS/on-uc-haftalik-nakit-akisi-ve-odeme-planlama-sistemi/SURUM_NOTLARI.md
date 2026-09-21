# SÜRÜM NOTLARI — 13 HAFTALIK NAKİT AKIŞI VE ÖDEME PLANLAMA SİSTEMİ

## v1.0 — 10.08.2026

**Ürün:** 13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi
**Slug:** `on-uc-haftalik-nakit-akisi-ve-odeme-planlama-sistemi`
**Satış fiyatı:** 2.490 ₺ (KDV dahil)

### Kapsam
- 13 haftalık nakit giriş-çıkış planı üzerinden para bitiş haftası tespiti.
- Para bitiş / risk haftasına göre karar (VERİ YOK / UYGUN / İNCELE / DURDUR) + gerekçe + önlem önerisi.
- Kayıt sayısı yetersizse karar motoru `VERİ YOK` üretir; boş dosya sessizce `UYGUN` demez (D11).

### Analitik derinlik (D03/D04)
- Temel (T1–T6): kıtlık haftası yaşlandırması, net akış eğimi, Pareto payı, dönem karşılaştırma, mutabakat farkı, kritik ödeme payı.
- Orta (O1–O6, O8): çıkış anomalisi (MAD), tornado duyarlılık, senaryo motoru, nakit dönüşüm oranı, kıtlık maliyeti, HHI yoğunlaşması, canlı veri kalite skoru.
- İleri (I1, I2, I7): 14. hafta bakiye tahmini (FORECAST.LINEAR), net akış P90, nakit açığı / dış finansman ihtiyacı.
- Toplam derinlik puanı: 240.

### Kapı durumu
- G01–G24: 0 KALDI (denetci.py)
- Ö1, Ö1b, Ö2: 0 KALDI (olcek_testi.py — 6.000 satır, 7 uç durum, 0 hata)
- D01–D14: 0 KALDI (deger_kapilari.py)
- Altın çıktı SHA-256: `RAPOR_ALTIN.md` içinde.

### Görünür yapı (17 sayfa)
KAPAK → HIZLI_BASLANGIC → NAKIT_GIRISLERI → NAKIT_CIKISLARI → HAFTALIK_PLAN → MOTOR → KONTROLLER → KARAR_MOTORU → PANO → SENARYO_DUYARLILIK → RAPOR → ORNEK_VERI → TESTLER → DEGISIKLIK_KAYDI → LISTELER → AYARLAR → KILAVUZ

### Önemli değişiklikler
- Tüm sayfalar `1234` şifresiyle korunur; giriş hücreleri (sarı) açıktır.
- HAFTALIK_PLAN tarih formülü `MaksTarihUfku` parametresiyle taşmaya karşı korunur.
- Dönem başı bakiye formülü `N()` ile başlık hücresine düşmeye karşı korunur.
- PANO grafikleri gerçek haftalık plan verisine bağlanmıştır; boş seri yoktur.

### Teslimat
- `OnUcHaftalikNakitAkisiVeOdemePlanlamaSistemi.xlsx`
- `KULLANIM_KILAVUZU.pdf`
- `ornek_veri.csv`
- `KANIT/` — üç denetim raporu, altın çıktı, parametre kaynakları, fiyat savunması

### Bilinen sınır
- Tüm eşik parametreleri (tampon, veri kalitesi, minimum veri sayısı vb.) varsayımdır; alıcı kendi nakit yapısına göre AYARLAR'dan günceller.
