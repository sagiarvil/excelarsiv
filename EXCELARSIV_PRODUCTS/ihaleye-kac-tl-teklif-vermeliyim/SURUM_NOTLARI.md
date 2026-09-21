# SÜRÜM NOTLARI — İHALEYE KAÇ TL TEKLİF VERMELİYİM? (SINIR DEĞER HESAPLAYICI)

**Sürüm:** 1.0.0
**Tarih:** 09.08.2026
**Ürün:** IhaleyeKacTlTeklifVermeliyim.xlsx
**Üretici betik:** `sinir_deger_hesaplayici.py`
**Denetim:** **0 KALDI** — G01–G24 (İşçilik) + Ö1, Ö1b, Ö2 (Dayanıklılık) + D01–D14 (Değer)
**SHA-256:** `512e324975c99143ec94750dc6fc1e17d2367c88bfafd4fa8f0e647dd838bc48`

## 1.0.0 — İlk sertifikalı sürüm

1. **Sayfa mimarisi ve akış (G01):** 17 sayfa kullanım sırasına göre: KAPAK → HIZLI_BASLANGIC →
   IHALE_KARTI → RAKIP_TEKLIFLER → ONCEKI_IHALELER → MALIYET_KALEMLERI → MOTOR → KONTROL →
   KARAR → PANO → SENARYO_DUYARLILIK → RAPOR → ORNEK_VERI → DEGISIKLIK_KAYDI → LISTELER →
   AYARLAR → KILAVUZ. Giriş başta, raporlama ortada, kılavuz/ayarlar/listeler sonda.
2. **Sınır değer motoru:** SD = YM×N + OrtAlt×(1−N); geçerli teklif ayrımı, aşırı düşük eşiği ve
   güvenlik payı ile önerilen teklif üretimi; UYGUN / İNCELE / DURDUR / VERİ YOK kararı.
3. **Analitik modüller:** T2 sınır değer eğilimi, T3 Pareto, T6 kırılım, O1 anomali (Z), O2 tornado,
   O3 senaryo, O8 kalite skoru, I1 tahmin (regresyon + alt/üst bandı), I2 P90 — 9 modül, 2 ileri
   (I1, I2), makine üretimli yorumlar. Derinlik puanı 140; motor 49 adımlı (D13).
4. **Karar motoru:** VERİ YOK / VERİ YETERSİZ / DURDUR / İNCELE / UYGUN; gerekçe satırı ve 3
   aksiyon önerisi; boş dosyada VERİ YOK (D11); uç değer (AzamiTutarSiniri) ve negatif tutar
   denetimleri KONTROL sayfasında.
5. **Ölçek ve uç durum (Ö1, Ö2):** LibreOffice ile gerçek yeniden hesaplama; 15.000 satır /
   14.997 formülde 0 hata, 2,9 sn; 7 uç durum senaryosu temiz.
6. **Tahmin düzeltmesi:** `FORECAST.LINEAR` bu LibreOffice sürümünde tanınmadığından (`#NAME?`)
   eski ad `FORECAST` kullanıldı; tahmin artık tarih ekseni üzerinde regresyonla pozitif değer
   üretir ve PANO'daki tahmin aralığı grafiği dolu görünür (D09b).
7. **Formül doğrulama:** `_xlfn.STDEV.P` popülasyon sapması; `STDEV.P(IF(...))` dizi formülü
   gerektirdiğinden IF'siz kolon referansına çevrildi; tablo hesaplanan kolon formülleri
   `excel_uretim/lo_geri_islem.py` ile geri yüklenir (G08 korunur).
8. **Çapraz tutarlılık (D12):** SENARYO_DUYARLILIK tornado değişken adları "Yaklaşık Maliyet (±%5)"
   ve "Ortalama Teklif (±%5)" olarak ayrıştırıldı; MOTOR ve RAPOR'daki aynı adlı KPI'larla
   çakışma giderildi.
9. **Kılavuz ve PDF:** KILAVUZ sayfasına karar kuralları, kaynaklar/varsayımlar, güvenlik/KVKK
   ve ortam beyanı eklendi; `excel_uretim/kilavuz_pdf.py` ile ekran görüntülü
   `KULLANIM_KILAVUZU.pdf` üretildi (PANO, IHALE_KARTI, RAPOR ekranları render'lı).

## Uyum karnesi

- **0 KALDI** — G01–G24 + Ö1, Ö1b, Ö2 + D01–D14
- Koşullu biçimlendirme: 6.058 kural (G13)
- Volatil fonksiyon: 1 (`raporTarihi` = `TODAY()` — AYARLAR'daki tek tarih kaynağı; sınır: 12)
- Taşan dizi fonksiyonu: yok; `_xlfn.` öneki: TEXTJOIN, STDEV.P
- Koruma: tüm sayfalar `1234`; giriş hücreleri `#FFF2CC` ve kilitsiz; formül hücreleri kilitli
- Ortam: Excel 2016–365 (Windows/Mac), LibreOffice; tamamen çevrimdışı, makro yok

## Teslim paketi

- `IhaleyeKacTlTeklifVermeliyim.xlsx` — ürün
- `KANIT/RAPOR_DENETIM.md` — işçilik kapıları (0 KALDI)
- `KANIT/RAPOR_OLCEK.md` — ölçek/uç durum (0 KALDI)
- `KANIT/RAPOR_DEGER.md` — değer kapıları (0 KALDI)
- `KANIT/PARAMETRE_KAYNAKLARI.md` — parametre kaynakları
- `KANIT/FIYAT_SAVUNMASI.md` — fiyat çapası (D01) ve ayrışma (D02)
- `KANIT/RAPOR_ALTIN.md` + `KANIT/altin_cikti_ihaleye_kac_tl.json` — regresyon kanıtı
- `KULLANIM_KILAVUZU.pdf` — kullanım kılavuzu
- `ornek_veri.csv` — örnek veri yapısı
- `SURUM_NOTLARI.md` — bu dosya

## Yasal not

Bu araç karar destek amaçlıdır; ihale mevzuatı ve hukuki danışmanlık yerine geçmez. Sınır
değer katsayısı (N), hesap yöntemi ve eşikler ihale dokümanına göre değişir; teklif öncesi
ihale dokümanındaki değerler AYARLAR'dan doğrulanmalı, gerektiğinde ihale uzmanı görüşü alınmalıdır.
