# SÜRÜM NOTLARI — AŞIRI DÜŞÜK TEKLİF SAVUNMA ROBOTU

**Sürüm:** 1.0.0
**Tarih:** 09.08.2026
**Ürün:** AsiriDusukTeklifSavunmaRobotu.xlsx
**Üretici betik:** `asiri_dusuk_teklif_savunma.py`
**Denetim:** **0 KALDI** — G01–G24 (İşçilik) + Ö1, Ö1b, Ö2 (Dayanıklılık) + D01–D14 (Değer)
**SHA-256:** `868d399a0c9c8511877bf154f6eb851d0712b34bb4e04783803b50f4ea70d2c8`

## 1.0.0 — İlk sertifikalı sürüm

1. **Sayfa mimarisi ve akış (G01):** 17 sayfa kullanım sırasına göre: KAPAK → HIZLI_BASLANGIC →
   IHALE_KARTI → RAKIP_TEKLIFLER → TEKLIF_BILESENLERI → FIYAT_FARKI_ENDEX → MOTOR → KONTROL →
   KARAR → PANO → SENARYO_DUYARLILIK → RAPOR → ORNEK_VERI → DEGISIKLIK_KAYDI → LISTELER →
   AYARLAR → KILAVUZ. Giriş başta, raporlama ortada, kılavuz/ayarlar/listeler sonda.
2. **Sınır değer motoru:** N katsayısı ve yöntem seçimiyle sınır değer hesabı; teklif
   AŞIRI DÜŞÜK/SINIRDA/GÜVENLİ bölgelerine otomatik yerleştirilir.
3. **Savunma katmanı:** Bileşen rayiç sapması (Z skoru), dayanak ağırlığı denetimi,
   ZAYIF/YETERLİ kararı; aşırı düşük kapsamındaki teklif için açıklama seti ve aksiyonlar.
4. **Analitik modüller:** T2 endeks eğilimi, T3 Pareto, T6 kırılım, O1 anomali, O2 tornado,
   O3 senaryo, O8 kalite skoru, I1 tahmin, I2 P90 — 9 modül, 2 ileri (I1, I2), makine üretimli
   yorumlar. Derinlik puanı 140.
5. **Karar motoru:** GÜVENLİ/İNCELE/VERİ YOK kararı, gerekçe satırı, 3 aksiyon önerisi;
   boş dosyada VERİ YOK (D11).
6. **Ölçek ve uç durum (Ö1, Ö2):** LibreOffice ile gerçek yeniden hesaplama; 15.000 satır /
   49.990 formülde 0 hata, 3,0 sn; 7 uç durum senaryosu temiz.
7. **Performans:** Giriş hücresi comment'leri yalnızca dolu hücrelere yazıldı; comment XML'i
   13 MB → ~90 KB. `_xlfn.STDEV.P` ile standart sapma hem Excel'de hem LibreOffice'te doğru
   hesaplanır (popülasyon sapması).
8. **Formül doğrulama:** LO yeniden hesaplamada 0 hata; tablo hesaplanan kolon formülleri
   `excel_uretim/lo_geri_islem.py` ile geri yüklenir (G08 korunur).
9. **Kılavuz ve PDF:** KILAVUZ sayfasına karar kuralları, sayfa haritası, SSS, güvenlik/KVKK
   ve ortam beyanı eklendi; `excel_uretim/kilavuz_pdf.py` ile 5 sayfalık ekran görüntülü
   `KULLANIM_KILAVUZU.pdf` üretildi (PANO, KARAR, RAPOR ekranları render'lı).

## Uyum karnesi

- **0 KALDI** — G01–G24 + Ö1, Ö1b, Ö2 + D01–D14
- Koşullu biçimlendirme: 64 kural
- Volatil fonksiyon: 0 (sınır: 12); rapor tarihi AYARLAR'dan sabit okunur
- Taşan dizi fonksiyonu: yok; `_xlfn.` öneki: TEXTJOIN, STDEV.P
- Koruma: tüm sayfalar `1234`; giriş hücreleri `#FFF2CC` ve kilitsiz
- Ortam: Excel 2016–365 (Windows/Mac), LibreOffice; tamamen çevrimdışı, makro yok

## Teslim paketi

- `AsiriDusukTeklifSavunmaRobotu.xlsx` — ürün
- `KANIT/RAPOR_DENETIM.md` — işçilik kapıları (0 KALDI)
- `KANIT/RAPOR_OLCEK.md` — ölçek/uç durum (0 KALDI)
- `KANIT/RAPOR_DEGER.md` — değer kapıları (0 KALDI)
- `KANIT/PARAMETRE_KAYNAKLARI.md` — parametre kaynakları
- `KANIT/FIYAT_SAVUNMASI.md` — fiyat çapası (D01) ve ayrışma (D02)
- `KANIT/RAPOR_ALTIN.md` + `KANIT/altin_cikti_asiri_dusuk_teklif.json` — regresyon kanıtı
- `KULLANIM_KILAVUZU.pdf` — kullanım kılavuzu
- `ornek_veri.csv` — örnek veri yapısı
- `SURUM_NOTLARI.md` — bu dosya

## Yasal not

Bu araç karar destek amaçlıdır; ihale mevzuatı ve hukuki danışmanlık yerine geçmez. Sınır
değer katsayısı (N), hesap yöntemi ve eşikler ihale dokümanına göre değişir; teklif öncesi
ihale dokümanındaki değerler AYARLAR'dan doğrulanmalı, gerektiğinde ihale uzmanı görüşü alınmalıdır.
