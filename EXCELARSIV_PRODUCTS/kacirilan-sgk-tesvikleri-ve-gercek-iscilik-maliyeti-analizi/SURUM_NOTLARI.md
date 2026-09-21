# SÜRÜM NOTLARI — KAÇIRILAN SGK TEŞVİKLERİ & GERÇEK İŞÇİLİK MALİYETİ ANALİZİ

**Sürüm:** 1.1.1
**Tarih:** 09.08.2026
**Ürün:** KacirilanSgkTesvikleriVeGercekIscilikMaliyetiAnalizi.xlsx
**Üretici betik:** `sgk_tesvikleri.py`
**Denetim:** **0 KALDI** — G01–G24 (İşçilik) + Ö1, Ö1b, Ö2 (Dayanıklılık) + D01–D14 (Değer)
**SHA-256:** `402b46e7647272dcaaa56161d61d3f4c70ae9100a052044064fc5fad438e56b0`

## 1.1.1 — Bu sürümde yapılanlar

1. **Ortam ve uyumluluk beyanı (KILAVUZ):** Excel 2016–365 (Windows/Mac) tam destek,
   LibreOffice Calc desteklenir, Google Sheets desteklenmez — manda §6 gereği yazılı beyan eklendi.
2. **Gizlilik / KVKK bölümü (KILAVUZ):** "Verileriniz cihazınızdan çıkmaz" ve KVKK uyumu
   açıkça yazıldı; bulut tabanlı rakiplere karşı satış argümanı güçlendirildi.
3. **Ekran görüntülü kılavuz (KULLANIM_KILAVUZU.pdf):** `kilavuz_pdf.py` yeniden yazıldı;
   kılavuz artık metin sayfalarının (KILAVUZ, HIZLI_BASLANGIC) yanına ürünün gerçek ekranlarını
   (PANO, CALISAN_KARTI, RAPOR) tek sayfaya sığdırılmış olarak basar. Formül bütünlüğü için
   ürün sayfaları silinmez, yalnızca gizlenir (LibreOffice gizli sayfaları PDF'e basmaz).
   PDF 7 sayfa (manda sınırı 5–8) ve ~178 KB.
4. **Denetim raporlarının yolu:** Üç rapor teslim edilen dosya üzerinden yeniden üretildi;
   "Dosya:" satırı artık `urunler/.../KacirilanSgkTesvikleriVeGercekIscilikMaliyetiAnalizi.xlsx`
   yolunu gösterir, ara `cikti/` yolunu değil. SHA-256 üç raporda da aynı.
5. **Altın çıktı güncellendi:** KILAVUZ içeriği değiştiği için altın dosya yeniden üretildi;
   sıfırdan yeniden üretimle içerik özeti birebir aynı (0 fark, `RAPOR_ALTIN.md`).
   Altın içerik SHA-256: `422457cbba323da64b97fbc9510fea46e25bac58ac9fe216e196dd3999527590`.

## Önceki sürüm (1.1.0) — Değer katmanı sertifikasyonu

1. **Değer katmanı sertifikasyonu (D01–D14):** Önceki sürümdeki 8 KALDI kapı düzeltildi:
   - **D10 (kaynak katmanı):** AYARLAR sayfasına `Birim`, `kaynak`, `yururluk_tarihi`,
     `dogrulama_tarihi` sütunları eklendi; tüm parametrelerde kaynak ve yürürlük tarihi dolu.
   - **D07 (ölü çıktı):** CALISAN_KARTI'na aylık işveren yükü kolonu formülle eklendi;
     `firmaUnvani` dahil 97 ad tanımının hedefi dolu.
   - **D06 (sahte KPI):** SENARYO_DUYARLILIK'teki değişim oranları gerçek hesaplamaya
     bağlandı; sabit sayıdan türeyen KPI kalmadı.
   - **D13 (motor derinliği):** MOTOR 49 isimli/hesaplanmış adıma bölündü; bölen sıfır
     koruması ile tek dev formül yok.
   - **D08/D09 (pano):** PANO'da 17 canlı KPI ve 8 grafik; tahmin aralığı grafiği dahil.
   - **D03/D04/D05 (analitik derinlik):** O1 anomali, O2 tornado, O3 senaryo, O8 kalite,
     I1 tahmin, I2 yüzdelik, I7 geri kazanım, T6 kırılım modülleri formül + makine yorumu ile
     çalışıyor; derinlik puanı 145, 3 ileri modül.
   - **D11 (karar dürüstlüğü):** Boş dosyada karar `VERİ YOK`; gerekçesiz karar üretilmez.
   - **D12 (çapraz tutarlılık):** teşvik türleri ve değerler sayfalar arası tutarlı.
2. **Ölçek ve uç durum (Ö1, Ö2):** LibreOffice ile gerçek yeniden hesaplama; 15.000 satır /
   44.991 formülde 0 hata, 4,1 sn; 7 uç durum senaryosu temiz.
3. **Formül doğrulama:** MOTOR yorum hücrelerinde kullanılan ad tanımları (örn. `DuyarliDegisim`,
   `P50Katsayi`) AYARLAR'dan doğru beslenir; LO yeniden hesaplamada `#NAME?` yok.
4. **Kontrol mantığı:** KONTROL satır 13 ("Maliyet/brüt oran tutarlı mı?") kıdem+yan faydayı
   hesaba katmadığı için her dolu dosyada yanlış FAIL üretiyordu; gerçekten hata yakalayan
   "Aylık teşvik toplam brütü aşıyor mu?" kontrolüyle değiştirildi. Demo dosya artık
   "SAĞLIKLI" kalite seviyesi ve DURDUR kararı üretir.
5. **Duyarlılık formülleri:** SENARYO_DUYARLILIK'teki Değişim (₺) / Etki (%) kolonları
   `$D13`/`$E13` ile kendine referans veriyordu (dairesel → 0); `C13`/`D13` referanslarıyla
   düzeltildi; tornado zirvesi artık canlı değer üretir.
6. **Üretim boru hattı:** `excel_uretim/lo_geri_islem.py` ile LO yeniden hesap sonrası tablo
   hesaplanan kolon formülleri geri yüklenir (G08 korunur).
7. **Kılavuz ve PDF:** KILAVUZ sayfasına karar kuralları, sayfa haritası, SSS ve güvenlik
   bölümleri eklendi; `excel_uretim/kilavuz_pdf.py` ile 8 sayfalık okunaklı
   `KULLANIM_KILAVUZU.pdf` üretildi.

## Önceki sürüm (1.0.0)

1. **Sayfa mimarisi ve akış (G01):** 16 sayfa kullanım sırasına göre; giriş başta, raporlama
   ortada, `KILAVUZ`/`AYARLAR`/`LISTELER` sonda.
2. **Gerçek işçilik maliyeti motoru:** SGK işveren payı + işsizlik + kıdem karşılığı + yan fayda
   ile aylık/yıllık gerçek maliyet; çalışan başı ve saatlik maliyet.
3. **SGK teşvik simülatörü:** 5 puan, ilave istihdam, kadın/genç, engelli, genç girişimci
   teşvikleri; kaçırılan işaretlemesi.
4. **Veri kalitesi ve karar motoru:** 8 canlı kontrol, kalite skoru (0-100), UYGUN/İNCELE/
   DURDUR kararı, gerekçe ve 3 aksiyon önerisi.

## Uyum karnesi

- **0 KALDI** — G01–G24 + Ö1 + Ö2 + D01–D14
- Koşullu biçimlendirme: 8.016 kural
- Volatil fonksiyon: yalnızca `raporTarihi` (sınır: 12)
- Taşan dizi fonksiyonu: yok; `_xlfn.` öneki: TEXTJOIN
- Koruma: tüm sayfalar `1234`; giriş hücreleri `#FFF2CC` ve kilitsiz
- Ortam: Excel 2016–365 (Windows/Mac), LibreOffice; tamamen çevrimdışı, makro yok

## Teslim paketi

- `KacirilanSgkTesvikleriVeGercekIscilikMaliyetiAnalizi.xlsx` — ürün
- `KANIT/RAPOR_DENETIM.md` — işçilik kapıları (0 KALDI)
- `KANIT/RAPOR_OLCEK.md` — ölçek/uç durum (0 KALDI)
- `KANIT/RAPOR_DEGER.md` — değer kapıları (0 KALDI)
- `KANIT/PARAMETRE_KAYNAKLARI.md` — parametre kaynakları
- `KANIT/FIYAT_SAVUNMASI.md` — fiyat çapası (D01) ve ayrışma (D02)
- `KANIT/RAPOR_ALTIN.md` + `KANIT/altin_cikti_kacirilan_sgk_tesvikleri.json` — regresyon kanıtı
- `KULLANIM_KILAVUZU.pdf` — kullanım kılavuzu
- `ornek_veri.csv` — örnek veri yapısı
- `SURUM_NOTLARI.md` — bu dosya

## Yasal not

Bu araç karar destek amaçlıdır; SGK mevzuatı ve mali danışmanlık yerine geçmez. Teşvik
oranları ve hak ediş koşulları için mali müşavir görüşü alınmalıdır.
