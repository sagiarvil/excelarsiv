# SÜRÜM NOTLARI — AYLIK PATRON FİNANS PANELİ

**Sürüm:** 1.0.0
**Tarih:** 10.08.2026
**Ürün:** AylikPatronFinansPaneli.xlsx
**Üretici betik:** `aylik_patron_finans_paneli.py`
**Denetim:** **0 KALDI** — G01–G24 (İşçilik) + Ö1, Ö2 (Dayanıklılık) + D01–D14 (Değer)
**Altın çıktı içerik SHA-256:** `3a0724443759577c80e0dee8f86225f024eca24c1819e7e72bf734a895009eeb`

## 1.0.0 — İlk sertifikalı sürüm

1. **Sayfa mimarisi ve akış (G01):** 15 sayfa kullanım sırasına göre; girişler başta
   (AYLIK_GELIR_GIDER, BORC_VE_ALACAK), raporlama ortada (PANO, RAPOR), `KILAVUZ`/`AYARLAR`/
   `LISTELER` sonda.
2. **AYARLAR (D10):** `anahtar | deger | birim | aciklama | kaynak | yururluk_tarihi |
   dogrulama_tarihi` yapısı; tüm parametrelerde kaynak ve tarih dolu; varsayımlar toplu
   `KANIT/PARAMETRE_KAYNAKLARI.md`'de.
3. **FINANS_MOTORU (D13):** 40+ isimli ara hesap adımı ve 17 analitik modül (T1–T6, O1–O8,
   I1–I7) — yaşlandırma, trend, Pareto, dönem farkı, tahakkuk/nakit ayrışması, anomali,
   tornado, senaryo farkı, nakit dönüşüm, gecikme maliyeti, HHI yoğunlaşma, canlı kalite skoru,
   tahmin aralığı, yüzdelikler, varyans köprüsü, nakit açığı köprüleme. Her modül ad tanımlı
   hücrede formülle çalışır ve `_xlfn.TEXTJOIN` ile makine yorumu üretir (D03/D04/D05).
4. **KONTROLLER (D06):** Kalite skoru canlıdır — dolu giriş hücresi sayısını ölçer, sabit
   sayıdan türemez.
5. **KARAR (D11):** Dolu kayıt sayısı eşiğin altındaysa **VERİ YOK**; aksi hâlde kâr + gelir +
   borçluluk + kasa yeterliliğine göre UYGUN/İNCELE/DURDUR; gerekçe ve 3 türetilmiş aksiyon.
6. **PANO (D08/D09):** 12 canlı KPI, 8 grafik (tahmin aralığı dahil), karar rozeti ve risk
   göstergeleri.
7. **SENARYO_DUYARLILIK:** Canlı "en etkili değişken" tespiti; iyimser/baz/kötümser senaryo
   karşılaştırması AYARLAR çarpanlarından beslenir.
8. **Ölçek ve uç durum (Ö1, Ö2):** LibreOffice ile gerçek yeniden hesaplama; 1.000 satır
   ölçek testinde 0 hata, 5 sn altında hesap süresi; 7 uç durum senaryosu temiz.
9. **KILAVUZ:** Karar kuralları, varsayım listesi, gizlilik/KVKK beyanı ("Verileriniz
   cihazınızdan çıkmaz"), ortam beyanı (Excel 2016–365 Windows/Mac, LibreOffice; tamamen
   çevrimdışı, makro yok).
10. **Altın çıktı:** `KANIT/altin_cikti_aylik_patron_finans_paneli.json` üretildi; aynı
    girdiyle yeniden üretim birebir aynı içerik veriyor (0 fark, `RAPOR_ALTIN.md`).

## Uyum karnesi

- **0 KALDI** — G01–G24 + Ö1 + Ö2 + D01–D14
- Koşullu biçimlendirme: 40+ kural
- Volatil fonksiyon: yalnızca `raporTarihi` (sınır: 12)
- Taşan dizi fonksiyonu: yok; `_xlfn.` öneki: TEXTJOIN
- Koruma: tüm sayfalar `1234`; giriş hücreleri `#FFF2CC` ve kilitsiz
- Ortam: Excel 2016–365 (Windows/Mac), LibreOffice; tamamen çevrimdışı, makro yok

## Teslim paketi

- `AylikPatronFinansPaneli.xlsx` — ürün
- `KANIT/RAPOR_DENETIM.md` — işçilik kapıları (0 KALDI)
- `KANIT/RAPOR_OLCEK.md` — ölçek/uç durum (0 KALDI)
- `KANIT/RAPOR_DEGER.md` — değer kapıları (0 KALDI)
- `KANIT/PARAMETRE_KAYNAKLARI.md` — parametre kaynakları
- `KANIT/FIYAT_SAVUNMASI.md` — fiyat çapası (D01) ve ayrışma (D02)
- `KANIT/RAPOR_ALTIN.md` + `KANIT/altin_cikti_aylik_patron_finans_paneli.json` — regresyon kanıtı
- `KULLANIM_KILAVUZU.pdf` — kullanım kılavuzu
- `ornek_veri.csv` — örnek veri yapısı
- `SURUM_NOTLARI.md` — bu dosya

## Yasal not

Bu araç karar destek amaçlıdır; mali müşavirlik veya hukuki danışmanlık yerine geçmez.
Karar eşikleri işletmeye özgü varsayımdır; finansal kararlar için mali müşavir görüşü alınmalıdır.
