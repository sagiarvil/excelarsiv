# SÜRÜM NOTLARI — ŞUBE KÂRLILIK VE NAKİT HESAPLAYICI

**Sürüm:** 1.0.0
**Tarih:** 10.08.2026
**Ürün:** SubeKarlilikVeNakitHesaplayici.xlsx
**Üretici betik:** `sube_karlilik_nakit.py`
**Denetim sonucu:** **0 KALDI** (G01–G24 + Ö1 + Ö2 GEÇTİ) — bkz. `KANIT/RAPOR_DENETIM.md`
**Ölçek testi:** 1.500 satır genişletmesi sonrası **0 KALDI** — bkz. `KANIT/RAPOR_OLCEK.md`
**Değer kapıları:** **0 KALDI** (D01–D14) — bkz. `KANIT/RAPOR_DEGER.md`
**LibreOffice doğrulaması:** Tüm kritik hücreler bağımsız motorla yeniden hesaplandı ve beklenen değerlerle birebir eşleşti

## Bu sürümde yapılanlar

1. **Sayfa mimarisi ve akış (G01):** 18 sayfa kullanım sırasına göre sıralandı; giriş
   sayfaları (HIZLI_BASLANGIC, SUBE_LISTESI, GELIR_GIDER, NAKIT_AKISI) başta, motor ve
   karar (HESAP, KONTROL, KARAR) ortada, raporlama (PANO, SENARYO_DUYARLILIK, RAPOR)
   sonda; `KILAVUZ` / `AYARLAR` / `LISTELER` daima en sonda.
2. **Şube kârlılık motoru:** Şube bazında gelir ve gider kayıtlarından işaretli tutar,
   net kâr ve kâr marjı formülle üretilir; hedef marj eşiği AYARLAR'dan okunur.
3. **Nakit akışı:** Tahsilat ve ödeme kayıtlarından aylık net nakit akışı, dönem sonu
   nakit pozisyonu ve açık fazlası otomatik hesaplanır; 6 aylık ciro senaryosu bandı
   (İyi/Baz/Kötü/Kritik) ile gelecek dönem nakit tahmini üretilir.
4. **Karar motoru:** Kâr marjı, nakit pozisyonu ve veri kalitesi eşikleriyle KARLI /
   İNCELE / ZARARLI kararı ve makine gerekçesi üretilir; aksiyon önerileri sunulur.
5. **Analitik modüller:** Ortalama marj ve nakit trendi (T1/T2), gider sapması (O1),
   en yüksek gider kalemi (O2), senaryo motoru (O3), gelir yoğunlaşması HHI (O6),
   canlı veri kalite skoru (O8), zaman serisi projeksiyonu + %90 güven bandı (I1) ve
   P90 yüzdelik dilimi (I2).

## Uyum karnesi

- **0 KALDI** — G01-G24 + Ö1 + Ö2 + D01-D14 kapılarının tamamı GEÇTİ
- Koşullu biçimlendirme: asgari 40 kural üzeri
- Volatil fonksiyon: yalnızca `raporTarihi` (sınır: 12)
- Taşan dizi fonksiyonu: yok; `_xlfn.` öneki: `MAXIFS` kullanıldı
- Koruma: tüm sayfalar `1234`; giriş hücreleri `#FFF2CC` ve kilitsiz

## Teslim paketi

- `SubeKarlilikVeNakitHesaplayici.xlsx` — ürün
- `KANIT/RAPOR_DENETIM.md` — denetim kanıtı (0 KALDI)
- `KANIT/RAPOR_OLCEK.md` — ölçek kanıtı (0 KALDI)
- `KANIT/RAPOR_DEGER.md` — değer kanıtı (0 KALDI)
- `KANIT/RAPOR_ALTIN.md` + `altin_cikti_sube.json` — altın çıktı
- `PARAMETRE_KAYNAKLARI.md` — parametre kaynakları
- `FIYAT_SAVUNMASI.md` — D01/D02 savunması
- `KULLANIM_KILAVUZU.pdf` — kullanım kılavuzu
- `ornek_veri.csv` — örnek veri yapısı
- `SURUM_NOTLARI.md` — bu dosya

## Yasal not

Bu araç karar destek amaçlıdır; mali müşavirlik hizmeti yerine geçmez. Şube kârlılığı
ve nakit yönetimi kararları için profesyonel finans danışmanlığına başvurulmalıdır.
