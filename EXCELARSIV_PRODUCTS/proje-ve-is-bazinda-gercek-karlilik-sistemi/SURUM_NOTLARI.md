# SÜRÜM NOTLARI — PROJE VE İŞ BAZINDA GERÇEK KÂRLILIK SİSTEMİ

**Ürün:** Proje ve İş Bazında Gerçek Kârlılık Sistemi
**Sürüm:** 1.0.0 (10.08.2026)
**Üretici betik:** `proje_is_karlilik.py`
**Denetim:** **0 KALDI** — G01–G24 (İşçilik) + Ö1, Ö1b, Ö2 (Dayanıklılık) + D01–D14 (Değer)
**Satış fiyatı:** 2.490 TL
**SHA-256:** `aa615f97dfa8f71c01cd2c4c9cb575618e6cf0c836ea02d40ac676ffa42e79f4`

## Bu sürümde neler var

- **İş giriş tablosu:** 1.000 satır kapasiteli `tblIs`; iş numarası, iş adı, toplam gelir, direkt maliyet ve süre girişi; kâr, marj, günlük kâr ve durum otomatik hesaplanır.
- **İş hesap motoru:** 40 adımlı hesap katmanı — toplam gelir/kâr, ağırlıklı marj, günlük kâr, kârlı/riskli/zararlı adetler, medyan, P10/P90, kâr aralığı, trend eğimi/kesişimi, sonraki iş kârı tahmini, dönem kâr farkı.
- **Karar motoru:** 7 uç durumda doğru karar — UYGUN / İNCELE / DURDUR; veri yokken VERİ YOK (boş dosyada sessiz onay yok).
- **Senaryo ve duyarlılık:** İyimser/baz/kötümser kâr bandı + tornado ile gelir/maliyet etkisi önceliklendirilir.
- **Analitik modüller:** T2, T4, O1, O2, O3, O8, I1, I2 kataloğundan 8 modül; derinlik puanı 130 (eşik ≥100, ileri modül 2) — tüm yorumlar makine üretimli.
- **Pano:** 48 canlı KPI, 8 grafik (tahmin katmanı dahil), karar şeridi, senaryo ve gelir-maliyet-kâr yapısı.
- **Veri kalitesi:** Boş/negatif giriş kontrolleri, kalite skoru, doğrulama + ipucu penceresi + hücre notu kapsamı %100.

## Değişiklik geçmişi

- **1.0.0** — İlk sürüm: manda v4 sertifikasyonu tamamlandı; D01–D14 dahil tüm kapılar temiz.

## Altın çıktı

`KANIT/altin_cikti_proje_karlilik.json` aynı girdiyle yeniden üretimde 0 fark verir (regresyon koruması).

## Bilinen sınır

Örnek veri 10 iş satırıdır; iş adedi ve dönem derinleştikçe tahmin ve yüzdelik katmanları daha güvenilir okunur. Günlük kâr, süre girişinin doğruluğuna bağlıdır.
