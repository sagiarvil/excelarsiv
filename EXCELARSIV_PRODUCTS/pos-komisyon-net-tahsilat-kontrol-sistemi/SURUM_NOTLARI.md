# SÜRÜM NOTLARI — POS, KOMİSYON VE NET TAHSİLAT KONTROL SİSTEMİ

**Ürün:** POS, Komisyon ve Net Tahsilat Kontrol Sistemi
**Sürüm:** 1.0.0 (10.08.2026)
**Üretici betik:** `pos_komisyon_net_tahsilat.py`
**Denetim:** **0 KALDI** — G01–G24 (İşçilik) + Ö1, Ö1b, Ö2 (Dayanıklılık) + D01–D14 (Değer)
**Satış fiyatı:** 990 TL
**SHA-256:** `889b589cf8b720a3f059d7cb8528c6b469ffc9ab9ee6ea654e57fc867ce9aab2`

## Bu sürümde neler var

- **POS işlem tablosu:** 1.000 satır kapasiteli `tblPos`; kanal, işlem tarihi, brüt satış, iade ve komisyon oranı girişi; net satış, komisyon tutarı ve net tahsilat otomatik hesaplanır.
- **Kanal hesap motoru:** 49 adımlı hesap katmanı — kanal bazlı brüt/iade/komisyon/net tahsilat, en pahalı kanal, iade ve komisyon oranları.
- **Karar motoru:** 7 uç durumda doğru karar — UYGUN / İNCELE / DURDUR; veri yokken VERİ YOK (boş dosyada sessiz onay yok).
- **Senaryo ve duyarlılık:** İyimser/baz/kötümser net tahsilat bandı + tornado ile hangi değişkenin etkisinin büyük olduğu.
- **Analitik modüller:** T2, T4, O1, O2, O3, O6, O8, I1, I2 kataloğundan 9 modül; derinlik puanı 145 (eşik ≥100, ileri modül 2) — tüm yorumlar makine üretimli.
- **Pano:** 35 canlı KPI, 6 grafik (tahmin katmanı dahil), karar şeridi, kalite skoru.
- **Veri kalitesi:** Boş/negatif giriş kontrolleri, kalite skoru, doğrulama + ipucu penceresi + hücre notu kapsamı %100.

## Değişiklik geçmişi

- **1.0.0** — İlk sürüm: manda v4 sertifikasyonu tamamlandı; D01–D14 dahil tüm kapılar temiz.

## Altın çıktı

`KANIT/altin_cikti_pos.json` aynı girdiyle yeniden üretimde 0 fark verir (regresyon koruması).

## Bilinen sınır

Örnek veri 12 işlem satırıdır; kanal bazlı aylık düzenlilik veri derinleştikçe daha güvenilir okunur. Uzun dönem veri girildikçe tahmin güvenilirliği artar.
