# SÜRÜM NOTLARI — KOBİ FİNANS YÖNETİM PAKETİ

**Ürün:** KOBİ Finans Yönetim Paketi
**Sürüm:** 1.0.0 (10.08.2026)
**Üretici betik:** `kobi_finans_yonetim_paketi.py`
**Denetim:** **0 KALDI** — G01–G24 (İşçilik) + Ö1, Ö1b, Ö2 (Dayanıklılık) + D01–D14 (Değer)
**Satış fiyatı:** 7.900 TL
**SHA-256:** `9938216c5d4858dfe9bb70aaed51d78718676d7ba9d40818676d56d190411700`

## Bu sürümde neler var

- **Aylık veri tablosu:** 1.000 satır kapasiteli `tblKobi`; dönem, dört tahsilat kanalı (kasa/banka/cari/POS), stok alışı, operasyon gideri, maaş, kredi taksidi ve risk skoru girişi; net nakit değişimi ve kapanış nakdi otomatik hesaplanır.
- **Konsolide nakit modeli:** 41 adımlı hesap motoru — toplam giriş/çıkış, kapanış, tahmini kâr, risk skoru, sapma, tahmin alt/üst sınırı, nakit tamponu, nakit akış oranı.
- **Karar motoru:** 7 uç durumda doğru karar — DURDUR / İNCELE / UYGUN; veri yokken VERİ YOK (boş dosyada sessiz onay yok).
- **Senaryo ve duyarlılık:** İyimser/baz/kötümser kapanış bandı + tornado ile hangi değişkenin kapanışı en çok etkilediği.
- **Analitik modüller:** T2, T4, O1, O2, O3, O4, O8, I1, I2 kataloğundan 9 modül; derinlik puanı 145 (eşik ≥100, ileri modül 2) — tüm yorumlar makine üretimli.
- **Pano:** 92 canlı KPI, 8 grafik (tahmin katmanı dahil), karar şeridi, kalite skoru.
- **Veri kalitesi:** Boş/negatif giriş kontrolleri, kalite skoru, doğrulama + ipucu penceresi + hücre notu kapsamı %100.

## Değişiklik geçmişi

- **1.0.0** — İlk sürüm: manda v4 sertifikasyonu tamamlandı; D01–D14 dahil tüm kapılar temiz.

## Altın çıktı

`KANIT/altin_cikti_kobi_finans.json` aynı girdiyle yeniden üretimde 0 fark verir (regresyon koruması).

## Bilinen sınır

Örnek veri 6 aylıktır; 12 ay üstü sezonellik kalıbı örnekte tam gözlenmez. Uzun dönem veri girildikçe tahmin güvenilirliği artar.
