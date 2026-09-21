# SÜRÜM NOTLARI — KIDEM–İHBAR YÜKÜ VE PERSONEL ÇIKARMA MALİYETİ HESAPLAYICI

**Ürün:** Kıdem–İhbar Yükü ve Personel Çıkarma Maliyeti Hesaplayıcı
**Sürüm:** 1.0.0 (09.08.2026)
**Satış fiyatı:** 1.490 TL
**SHA-256:** `3d0c54d43791ff6cbc35f87224ac3661562ccb7598e36c9aa2de1a6e083f0b6a`

## Bu sürümde neler var

- **Çalışan kartı:** Ad, işe giriş/çıkış tarihi, brüt ücret ve ücret türü (aylık/günlük) girişi;
  tablo kapasitesi 1.000 satır.
- **Kıdem tazminatı motoru:** 1475 s. İş Kanunu md.14 kuralıyla tam yıl + artan gün oranı,
  tavan sınırlı otomatik hesap.
- **İhbar tazminatı:** 4857 s. İş Kanunu md.17 kıdem dilimlerine göre bildirim süresi
  (2/4/6/8 hafta) otomatik seçimi ve günlük ücretten türetme.
- **İzin ücreti:** 4857 s. İş Kanunu md.53 izin hakları (14/20/26 gün) üzerinden kullanılmamış
  izin hesabı.
- **Karar motoru:** UYGUN / İNCELE / DURDUR kararı, makine üretimi gerekçe ve 4 aksiyon;
  veri yokken VERİ YOK kararı.
- **Analitik modüller:** T1-T6, O1-O8, I1-I7 katalogundan 15 modül; derinlik puanı 240
  (eşik ≥100, ileri modül ≥2).
- **Pano:** 18 canlı KPI, 7 grafik, karar ve gerekçe şeridi, en kritik 3 içgörü.
- **Senaryo/duyarlılık:** Kötümser/baz/iyimser senaryo motoru, tornado, kırılım analizi.
- **Rapor:** PDF'e hazır tek sayfa yönetici raporu; varsayımlar, uyarılar ve aksiyonlar.

## Kalite kanıtları (manda v4)

- G kapıları (G01-G24): **0 KALDI**
- Ölçek kapıları (Ö1, Ö1b, Ö2): **0 KALDI** — 1.200 satır / 9.592 formül, 2.3 sn hesap
- Değer kapıları (D01-D14): **0 KALDI** — derinlik 240, D01 7.000 ≥ 4.470, D02 8 ayrım
- Altın çıktı regresyonu: 0 fark

## Bilinen sınırlamalar ve varsayımlar

- Kıdem tazminatı tavanı yıllık Resmi Gazete duyurusundan güncellenmelidir; dosyadaki değer
  kullanıcı doğrulamalıdır (AYARLAR → KidemTavan).
- Bütçe eşiği, anomali eşiği, senaryo çarpanları varsayımdır; kullanıcı kendi işletme
  koşullarına göre AYARLAR'dan değiştirebilir.
- Hesaplar bordro sisteminin ve iş mahkemesinin yerine geçmez; karar destek amaçlıdır.
- Kişisel veri işlenmez; dosya tamamen çevrimdışı çalışır.
