# SÜRÜM NOTLARI — HAKEDİŞ FİYAT FARKI VE HAK KAYBI CETVELİ

## 1.0.0 (09.08.2026) — İlk sürüm

**SHA-256:** `7ae441047da1603bc1f5b33e128a0641d8ef4b419e5ec152a59d971f8ba75a65`

### Yeni
- Sözleşme kartı, hakediş kalemleri, endeks kayıtları ve dönem kayıtları için yapılandırılmış
  giriş sayfaları (sarı hücreler giriş, açılır ipucu ve hücre notu her girişte).
- Fiyat farkı hesabı (F = a + b×(Gk/G0) + c×(Bk/B0) + d×(Yk/Y0) − 1), endeks katsayıları
  tablodan okunur.
- Hak kaybı analizi: piyasa birim fiyatı ile sözleşme fiyatı farkından kalem bazlı kayıp,
  toplam hak kaybı oranı ve geç ödeme maliyeti.
- Karar motoru: `UYGUN / İNCELE / DURDUR` + gerekçe; veri yoksa `VERİ YOK`.
- Analitik modüller: trend ve tahmin aralığı (I1), yüzdelik P10/P50/P90 (I2), anomali (MAD)
  (O1), duyarlılık/tornado (O2), senaryo motoru (O3), yoğunlaşma (HHI) (O6), canlı veri
  kalite skoru (O8), pareto kırılımı (T3), dönem karşılaştırma (T4), mutabakat farkı (T5).
- Kontrol paneli: 12 denetim satırı, dolu giriş sayacı, formül hata sayısı.
- Pano: 12+ canlı KPI, 8 grafik (biri tahmin aralıklı), karar rozeti, güven/risk göstergesi.
- Rapor: A4 yazdırılabilir tek sayfa, yönetici özeti, varsayımlar, aksiyonlar, disclaimer.
- Örnek veri modu, değişiklik kaydı, liste kaynakları, ayarlar (kaynak/yürürlük tarihli),
  kılavuz sayfası.

### Koruma
- Tüm sayfalar `1234` şifresiyle korunur. Giriş (sarı) hücreleri kilitsiz; formül ve başlık
  hücreleri kilitlidir. Kılavuz sayfasında şifre ve kilit açma talimatı yer alır.

### Mevzuat dayanağı
- Fiyat farkı katsayı ve ağırlık yöntemi sözleşme hükümlerine ve ilgili Hazine ve Maliye
  Bakanlığı tebliğlerine; geç ödeme eşiği Türk Borçlar Kanunu md.117'ye dayanır. Tüm
  parametre kaynakları KANIT/PARAMETRE_KAYNAKLARI.md ve AYARLAR sayfasında listelenir.

### Uyum
- Manda v4 denetimleri: G01–G24 → 0 KALDI, Ö1/Ö2 → 0 KALDI, D01–D14 → 0 KALDI.
  Raporlar KANIT/ dizinindedir (SHA-256: e97bfbcbcf9c139d167239676f1d5367d695edbd3bbb53eebe557b7b1f26f2c8).
- Altın çıktı regresyon karşılaştırması `altin_cikti.json` ile yapılır.
