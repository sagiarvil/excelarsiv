SURUM_NOTLARI.md
# SÜRÜM NOTLARI — YILLARA SARİ İNŞAAT STOPAJ & NAKİT AKIŞ PLANLAYICI

## 1.0.0 (09.08.2026) — İlk sürüm

**SHA-256:** `bb96fe4410f175a2ec7a8c842e4bbdbd88ab07c24a7b6c76c7828132a5ed0fde`

### Yeni
- Proje kartı, yıllık plan, hakedişler, nakit akışı ve vergi planı için yapılandırılmış giriş
  sayfaları (sarı hücreler giriş, açılır ipucu ve hücre notu her girişte).
- Kesinti hesabı: KDV = Brüt × KdvOrani; KDV Tevkifatı = Brüt × KdvOrani × KdvTevkifatOrani;
  Stopaj = Brüt × GelirVergisiStopajOrani; Net = Brüt + KDV − tevkifat − stopaj.
- Vergi planı: yıllık tahmini matrah, geçici ve kurumlar vergisi, stopaj mahsubu, ödenecek
  vergi ve vergi yükü oranı.
- Nakit planı: dönemlik net nakit akışı, kümülatif pozisyon (N tabanlı O(n) formül), nakit
  açığı ve açık dönemi sayısı.
- Karar motoru: `UYGUN / İNCELE / DURDUR` + gerekçe; veri yoksa `VERİ YOK`.
- Analitik modüller: hakediş zaman kapsamı (T1), eğilim (T2), pareto pay (T3), plan-gerçekleşme
  farkı (T5), başabaş (T6), anomali Z skoru (O1), tornado (O2), senaryo motoru (O3), pozitif
  nakit oranı (O4), yoğunlaşma HHI (O6), veri kalite skoru (O8), tahmin aralığı (I1), P90 (I2),
  iskonto edilmiş nakit NPV (I3), nakit açığı köprüsü (I7). Her modülün makine üretimi yorumu
  `_xlfn.TEXTJOIN` ile üretilir.
- Kontrol paneli: 10 denetim satırı, PASS/FAIL sayacı, veri kalite skoru ve güven seviyesi.
- Pano: 12+ canlı KPI, 8 grafik, karar rozeti, güven/risk göstergesi.
- Rapor: A4 yazdırılabilir tek sayfa, yönetici özeti, varsayımlar, uyarılar, aksiyonlar,
  disclaimer.
- Örnek veri modu, değişiklik kaydı, liste kaynakları, ayarlar (kaynak/yürürlük tarihli),
  kılavuz sayfası.

### Koruma
- Tüm sayfalar `1234` şifresiyle korunur. Giriş (sarı) hücreleri kilitsiz; formül ve başlık
  hücreleri kilitlidir. Kılavuz sayfasında şifre ve kilit açma talimatı yer alır.

### Mevzuat dayanağı
- GVK md.42-44: yıllara sari inşaat ve onarma işlerinde kârın işin bittiği yıl tespiti.
- GVK md.94: hakediş ödemelerinde gelir vergisi stopajı (yıllara sari işlerde %5).
- KDV Genel Uygulama Tebliği: yapım işlerinde sorumlu sıfatıyla KDV tevkifatı (4/10).
- Kurumlar Vergisi Kanunu md.21, md.32: geçici ve kurumlar vergisi oranları.
- Tüm parametre kaynakları KANIT/PARAMETRE_KAYNAKLARI.md ve AYARLAR sayfasında listelenir.

### Performans notu
- Kümülatif nakit toplamı `N($K5)+$J6` ile O(n) yapıldı; vergi planı SUMIFS'leri tarih koşullu
  tek seferlik taramaya indirildi. Ölçek testi (4.800 satır, 17.985 formül) 3,2 sn'de temizlendi.
