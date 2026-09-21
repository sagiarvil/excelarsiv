# SÜRÜM NOTLARI — VERGİ/SGK BORCUNU TECİL ETMELİ MİYİM? (KREDİ Mİ TECİL Mİ?)

## 1.0.0 — 10.08.2026

### Bu sürümde olanlar
- Borç listesi tablosu (borç kodu, borç türü, borç aslı, gecikme zammı, vade tarihi, aylık zam oranı); geciken ay, toplam borç ve aylık zam yükü hesabı.
- Tecil (6183/48 vadeye yayma), banka kredisi ve peşin ödeme (terkinli) seçeneklerinin taksit formülüyle toplam maliyet hesapları ve iskonto oranıyla bugünkü değere (NBD) indirgenmiş karşılaştırma.
- En düşük maliyetli ödeme yolunu seçen ve kredi onayı yoksa kredi seçeneğini devre dışı bırakan karar kapısı (TECİL YAP / KREDİ KULLAN / PEŞİN ÖDE / VERİ YOK / DURDUR), gerekçesi ve önerilen aksiyonlar.
- İyimser, baz, kötümser ve kritik tecil faizi senaryolarında maliyet bant aralığı; tecil ve kredi faizi duyarlılığı (tornado); gelecek dönem borç tahmini + güven bandı; P90 yüzdelik.
- Borç türü bazında yoğunlaşma (HHI), anomali (MAD), ortalama zam oranı ve veri kalite skoru ile makine üretimi yorum katmanı (analitik motor).
- 9 grafikli yönetim paneli, yönetici özeti raporu ve 1234 koruma şifresi; tüm giriş hücrelerinde açılır ipucu ve açıklama notu.
- Gerçekçilik üst sınırı: geciken ay, zam oranı ve taksit sayıları doğrulama aralıklarıyla sınırlanır; borç girişi yokken karar "VERİ YOK" kalır, negatif girişte "DURDUR" üretir.

### Denetim sonucu
- G kapıları (G01–G24): 0 KALDI
- Ö kapıları (Ö1, Ö1b, Ö2): 0 KALDI
- D kapıları (D01–D14): 0 KALDI
- SHA-256: `be6f5103d53ad8aec25602b5bd1c52542ede5f820ff984026b82b413a34da1ca`

### Değişen girdi / çıktı
İlk üretim sürümüdür; altın dosya `KANIT/altin_cikti_tecil.json` ile aynı girdide aynı çıktıyı üretmesi doğrulanır.
