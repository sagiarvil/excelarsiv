# SÜRÜM NOTLARI — YENİDEN DEĞERLEME YAPMALI MIYIM? (VERGİ TASARRUF ANALİZİ)

## 1.0.0 — 10.08.2026

### Bu sürümde olanlar
- Kıymet listesi tablosu (kıymet türü, defter değeri, birikmiş amortisman, amortisman oranı, kalan faydalı ömür, yeniden değerleme katsayısı); net defter değeri, değerlenmiş net değer, değer artışı, %2 fon vergisi ve yıllık ek amortisman hesabı.
- Ek amortismanın kalan ömür boyunca sağladığı vergi tasarrufunun iskonto edilerek bugünkü değerine indirgenmesi ve kıymet bazında net fayda (NBD − fon vergisi) hesabı.
- Net fayda ve değer artışı eşiklerine göre UYGUN / İNCELE / DURDUR karar kapısı, gerekçesi ve önerilen aksiyonlar.
- İyimser, baz, kötümser ve kritik senaryo bant aralığı; iskonto oranı ve fon oranı duyarlılığı (tornado); gelecek dönem tasarruf tahmini + güven bandı; P90 yüzdelik.
- Kıymet türü bazında değer artışı yoğunlaşması (HHI), anomali (MAD), trend/hareketli ortalama ve veri kalite skoru ile makine üretimi yorum katmanı (analitik motor).
- 8 grafikli yönetim paneli, yönetici özeti raporu ve 1234 koruma şifresi; tüm giriş hücrelerinde açılır ipucu ve açıklama notu.
- Gerçekçilik üst sınırı: katsayı ve değerler doğrulama aralıklarıyla sınırlanır; veri azlığında karar "VERİ YOK" kalır.

### Denetim sonucu
- G kapıları (G01–G24): 0 KALDI
- Ö kapıları (Ö1, Ö1b, Ö2): 0 KALDI
- D kapıları (D01–D14): 0 KALDI
- SHA-256: `f893736372271bedcd70aa9c1c04e48a7b75ced29aca2a51fe8066c4bb1c7b72`

### Değişen girdi / çıktı
İlk üretim sürümüdür; altın dosya `KANIT/altin_cikti_yenideger.json` ile aynı girdide aynı çıktıyı üretmesi doğrulanır.
