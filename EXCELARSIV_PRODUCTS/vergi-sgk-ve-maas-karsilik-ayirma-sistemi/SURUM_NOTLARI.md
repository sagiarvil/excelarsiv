# SÜRÜM NOTLARI — VERGİ, SGK VE MAAŞ KARŞILIK AYIRMA SİSTEMİ

## 1.0.0 — 10.08.2026

### Bu sürümde olanlar
- Dönem geliri ve gideri üzerinden KDV, kurumlar vergisi, SGK (işçi+işveren), işsizlik, stopaj, damga ve maaş karşılıklarını tek motorda ayıran sistem.
- 1.000 satır kapasiteli dönem ve bordro tabloları; karşılık motoru, nakit projeksiyonu ve ay bazlı karşılık grafiği.
- UYGUN / İNCELE / DURDUR karar kapısı, gerekçe ve aksiyon önerileri.
- İyimser, baz, kötümser senaryolar; KDV/KV/SGK/gelir tornado analizi; tahmin + güven bandı.
- Dönemler arası trend, anomali oranı, veri kalite skoru ve makine üretimi yorum katmanı (analitik motor).
- 8 grafikli yönetim paneli, yönetici özeti raporu ve 1234 koruma şifresi.
- AYARLAR sayfası kaynak/mevzuat sütunlu yapıda; oranlar tek noktadan güncellenir (D10).

### Denetim sonucu
- G kapıları (G01–G24): 0 KALDI
- Ö kapıları (Ö1, Ö2): 0 KALDI
- D kapıları (D01–D14): 0 KALDI
- SHA-256: `dce740dec6ed2a4cec2ad85638905b8f169be63cdcf2c4e1b13b18105b1132a3`

### Değişen girdi / çıktı
İlk üretim sürümüdür; altın dosya `ornek/altin_cikti_vergi_sgk.json` ile aynı girdide aynı çıktıyı üretmesi doğrulanır.
