# SÜRÜM NOTLARI — AMORTİSMAN & SABİT KIYMET SATIŞ ZAMANLAMA STRATEJİSTİ

## 1.0.0 — İlk sürüm

### Ürün amacı
Sabit kıymetlerin amortisman hesabını normal ve azalan bakiyeli yöntemlerle yapar; yıl içinde dört çeyrek satış senaryosunun kıst amortisman, satış kârı/zararı, kurumlar vergisi etkisi ve net nakit değerlerini üretir; çeyrek nakit akışlarını bugünkü değere indirgeyerek satış için en uygun dönemi gerekçesiyle önerir.

### Özellikler
- **Kıymet listesi:** 1.000 satır kapasiteli tabloda kıymet kodu, tür, edinim tarihi, edinim bedeli, yöntem, oran ve satış bedeli girişi; geçen yıl, birikmiş amortisman, net defter değeri, yıllık amortisman ve satış kârı/zararı canlı formülle.
- **Satış zamanlama motoru:** Dört çeyrek senaryosunda kıst amortisman, birikmiş amortisman, net defter değeri, kâr/zarar, vergi etkisi ve net nakit; NBD ile bugünkü değer karşılaştırması.
- **Karar kapısı:** SAT / BEKLE / VERİ YOK / DURDUR kararı ve gerekçesi; aksiyon önerileri.
- **Analitik motor:** Ortalama oran, NDD sapması (MAD), en yüksek satış kârı, senaryo bant genişliği, HHI yoğunlaşması, kalite skoru, tahmin + güven bandı, P90 yüzdelik, NBD farkı (10 modül, 170 puan; 3 ileri modül).
- **Panel:** 9 grafik (tür bazlı edinim/birikmiş/NDD/amortisman, kalem bazlı NDD trendi, projeksiyon bant aralığı, çeyrek net nakit ve NBD, senaryo net nakitleri) ve 62 canlı KPI.
- **Senaryo ve duyarlılık:** İyi/Baz/Kötü/Kritik senaryo bant aralığı ve satış bedeli/amortisman oranı tornado analizi.

### Denetim sonuçları
- G kapıları (G01–G24): **0 KALDI**
- Ö kapıları (Ö1, Ö1b, Ö2): **0 KALDI** (3.000 satır / 13.491 formül, 0 hata, 3,2 sn; 7 uç durum temiz)
- D kapıları (D01–D14): **0 KALDI** (D01 ikame 8.500 ₺ ≥ 4.470 ₺; D02 5 ayrım; D03 170 puan; D04 3 ileri modül)

### Altın çıktı
- Altın dosya: `KANIT/altin_cikti_amortisman.json`
- İçerik SHA-256: `abdff9a9a02fa2b680b0a1f8e5ce8ede601a4211198855743ffd0c644b13ea0a`
- Fark sayısı: 0

### Dosya parmak izi
- SHA-256 (teslim edilen dosya): `6f4c18582a725cb8d82058a0fc3e925cc7725c5984a1c33fb4371527a6e97510`
- KANIT'taki üç rapor da aynı parmak iziyle imzalanmıştır (RAPOR_DENETIM, RAPOR_OLCEK, RAPOR_DEGER).

### Değişiklik günlüğü
| Tarih | Sürüm | Değişiklik | Neden |
|---|---|---|---|
| 10.08.2026 | 1.0.0 | İlk üretim | Ürünün yayına hazırlanması |
