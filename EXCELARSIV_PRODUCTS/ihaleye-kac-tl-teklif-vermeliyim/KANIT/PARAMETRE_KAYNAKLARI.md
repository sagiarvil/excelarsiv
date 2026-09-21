# PARAMETRE KAYNAKLARI — İHALEYE KAÇ TL TEKLİF VERMELİYİM? (SINIR DEĞER HESAPLAYICI)

**Dosya:** `IhaleyeKacTlTeklifVermeliyim.xlsx`
**Sürüm:** 1.0.0
**Üretim tarihi:** 09.08.2026

Bu belge, dosyadaki tüm sayısal parametrelerin kaynağını ve yürürlük/doğrulama tarihini
gösterir. Kaynağı belirlenemeyen her değer açıkça **Varsayım** olarak işaretlenir ve aşağıda
toplu listelenir. AYARLAR sayfasında bu bilgilerin tamamı `kaynak`, `yururluk_tarihi` ve
`dogrulama_tarihi` sütunlarında yer alır (D10).

## 1. Mevzuat parametreleri

| Parametre | Değer | Birim | Kaynak | Yürürlük |
|---|---|---|---|---|
| Aşırı Düşük Teklif Eşiği (Teklif/SD oranı) | 1,00 | Oran | 4734 s. KİK; Yapım İşleri İhaleleri Uygulama Yönetmeliği md.61-62 | 01.01.2003 |

Sınır değer formülü (SD = YM×N + OrtAlt×(1−N)) ve N katsayısı idari şartnameye göre değişir;
alıcı ihale dokümanındaki değerleri AYARLAR'dan işlemelidir. Bu süreç KILAVUZ sayfasında anlatılır.

## 2. Sistem parametreleri

| Parametre | Değer | Birim | Kaynak |
|---|---|---|---|
| Rapor Tarihi | 09.08.2026 | Tarih | Kullanıcı (rapor günü) |
| Firma Ünvanı | Örnek Yapı A.Ş. | Metin | Kullanıcı |
| Değer Ölçer: Kazanılan Saat / Ay | 12 | Saat/ay | Kullanıcı |
| Değer Ölçer: Saat Ücreti | 1.500 | ₺ | Kullanıcı |

## 3. Varsayım parametreleri (toplu liste)

Aşağıdaki parametrelerin resmi bir kaynağı yoktur; sektör pratiği veya kurulum sırasında
kullanıcı onayı ile belirlenir. Her biri AYARLAR sayfasında sarı hücrede olup kullanıcı
tarafından güncellenebilir:

| Parametre | Değer | Birim | Gerekçe |
|---|---|---|---|
| Güvenlik Payı | 0,03 | Oran | Sınır değerin %3 üzeri güvenlik payı varsayımı |
| Anomali Z Eşiği | 2,00 | Adet | Z skoru eşiği varsayımı (istatistik pratiği) |
| N Katsayısı Duyarlılık Üst Sınırı | 0,50 | Oran | Tornado analizinde N üst sınırı varsayımı |
| Mutabakat Toleransı | 10.000 | ₺ | Maliyet kalemleri/toplam maliyet kabul toleransı varsayımı |
| Kalite Skoru İyi/Orta Eşiği | 90 / 70 | Puan | Skorlama varsayımı |
| Kalite Skoru Uyarı Cezası | 10 | Puan | Her WARN için skor kesintisi varsayımı |
| Azami Tutar Sınırı | 100.000.000.000 | ₺ | Uç değer denetimi mutlak üst sınır varsayımı |
| Senaryo Katsayıları (Kötümser/İyimser) | 0,15 / 0,15 | Oran | Senaryo aralığı varsayımı |
| Duyarlılık Değişim Oranı | 0,05 | Oran | Tornado adımı (±%5) varsayımı |
| Tahmin Ufku | 6 | Adet | Gelecek tahmin adımı (ay cinsinden) varsayımı |
| Tahmin Bandı Genişliği | 0,10 | Oran | Tahmin aralığı güven bandı varsayımı |
| Fiyat Farkı Etki Oranı | 0,03 | Oran | Fiyat farkı maddesinin tahmini katkı oranı varsayımı |
