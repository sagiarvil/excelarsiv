# PARAMETRE KAYNAKLARI — AYLIK PATRON FİNANS PANELİ

**Dosya:** `AylikPatronFinansPaneli.xlsx`
**Sürüm:** 1.0.0
**Üretim tarihi:** 10.08.2026

Bu belge, dosyadaki tüm sayısal parametrelerin kaynağını ve yürürlük/doğrulama tarihini
gösterir. Kaynağı belirlenemeyen her değer açıkça **Varsayım** olarak işaretlenir ve aşağıda
toplu listelenir. AYARLAR sayfasında bu bilgilerin tamamı `kaynak`, `yururluk_tarihi` ve
`dogrulama_tarihi` sütunlarında yer alır (D10).

## 1. Mevzuat dayanağı

| Parametre | Değer | Birim | Kaynak | Yürürlük |
|---|---|---|---|---|
| Kayıt düzeni ve gelir-gider tespiti | — | — | VUK (213 s.) md.182–198 — defter tutma ve kayıt düzeni (Gelir İdaresi Başkanlığı) | 01.01.2026 |

Dosyada kullanılan tüm oranlar işletme eşiği niteliğindedir (ör. borç/gelir eşiği, kasa yeterlilik
katı, senaryo çarpanları). Bu eşikler doğrudan kanundan gelmez; kullanıcının kendi işletmesine
göre kurulum sırasında onayladığı parametrelerdir ve aşağıda **Varsayım** olarak toplu listelenir.

## 2. Sistem parametreleri

| Parametre | Değer | Birim | Kaynak |
|---|---|---|---|
| Rapor tarihi | 08.08.2026 | Tarih | AYARLAR (kullanıcı değiştirir) |
| Yıllık hesap ay sayısı | 12 | Adet | Sistem |

## 3. Varsayım parametreleri (toplu liste)

Aşağıdaki parametrelerin resmi bir kaynağı yoktur; işletme eşiği, sektör ortalaması veya
kullanıcı onayı ile belirlenir. Her biri AYARLAR sayfasında sarı hücrede olup kullanıcı
tarafından güncellenebilir:

| Parametre | Değer | Birim | Gerekçe |
|---|---|---|---|
| Ay başı kasa girişi | 1.500.000 | ₺ | İşletme başlangıç bakiyesi varsayımı |
| Kâr eşiği | 0 | ₺ | Net kârın altına düşmemesi gereken eşik varsayımı |
| Borç/gelir eşiği | 0,75 | Oran | Borçluluk karar eşiği varsayımı |
| Kasa/gider katı | 1 | Kat | Kasanın aylık gideri karşılama katı varsayımı |
| Asgari ay sonu kasa | 250.000 | ₺ | Kasa yeterlilik eşiği varsayımı |
| İyimser senaryo çarpanı | 1,10 / 0,90 | Çarpan | Senaryo varsayımı |
| Baz senaryo çarpanı | 1,00 / 1,00 | Çarpan | Senaryo varsayımı |
| Kötümser senaryo çarpanı | 0,85 / 1,15 | Çarpan | Senaryo varsayımı |
| Tornado değişim oranı | 0,10 / 0,01 | Oran/puan | Duyarlılık varsayımı |
| Tahmin güven katsayısı | 1,645 | Katsayı | %90 güven aralığı varsayımı |
| Kalite skoru iyi/orta eşiği | 90 / 70 | Puan | Skorlama varsayımı |
| Asgari veri sayısı | 6 | Adet | Karar üretimi için gereken kayıt eşiği varsayımı |
| Anomali eşiği | 1,5 | Katsayı | Z-skor eşiği varsayımı |
| Gecikme maliyeti oranı | 0,02 | Oran | Yıllık gecikme maliyeti varsayımı |
| Yoğunlaşma (HHI) eşiği | 0,25 | Oran | Tek müşteri bağımlılığı eşiği varsayımı |
| Nakit dönüşüm eşiği | 0,85 | Oran | Tahsilata dönüşme eşiği varsayımı |
| P10/P50/P90 katsayıları | 0,10 / 0,50 / 0,90 | Katsayı | Dağılım varsayımı |
| Aksiyon öneri metinleri | 3 adet | Metin | Karar motoru aksiyon metinleri varsayımı |

## 4. Doğrulama yöntemi

- Mevzuat dayanağı VUK (213 s.) md.182–198 üzerinden doğrulanmıştır.
- Varsayım parametreleri LibreOffice ile yeniden hesaplanmış dosyada formül bütünlüğü
  denetlenerek doğrulanmıştır (0 hata; ölçek ve uç durum senaryoları temiz).
- Dosya tamamen çevrimdışıdır; dış bağlantı, telemetri veya makro içermez.
