# PARAMETRE KAYNAKLARI — KAÇIRILAN SGK TEŞVİKLERİ & GERÇEK İŞÇİLİK MALİYETİ ANALİZİ

**Dosya:** `KacirilanSgkTesvikleriVeGercekIscilikMaliyetiAnalizi.xlsx`
**Sürüm:** 1.1.0
**Üretim tarihi:** 09.08.2026

Bu belge, dosyadaki tüm sayısal parametrelerin kaynağını ve yürürlük/doğrulama tarihini
gösterir. Kaynağı belirlenemeyen her değer açıkça **Varsayım** olarak işaretlenir ve aşağıda
toplu listelenir. AYARLAR sayfasında bu bilgilerin tamamı `kaynak`, `yururluk_tarihi` ve
`dogrulama_tarihi` sütunlarında yer alır (D10).

## 1. Mevzuat parametreleri

| Parametre | Değer | Birim | Kaynak | Yürürlük |
|---|---|---|---|---|
| İşveren SGK payı oranı | 0,205 | Oran | 5510 sayılı Kanun md.81 (SGK) | 01.01.2026 |
| İşsizlik sigortası işveren payı | 0,02 | Oran | 4447 sayılı Kanun md.49 (İŞKUR) | 01.01.2026 |
| 5 puan prim indirimi teşviki | 0,05 | Oran | 5510 sayılı Kanun md.81 (SGK) | 01.01.2026 |
| İlave istihdam teşviki | 0,145 | Oran | 4447 sayılı Kanun geç.10 (SGK) | 01.01.2026 |
| Kadın/genç/mesleki belgeli teşviki | 0,06 | Oran | 4447 sayılı Kanun geç.10 (SGK) | 01.01.2026 |
| Engelli istihdam teşviki | 0,20 | Oran | 4857 sayılı Kanun md.30 (SGK) | 01.01.2026 |
| Genç girişimci istisnası | 0,225 | Oran | 5510 sayılı Kanun geç.28 (SGK) | 01.01.2026 |

Teşvik oranları dönemsel olarak değişebilir; üretim tarihi itibarıyla güncel resmi oranlar
kullanılmıştır. Alıcı, oranların güncelliğini SGK e-Sigorta portalı veya mali müşaviri
üzerinden teyit etmelidir. Bu teyit süreci KILAVUZ sayfasında da anlatılır.

## 2. Sistem parametreleri

| Parametre | Değer | Birim | Kaynak |
|---|---|---|---|
| Rapor tarihi | =TODAY() | Tarih | Sistem (kullanıcı değiştirir) |
| Yıllık hesap ay sayısı | 12 | Adet | Sistem |

## 3. Varsayım parametreleri (toplu liste)

Aşağıdaki parametrelerin resmi bir kaynağı yoktur; sektör ortalaması, yasal katsayı üzerinden
türetme veya kurulum sırasında kullanıcı onayı ile belirlenir. Her biri AYARLAR sayfasında
sarı hücrede olup kullanıcı tarafından güncellenebilir:

| Parametre | Değer | Birim | Gerekçe |
|---|---|---|---|
| Kıdem karşılığı oranı | 0,0833 | Oran/ay | Yasal kıdem tazminatı katsayısı üzerinden aylık karşılık varsayımı |
| Yan fayda (kişi başı/ay) | 3.000 | ₺/ay | Yemek/yol/yan hak ortalaması varsayımı |
| Aylık çalışma saati | 180 | Saat/ay | Sektör ortalaması varsayımı |
| Kaçırılan teşvik uyarı eşiği | 50.000 | ₺/yıl | Kullanıcı karar eşiği varsayımı |
| Çalışan başı aylık maliyet eşiği | 40.000 | ₺/ay | Kullanıcı karar eşiği varsayımı |
| Teşvik kullanım oranı alt eşiği | 0,15 | Oran | Kullanıcı karar eşiği varsayımı |
| İşveren yükü makul alt/üst sınırı | 0,10 / 0,50 | Oran | Kontrol toleransı varsayımı |
| Maliyet/brüt tutarlılık toleransı | 0,02 | Oran | Kontrol toleransı varsayımı |
| İşveren yükü karar eşiği | 0,30 | Oran | Kullanıcı karar eşiği varsayımı |
| Veri kalitesi iyi/orta eşiği | 90 / 70 | Puan | Skorlama varsayımı |
| Her WARN için skor cezası | 10 | Puan | Skorlama varsayımı |
| Kötümser/iyimser senaryo çarpanı | 0,10 | Oran | Senaryo varsayımı |
| Tornado değişim oranı | 0,05 | Oran | Duyarlılık varsayımı |
| P10/P50/P90 katsayıları | 0,10 / 0,50 / 0,90 | Katsayı | Dağılım varsayımı |
| Tahmin güven aralığı katsayısı | 1,65 | Katsayı | %90 güven aralığı varsayımı |
| Anomali eşik katı | 2,00 | Kat | Ortanca sapma eşiği varsayımı |
| Teşvik geri kazanım hedef ayı | 12 | Ay | Nakit planlama varsayımı |

## 4. Doğrulama yöntemi

- Mevzuat parametreleri SGK resmi duyuruları ve ilgili kanun maddeleri üzerinden doğrulanmıştır.
- Varsayım parametreleri LibreOffice ile yeniden hesaplanmış dosyada formül bütünlüğü
  denetlenerek doğrulanmıştır (0 hata, 7 uç durum senaryosu temiz).
- Dosya tamamen çevrimdışıdır; dış bağlantı, telemetri veya makro içermez.
