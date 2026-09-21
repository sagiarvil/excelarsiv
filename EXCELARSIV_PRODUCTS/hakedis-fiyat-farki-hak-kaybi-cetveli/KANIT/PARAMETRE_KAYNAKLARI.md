# PARAMETRE KAYNAKLARI — HAKEDİŞ FİYAT FARKI VE HAK KAYBI CETVELİ

**Dosya:** `HakedisFiyatFarkiVeHakKaybiCetveli.xlsx`
**Sürüm:** 1.0.0
**Üretim tarihi:** 09.08.2026

Bu belge, dosyadaki tüm sayısal parametrelerin kaynağını ve yürürlük/doğrulama tarihini
gösterir. Kaynağı belirlenemeyen her değer açıkça **Varsayım** olarak işaretlenir ve aşağıda
toplu listelenir. AYARLAR sayfasında bu bilgilerin tamamı `kaynak`, `yururluk_tarihi` ve
`dogrulama_tarihi` sütunlarında yer alır (D10).

## 1. Mevzuat parametreleri

| Parametre | Değer | Birim | Kaynak | Yürürlük |
|---|---|---|---|---|
| Ağırlık toplamı toleransı (a+b+c+d ≠ 1 kontrolü) | 0,01 | Oran | 4734/4735 s. Kamu İhale Kanunu; sözleşme hükümleri | 01.01.2003 |
| Geç ödeme eşiği | 30 | Gün | Türk Borçlar Kanunu md.117; sözleşme hükmü | 11.01.2011 |

Fiyat farkı formülü (F = a + b×(Gk/G0) + c×(Bk/B0) + d×(Yk/Y0) − 1) sözleşmedeki ağırlıklara ve
Hazine ve Maliye Bakanlığı tebliğlerinde tanımlı endeks yöntemine dayanır; katsayı ve ağırlıklar
her sözleşmeden okunur. Endeks değerleri TÜİK veya sözleşmede tanımlı endekslerden alınır.

## 2. Sistem parametreleri

| Parametre | Değer | Birim | Kaynak |
|---|---|---|---|
| Rapor Tarihi | 09.08.2026 | Tarih | Kullanıcı (rapor günü) |
| Firma Ünvanı | Örnek Yapı A.Ş. | Metin | Kullanıcı |

## 3. Varsayım parametreleri (toplu liste)

Aşağıdaki parametreler belirli bir mevzuat maddesine dayanmaz; ürünün çalışması için tanımlı
genel eşiklerdir. Alıcı, sözleşmesine ve işin gerçeğine göre AYARLAR'dan güncellemelidir.

| Parametre | Değer | Birim | Doğrulama yolu |
|---|---|---|---|
| KayipEsikOran | 0,05 | Oran | Sözleşmedeki kayıp/kâr analizi |
| KayipDikkatOran | 0,15 | Oran | Sözleşmedeki kayıp/kâr analizi |
| MutabakatTolerans | 100 | ₺ | Dönem kayıt mutabakatı |
| AnomaliZEsik | 2,00 | Oran | Endeks serisi istatistiği |
| VeriKaliteIyiEsik / OrtaEsik | 90 / 70 | Puan | Kullanıcı kalite hedefi |
| VeriKaliteUyariCeza | 10 | Puan | Skor duyarlılığı |
| SenaryoKotumserKat / IyimserKat | 0,15 / 0,15 | Oran | Senaryo aralığı |
| DuyarliDegisim | 0,05 | Oran | Tornado değişim aralığı |
| ParetoEsikOran | 0,50 | Oran | Yoğunlaşma eşiği |
| P90Kuantil | 0,90 | Oran | Yüzdelik dağılım |
| EndeksBantOrani | 0,10 | Oran | Tahmin bant genişliği |
| KayipOraniEsik | 0,02 | Oran | Geç ödeme maliyeti tahmini |

Bu liste KILAVUZ sayfasında "Kaynaklar ve varsayımlar" bölümüyle eşleşir.
