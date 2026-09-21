# PARAMETRE KAYNAKLARI — TAŞERON/ALT YÜKLENİCİ HAKEDİŞ & KESİNTİ MUTABAKATI

**Dosya:** `TaseronHakedisVeKesintiMutabakati.xlsx`
**Sürüm:** 1.0.0
**Üretim tarihi:** 09.08.2026

Bu belge, dosyadaki tüm sayısal parametrelerin kaynağını ve yürürlük/doğrulama tarihini
gösterir. Kaynağı belirlenemeyen her değer açıkça **Varsayım** olarak işaretlenir ve aşağıda
toplu listelenir. AYARLAR sayfasında bu bilgilerin tamamı `kaynak`, `yururluk_tarihi` ve
`dogrulama_tarihi` sütunlarında yer alır (D10).

## 1. Mevzuat parametreleri

| Parametre | Değer | Birim | Kaynak | Yürürlük |
|---|---|---|---|---|
| KDV oranı | %20 | Oran | KDV Kanunu md.28 | 25.10.1984 |
| Yapım işi KDV tevkifatı (4/10) | %4 | Oran | KDV Genel Uygulama Tebliği | 26.04.2018 |
| Hakediş gelir vergisi stopajı | %5 | Oran | GVK md.94 | 01.01.2006 |
| SGK kesinti oranı | %20 | Oran | 5510 s. Kanun; sözleşme hükmü | 01.10.2008 |
| Gecikme faizi oranı | %15 | Oran (yıllık) | TBK md.117-120; sözleşme hükmü | 11.01.2011 |

Yapım işlerinde sorumlu sıfatıyla KDV tevkifatı, GVK md.94 kapsamında hakediş stopajı ve
5510 s. Kanun kapsamında işveren/alt işveren prim sorumluluğu sözleşme ve maliye uygulamasına
göre kurgulanır. SGK kesintisi, alt işveren prim sorumluluğu kapsamında sözleşmede ayrıştırılan
tutardır; bu oranlar her sözleşmede mali müşavir teyidiyle güncellenmelidir.

## 2. Sözleşme kaynaklı parametreler

| Parametre | Değer | Birim | Kaynak |
|---|---|---|---|
| Teminat kesinti oranı | %5 | Oran | Sözleşme hükmü |
| Vade günü | 60 | Gün | Sözleşme hükmü |

## 3. Sistem parametreleri

| Parametre | Değer | Birim | Kaynak |
|---|---|---|---|
| Rapor Tarihi | 09.08.2026 | Tarih | Kullanıcı (rapor günü; tek TODAY kaynağı) |
| Firma Ünvanı | Örnek Yapı A.Ş. | Metin | Kullanıcı |

## 4. Varsayım parametreleri (toplu liste)

Aşağıdaki parametreler belirli bir mevzuat maddesine dayanmaz; ürünün çalışması için tanımlı
genel eşiklerdir. Alıcı, sözleşmesine ve işin gerçeğine göre AYARLAR'dan güncellemelidir.

| Parametre | Değer | Birim | Doğrulama yolu |
|---|---|---|---|
| MutabakatTolerans | 100 | ₺ | Firma mutabakat politikası |
| MutabakatFarkiEsik | 0,02 | Oran | Firma mutabakat eşiği |
| GecOdemeEsik | 30 | Gün | Firma tahsilat politikası |
| KesintiEsik | 0,25 | Oran | Sözleşme kesinti karşılaştırması |
| AnomaliZEsik | 2,00 | Oran | Hakediş serisi istatistiği |
| VeriKaliteIyiEsik / OrtaEsik | 90 / 70 | Puan | Kullanıcı kalite hedefi |
| VeriKaliteUyariCeza | 10 | Puan | Skor duyarlılığı |
| SenaryoKotumserKat / IyimserKat | 0,10 / 0,10 | Oran | Senaryo aralığı |
| DuyarliDegisim | 0,05 | Oran | Tornado değişim aralığı |
| ParetoEsikOran | 0,50 | Oran | Yoğunlaşma eşiği |
| P90Kuantil | 0,90 | Oran | Yüzdelik dağılım |
| TahminBantOrani | 0,10 | Oran | Tahmin bant genişliği |
| IskontoOrani | 0,10 | Oran | Nakit NPV iskonto oranı |

Bu liste KILAVUZ sayfasındaki varsayım açıklamalarıyla eşleşir.
