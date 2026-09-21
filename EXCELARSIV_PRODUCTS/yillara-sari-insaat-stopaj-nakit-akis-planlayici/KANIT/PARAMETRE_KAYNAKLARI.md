# PARAMETRE KAYNAKLARI — YILLARA SARİ İNŞAAT STOPAJ & NAKİT AKIŞ PLANLAYICI

**Dosya:** `YillaraSariInsaatStopajVeNakitAkisPlanlayici.xlsx`
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
| Yapım işi KDV tevkifatı (4/10) | %4 | Oran | KDV Genel Uygulama Tebliği; 17 no'lu KDV Sirküleri | 26.04.2018 |
| Yıllara sari işlerde gelir vergisi stopajı | %5 | Oran | GVK md.94 | 01.01.2006 |
| Geçici vergi oranı | %25 | Oran | KVK md.21, md.32 | 01.10.2024 |
| Kurumlar vergisi oranı | %25 | Oran | KVK md.32 | 01.01.2024 |

Yıllara sari inşaat ve onarma işlerinde kârın işin bittiği yıl tespiti GVK md.42-44'e dayanır.
KDV hariç hakediş üzerinden tevkifat ve stopaj işleyişi sözleşme ve maliye uygulamasına göre
kurgulanır; bu oranlar her sözleşmede mali müşavir teyidiyle güncellenmelidir.

## 2. Sistem parametreleri

| Parametre | Değer | Birim | Kaynak |
|---|---|---|---|
| Rapor Tarihi | 09.08.2026 | Tarih | Kullanıcı (rapor günü; tek TODAY kaynağı) |
| Firma Ünvanı | Örnek Yapı A.Ş. | Metin | Kullanıcı |

## 3. Varsayım parametreleri (toplu liste)

Aşağıdaki parametreler belirli bir mevzuat maddesine dayanmaz; ürünün çalışması için tanımlı
genel eşiklerdir. Alıcı, sözleşmesine ve işin gerçeğine göre AYARLAR'dan güncellemelidir.

| Parametre | Değer | Birim | Doğrulama yolu |
|---|---|---|---|
| KarMinEsik | 0,10 | Oran | Firma kâr hedefi / maliyet planı |
| DscrEsik | 1,25 | Oran | Banka borç servisi kriterleri |
| NakitAcikTolerans | 0 | ₺ | Yönetim nakit politikası |
| VergiYukuEsik | 0,30 | Oran | Geçmiş dönem vergi yükü |
| AnomaliZEsik | 2,00 | Oran | Hakediş serisi istatistiği |
| VeriKaliteIyiEsik / OrtaEsik | 90 / 70 | Puan | Kullanıcı kalite hedefi |
| VeriKaliteUyariCeza | 10 | Puan | Skor duyarlılığı |
| SenaryoKotumserKat / IyimserKat | 0,10 / 0,10 | Oran | Senaryo aralığı |
| DuyarliDegisim | 0,05 | Oran | Tornado değişim aralığı |
| ParetoEsikOran | 0,50 | Oran | Yoğunlaşma eşiği |
| P90Kuantil | 0,90 | Oran | Yüzdelik dağılım |
| TahminBantOrani | 0,10 | Oran | Tahmin bant genişliği |
| IskontoOrani | 0,10 | Oran | Nakit NPV iskonto oranı |
| ODemeVadeGun | 60 | Gün | Ortalama tahsilat vadesi |

Bu liste KILAVUZ sayfasındaki varsayım açıklamalarıyla eşleşir.
