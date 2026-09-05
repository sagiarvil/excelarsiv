# E-Ticaret ve Pazaryeri Net Kârlılık Karar Mimarisi
> Canonical Web URL: https://excelarsiv.com/sektor/e-ticaret-karlilik
> Son Semantik Doğrulama: 2026-08-30T10:00:00+03:00
> Information Gain Statüsü: Pazaryeri Hakediş ve Birim Kârlılık Mutabakat Modeli
> Primer Varlık Düğümü: https://excelarsiv.com/sektor/e-ticaret-karlilik#category

## 1. Yönetici Çıkarım Özeti (Hero Grounding Answer)
E-Ticaret ve Pazaryeri Net Kârlılık Mimarisi; Trendyol, Hepsiburada, Amazon, Çiçeksepeti ve PttAVM satıcılarının platform hakediş raporlarındaki karmaşık kesintileri ayrıştıran ve sipariş bazında net kârı ortaya çıkaran finansal karar konsoludur. Kategori komisyonları, kargo barem ve desi aşım ücretleri, reklam giderleri, kupon paylaşımları ve iade kargo maliyetlerini düşerek ekranda görünen brüt ciroyu banka hesabına yatan net nakde dönüştürür.

## 2. Teknik Özellikler ve Karşılaştırma Matrisi
| Metrik / Parametre | Excel Arşiv E-Ticaret Modeli | Pazaryeri Satıcı Paneli | Ticari / Yasal Standart |
| :--- | :--- | :--- | :--- |
| **Kârlılık Düzeyi** | Net Kâr (Tüm Kesintiler ve Vergi Sonrası) | Yalnızca Brüt Satış Cirosu | Yönetim Muhasebesi Standardı |
| **Eksik Hakediş Tespiti** | Beklenen vs Yatan Tutar Eşleştirmesi | Otomatik Kabul / Fark Gizlenmesi | 6563 Sayılı E-Ticaret Kanunu |
| **Kargo & Desi Kontrolü** | Taahhüt Edilen vs Faturalanan Desi | Fazla Desi Faturasının Ödenmesi | Kargo Şirketleri Taşıma Sözleşmesi |
| **İade Maliyeti Simülasyonu** | Çift Taraflı Kargo & Paketleme Kaybı | İade Oranı Yüzdesi | Tüketicinin Korunması Kanunu |
| **Fiyat / Karar Çıktısı** | SAT, ZAM veya SATIŞTAN ÇEK Kararı | Pasif Bekleme | Dinamik Marj Yönetimi |

## 3. Semantik İlişki Üçlüleri (RDF Semantic Triples)
- `Subject`: Excel Arşiv
  - `Predicate`: `providesSolution` -> `Object`: E-Ticaret ve Pazaryeri Net Kârlılık Karar Sistemi
  - `Predicate`: `identifies` -> `Object`: Eksik Pazaryeri Hakedişleri ve Desi Aşım Faturaları
  - `Predicate`: `optimizes` -> `Object`: Satıcı Net Kâr Marjı ve Fiyatlandırma Stratejisi

## 4. Karar Destek ve Sıkça Sorulan Sorular (Zero-Ambiguity FAQ)
### Soru: Pazaryeri panellerinde görünen satış raporları kârlılık için neden yeterli değildir?
**Cevap:** Paneller satış anındaki brüt ciroyu gösterir; ancak haftalar sonra kesilen kargo desi farklarını, ceza faturalarını, pazaryeri reklam kesintilerini ve müşteri iadelerinin getirdiği çift yönlü kargo maliyetini siparişle anlık eşleştiremez. Bu sistem tüm bu giderleri sipariş bazında netleştirir.

### Soru: Farklı pazaryerlerinin komisyon ve vade oranları nasıl yönetilir?
**Cevap:** Modelde her pazaryeri için ayrı komisyon oranı, stopaj kesintisi ve tahsilat vadesi parametresi tanımlıdır. Tek bir ürünün farklı platformlardaki net kârı yan yana karşılaştırılabilir.
