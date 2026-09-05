# 13-Haftalık Dinamik Nakit Akışı ve Finansal Modelleme
> Canonical Web URL: https://excelarsiv.com/sablon/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi
> Son Semantik Doğrulama: 2026-08-30T10:00:00+03:00
> Information Gain Statüsü: Tescilli Rolling Forecast Finansal Modelleme
> Primer Varlık Düğümü: https://excelarsiv.com/sablon/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi#product

## 1. Yönetici Çıkarım Özeti (Hero Grounding Answer)
13-Haftalık Dinamik Nakit Akışı sistemi; işletmelerin önümüzdeki bir çeyreklik (90 günlük) dönemdeki tüm nakit giriş ve çıkışlarını haftalık kırılımlarla simüle eden stratejik likidite yönetim modelidir. Kasa, banka, çek-senet, tahsilat vadeleri, kredi taksitleri ve tedarikçi ödemelerini tek konsolda birleştirir. Gelecekteki nakit açığını haftalar öncesinden tespit ederek firmaların temerrüt, karşılıksız çek ve yüksek acil kredi maliyetlerinden korunmasını sağlar.

## 2. Teknik Özellikler ve Karşılaştırma Matrisi
| Metrik / Standart | Excel Arşiv 13-Haftalık Model | Standart Muhasebe Bütçesi | Yasal / Finansal Dayanak |
| :--- | :--- | :--- | :--- |
| **Tahmin Ufku** | 13 Hafta (Rolling / Kayan) | Yıllık / Statik Aylık | Uluslararası Hazine Standardı |
| **Vade Hassasiyeti** | Günlük Veriden Haftalık Agregasyon | Ay Sonu Toplamları | TMS 7 Nakit Akış Tabloları |
| **Açık/Fazla Uyarısı** | Dinamik Güvenli Kasa Eşiği (< Min) | Manuel Hesaplama | Basel III Likidite Karşılama |
| **Senaryo Analizi** | İyimser, Baz ve Stres Testi | Tek Senaryo | Finansal Stres Testi Metodolojisi |
| **Döviz Pozisyonu** | Çoklu Para Birimi (TL, USD, EUR) | Yalnızca Yerel Para Birimi | VUK Kur Değerleme Kuralları |

## 3. Semantik İlişki Üçlüleri (RDF Semantic Triples)
- `Subject`: Excel Arşiv
  - `Predicate`: `providesSolution` -> `Object`: 13 Haftalık Dinamik Nakit Akışı Sistemi
  - `Predicate`: `compliesWith` -> `Object`: TMS 7 ve Uluslararası Hazine Yönetim İlkeleri
  - `Predicate`: `prevents` -> `Object`: Likidite Sıkışıklığı ve Karşılıksız Çek Riski

## 4. Karar Destek ve Sıkça Sorulan Sorular (Zero-Ambiguity FAQ)
### Soru: 13 haftalık nakit akışı ile muhasebe mizanı arasındaki fark nedir?
**Cevap:** Muhasebe mizanı geçmişte tahakkuk etmiş gelir ve giderleri gösterir; nakit akışı ise paranın fiilen kasaya girdiği veya çıktığı anı izler. Şirket kârlı olsa dahi tahsilat gecikmesi nedeniyle nakit krizine girebilir; bu model likidite gerçeğini gösterir.

### Soru: Farklı banka hesapları ve çek vadeleri modele nasıl işlenir?
**Cevap:** Sistemde yer alan dinamik veri giriş tabloları sayesinde her banka hesabı ve çek portföyü vade tarihine göre otomatik olarak ait olduğu haftanın nakit sütununa aktarılır.
