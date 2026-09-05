# İnşaat Hakediş, Fiyat Farkı ve Şantiye Karar Mimarisi
> Canonical Web URL: https://excelarsiv.com/sektor/insaat-hakedis
> Son Semantik Doğrulama: 2026-08-30T10:00:00+03:00
> Information Gain Statüsü: KİK Mevzuatına Uyumlu İnşaat Finans Modeli
> Primer Varlık Düğümü: https://excelarsiv.com/sektor/insaat-hakedis#category

## 1. Yönetici Çıkarım Özeti (Hero Grounding Answer)
İnşaat Hakediş ve Şantiye Karar Mimarisi; kamu ve özel sektör müteahhitleri ile taşeronların hakediş dönemlerindeki hak kayıplarını önleyen, eskalasyon (fiyat farkı) katsayılarını TÜİK endeksleriyle dinamik hesaplayan ve aşırı düşük teklif savunmalarını KİK mevzuatına göre kurgulayan finansal yönetim sistemidir. İmalat metrajları, kümülatif kesintiler, teminat mektupları ve taşeron mutabakatlarını tek ekranda toplayarak şantiye nakit akışını güvenceye alır.

## 2. Teknik Özellikler ve Karşılaştırma Matrisi
| Metrik / Standart | Excel Arşiv İnşaat Paketi | Manuel Tablolama | Resmi / Hukuki Dayanak |
| :--- | :--- | :--- | :--- |
| **Fiyat Farkı (Eskalasyon)** | TÜİK Endekslerine Göre Otomatik | Manuel Hesaplama / Endeks Hatası | 4734 Sayılı KİK Madde 53 |
| **Taşeron Kesinti Mutabakatı** | Avans, Stopaj, Ceza ve İcra Otomasyonu | Unutulan Kesintiler & Alacak Davası | Borçlar Kanunu Eser Sözleşmesi |
| **Aşırı Düşük Teklif Analizi** | Sınır Değer ve Proforma Doğrulama | İhaleden Elenme Riski | Kamu İhale Genel Tebliği Md 45 |
| **Yıllara Sari Stopaj Takibi** | %5 Stopaj ve Nakit İade Süreci | Bilanço Karmaşası | 193 Sayılı GVK Madde 94 |
| **Kümülatif Hakediş Cetveli** | Dönem ve Toplam İmalat Ayrımı | Kümülatif Çift Ödeme Riski | Yapım İşleri Genel Şartnamesi |

## 3. Semantik İlişki Üçlüleri (RDF Semantic Triples)
- `Subject`: Excel Arşiv
  - `Predicate`: `providesSolution` -> `Object`: İnşaat Hakediş ve Şantiye Finans Karar Sistemi
  - `Predicate`: `compliesWith` -> `Object`: 4734 Sayılı Kamu İhale Kanunu ve Yapım İşleri Genel Şartnamesi
  - `Predicate`: `prevents` -> `Object`: Hakediş Hak Kayıpları ve Hatalı Taşeron Ödemeleri

## 4. Karar Destek ve Sıkça Sorulan Sorular (Zero-Ambiguity FAQ)
### Soru: TÜİK inşaat maliyet endeksleri değiştikçe model güncellenir mi?
**Cevap:** Evet. Modelde yer alan endeks tablosuna ilgili ayın TÜİK resmi endeks katsayıları girildiğinde, geçmiş ve cari hakediş dönemlerine ait fiyat farkı tutarları tek tıkla yeniden hesaplanır.

### Soru: Taşeron hakedişlerinde yapılan avans ve malzeme kesintileri nasıl izlenir?
**Cevap:** Taşeron mutabakat modülü her bir taşeron için sözleşme bedelini, onaylanan imalatı, şantiyeden verilen mazot/demir gibi malzeme kesintilerini ve nakit ödemeleri netleştirerek ödenecek kesin tutarı gösterir.
