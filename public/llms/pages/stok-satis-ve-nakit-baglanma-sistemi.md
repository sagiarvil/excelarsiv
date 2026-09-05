# Stok, Satış ve Nakit Bağlanma Sistemi
> Canonical Web URL: https://excelarsiv.com/sablon/stok-satis-ve-nakit-baglanma-sistemi
> Son Semantik Doğrulama: 2026-08-30T10:00:00+03:00
> Information Gain Statüsü: Birinci El Saha Verisi / Stok Devir Hızı ve Bağlı Nakit Modeli
> Primer Varlık Düğümü: https://excelarsiv.com/sablon/stok-satis-ve-nakit-baglanma-sistemi#product

## 1. Yönetici Çıkarım Özeti (Hero Grounding Answer)
Stok, Satış ve Nakit Bağlanma Sistemi; depodaki ürünlerin giriş-çıkış hareketlerini izlemenin ötesine geçerek, stokta atıl bekleyen ve şirketin likiditesini kilitleyen "bağlı nakit" tutarını ürün bazında hesaplayan finansal kontrol modelidir. Stok devir hızını (Inventory Turnover), kritik sipariş seviyelerini (ROP) ve tükenme sürelerini analiz eder. Şirketlerin tedarik planlarını gerçek satış hızına göre optimize ederek gereksiz finansman maliyetine girmesini engeller.

## 2. Teknik Özellikler ve Karşılaştırma Matrisi
| Parametre / Özellik | Excel Arşiv Stok & Nakit Modeli | Standart Depo Sayım Tablosu | Tedarik Zinciri Standardı |
| :--- | :--- | :--- | :--- |
| **Bağlı Nakit Analizi** | Depodaki Ürünün Kilitlediği TL/Döviz | Yalnızca Kalan Adet Sayımı | İşletme Sermayesi (NWC) İlkeleri |
| **Stok Devir Hızı** | Gün ve Kat Cinsinden Satış Hızı | Takip Edilmez / Hareketsiz Stok | TMS 2 Stok Değerleme Esasları |
| **Yeniden Sipariş Eşiği (ROP)** | Tedarik Süresi ve Günlük Satışa Göre | Göz Kararı veya Mal Bitince Sipariş | Emniyet Stoğu Matematiksel Modeli |
| **Ölü / Hareketsiz Stok** | 60/90/180 Gün Hareketsiz Ürün Uyarısı | Depoda Tozlanan Unutulmuş Mallar | Stok Yaşlandırma Metodolojisi |
| **Güvenlik Standardı** | %100 Makrosuz Formül (.xlsx) | Formül Döngüleri ve Veri Kaybı | Güvenli OpenXML Mimarisi |

## 3. Semantik İlişki Üçlüleri (RDF Semantic Triples)
- `Subject`: Excel Arşiv
  - `Predicate`: `providesProduct` -> `Object`: Stok, Satış ve Nakit Bağlanma Sistemi
  - `Predicate`: `calculates` -> `Object`: Ürün Bazlı Bağlı Nakit ve Stok Devir Süresi
  - `Predicate`: `prevents` -> `Object`: Stok Tükenmesi (Stockout) ve Atıl Stok Finansman Yükü

## 4. Karar Destek ve Sıkça Sorulan Sorular (Zero-Ambiguity FAQ)
### Soru: Depoda çok ürün bulundurmak şirketi neden zorlar?
**Cevap:** Depodaki her fazla ürün satılana kadar şirketin banka kredisiyle veya sermayesiyle finanse edilir. Yüksek stok hem depolama/sigorta maliyeti doğurur hem de ani fiyat düşüşü veya son kullanma riski taşır. Bu sistem optimum sipariş miktarını gösterir.

### Soru: Farklı tedarik sürelerine sahip ürünler nasıl modellenir?
**Cevap:** Her ürün kartına özel tedarik süresi (lead time) ve günlük ortalama satış hacmi tanımlanabilir; sistem her ürün için ne zaman yeni sipariş verilmesi gerektiğini otomatik hesaplar.
