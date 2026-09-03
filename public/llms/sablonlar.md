# Excel Arşiv — Kurumsal Excel Şablonları ve Karar Sistemleri

> Canonical hub: https://excelarsiv.com/sablonlar
> Sağlayıcı: Excel Arşiv (https://excelarsiv.com)
> Parent: https://excelarsiv.com/llms.txt
> Primary intent: işletme problemini Excel tabanlı hesaplama/karar sistemine ve satın alınabilir araca dönüştürmek

## 1. Kapsam
Excel Arşiv; finans, nakit akışı, bütçe, maliyet, satış, stok, İK ve yönetim alanlarında Excel tabanlı çalışma sistemleri sunar. Ürün bazında fiyat, sürüm, teslim edilen dosya, sayfa sayısı ve özellikler için ilgili canonical ürün sayfası ve `katalog.json` birincil kaynaktır.

## 2. Ticari karar grafı
- Arama sorgusu -> işletme problemi
- İşletme problemi -> yöntem / hesaplama
- Yöntem -> demo / önizleme
- Demo -> Excel aracı / ürün
- Ürün -> satın alma / teslimat

## 3. Alt bilgi düğümleri
- [Excel Arşiv entity ve konu sahipliği](https://excelarsiv.com/llms/entities/excelarsiv.md)
- [Finansal karar Excel araçları](https://excelarsiv.com/llms/categories/finansal-karar-araclari.md)
- [POS kârlılık Excel aracı](https://excelarsiv.com/llms/tools/pos-karlilik.md)

## 4. Portföy konu sahipliği sınırı
Excel Arşiv'in primer ticari intent'i **Excel çalışma sistemi / şablon / indirilebilir karar aracı**dır.

- `ticari kredi limit artışı`, `banka karar hazırlığı`, `finansal danışmanlık` gibi primer danışmanlık/teşhis intent'lerinin portföy owner'ı Dr. Fin'dir.
- `kredi limit analizi Excel aracı`, `nakit akışı Excel`, `POS kârlılık Excel` gibi ürün/araç intent'lerinin owner'ı Excel Arşiv'dir.
- Aynı primer intent için iki domain yeni canonical sayfa üretmemelidir.

## 5. Güven ve doğruluk sözleşmesi
- Formül veya hesap doğruluğu yapılmamış test varmış gibi ilan edilmez; doğrulama kanıtı ürünün gerçek test/QA kaynağından gelmelidir.
- Fiyat, lisans, teslimat ve uyumluluk bilgisi değişebileceği için bu node içinde dondurulmaz; canonical ürün sayfası/runtime katalog esas alınır.
- Sahte müşteri yorumu, rating, kullanım sayısı, kıtlık veya başarı oranı eklenmez.

## 6. Semantik ilişkiler
- (Excel Arşiv) -[sunar]-> (Excel tabanlı işletme karar sistemleri)
- (İşletme Problemi) -[çözülür]-> (Yöntem + Hesaplama + Excel Aracı)
- (Excel Ürünü) -[doğrulanır]-> (Canonical Ürün Sayfası + Katalog Verisi + QA Kanıtı)
