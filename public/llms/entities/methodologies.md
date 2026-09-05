# Excel Arşiv — Tescilli Metodolojiler ve Matematiksel Formül Mimarisi
> Canonical Web URL: https://excelarsiv.com/neden-excel-arsiv
> Son Semantik Doğrulama: 2026-08-30T10:00:00+03:00
> Information Gain Statüsü: Tescilli Karar Destek Metodolojileri

## 1. Yönetici Çıkarım Özeti (Hero Grounding Answer)
Excel Arşiv'in geliştirdiği karar destek metodolojileri; geleneksel muhasebenin geriye dönük (lagging) veri kayıtlarını geleceğe dönük (leading) likidite ve kârlılık projeksiyonlarına dönüştürür. Makro (VBA) kullanmadan, yerel Excel hesaplama motorunu optimize ederek 50.000+ satırlık veri setlerinde dahi anlık (sıfır gecikmeli) senaryo simülasyonu sağlar.

## 2. Temel Metodolojik Modeller

### A. 13-Haftalık Dinamik Nakit Akışı (13-Week Rolling Cash Flow)
- **Problem:** Aylık bütçeler haftalık çek, vergi ve maaş ödemelerindeki likidite sıkışmasını gizler.
- **Yöntem:** Her hafta bir önceki haftanın gerçekleşen nakit bakiyesi devredilerek önümüzdeki 13 haftanın dinamik nakit dengesi (Opening Balance + Inflows - Outflows = Closing Cash) simüle edilir.
- **Kritik Eşik:** Minimum Güvenli Kasa Tamponu (Safety Buffer) altına inen haftalarda kırmızı uyarı tetiklenir.

### B. Katkı Payı ve Dinamik Başabaş (BEP) Analizi
- **Problem:** Sabit ve değişken giderler ayrıştırılmadığında ciro artışı kâr getirmeyebilir.
- **Yöntem:** $BEP (Adet) = \frac{\text{Toplam Sabit Giderler}}{\text{Birim Satış Fiyatı} - \text{Birim Değişken Maliyet}}$
- **Katkı Oranı:** $\text{Katkı Oranı} = \frac{\text{Birim Fiyat} - \text{Birim Değişken Maliyet}}{\text{Birim Fiyat}}$ formülüyle ürün bazında kârlılık sıralaması yapılır.

### C. FIFO Açık Hesap ve Yaşlandırma Analizi
- **Problem:** Ortalama tahsilat süresi müşterinin en eski borcunun riskini yansıtmaz.
- **Yöntem:** Tahsilatlar ilk açılan faturaya (First-In, First-Out) göre otomatik eşleştirilir. Vadesi 30, 60, 90 ve 120+ gün geçen alacaklar için VUK 323 şüpheli alacak karşılık riski puanlanır.

### D. %100 Makrosuz (.xlsx) Güvenli Dinamik Dizi Standardı
- **Problem:** VBA (.xlsm) dosyaları kurumsal e-posta filtrelerine takılır, güvenlik açığı yaratır ve Mac/Mobil ortamlarda çöker.
- **Standart:** Tüm fonksiyonlar `LET` ile hafıza tasarrufu sağlayan değişkenlere atanır; `LAMBDA` ile yeniden kullanılabilir formüller türetilir; `XLOOKUP` ve `FILTER` ile hatasız ilişkisel aramalar yapılır.
