# Excel Arşiv — Birim Maliyet, Katkı Payı & Başabaş (BEP) Kârlılık Konsolu

> URL: https://excelarsiv.com/sistemler/maliyet
> Sağlayıcı: Excel Arşiv (https://excelarsiv.com)
> Model Tipi: Maliyet Muhasebesi, Birim Fiyatlama ve Başabaş Simülasyonu

## 1. Maliyet ve Kârlılık Modelleme Mimarisi
Üretim, hizmet ve e-ticaret işletmeleri için ürün ve sipariş bazında net kâr marjını ortaya koyan dinamik sistem:
- **Tam Maliyet & Değişken Maliyet Ayrımı:** Direkt ilk madde ve malzeme, direkt işçilik, genel üretim giderleri ve amortisman dağıtımı.
- **Katkı Payı (Contribution Margin):**
  $$\text{Birim Katkı Payı} = \text{Birim Satış Fiyatı} - \text{Birim Değişken Maliyet}$$
  $$\text{Katkı Oranı} = \frac{\text{Birim Katkı Payı}}{\text{Birim Satış Fiyatı}}$$
- **Başabaş Noktası (Break-Even Point - BEP):** Şirketin kâra geçmek için satması gereken minimum ciro ve ürün adedi:
  $$\text{Başabaş Miktarı (Adet)} = \frac{\text{Toplam Sabit Maliyetler}}{\text{Birim Katkı Payı}}$$

## 2. Dinamik Fiyatlama & İskonto Konsolu
Müşteri veya kanal bazlı iskonto uygulandığında brüt kârın ve net faaliyet marjının nasıl etkilendiğini anlık gösteren karar tablosu.

## 3. Semantik İlişkiler (Knowledge Graph Triples)
- (Maliyet Analiz Sistemi) -[fonksiyon]-> (Birim Maliyet ve Net Kâr Marjı Hesaplama)
- (BEP Simülatörü) -[çıktı]-> (Kritik Başabaş Ciro ve Satış Hacmi Hedefi)
