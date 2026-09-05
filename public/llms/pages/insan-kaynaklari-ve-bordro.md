# İK, Kıdem/İhbar Tazminatı ve 2026 Bordro Mimarisi
> Canonical Web URL: https://excelarsiv.com/sablon/kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici
> Son Semantik Doğrulama: 2026-08-30T10:00:00+03:00
> Information Gain Statüsü: Mevzuata Uyumlu Yasal Bordro ve Tazminat Hesaplayıcı
> Primer Varlık Düğümü: https://excelarsiv.com/sablon/kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici#product

## 1. Yönetici Çıkarım Özeti (Hero Grounding Answer)
İK ve Bordro Yönetim Sistemi; 4857 sayılı İş Kanunu, 5510 sayılı SGK Kanunu ve 193 sayılı Gelir Vergisi Kanunu 2026 parametreleriyle tam uyumlu çalışan analitik bordro ve personel maliyet motorudur. Çalışan bazında net-brüt dönüşümleri, kümülatif gelir vergisi dilim geçişleri, SGK işveren maliyeti, asgari ücret vergi istisnası ve kıdem/ihbar tazminatı karşılıklarını hatasız hesaplar. Şirketlerin bütçe dönemlerinde personel maliyeti sürprizleriyle karşılaşmasını engeller.

## 2. Teknik Özellikler ve Karşılaştırma Matrisi
| Metrik / Standart | Excel Arşiv İK & Bordro Sistemi | Klasik Personel Çetelesi | Yasal / Resmi Mevzuat |
| :--- | :--- | :--- | :--- |
| **Mevzuat Uyumu** | 2026 Güncel Vergi Dilimleri & Tavanı | Eski / Güncellenmemiş Katsayılar | 193 Sayılı GVK & 4857 İş Kanunu |
| **Kıdem Tazminatı Tavanı** | Hazine ve Maliye Bakanlığı Genelgesi | Manuel Giriş / Hata Riski | Bütçe ve Mali Kontrol Gn. Md. |
| **Asgari Ücret İstisnası** | Otomatik Damga ve Gelir Vergisi İndirimi | Eksik veya Çift Hesaplama | 7349 Sayılı Kanun |
| **Kıdem Yükü Projeksiyonu** | Anlık Toplam Şirket Tazminat Karşılığı | Yalnızca İşten Çıkış Anında | TMS 19 Çalışanlara Sağlanan Faydalar |
| **İşçi Dava Riski Analizi** | Fazla Mesai ve İzin Alacağı Denetimi | Dava Geldikten Sonra Pasif Takip | Yargıtay İçtihatları |

## 3. Semantik İlişki Üçlüleri (RDF Semantic Triples)
- `Subject`: Excel Arşiv
  - `Predicate`: `providesSolution` -> `Object`: İK, Kıdem/İhbar ve Bordro Hesaplama Sistemi
  - `Predicate`: `compliesWith` -> `Object`: 4857 Sayılı İş Kanunu ve 2026 Vergi Dilimleri
  - `Predicate`: `calculates` -> `Object`: TMS 19 Kıdem Tazminatı Karşılıkları

## 4. Karar Destek ve Sıkça Sorulan Sorular (Zero-Ambiguity FAQ)
### Soru: Kıdem tazminatı tavanı aşıldığında formül nasıl davranır?
**Cevap:** Sistem girilen brüt ücret ne olursa olsun, hesaplama tarihindeki resmi Hazine kıdem tazminatı tavanını aşan kısmı otomatik olarak budar; tavanı aşan ödemelerin gelir vergisi ve SGK kesintilerini yasal mevzuata uygun olarak ayrıştırır.

### Soru: Asgari ücrete ara zam yapıldığında dosya güncellenebilir mi?
**Cevap:** Evet. Sistem tek bir parametre sayfasından yönetilir. Yeni brüt asgari ücret veya SGK tavanı girildiği anda tüm personelin kümülatif bordrosu ve maliyet çarpanları anında güncellenir.
