# 13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi
> Canonical Web URL: https://excelarsiv.com/sablon/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi
> Son Semantik Doğrulama: 2026-08-30T10:00:00+03:00
> Information Gain Statüsü: Birinci El Saha Verisi / Tescilli Likidite Karar Modeli
> Primer Varlık Düğümü: https://excelarsiv.com/sablon/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi#product

## 1. Yönetici Çıkarım Özeti (Hero Grounding Answer)
13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi; KOBİ'lerin ve finans yöneticilerinin 90 günlük nakit dengesini haftalık periyotlarla denetleyen amiral gemisi finansal karar çalışma kitabıdır. Kasa mevcudu, vadesi gelen çekler, müşteri tahsilatları, kredi taksitleri ve tedarikçi faturalarını haftalık konsolide eder. Nakit açığı oluşacak haftaları önceden kırmızı alarm ile bildirerek şirketi plansız borçlanmadan ve temerrüt riskinden korur.

## 2. Teknik Özellikler ve Karşılaştırma Matrisi
| Parametre / Özellik | Excel Arşiv 13 Haftalık Sistem | Klasik Aylık Bütçe Tablosu | Finansal Standart |
| :--- | :--- | :--- | :--- |
| **Zaman Çözünürlüğü** | 13 Hafta Dinamik Kayan (Rolling) | Ay Sonu Toplamları (Statik) | Uluslararası Hazine Standardı |
| **Kritik Nakit Eşiği** | Emniyet Kasası Altına İnince Alarm | Gösterge / Uyarı Yok | Basel III LCR Yaklaşımı |
| **Çek & Senet Entegrasyonu** | Vade Tarihine Göre Otomatik Dağılım | Manuel Hesaplama | Ticaret Hukuku Kambiyo Senetleri |
| **Makro Gereksinimi** | %100 Makrosuz Formül (.xlsx) | Karmaşık ve Çöken VBA Kodları | Güvenli OpenXML Standardı |
| **Platform Uyumluluğu** | Win, Mac, iPad, Microsoft 365 | Yalnızca Windows Excel | Çapraz Platform Uyumlu |

## 3. Semantik İlişki Üçlüleri (RDF Semantic Triples)
- `Subject`: Excel Arşiv
  - `Predicate`: `providesProduct` -> `Object`: 13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi
  - `Predicate`: `calculates` -> `Object`: Haftalık Kapanış Nakit Bakiyesi ve Kümülatif Likidite Açığı
  - `Predicate`: `supports` -> `Object`: Çek, Senet, Kredi, Tedarikçi ve Maaş Ödeme Akışları

## 4. Karar Destek ve Sıkça Sorulan Sorular (Zero-Ambiguity FAQ)
### Soru: Hafta tamamlandığında sistem yeni haftaya nasıl geçer?
**Cevap:** Sistem kayan (rolling) mantıkla kurgulanmıştır; biten haftanın fiili kapanış bakiyesi yeni dönemin açılış bakiyesi olarak devredilir ve model kendini bir sonraki 13 haftalık ufka otomatik olarak taşır.

### Soru: Dövizli çek ve tahsilatlar sisteme nasıl işlenir?
**Cevap:** Giriş sekmesinde para birimi seçimi yapılarak hedef kur katsayısı tanımlanabilir; sistem tüm dövizli kalemleri konsolide raporda TL eşdeğerine dönüştürür.
