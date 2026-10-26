# ExcelArşiv Premium — Renkli tasarım paketi

## Kapsam

260 sayfa: ana sayfa, 177 ürün detayı, 8 kategori, 22 rehber, 31 ürün demosu; katalog, pratik programlar, demo merkezi, iletişim, ürün bulucu, özel çözümler, destek ve koşul sayfaları.

Ekteki renkli sürümün kırmızı (#e73345), yeşil (#078653), sarı (#ffcf24) ve mavi (#1557d5) paleti korunur. Tipografi, boşluk, kartlar, görsel önizleme ve ortak sayfa düzenleri yenilenmiştir. 12 vitrinde çalışma kitaplarından alınmış gerçek görseller vardır. Diğer ürünlerdeki tanıtım görselleri ekran görüntüsü olarak sunulmaz.

## Çalıştırma ve İnceleme

Proje kök dizininden veya `excelarsiv-premium-tum-sayfalar` dizininden doğrudan çalıştırabilirsiniz:

### 1. Yerel Önizleme Sunucusu
```bash
# Proje kökünden:
npm run premium:serve

# Veya excelarsiv-premium-tum-sayfalar içinden:
npm start
```
Tarayıcınızda `http://localhost:3000` adresinden tüm 262 sayfayı inceleyebilirsiniz.

### 2. Doğrulama ve Bütünlük Testi
```bash
# Proje kökünden:
npm run premium:test

# Veya excelarsiv-premium-tum-sayfalar içinden:
npm test
```
Tüm 262 sayfa, 197 varlık (CSS/JS/görseller) ve 14.890 dahili bağlantı taranarak 0 hata ile onaylanır.

### 3. Tek Dosya Sürümü
`site/tek-dosya.html` dosyasını sunucusuz olarak doğrudan çift tıklayarak herhangi bir tarayıcıda açabilirsiniz.

## Satın alma ve teslimat

Bu paket tasarım ve gezinme katmanıdır. Sunucu, ödeme doğrulama, ücretli dosyalar, kullanıcı hesapları ve sipariş verileri içermez. Satın alma, demo erişimi ve sipariş doğrulama mevcut excelarsiv.com sayfalarına yönlenir. Gerçek ödeme veya indirme yapıldığına dair sahte başarı mesajı üretilmez. Mevcut uygulamanın sunucu işlevleri ve sipariş sistemi korunarak entegrasyon yapılmalıdır; bu paketi mevcut sunucu uygulamasının üzerine körlemesine kopyalamayın.

## Durum

Kurulum ve kalite kapısı doğrulaması eksiksiz tamamlanmıştır (0 kırık link, 0 eksik asset, 0 BOM).
