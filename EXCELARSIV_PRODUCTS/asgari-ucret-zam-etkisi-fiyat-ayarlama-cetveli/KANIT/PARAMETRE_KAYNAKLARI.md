# PARAMETRE KAYNAKLARI — ASGARİ ÜCRET ZAM ETKİSİ & FİYAT AYARLAMA CETVELİ

Bu belge, dosyadaki `AYARLAR` sayfasında bulunan tüm parametrelerin kaynağını,
yürürlük tarihini ve doğrulama tarihini listeler. Manda v4 D10 gereği kaynağı olmayan
hiçbir sayı `kaynak = Varsayım` ile işaretlenir ve aşağıda ayrıca toplu gösterilir.

## 1. Mevzuat kaynaklı parametreler

| Parametre | Değer | Kaynak | Yürürlük | Doğrulama |
|---|---|---|---|---|
| AsgariUcretEski | 12.000 ₺ | Asgari Ücret Tespit Komisyonu kararı; kullanıcı doğrulamalı | 01.01.2025 | 09.08.2026 |
| AsgariUcretYeni | 15.000 ₺ | Asgari Ücret Tespit Komisyonu kararı; kullanıcı doğrulamalı | 01.01.2026 | 09.08.2026 |
| SgkIsverenOran | %20,5 | 5510 s. SGK md.80 | 01.10.2008 | 09.08.2026 |
| IssizlikIsverenOran | %2 | 4447 s. İşsizlik Sigortası md.46 | 01.09.2000 | 09.08.2026 |
| SgkIsciOran | %14 | 5510 s. SGK md.80 | 01.10.2008 | 09.08.2026 |
| IssizlikIsciOran | %1 | 4447 s. İşsizlik Sigortası md.46 | 01.09.2000 | 09.08.2026 |
| GunlukCalismaSaat | 8 saat | 4857 s. İş Kanunu md.63 | 10.06.2003 | 09.08.2026 |

## 2. Yaygın uygulama / kullanıcı kaynaklı parametreler

| Parametre | Değer | Kaynak | Yürürlük | Doğrulama |
|---|---|---|---|---|
| UcretHesaplamaGun | 30 gün | Yaygın uygulama | 01.01.1970 | 09.08.2026 |
| firmaUnvani | Örnek Şirket A.Ş. | Kullanıcı | 09.08.2026 | 09.08.2026 |
| raporTarihi | =TODAY() | Kullanıcı | 09.08.2026 | 09.08.2026 |

## 3. Tasarım kaynaklı parametreler

| Parametre | Değer | Kaynak | Yürürlük | Doğrulama |
|---|---|---|---|---|
| KayitKapasite | 3.000 adet | Tasarım | 09.08.2026 | 09.08.2026 |

## 4. Varsayım kaynaklı parametreler (D10 toplu listesi)

Aşağıdaki parametrelerde kaynak `Varsayım`dır; kullanıcı kendi işletme verisiyle
doğrulamalıdır. Hepsi AYARLAR sayfasında sarı giriş hücresinde tutulur:

| Parametre | Değer | Açıklama |
|---|---|---|
| BolumPaydaTaban | 0,0000001 | Bölmede sıfır payda koruması |
| MaliyetArtisEsik | %10 | Maliyet artışı uyarı eşiği |
| FiyatYansitmaEsik | %15 | Fiyat yansıtma yüksek kabul eşiği |
| UcretAnomaliZEsik | 2,00 | Ücret anomali |Z| eşiği |
| VeriKaliteIyiEsik | 90 | Kalite skoru iyi eşiği |
| VeriKaliteOrtaEsik | 70 | Kalite skoru orta eşiği |
| VeriKaliteUyariCeza | 10 | Her WARN için skor cezası |
| SenaryoKotumserKat | %10 | Kötümser senaryo çarpanı |
| SenaryoIyimserKat | %10 | İyimser senaryo çarpanı |
| DuyarliDegisim | %5 | Tornado değişim oranı |
| ParetoEsikOran | %30 | Ücret yoğunlaşması yüksek eşiği |
| P90Kuantil | 0,90 | P90 kuantil değeri |
| IskontoOrani | %10 | NPV iskonto oranı |
| KidemSinirGun | 1.800 | Beş yıl eşiği (gün) |
| VeriKaliteTaban | 100 | Skor başlangıç puanı |
| VeriKaliteFailCeza | 20 | Her FAIL için ceza |
| ButceInceleEsikOran | %60 | İNCELE kararı eşiği |
| ButceEtkiEsikOran | %10 | Tornado zirve etkisi eşiği |
| HhiYuksekEsik | 0,50 | Yoğunlaşma yüksek eşiği |
| NpvEsikOran | %50 | NPV/artış yüksek kabul eşiği |

## 5. Doğrulama notu

- Mevzuat oranları (5510, 4447, 4857) birincil kaynak metinlerinden alınmıştır;
  oran değişikliklerinde kullanıcı AYARLAR'dan güncellemelidir.
- Asgari ücret değerleri Resmi Gazete duyurusu sonrası güncellenir; dosya duyurudan
  bağımsız çalışır ve güncel değerler girilmelidir.
- Varsayım parametreleri kullanıcının işletme gerçeğine göre revize edilmelidir;
  güncellenmeyen varsayımla üretilen kararlar düşük güven işareti taşır.
