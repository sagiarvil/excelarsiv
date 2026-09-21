# PARAMETRE KAYNAKLARI — FAZLA MESAİ VE İŞÇİ DAVA RİSKİ TESPİT DOSYASI

**Ürün:** Fazla Mesai ve İşçi Dava Riski Tespit Dosyası
**Dosya:** `FazlaMesaiVeIsciDavaRiskiTespitDosyasi.xlsx`
**Üretim tarihi:** 09.08.2026

Bu belge, AYARLAR sayfasındaki her parametrenin kaynağını ve yürürlük tarihini listeler.
Kaynağı `Varsayım` işaretli değerler kullanıcı tarafından kendi sözleşme/işletme koşullarına göre
doğrulanmalıdır; mevzuat parametreleri ilgili kanun/tebliğden gelir.

## 1. Mevzuat parametreleri

| Parametre | Değer | Birim | Kaynak | Yürürlük |
|---|---|---|---|---|
| MesaiKatsayi | 1,50 | katsayı | 4857 s. İş Kanunu md.41 (saat ücreti + %50) | 10.06.2003 |
| FazlaSureKatsayi | 1,25 | katsayı | 4857 s. İş Kanunu md.42 (saat ücreti + %25) | 10.06.2003 |
| YillikMesaiSinir | 270 | saat | 4857 s. İş Kanunu md.41 (yıllık fazla çalışma üst sınırı) | 10.06.2003 |
| UcretHesaplamaGun | 30 | gün | Yaygın uygulama (aylık ücretin günlük karşılığı) | 01.01.1970 |
| GunlukCalismaSaat | 8 | saat | 4857 s. İş Kanunu md.63 (günlük normal çalışma süresi) | 10.06.2003 |
| HaftalikCalismaSinir | 45 | saat | 4857 s. İş Kanunu md.63 (haftalık normal çalışma süresi üst sınırı) | 10.06.2003 |
| ZamanAsimiYil | 5 | yıl | 4857 s. İş Kanunu md.32/8 (ücret ve fazla mesai alacaklarında zamanaşımı) | 10.06.2003 |
| AySayisi | 12 | adet | Takvim yılı | 01.01.1970 |

## 2. Varsayım parametreleri (kullanıcı doğrulamalı)

| Parametre | Değer | Birim | Açıklama |
|---|---|---|---|
| BolumPaydaTaban | 0,0000001 | katsayı | Bölme işlemlerinde sıfır payda koruması taban değeri |
| DavaRiskButceEsik | 500.000 | ₺ | Dava riski tutarı bütçe eşiği |
| DavaRiskEsik | 0,15 | oran | Fazla mesai ücretinin yıllık ücret yüküne oranı uyarı eşiği |
| KayitFarkiEsik | 0,20 | oran | Bildirilen-beldenen fark oranı uyarı eşiği |
| UcretAnomaliZEsik | 2,00 | oran | Mesai anomali \|Z\| eşiği |
| VeriKaliteIyiEsik | 90 | puan | Kalite skoru iyi eşiği |
| VeriKaliteOrtaEsik | 70 | puan | Kalite skoru orta eşiği |
| VeriKaliteUyariCeza | 10 | puan | Her WARN için skor cezası |
| SenaryoKotumserKat | 0,10 | oran | Kötümser senaryo mesai çarpanı |
| SenaryoIyimserKat | 0,10 | oran | İyimser senaryo mesai çarpanı |
| DuyarliDegisim | 0,05 | oran | Tornado analizinde değişim oranı |
| ParetoEsikOran | 0,30 | oran | Mesai yoğunlaşması yüksek kabul eşiği |
| P90Kuantil | 0,90 | oran | Yüzdelik dağılım analizinde P90 kuantil değeri |
| IskontoOrani | 0,10 | oran | Risk tutarı NPV iskonto oranı |
| KidemSinirGun | 1.800 | gün | Kıdem ortalamasında beş yıl eşiği (yaklaşık) |
| VeriKaliteTaban | 100 | puan | Veri kalite skoru başlangıç puanı |
| VeriKaliteFailCeza | 20 | puan | Her FAIL için skor cezası |
| ButceInceleEsikOran | 0,60 | oran | Risk bütçesinin %60 üzerinde İNCELE kararı eşiği |
| ButceEtkiEsikOran | 0,10 | oran | Tornado zirve etkisinin bütçeye oranı uyarı eşiği |
| HhiYuksekEsik | 0,50 | oran | Mesai yoğunlaşma endeksi yüksek kabul eşiği |
| NpvEsikOran | 0,50 | oran | Risk NPV değerinin bütçeye oranı yüksek kabul eşiği |

## 3. Tasarım parametreleri

| Parametre | Değer | Birim | Açıklama |
|---|---|---|---|
| KayitKapasite | 3.000 | adet | Çalışan ve kayıt tablolarının toplam giriş kapasitesi |

## 4. Kullanıcı parametreleri

| Parametre | Değer | Birim | Açıklama |
|---|---|---|---|
| firmaUnvani | Örnek Şirket A.Ş. | metin | Rapor başlığında görünen firma adı |
| raporTarihi | =TODAY() | tarih | Dosyadaki tek tarih kaynağı; tüm rapor tarihleri buradan okunur |
