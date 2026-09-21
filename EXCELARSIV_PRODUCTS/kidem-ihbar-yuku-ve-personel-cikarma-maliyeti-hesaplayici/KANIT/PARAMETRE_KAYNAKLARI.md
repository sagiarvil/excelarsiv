# PARAMETRE KAYNAKLARI — KIDEM–İHBAR YÜKÜ VE PERSONEL ÇIKARMA MALİYETİ HESAPLAYICI

**Ürün:** Kıdem–İhbar Yükü ve Personel Çıkarma Maliyeti Hesaplayıcı
**Dosya:** `KidemIhbarYukuVePersonelCikarmaMaliyetiHesaplayici.xlsx`
**Üretim tarihi:** 09.08.2026

Bu belge, AYARLAR sayfasındaki her parametrenin kaynağını ve yürürlük tarihini listeler.
Kaynağı `Varsayım` işaretli değerler kullanıcı tarafından kendi sözleşme/işletme koşullarına göre
doğrulanmalıdır; mevzuat parametreleri ilgili kanun/tebliğden gelir.

## 1. Mevzuat parametreleri

| Parametre | Değer | Birim | Kaynak | Yürürlük |
|---|---|---|---|---|
| KidemTavan | 45.000 | ₺/ay | Hazine ve Maliye Bakanlığı yıllık duyurusu; güncel değer kullanıcı doğrulamalı | 09.08.2026 |
| UcretGunSayisi | 30 | gün | 1475 s. İş Kanunu md.14 uygulaması (aylık ücretin günlük karşılığı) | 01.09.1971 |
| IhbarSure1 | 14 | gün | 4857 s. İş Kanunu md.17 (2 hafta) | 10.06.2003 |
| IhbarSure2 | 28 | gün | 4857 s. İş Kanunu md.17 (4 hafta) | 10.06.2003 |
| IhbarSure3 | 42 | gün | 4857 s. İş Kanunu md.17 (6 hafta) | 10.06.2003 |
| IhbarSure4 | 56 | gün | 4857 s. İş Kanunu md.17 (8 hafta) | 10.06.2003 |
| IhbarEsik1Gun | 180 | gün | 4857 s. İş Kanunu md.17 (6 ay) | 10.06.2003 |
| IhbarEsik2Gun | 547 | gün | 4857 s. İş Kanunu md.17 (1,5 yıl) | 10.06.2003 |
| IhbarEsik3Gun | 1095 | gün | 4857 s. İş Kanunu md.17 (3 yıl) | 10.06.2003 |
| IzinHakki1 | 14 | gün | 4857 s. İş Kanunu md.53 (1-5 yıl) | 10.06.2003 |
| IzinHakki2 | 20 | gün | 4857 s. İş Kanunu md.53 (5-15 yıl) | 10.06.2003 |
| IzinHakki3 | 26 | gün | 4857 s. İş Kanunu md.53 (15 yıl ve üzeri) | 10.06.2003 |
| IzinEsik1 | 5 | yıl | 4857 s. İş Kanunu md.53 geçiş eşiği | 10.06.2003 |
| IzinEsik2 | 15 | yıl | 4857 s. İş Kanunu md.53 geçiş eşiği | 10.06.2003 |

## 2. Varsayım parametreleri (kullanıcı doğrulamalı)

| Parametre | Değer | Birim | Açıklama |
|---|---|---|---|
| CikarmaButceEsik | 500.000 | ₺ | Çıkarma maliyeti bütçe eşiği |
| ButceKullanimEsik | 0,80 | oran | Bütçe kullanım oranı uyarı eşiği |
| KidemYilEsik | 10 | yıl | Ortalama kıdem uyarı eşiği |
| UcretAnomaliZEsik | 2,00 | oran | Ücret anomali \|Z\| eşiği |
| VeriKaliteIyiEsik | 90 | puan | Kalite skoru iyi eşiği |
| VeriKaliteOrtaEsik | 70 | puan | Kalite skoru orta eşiği |
| VeriKaliteUyariCeza | 10 | puan | Her WARN için skor cezası |
| SenaryoKotumserKat | 0,10 | oran | Kötümser senaryo çarpanı |
| SenaryoIyimserKat | 0,10 | oran | İyimser senaryo çarpanı |
| DuyarliDegisim | 0,05 | oran | Tornado analizinde değişim oranı |
| ParetoEsikOran | 0,50 | oran | Maliyet yoğunlaşması yüksek kabul eşiği |
| P90Kuantil | 0,90 | oran | P90 yüzdelik kuantil değeri |
| TahminBantOrani | 0,10 | oran | Maliyet tahmin aralığı bant oranı |
| IskontoOrani | 0,10 | oran | Maliyet yükü NPV iskonto oranı |

## 3. Kullanıcı parametreleri

| Parametre | Değer | Birim | Açıklama |
|---|---|---|---|
| firmaUnvani | Örnek Şirket A.Ş. | metin | Rapor başlığında görünen firma adı |
| raporTarihi | =TODAY() | tarih | Dosyadaki tek tarih kaynağı |
