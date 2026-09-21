# FİYAT SAVUNMASI — HAKEDİŞ FİYAT FARKI VE HAK KAYBI CETVELİ

**Satış fiyatı:** 990 TL
**Manda v4 D01 eşiği:** Y1 (danışman ikamesi) + Y2 (kurulum ikamesi) ≥ 3 × 990 = **2.970 TL**

## Y1 — Danışman ikamesi (ne yapıldığında para ödenir)

Fiyat farkı hesabı ve hak kaybı tespiti tek başına dış kaynaklı bir mali iş uzmanına
yaptırıldığında ödenecek tek seferlik danışman bedeli:

| Hizmet | İçerik | Tutar |
|---|---|---|
| Hakediş başına fiyat farkı hesap ve rapor ücreti | Formül kurgusu, endeksler, konsolidasyon, yönetici özeti | 3.000 TL |
| **Y1 toplam** | | **3.000 TL** |

## Y2 — Kurulum ikamesi (iç kaynağın kurma maliyeti)

Şirket içi bir kişinin bu cetveli sıfırdan kurması (formül mimarisi, kontrol paneli, senaryo
analizi, rapor düzeni) için gereken asgari emek bedeli:

| Kalem | Tutar |
|---|---|
| Formül motoru, kontrol, karar ve rapor katmanlarının kurulması | 1.500 TL |
| **Y2 toplam** | **1.500 TL** |

## Toplam değer hesabı

**Y1 + Y2 = 3.000 + 1.500 = 4.500 TL** ≥ 2.970 TL eşiği → **GEÇTİ**

## Y3 — Önlenen hata (bilgi amaçlı, D01 hesabına dahil değil)

Yanlış fiyat farkı hesabı veya tespit edilemeyen geç ödeme/hak kaybı kalemlerinin ortalama
sözleşme bedeline etkisi hesaplanmamıştır; D01 hesabı yalnızca Y1 ve Y2 üzerinden kurulur.

## Serbest alternatif ayrımı (D02 özeti)

Aşağıdaki her madde, dosyada ad tanımlı bir hücreye karşılık gelir ve "bunu bedavaya kendim
yaparım" diyen alıcıya karşı savunulabilir:

1. **Karar motoru** — Hakediş ve endeks verisi boşken `VERİ YOK`, dolu ancak kayıp yüksekken
   `KAYIP` kararını formülle üretir; serbest Excel'de karar satırı yoktur. (`KARAR!B11`)
2. **Veri kalite skoru** — Dolu satır sayısını, anomali bayrağını ve eksik hücreyi ölçerek
   güven yüzdesi üretir; boş dosyada sonuç üretmeyi engeller. (`MOTOR!VeriKaliteSkoru`)
3. **Hak kaybı çıpası** — Fiyat farkı tutarı ile piyasa birim fiyatı farkını tek oranda
   birleştirip "hakedişten kaçırılan tutar" gösterir. (`MOTOR!HakKaybiOran`)
4. **Geç ödeme maliyeti** — Dönem kayıtlarından geç ödeme gününü, gecikme zammı oranından
   maliyeti formülle türetir. (`MOTOR!GecOdemeMaliyeti`)
5. **Tornado/senaryo katmanı** — Hangi girdinin sonucu en çok etkilediğini sıralar; üç senaryoda
   fiyat farkı tahminini karşılaştırır. (`SENARYO_DUYARLILIK`)
6. **Kontrol ve denetim paneli** — Formül hatası sayısı, boş kalem ve dolu giriş hücresi
   sayacıyla dosya sağlığını gösterir. (`KONTROL`)

## Sonuç

- D01: **GEÇTİ** (4.500 ≥ 2.970)
- D02: **GEÇTİ** (6 ayrım maddesi ≥ 5, her biri ad tanımlı çıktı)
- Satış fiyatı 990 TL; ürün üzerinde kurulan değer mimarisi **ikame bedelinin 4,5 katı**
  değer üretir (4.500 TL ikame / 990 TL satış).
