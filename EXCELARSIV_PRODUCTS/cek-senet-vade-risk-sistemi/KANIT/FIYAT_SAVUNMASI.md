# FİYAT SAVUNMASI — ÇEK SENET VADE VE KARŞILIK RİSK SİSTEMİ

**Satış fiyatı:** 1.490 TL
**Manda v4 D01 eşiği:** Y1 (danışman ikamesi) + Y2 (kurulum ikamesi) ≥ 3 × 1.490 = **4.470 TL**

## Y1 — Danışman ikamesi (ne yapıldığında para ödenir)

Çek/senet portföyü için vade planlaması, karşılık yeterliliği analizi ve risk ağırlıklı tahsil
değerlendirmesi dış kaynaklı bir finans/nakit yönetimi danışmanına yaptırıldığında ödenecek bedel:

| Hizmet | İçerik | Tutar |
|---|---|---|
| Nakit/vade risk danışmanlığı | Belge bazında vade takvimi, risk ağırlıklı tahsil değeri, karşılık açığı analizi, müşteri/banka yoğunlaşması, karşılıksız çek riski skorlaması, senaryo/duyarlılık ve yönetici raporu | 6.000 TL |
| **Y1 toplam** | | **6.000 TL** |

## Y2 — Kurulum ikamesi (iç kaynağın kurma maliyeti)

Şirket içi bir kişinin bu sistemi sıfırdan kurması (belge kayıt düzeni, vade takvimi motoru, karşılık
ve yoğunlaşma hesapları, karar kuralları, analitik modüller, pano ve rapor düzeni) için gereken
asgari emek bedeli:

| Kalem | Tutar |
|---|---|
| Vade takvimi ve karşılık açığı motorunun formülle kurulması ve doğrulanması | 1.700 TL |
| Pano, tahmin/yüzdelik modülleri ve rapor katmanlarının kurulması | 800 TL |
| **Y2 toplam** | **2.500 TL** |

## Toplam değer hesabı

**Y1 + Y2 = 6.000 + 2.500 = 8.500 TL** ≥ 4.470 TL eşiği → **GEÇTİ**

## D02 — Serbest alternatifle örtüşme riski

Alıcının "bunu bedavaya kendim yaparım" diyebileceği en güçlü nokta: **vadeleri banka listelerinden
ve çek defterinden tek tek takip etmek**. Sistemin bu alternatife karşı ayrıştığı maddeler dosyada ad
tanımlarıyla kanıtlanır: risk ağırlıklı alınan toplamı (`RiskAglikliAlinan`), günlük karşılık açığı
(`MaxAcik`), müşteri yoğunlaşma payı (`modulT3Yogunlasma`), gecikme maliyeti (`modulO5GecikmeMaliyeti`)
ve veri yetersizken "VERİ YOK" diyen karar katmanı. Defter takibi tek başına "hangi hafta ödeme
baskısı var, karşılıksız çıkarsa nakit açığı ne olur" sorusuna cevap vermez; bu dosya verir.
